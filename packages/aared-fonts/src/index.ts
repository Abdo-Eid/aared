/**
 * Metric-compatible font substitution + a CDN-backed default resolver.
 * See docs/ARCHITECTURE.md § Fonts.
 *
 * Frozen and unpublished (Notes §53): on a real phone the browser's
 * own fonts opened documents 2-15x faster and read better, so aared's default
 * is no resolver. Kept for a consumer that needs Office's line breaks exactly.
 *
 * No font binaries ship in this package. `defaultResolver` maps each requested Office family to its
 * substitute, fetches that substitute's `@font-face` CSS from Google Fonts
 * once per family, and registers every subset face under `fontAlias(requested)`
 * (never the requested name itself, never the substitute's own name - see
 * `fontAlias`) via `FontFace` on `document.fonts`, so faces cross shadow-DOM
 * boundaries and the browser fetches only the subsets actually painted. The
 * resolver is main-thread-only: `core` receives it through `open()` but never
 * posts it to the Worker (see the API reference page § Font resolution).
 */

import { fontAlias } from "aared/core";
import type { FontResolver } from "aared/core";

/** Requested Office font family -> delivered substitute family. */
export const FONT_SUBSTITUTIONS: Record<string, string> = {
  Calibri: "Carlito",
  "Calibri Light": "Carlito",
  Cambria: "Caladea",
  Arial: "Arimo",
  "Times New Roman": "Tinos",
  "Courier New": "Cousine",
  "Traditional Arabic": "Amiri",
  "Sakkal Majalla": "Markazi Text",
  "Simplified Arabic": "Lateef",
  "Arabic Typesetting": "Amiri",
  Aptos: "Radio Canada",
  "Aptos Display": "Roboto",
  "Aptos Narrow": "IBM Plex Sans Condensed",
};

/**
 * `FontFace` metric-override descriptors for a substitution that is not an
 * exact clone. Carlito/Caladea/Arimo/Tinos/Cousine share their Office
 * counterpart's advance widths glyph for glyph (verified: `tools/font-calibration/`, 0 of 95 ASCII glyphs differ for the Calibri/Arial/Times New
 * Roman/Courier New pairs; Cambria/Caladea differ only by sub-pixel
 * units-per-em quantization, under 0.13%) - nothing to correct.
 *
 * Aptos (Office's default theme font since 2023: Aptos for body text, Aptos
 * Display for headings, Aptos Narrow in Excel) has no open clone either, and
 * browsers never have it: Office downloads it into a per-user cache and never
 * installs it as a system font. Its three families were measured the way the
 * Arabic fonts were, against the real files from that cache on every Latin
 * run in the fixtures (`tools/font-calibration/latin-survey.ts`), among 61
 * open candidates delivered by this resolver's own request. Once widths match,
 * a look-alike must also keep Aptos's letter heights: of the 61, one face per
 * family lands within 4% on x-height and cap height at both weights - Radio
 * Canada for Aptos, Roboto for Aptos Display, IBM Plex Sans Condensed for
 * Aptos Narrow. They also break real paragraphs where Aptos does
 * (`latin-fit.ts`, eight widths: 96%/93%, 95%/94% and 96%/95% of paragraphs
 * broken identically, regular/bold): the best average of the 16 closest
 * candidates for Aptos and Aptos Narrow, 0.8 points under the best for Aptos
 * Display. Unsubstituted, the generic sans-serif (Arial's widths) breaks
 * 38-75% of them the same way. Aptos Display is Aptos spaced
 * tighter for large sizes, which a size-adjust cannot reproduce, so it takes
 * a face of its own. All three carry bold numbers although their weights
 * differ by only 1-3%: Latin paragraphs are long enough that Aptos Narrow's
 * bold at the regular numbers (1.3% wide) broke 91% of them like Aptos, not
 * 95%. Italics are 2-3% narrower than Aptos's own for Aptos and Aptos Display
 * (one calibration covers both styles of a weight). Notes §49.
 *
 * Arabic has no
 * such clone: every Office Arabic font (Simplified/Traditional Arabic, Sakkal
 * Majalla, Arabic Typesetting) is Microsoft-licensed, Arabic's contextual
 * letterforms make an unlicensed clone much harder than Latin's, and a search
 * for one found none. So each is calibrated instead. The candidate is the open
 * face with the smallest spread of per-paragraph width ratios against real
 * shaped Arabic text (every Arabic-only run in the fixtures, 59 samples, in
 * the font's real Windows file; line breaks depend on that spread, while the
 * mean is what `size-adjust` removes), among eleven candidates loaded from the
 * same Google Fonts files this resolver delivers (`tools/font-calibration/
 * arabic-survey.ts`, Notes §45). A candidate had to beat the incumbent
 * (Amiri, the original choice) by a quarter to displace it: Lateef does for
 * Simplified Arabic (47% less spread), Markazi Text for Sakkal Majalla (41%).
 * Traditional Arabic and Arabic Typesetting keep Amiri (Lateef: 12% less and
 * 10% more). A large `size-adjust` is not itself a problem: the height of alef
 * and of the seen body was measured after adjusting, and Lateef lands as close
 * to the original's proportions as Amiri does. `sizeAdjustPercent` scales the
 * substitute's advances to match; `ascentOverridePercent`/`descentOverridePercent`
 * are computed to compensate for `size-adjust` scaling them too - a browser
 * multiplies an override percentage by size-adjust's own factor rather than
 * applying it to the font's un-adjusted em (verified empirically, not assumed
 * from spec text: `tools/font-calibration/size-adjust-interaction.ts`). The
 * ascent and descent are the original's own, as a browser on Windows paints
 * the real font (its `usWin` metrics; the typographic ones for a font that
 * sets USE_TYPO_METRICS, as Aptos does: 0.939/0.282 of the em, not its
 * 1.010/0.275), read in the browser and unrounded: canvas
 * metrics at a small size come back in whole pixels, and numbers read that way
 * were off by up to 2.5% of the em (Notes §46).
 * Widths were measured per weight: bold differs from regular by up to 5.6%, so
 * `bold` carries its own numbers (used by faces of weight 600 and up) wherever
 * the difference exceeds about 1.5% for the Arabic fonts, and always for the
 * Aptos families (see above).
 */
export interface MetricOverrides {
  sizeAdjustPercent: number;
  ascentOverridePercent: number;
  descentOverridePercent: number;
}

export interface FontCalibration extends MetricOverrides {
  /** Numbers for bold faces (weight 600 and up); absent when bold measures the same as regular. */
  bold?: MetricOverrides;
}

export const FONT_CALIBRATION: Record<string, FontCalibration> = {
  "Simplified Arabic": { sizeAdjustPercent: 130.3, ascentOverridePercent: 90.5, descentOverridePercent: 36.7 },
  "Traditional Arabic": {
    sizeAdjustPercent: 80.0,
    ascentOverridePercent: 124.3,
    descentOverridePercent: 62.6,
    bold: { sizeAdjustPercent: 84.5, ascentOverridePercent: 121.1, descentOverridePercent: 60.3 },
  },
  "Sakkal Majalla": { sizeAdjustPercent: 91.0, ascentOverridePercent: 97.1, descentOverridePercent: 56.3 },
  "Arabic Typesetting": { sizeAdjustPercent: 67.1, ascentOverridePercent: 104.4, descentOverridePercent: 65.4 },
  Aptos: {
    sizeAdjustPercent: 92.2,
    ascentOverridePercent: 101.8,
    descentOverridePercent: 30.6,
    bold: { sizeAdjustPercent: 93.2, ascentOverridePercent: 100.7, descentOverridePercent: 30.2 },
  },
  "Aptos Display": {
    sizeAdjustPercent: 90.7,
    ascentOverridePercent: 103.5,
    descentOverridePercent: 31.1,
    bold: { sizeAdjustPercent: 93.7, ascentOverridePercent: 100.2, descentOverridePercent: 30.1 },
  },
  "Aptos Narrow": {
    sizeAdjustPercent: 96.9,
    ascentOverridePercent: 96.9,
    descentOverridePercent: 29.1,
    bold: { sizeAdjustPercent: 95.7, ascentOverridePercent: 98.1, descentOverridePercent: 29.4 },
  },
};

/**
 * Unicode ranges of the Arabic script: Arabic, Supplement, Extended-B and
 * Extended-A, and Presentation Forms A and B. Limits an Arabic-script fallback
 * face to Arabic text so it never touches Latin, digits, or punctuation
 * (verified: `tools/font-calibration/alias-composition.ts`).
 */
export const ARABIC_UNICODE_RANGE = "U+0600-06FF, U+0750-077F, U+0870-08FF, U+FB50-FDFF, U+FE70-FEFF";

/**
 * A Latin-only substitute (Arimo, Tinos, Caladea, Carlito) has no Arabic
 * glyphs, but its Office original often does, and Word sets a great deal of
 * Arabic text in those fonts: in the fixtures Arial carries 183 Arabic runs and
 * Times New Roman 12, while the four Arabic-named fonts above carry none.
 * Without an Arabic face under the alias, that text falls through to whatever
 * `sans-serif` the reader's system has (on Windows that happens to be Arial's
 * own Arabic, which is why this gap never showed on the development machine).
 * So the resolver adds a second face to the same alias, limited to
 * `ARABIC_UNICODE_RANGE` and calibrated to the original's Arabic widths.
 * Arial's and Times New Roman's Arabic match Simplified Arabic's within 1.3%
 * (the face Windows pairs with them, per Wikipedia's history of it), so they
 * take the same Lateef face. Calibri 6 (Windows 11 / current Office) carries
 * its own Arabic and is measured directly. Cambria has no Arabic glyphs; Word
 * draws that text in Times New Roman (read off Word's PDF export:
 * `tools/font-calibration/word-arabic-fallback.ps1`), so Cambria's Arabic is
 * measured against Times New Roman's. Courier New is deliberately absent: its
 * Arabic is monospace, every proportional candidate is off by 45% or more on
 * individual paragraphs, and no calibration can fix that. Tahoma and Segoe UI
 * are not in the table at all.
 *
 * Aptos has no Arabic glyphs, and Excel and Word both draw Arabic set in any
 * Aptos family in Arial, at Arial's own widths (read off each app's PDF
 * export: `tools/font-calibration/aptos-arabic-fallback.ps1`). That is the
 * usual case in a spreadsheet: a cell has one font and no complex-script
 * slot, and Aptos Narrow is Excel's default. So the three Aptos families take
 * Arial's face, fitted through their own alias: the spaces there come from
 * the size-adjusted Latin face, narrower than Arial's, so Lateef is set
 * 3-5% larger than under Arial to keep the total width Office draws.
 *
 * Each calibration is measured the way the face is painted: through the alias,
 * next to its Latin substitute. The Arabic face covers only the Arabic script,
 * so the spaces, Western digits and punctuation inside Arabic text come from
 * Arimo/Tinos/Caladea/Carlito at their own size; calibrating the Arabic
 * substitute on its own counts its own spaces instead, and left Arial's Arabic
 * about 3% too wide (`tools/font-calibration/arabic-fallback-fit.ts`,
 * Notes §46). That is also why Cambria, fitted through Caladea (which
 * paints its spaces here), has numbers of its own rather than Times New
 * Roman's.
 */
export interface ArabicScriptFallback {
  substitute: string;
  calibration: FontCalibration;
}

export const ARABIC_SCRIPT_FALLBACKS: Record<string, ArabicScriptFallback> = {
  Arial: { substitute: "Lateef", calibration: { sizeAdjustPercent: 126.9, ascentOverridePercent: 71.3, descentOverridePercent: 16.7 } },
  "Times New Roman": { substitute: "Lateef", calibration: { sizeAdjustPercent: 126.9, ascentOverridePercent: 70.2, descentOverridePercent: 17.0 } },
  Cambria: { substitute: "Lateef", calibration: { sizeAdjustPercent: 128.9, ascentOverridePercent: 69.1, descentOverridePercent: 16.8 } },
  Calibri: {
    substitute: "Markazi Text",
    calibration: {
      sizeAdjustPercent: 104.3,
      ascentOverridePercent: 91.3,
      descentOverridePercent: 25.8,
      bold: { sizeAdjustPercent: 102.1, ascentOverridePercent: 93.3, descentOverridePercent: 26.3 },
    },
  },
  "Calibri Light": {
    substitute: "Markazi Text",
    calibration: { sizeAdjustPercent: 101.9, ascentOverridePercent: 93.4, descentOverridePercent: 26.4 },
  },
  Aptos: { substitute: "Lateef", calibration: { sizeAdjustPercent: 130.9, ascentOverridePercent: 69.1, descentOverridePercent: 16.2 } },
  "Aptos Display": { substitute: "Lateef", calibration: { sizeAdjustPercent: 131.4, ascentOverridePercent: 68.9, descentOverridePercent: 16.1 } },
  "Aptos Narrow": { substitute: "Lateef", calibration: { sizeAdjustPercent: 132.6, ascentOverridePercent: 68.3, descentOverridePercent: 16.0 } },
};

const LOOKUP = new Map(
  Object.entries(FONT_SUBSTITUTIONS).map(([requested, substitute]) => [requested.toLowerCase(), substitute] as const),
);
const CALIBRATION_LOOKUP = new Map(
  Object.entries(FONT_CALIBRATION).map(([requested, calibration]) => [requested.toLowerCase(), calibration] as const),
);
const ARABIC_FALLBACK_LOOKUP = new Map(
  Object.entries(ARABIC_SCRIPT_FALLBACKS).map(([requested, fallback]) => [requested.toLowerCase(), fallback] as const),
);

function normalizeFamily(requestedFamily: string): string {
  return requestedFamily.trim().replace(/^["']+|["']+$/g, "");
}

/** Normalize a requested family and map it to its substitute, if known. */
export function substituteFamily(requestedFamily: string): string | undefined {
  const normalized = normalizeFamily(requestedFamily);
  if (normalized === "") return undefined;
  return LOOKUP.get(normalized.toLowerCase());
}

/** The metric-override descriptors for a requested family, if it needs any (see `FONT_CALIBRATION`). */
export function calibrationFor(requestedFamily: string): FontCalibration | undefined {
  return CALIBRATION_LOOKUP.get(normalizeFamily(requestedFamily).toLowerCase());
}

/** The Arabic-script face to add under a Latin-only substitute's alias, if its Office original also covers Arabic. */
export function arabicScriptFallbackFor(requestedFamily: string): ArabicScriptFallback | undefined {
  return ARABIC_FALLBACK_LOOKUP.get(normalizeFamily(requestedFamily).toLowerCase());
}

/** Google's own Arabic subset always claims the basic Arabic block; the Latin and other subsets never do. */
const ARABIC_SUBSET_RANGE = /U\+0600-06FF/;

/** A face's numbers for its weight: bold faces use `bold` when the calibration has separate bold numbers. */
function overridesFor(calibration: FontCalibration | undefined, weight: string): MetricOverrides | undefined {
  if (calibration === undefined) return undefined;
  return Number.parseInt(weight, 10) >= 600 && calibration.bold !== undefined ? calibration.bold : calibration;
}

/** Google Fonts CSS endpoint serving every subset (latin, arabic, ...) the browser may paint. */
export function fontCssUrl(families: readonly string[]): string {
  const params = families.map(
    (family) => `family=${encodeURIComponent(family).replace(/%20/g, "+")}:ital,wght@0,400;0,700;1,400;1,700`,
  );
  return `https://fonts.googleapis.com/css2?${params.join("&")}&display=swap`;
}

export interface ParsedFontFace {
  family: string;
  style: string;
  weight: string;
  unicodeRange: string;
  url: string;
}

/** Parse `@font-face` blocks out of a Google Fonts CSS response. Pure and headless-testable. */
export function parseFontCss(css: string): ParsedFontFace[] {
  const out: ParsedFontFace[] = [];
  for (const block of css.matchAll(/@font-face\s*{([^}]*)}/g)) {
    const body = block[1] ?? "";
    const field = (name: string): string | undefined => {
      const match = new RegExp(`${name}\\s*:\\s*([^;]+);`).exec(body);
      return match?.[1]?.trim();
    };
    const family = field("font-family")?.replace(/^['"]|['"]$/g, "");
    const url = /url\((https:[^)]+)\)/.exec(body)?.[1];
    if (family === undefined || family === "" || url === undefined) continue;
    out.push({
      family,
      style: field("font-style") ?? "normal",
      weight: field("font-weight") ?? "400",
      unicodeRange: field("unicode-range") ?? "",
      url,
    });
  }
  return out;
}

export interface ResolverDependencies {
  fetchFn?: typeof fetch;
  FontFaceCtor?: typeof FontFace;
  documentFonts?: FontFaceSet;
}

/**
 * Build a `FontResolver` over injectable globals. The default (`{}`) reads
 * `fetch`/`FontFace`/`document.fonts` from `globalThis`, so headless runtimes
 * without them resolve `undefined` instead of throwing.
 *
 * The returned promise never rejects: when delivery fails (offline CDN,
 * unknown subset) it resolves `undefined` and the browser falls back to its
 * own fonts. This matters because the format renderer awaits every resolver
 * result inside `read()` - a rejection there would fail the whole read.
 * The renderer ignores the resolved value; it only awaits settlement. Faces
 * are registered, never loaded here: the browser fetches each one when text
 * first needs it, and a renderer that must paint with the delivered face
 * before measuring or checking it (DOCX layout pagination, the DOCX and PPTX
 * fidelity warnings) loads the faces it paints itself (`awaitDocumentFonts`;
 * Notes §45, §50).
 */
export function createResolver(dependencies: ResolverDependencies = {}): FontResolver {
  // One CSS fetch per *substitute* family, shared by every requested family
  // that maps to it, whether as its main face or as an Arabic fallback (Arial,
  // Times New Roman and Simplified Arabic all deliver through Lateef).
  const cssCache = new Map<string, Promise<ParsedFontFace[] | undefined>>();
  // One set of registered `FontFace`s per *requested* family, each aliased
  // under `fontAlias(requestedFamily)` - never the substitute's own name, so
  // two requested families sharing one substitute (Arial and Simplified
  // Arabic, both through Lateef for Arabic text) still register separately,
  // each with its own calibration.
  const resolvedCache = new Map<string, Promise<FontFace | undefined>>();

  return (requestedFamily: string): Promise<FontFace> | undefined => {
    const substitute = substituteFamily(requestedFamily);
    if (substitute === undefined) return undefined;
    const cached = resolvedCache.get(requestedFamily);
    if (cached !== undefined) return cached as Promise<FontFace>;

    const fetchFn = dependencies.fetchFn
      ?? (typeof globalThis.fetch === "function" ? globalThis.fetch.bind(globalThis) : undefined);
    const FontFaceCtor = dependencies.FontFaceCtor
      ?? (typeof globalThis.FontFace === "function" ? globalThis.FontFace : undefined);
    const fonts = dependencies.documentFonts
      ?? (typeof globalThis.document === "object" ? globalThis.document.fonts : undefined);
    if (fetchFn === undefined || FontFaceCtor === undefined || fonts === undefined) return undefined;

    const facesOf = (family: string): Promise<ParsedFontFace[] | undefined> => {
      let task = cssCache.get(family);
      if (task === undefined) {
        task = (async (): Promise<ParsedFontFace[] | undefined> => {
          try {
            const response = await fetchFn(fontCssUrl([family]));
            if (!response.ok) return undefined;
            return parseFontCss(await response.text()).filter((face) => face.family === family);
          } catch {
            return undefined;
          }
        })();
        cssCache.set(family, task);
      }
      return task;
    };

    const alias = fontAlias(requestedFamily);
    // Register each face under the alias and return the regular one. `arabicOnly`
    // faces are a second, script-limited face on the same alias: they carry
    // `ARABIC_UNICODE_RANGE` (never the CDN's own subset range, which also claims
    // punctuation the Latin face owns).
    //
    // No face is loaded here: the browser fetches one when text first needs it
    // (its style, weight and `unicode-range`), so a document downloads only the
    // subsets and styles it paints — English in regular weight is the basic
    // Latin file alone, 10-30 KB, where loading every face cost 175-784 KB per
    // family (Notes §50). A renderer that measures or checks painted text
    // loads exactly what it paints first (`awaitDocumentFonts`).
    const register = (faces: readonly ParsedFontFace[], calibration: FontCalibration | undefined, arabicOnly: boolean): FontFace | undefined => {
      let primary: FontFace | undefined;
      for (const face of faces) {
        const overrides = overridesFor(calibration, face.weight);
        const unicodeRange = arabicOnly ? ARABIC_UNICODE_RANGE : face.unicodeRange;
        try {
          const constructed = new FontFaceCtor(alias, `url(${face.url})`, {
            style: face.style,
            weight: face.weight,
            ...(unicodeRange === "" ? {} : { unicodeRange }),
            ...(overrides === undefined ? {} : {
              sizeAdjust: `${overrides.sizeAdjustPercent}%`,
              ascentOverride: `${overrides.ascentOverridePercent}%`,
              descentOverride: `${overrides.descentOverridePercent}%`,
            }),
            display: "swap",
          });
          fonts.add(constructed);
          if (primary === undefined && face.style === "normal" && face.weight === "400") primary = constructed;
        } catch {
          continue;
        }
      }
      return primary;
    };

    const primaryTask = facesOf(substitute).then((faces) => faces === undefined ? undefined : register(faces, calibrationFor(requestedFamily), false));
    const arabic = arabicScriptFallbackFor(requestedFamily);
    const arabicTask = arabic === undefined
      ? Promise.resolve()
      : facesOf(arabic.substitute).then((faces) => {
        // Upright faces only: Word slants an Arabic run in Arial or Times New
        // Roman synthetically (neither has an Arabic italic), so an italic
        // request should find the upright face and be slanted the same way.
        if (faces !== undefined) register(faces.filter((face) => face.style === "normal" && ARABIC_SUBSET_RANGE.test(face.unicodeRange)), arabic.calibration, true);
      });
    const task: Promise<FontFace | undefined> = Promise.all([primaryTask, arabicTask]).then(([primary]) => primary);
    resolvedCache.set(requestedFamily, task);
    return task as Promise<FontFace>;
  };
}

/**
 * CDN-backed default resolver for the substitution table above. Lazy per
 * family: the first request for `Calibri` fetches Carlito's CSS once; later
 * families (and later documents) reuse the cache.
 */
export const defaultResolver: FontResolver = createResolver();
