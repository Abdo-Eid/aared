/**
 * Metric-compatible font substitution + a CDN-backed default resolver.
 * See docs/ARCHITECTURE.md § Fonts.
 *
 * Delivery decision: no font binaries ship in this package (see the delivery
 * report). `defaultResolver` maps each requested Office family to its
 * metric-compatible OFL substitute, fetches that substitute's `@font-face`
 * CSS from Google Fonts once per family, and registers every subset face via
 * `FontFace` on `document.fonts` — so faces cross shadow-DOM boundaries and
 * the browser fetches only the subsets actually painted. The resolver is
 * main-thread-only: `core` receives it through `open()` but never posts it
 * to the Worker (see the API reference page § Font resolution).
 */

// TODO: import { FontResolver } from "aared/core" once `aared` has a
// real build to resolve against — `aared`'s package.json exports point at
// dist/, which doesn't exist before the first build, so a cross-package
// type import here fails typecheck until then. Structurally identical stub
// in the meantime; keep in sync with packages/aared/src/core/index.ts.
type FontResolver = (
  requestedFamily: string,
) => Promise<FontFace> | undefined;

/** Requested Office font family -> metric-compatible OFL substitute. */
export const FONT_SUBSTITUTIONS: Record<string, string> = {
  Calibri: "Carlito",
  "Calibri Light": "Carlito",
  Cambria: "Caladea",
  Arial: "Liberation Sans",
  "Times New Roman": "Liberation Serif",
  "Courier New": "Liberation Mono",
  "Traditional Arabic": "Noto Naskh Arabic",
  "Sakkal Majalla": "Noto Naskh Arabic",
  "Simplified Arabic": "Noto Sans Arabic",
  "Arabic Typesetting": "Amiri",
};

/** Substitute families this resolver can deliver, for diagnostics. */
export const SUBSTITUTE_FAMILIES: readonly string[] = [
  "Carlito",
  "Caladea",
  "Liberation Sans",
  "Liberation Serif",
  "Liberation Mono",
  "Noto Naskh Arabic",
  "Noto Sans Arabic",
  "Amiri",
];

const LOOKUP = new Map(
  Object.entries(FONT_SUBSTITUTIONS).map(([requested, substitute]) => [requested.toLowerCase(), substitute] as const),
);

/** Normalize a requested family and map it to its OFL substitute, if known. */
export function substituteFamily(requestedFamily: string): string | undefined {
  const normalized = requestedFamily.trim().replace(/^["']+|["']+$/g, "");
  if (normalized === "") return undefined;
  return LOOKUP.get(normalized.toLowerCase());
}

/** Google Fonts CSS endpoint serving every subset (latin, arabic, …) the browser may paint. */
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
 * result inside `read()` — a rejection there would fail the whole read.
 * The renderer ignores the resolved value; it only awaits settlement.
 */
export function createResolver(dependencies: ResolverDependencies = {}): FontResolver {
  const cache = new Map<string, Promise<FontFace | undefined>>();
  return (requestedFamily: string): Promise<FontFace> | undefined => {
    const substitute = substituteFamily(requestedFamily);
    if (substitute === undefined) return undefined;
    const cached = cache.get(substitute);
    if (cached !== undefined) return cached as Promise<FontFace>;
    const fetchFn = dependencies.fetchFn
      ?? (typeof globalThis.fetch === "function" ? globalThis.fetch.bind(globalThis) : undefined);
    const FontFaceCtor = dependencies.FontFaceCtor
      ?? (typeof globalThis.FontFace === "function" ? globalThis.FontFace : undefined);
    const fonts = dependencies.documentFonts
      ?? (typeof globalThis.document === "object" ? globalThis.document.fonts : undefined);
    if (fetchFn === undefined || FontFaceCtor === undefined || fonts === undefined) return undefined;
    const task: Promise<FontFace | undefined> = (async (): Promise<FontFace | undefined> => {
      try {
        const response = await fetchFn(fontCssUrl([substitute]));
        if (!response.ok) return undefined;
        const faces = parseFontCss(await response.text()).filter((face) => face.family === substitute);
        let primary: FontFace | undefined;
        for (const face of faces) {
          try {
            const constructed = new FontFaceCtor(face.family, `url(${face.url})`, {
              style: face.style,
              weight: face.weight,
              ...(face.unicodeRange === "" ? {} : { unicodeRange: face.unicodeRange }),
              display: "swap",
            });
            fonts.add(constructed);
            // Load in the background: painting pulls used subsets without
            // blocking `read()` on every font file.
            constructed.load().catch(() => undefined);
            if (primary === undefined && face.style === "normal" && face.weight === "400") primary = constructed;
          } catch {
            continue;
          }
        }
        return primary;
      } catch {
        return undefined;
      }
    })();
    cache.set(substitute, task);
    return task as Promise<FontFace>;
  };
}

/**
 * CDN-backed default resolver for the metric-compatible mapping above.
 * Lazy per family: the first request for `Calibri` fetches Carlito's CSS
 * once; later families (and later documents) reuse the cache.
 */
export const defaultResolver: FontResolver = createResolver();
