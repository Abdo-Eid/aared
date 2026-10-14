import { describe, expect, test } from "bun:test";
import { fontAlias } from "@aared/viewer/core";
import {
  ARABIC_SCRIPT_FALLBACKS,
  ARABIC_UNICODE_RANGE,
  arabicScriptFallbackFor,
  calibrationFor,
  createResolver,
  defaultResolver,
  FONT_CALIBRATION,
  FONT_SUBSTITUTIONS,
  fontCssUrl,
  parseFontCss,
  substituteFamily,
} from "../src/index.js";

const SAMPLE_CSS = `/* arabic */
@font-face {
  font-family: 'Amiri';
  font-style: normal;
  font-weight: 400;
  font-display: swap;
  src: url(https://fonts.gstatic.com/s/amiri/v30/J7aRnpd8CGxBHpUrtLMA7w.woff2) format('woff2');
  unicode-range: U+0600-06FF, U+0750-077F;
}
/* latin */
@font-face {
  font-family: 'Amiri';
  font-style: normal;
  font-weight: 700;
  font-display: swap;
  src: url(https://fonts.gstatic.com/s/amiri/v30/J7acnpd8CGxBHp2VkaY_zp4.woff2) format('woff2');
  unicode-range: U+0000-00FF;
}`;

describe("font substitution table", () => {
  test("covers the fonts named in docs/ARCHITECTURE.md", () => {
    expect(FONT_SUBSTITUTIONS.Calibri).toBe("Carlito");
    expect(FONT_SUBSTITUTIONS.Arial).toBe("Arimo");
    expect(FONT_SUBSTITUTIONS["Times New Roman"]).toBe("Tinos");
    expect(FONT_SUBSTITUTIONS["Courier New"]).toBe("Cousine");
    // The four Office Arabic fonts each go to whichever open face measured
    // best for them against real Arabic text, with a per-family calibration,
    // not per-name guesses (tools/font-calibration/arabic-survey.ts).
    expect(FONT_SUBSTITUTIONS["Traditional Arabic"]).toBe("Amiri");
    expect(FONT_SUBSTITUTIONS["Sakkal Majalla"]).toBe("Markazi Text");
    expect(FONT_SUBSTITUTIONS["Simplified Arabic"]).toBe("Lateef");
    expect(FONT_SUBSTITUTIONS["Arabic Typesetting"]).toBe("Amiri");
    // Aptos has no clone: each family goes to the one open face that keeps its
    // letter heights once widths match (tools/font-calibration/latin-survey.ts).
    expect(FONT_SUBSTITUTIONS.Aptos).toBe("Radio Canada");
    expect(FONT_SUBSTITUTIONS["Aptos Display"]).toBe("Roboto");
    expect(FONT_SUBSTITUTIONS["Aptos Narrow"]).toBe("IBM Plex Sans Condensed");
  });

  test("lookup is case-insensitive and tolerates quotes and whitespace", () => {
    expect(substituteFamily("calibri")).toBe("Carlito");
    expect(substituteFamily("  \"Calibri\"  ")).toBe("Carlito");
    expect(substituteFamily("TRADITIONAL ARABIC")).toBe("Amiri");
    expect(substituteFamily("Cambria")).toBe("Caladea");
    expect(substituteFamily("aptos narrow")).toBe("IBM Plex Sans Condensed");
    expect(substituteFamily("Segoe UI")).toBeUndefined();
    expect(substituteFamily("")).toBeUndefined();
  });

  test("only the non-clone substitutions (the Arabic fonts and Aptos) carry a metric calibration", () => {
    expect(calibrationFor("Calibri")).toBeUndefined();
    expect(calibrationFor("Arial")).toBeUndefined();
    for (const family of ["Simplified Arabic", "Traditional Arabic", "Sakkal Majalla", "Arabic Typesetting", "Aptos", "Aptos Display", "Aptos Narrow"]) {
      const calibration = calibrationFor(family);
      expect(calibration, family).toBeDefined();
      // Sane percentages: a size-adjust within a plausible scaling range, and
      // overrides that land in a normal ascent/descent range once you divide
      // out size-adjust's own scaling (see the interaction probe).
      expect(calibration!.sizeAdjustPercent).toBeGreaterThan(50);
      expect(calibration!.sizeAdjustPercent).toBeLessThanOrEqual(150);
      expect(calibration!.ascentOverridePercent).toBeGreaterThan(50);
      expect(calibration!.descentOverridePercent).toBeGreaterThan(20);
      if (calibration!.bold !== undefined) {
        // Bold measured differently enough to need its own numbers, but never wildly so.
        expect(Math.abs(calibration!.bold.sizeAdjustPercent / calibration!.sizeAdjustPercent - 1)).toBeLessThan(0.1);
      }
    }
    expect(calibrationFor("simplified arabic")).toEqual(FONT_CALIBRATION["Simplified Arabic"]);
    // Aptos's own vertical metrics are its typographic ones (it sets USE_TYPO_METRICS):
    // 0.939/0.282 of the em, what the overrides carry once size-adjust is divided out.
    for (const family of ["Aptos", "Aptos Display", "Aptos Narrow"]) {
      const calibration = calibrationFor(family)!;
      for (const numbers of [calibration, calibration.bold!]) {
        expect(numbers.ascentOverridePercent * numbers.sizeAdjustPercent / 1e4, family).toBeCloseTo(0.939, 2);
        expect(numbers.descentOverridePercent * numbers.sizeAdjustPercent / 1e4, family).toBeCloseTo(0.282, 2);
      }
    }
  });

  test("Latin-only substitutes whose Office original also has Arabic get an Arabic-script fallback", () => {
    // Arial carries most Arabic runs in real documents, and Arimo has no
    // Arabic glyphs: without this the text falls to the reader's own font.
    for (const family of ["Arial", "Times New Roman", "Cambria", "Calibri", "Calibri Light", "Aptos", "Aptos Display", "Aptos Narrow"]) {
      const fallback = arabicScriptFallbackFor(family);
      expect(fallback, family).toBeDefined();
      expect(fallback!.calibration.sizeAdjustPercent).toBeGreaterThan(50);
      expect(fallback!.calibration.sizeAdjustPercent).toBeLessThanOrEqual(150);
    }
    expect(arabicScriptFallbackFor("  \"arial\" ")).toBe(ARABIC_SCRIPT_FALLBACKS.Arial);
    // Word draws Cambria's Arabic in Times New Roman (Cambria has none): the same face,
    // calibrated against Times New Roman's Arabic but through Caladea, which paints its spaces.
    expect(arabicScriptFallbackFor("Cambria")!.substitute).toBe(arabicScriptFallbackFor("Times New Roman")!.substitute);
    // Courier New's Arabic is monospace: no proportional face can match it, so none is offered.
    expect(arabicScriptFallbackFor("Courier New")).toBeUndefined();
    // The four Arabic-named fonts already deliver an Arabic face as their main substitute.
    expect(arabicScriptFallbackFor("Simplified Arabic")).toBeUndefined();
    expect(arabicScriptFallbackFor("Segoe UI")).toBeUndefined();
    // Aptos has no Arabic glyphs; Excel and Word draw that text in Arial
    // (aptos-arabic-fallback.ps1), so Aptos takes Arial's face, fitted through its own alias.
    for (const family of ["Aptos", "Aptos Display", "Aptos Narrow"]) {
      expect(arabicScriptFallbackFor(family)!.substitute, family).toBe(ARABIC_SCRIPT_FALLBACKS.Arial!.substitute);
    }
  });
});

describe("CDN font CSS", () => {
  test("requests regular, bold, and italics in one stylesheet per family", () => {
    const url = fontCssUrl(["Amiri", "Carlito"]);
    expect(url.startsWith("https://fonts.googleapis.com/css2?")).toBe(true);
    expect(url).toContain("family=Amiri:ital,wght@0,400;0,700;1,400;1,700");
    expect(url).toContain("family=Carlito:ital,wght@0,400;0,700;1,400;1,700");
    expect(url).toContain("display=swap");
  });

  test("parses subset faces out of a stylesheet", () => {
    const faces = parseFontCss(SAMPLE_CSS);
    expect(faces).toHaveLength(2);
    expect(faces[0]).toMatchObject({
      family: "Amiri",
      style: "normal",
      weight: "400",
      url: "https://fonts.gstatic.com/s/amiri/v30/J7aRnpd8CGxBHpUrtLMA7w.woff2",
    });
    expect(faces[0]?.unicodeRange).toContain("U+0600-06FF");
    expect(faces[1]).toMatchObject({ weight: "700" });
    expect(parseFontCss("body { color: red; }")).toEqual([]);
  });
});

/** A `FontFace` fake that records every descriptor it was constructed with, and every `load()`. */
class FakeFontFace {
  family: string;
  source: string;
  descriptors: Record<string, string> | undefined;
  loads = 0;
  constructor(family: string, source: string, descriptors?: Record<string, string>) {
    this.family = family;
    this.source = source;
    this.descriptors = descriptors;
  }
  async load(): Promise<this> {
    this.loads++;
    return this;
  }
}

/** A Google Fonts stylesheet with the subsets each real family serves, per style and weight. */
function familyCss(family: string, options: { arabic?: boolean } = {}): string {
  const slug = family.toLowerCase().replace(/ /g, "");
  const subsets: Array<[string, string]> = [
    ["latin", "U+0000-00FF, U+2000-206F"],
    ["latin-ext", "U+0100-02BA"],
  ];
  if (options.arabic !== false) subsets.push(["arabic", "U+0600-06FF, U+0750-077F, U+200C-200E, U+2010-2011, U+FB50-FDFF"]);
  const blocks: string[] = [];
  for (const style of ["normal", "italic"]) {
    for (const weight of [400, 700]) {
      for (const [name, range] of subsets) {
        blocks.push(
          `/* ${name} */\n@font-face {\n  font-family: '${family}';\n  font-style: ${style};\n  font-weight: ${weight};\n`
          + `  src: url(https://fonts.gstatic.com/s/${slug}/${style}${weight}${name}.woff2) format('woff2');\n  unicode-range: ${range};\n}`,
        );
      }
    }
  }
  return blocks.join("\n");
}

function requestedFamily(input: unknown): string {
  return decodeURIComponent(/family=([^:&]+)/.exec(String(input))![1]!.replace(/\+/g, " "));
}

/** A `fetch` that answers each Google Fonts CSS request with that family's stylesheet and records the requests. */
function cssFetch(requested: string[], options: { arabic?: (family: string) => boolean } = {}): typeof fetch {
  return (async (input: unknown) => {
    const family = requestedFamily(input);
    requested.push(family);
    return new Response(familyCss(family, { arabic: options.arabic?.(family) ?? true }), { status: 200 });
  }) as unknown as typeof fetch;
}

describe("default resolver", () => {
  test("returns undefined for unknown families without touching the network", async () => {
    let fetched = 0;
    const resolver = createResolver({
      fetchFn: (async () => {
        fetched++;
        throw new Error("must not fetch");
      }) as unknown as typeof fetch,
    });
    expect(resolver("Segoe UI")).toBeUndefined();
    expect(fetched).toBe(0);
  });

  test("registers a delivered substitute under the requested family's alias, never its own name", async () => {
    const added: Array<{ family: string }> = [];
    const fonts = { add(face: { family: string }): void { added.push({ family: face.family }); } };
    const resolver = createResolver({
      fetchFn: (async () => new Response(SAMPLE_CSS, { status: 200 })) as unknown as typeof fetch,
      FontFaceCtor: FakeFontFace as unknown as typeof FontFace,
      documentFonts: fonts as unknown as FontFaceSet,
    });
    const face = await resolver("Arabic Typesetting");
    expect(face).toBeInstanceOf(FakeFontFace);
    // Never "Amiri" (the substitute's own name — that would leave the
    // browser to pick it only by coincidence) and never "Arabic Typesetting"
    // (that would shadow a reader who has the real font installed).
    expect(added.every((face) => face.family === fontAlias("Arabic Typesetting"))).toBe(true);
    expect(added.some((face) => face.family === "Amiri")).toBe(false);
    expect(added.some((face) => face.family === "Arabic Typesetting")).toBe(false);
  });

  test("two requested families sharing one substitute fetch its CSS once but register separately, each with its own calibration", async () => {
    const requested: string[] = [];
    const added: FakeFontFace[] = [];
    const resolver = createResolver({
      fetchFn: cssFetch(requested),
      FontFaceCtor: FakeFontFace as unknown as typeof FontFace,
      documentFonts: { add(face: FakeFontFace): void { added.push(face); } } as unknown as FontFaceSet,
    });
    const traditional = await resolver("Traditional Arabic");
    const typesetting = await resolver("Arabic Typesetting");
    expect(requested).toEqual(["Amiri"]);
    expect(traditional).not.toBe(typesetting);
    const traditionalFaces = added.filter((face) => face.family === fontAlias("Traditional Arabic"));
    const typesettingFaces = added.filter((face) => face.family === fontAlias("Arabic Typesetting"));
    expect(traditionalFaces.length).toBeGreaterThan(0);
    expect(typesettingFaces.length).toBeGreaterThan(0);
    // Each alias carries its own calibration, computed for its own font.
    expect(traditionalFaces[0]?.descriptors?.sizeAdjust).toBe(`${FONT_CALIBRATION["Traditional Arabic"]!.sizeAdjustPercent}%`);
    expect(typesettingFaces[0]?.descriptors?.sizeAdjust).toBe(`${FONT_CALIBRATION["Arabic Typesetting"]!.sizeAdjustPercent}%`);
    expect(traditionalFaces[0]?.descriptors?.sizeAdjust).not.toBe(typesettingFaces[0]?.descriptors?.sizeAdjust);
  });

  test("bold faces take the calibration's bold numbers when it has them, regular faces the regular ones", async () => {
    const added: FakeFontFace[] = [];
    await createResolver({
      fetchFn: cssFetch([]),
      FontFaceCtor: FakeFontFace as unknown as typeof FontFace,
      documentFonts: { add(face: FakeFontFace): void { added.push(face); } } as unknown as FontFaceSet,
    })("Traditional Arabic");
    const calibration = FONT_CALIBRATION["Traditional Arabic"]!;
    const regular = added.find((face) => face.descriptors?.weight === "400")!;
    const bold = added.find((face) => face.descriptors?.weight === "700")!;
    expect(regular.descriptors?.sizeAdjust).toBe(`${calibration.sizeAdjustPercent}%`);
    expect(bold.descriptors?.sizeAdjust).toBe(`${calibration.bold!.sizeAdjustPercent}%`);
    expect(bold.descriptors?.ascentOverride).toBe(`${calibration.bold!.ascentOverridePercent}%`);

    // A calibration with no separate bold numbers applies its one set to both weights.
    const flat: FakeFontFace[] = [];
    await createResolver({
      fetchFn: cssFetch([]),
      FontFaceCtor: FakeFontFace as unknown as typeof FontFace,
      documentFonts: { add(face: FakeFontFace): void { flat.push(face); } } as unknown as FontFaceSet,
    })("Arabic Typesetting");
    expect(new Set(flat.map((face) => face.descriptors?.sizeAdjust)).size).toBe(1);
  });

  test("a metric-clone substitution (no calibration entry) carries no size-adjust or overrides", async () => {
    const added: FakeFontFace[] = [];
    const resolver = createResolver({
      fetchFn: (async (input: unknown) => new Response(
        requestedFamily(input) === "Carlito"
          ? `@font-face{font-family:'Carlito';font-style:normal;font-weight:400;src:url(https://fonts.gstatic.com/s/carlito/x.woff2)}`
          : "",
        { status: 200 },
      )) as unknown as typeof fetch,
      FontFaceCtor: FakeFontFace as unknown as typeof FontFace,
      documentFonts: { add(face: FakeFontFace): void { added.push(face); } } as unknown as FontFaceSet,
    });
    await resolver("Calibri");
    // Carlito's own face only: Markazi Text's (Calibri's Arabic fallback) stylesheet came back empty.
    expect(added).toHaveLength(1);
    expect(added[0]?.descriptors?.sizeAdjust).toBeUndefined();
    expect(added[0]?.descriptors?.ascentOverride).toBeUndefined();
  });

  test("Arial gets Arimo for Latin plus a calibrated, Arabic-limited Lateef face on the same alias", async () => {
    const requested: string[] = [];
    const added: FakeFontFace[] = [];
    const resolver = createResolver({
      // Arimo has no Arabic subset, exactly like the real family.
      fetchFn: cssFetch(requested, { arabic: (family) => family !== "Arimo" }),
      FontFaceCtor: FakeFontFace as unknown as typeof FontFace,
      documentFonts: { add(face: FakeFontFace): void { added.push(face); } } as unknown as FontFaceSet,
    });
    const primary = await resolver("Arial");
    expect(requested.sort()).toEqual(["Arimo", "Lateef"]);
    expect(added.every((face) => face.family === fontAlias("Arial"))).toBe(true);
    // The face the promise resolves to is still Arimo's regular Latin face.
    expect((primary as unknown as FakeFontFace).source).toContain("/arimo/");

    const arimo = added.filter((face) => face.source.includes("/arimo/"));
    const lateef = added.filter((face) => face.source.includes("/lateef/"));
    // Arimo is a metric clone: no calibration. Registered, never loaded: the browser fetches a
    // face when text needs it, so a document downloads only what it paints.
    expect(arimo.length).toBeGreaterThan(0);
    for (const face of arimo) {
      expect(face.descriptors?.sizeAdjust).toBeUndefined();
      expect(face.loads).toBe(0);
    }
    // Lateef contributes only its upright Arabic faces (400 and 700), limited to the Arabic
    // script — never the CDN's own range, which also claims punctuation Arimo owns —
    // calibrated to Arial's Arabic, and likewise left to load on demand.
    expect(lateef.map((face) => face.descriptors?.weight).sort()).toEqual(["400", "700"]);
    for (const face of lateef) {
      expect(face.source).toContain("arabic");
      expect(face.descriptors?.style).toBe("normal");
      expect(face.descriptors?.unicodeRange).toBe(ARABIC_UNICODE_RANGE);
      expect(face.descriptors?.sizeAdjust).toBe(`${ARABIC_SCRIPT_FALLBACKS.Arial!.calibration.sizeAdjustPercent}%`);
      expect(face.loads).toBe(0);
    }
  });

  test("Aptos gets a calibrated Radio Canada for Latin plus Arial's Arabic face, fitted through its own alias", async () => {
    const requested: string[] = [];
    const added: FakeFontFace[] = [];
    await createResolver({
      fetchFn: cssFetch(requested, { arabic: (family) => family !== "Radio Canada" }),
      FontFaceCtor: FakeFontFace as unknown as typeof FontFace,
      documentFonts: { add(face: FakeFontFace): void { added.push(face); } } as unknown as FontFaceSet,
    })("Aptos");
    expect(requested.sort()).toEqual(["Lateef", "Radio Canada"]);
    expect(added.every((face) => face.family === fontAlias("Aptos"))).toBe(true);
    const calibration = FONT_CALIBRATION.Aptos!;
    const latin = added.filter((face) => face.source.includes("/radiocanada/"));
    // Every Latin face, upright and italic, carries its weight's numbers, and none is loaded up front.
    expect(latin.length).toBeGreaterThan(0);
    for (const face of latin) {
      const numbers = face.descriptors?.weight === "700" ? calibration.bold! : calibration;
      expect(face.descriptors?.sizeAdjust).toBe(`${numbers.sizeAdjustPercent}%`);
      expect(face.descriptors?.ascentOverride).toBe(`${numbers.ascentOverridePercent}%`);
      expect(face.descriptors?.descentOverride).toBe(`${numbers.descentOverridePercent}%`);
      expect(face.loads).toBe(0);
    }
    const lateef = added.filter((face) => face.source.includes("/lateef/"));
    expect(lateef.map((face) => face.descriptors?.weight).sort()).toEqual(["400", "700"]);
    for (const face of lateef) {
      expect(face.descriptors?.unicodeRange).toBe(ARABIC_UNICODE_RANGE);
      expect(face.descriptors?.sizeAdjust).toBe(`${ARABIC_SCRIPT_FALLBACKS.Aptos!.calibration.sizeAdjustPercent}%`);
      expect(face.loads).toBe(0);
    }
  });

  test("Calibri's Arabic faces use Markazi Text with per-weight numbers", async () => {
    const added: FakeFontFace[] = [];
    await createResolver({
      fetchFn: cssFetch([], { arabic: (family) => family !== "Carlito" }),
      FontFaceCtor: FakeFontFace as unknown as typeof FontFace,
      documentFonts: { add(face: FakeFontFace): void { added.push(face); } } as unknown as FontFaceSet,
    })("Calibri");
    const markazi = added.filter((face) => face.source.includes("/markazitext/"));
    const calibration = ARABIC_SCRIPT_FALLBACKS.Calibri!.calibration;
    expect(markazi.find((face) => face.descriptors?.weight === "400")?.descriptors?.sizeAdjust).toBe(`${calibration.sizeAdjustPercent}%`);
    expect(markazi.find((face) => face.descriptors?.weight === "700")?.descriptors?.sizeAdjust).toBe(`${calibration.bold!.sizeAdjustPercent}%`);
  });

  test("Arial, Times New Roman, Cambria and Simplified Arabic share one Lateef stylesheet", async () => {
    const requested: string[] = [];
    const resolver = createResolver({
      fetchFn: cssFetch(requested),
      FontFaceCtor: FakeFontFace as unknown as typeof FontFace,
      documentFonts: { add(_face: unknown): void {} } as unknown as FontFaceSet,
    });
    await Promise.all([resolver("Arial"), resolver("Times New Roman"), resolver("Simplified Arabic"), resolver("Cambria")]);
    expect(requested.filter((family) => family === "Lateef")).toHaveLength(1);
  });

  test("a failed Arabic stylesheet still delivers the Latin face; a failed Latin one still delivers the Arabic face", async () => {
    const added: FakeFontFace[] = [];
    const fonts = { add(face: FakeFontFace): void { added.push(face); } } as unknown as FontFaceSet;
    const failing = (only: string): typeof fetch => (async (input: unknown) => {
      const family = requestedFamily(input);
      return family === only ? new Response("no", { status: 500 }) : new Response(familyCss(family), { status: 200 });
    }) as unknown as typeof fetch;

    const noArabic = await createResolver({ fetchFn: failing("Lateef"), FontFaceCtor: FakeFontFace as unknown as typeof FontFace, documentFonts: fonts })("Arial");
    expect(noArabic).toBeDefined();
    expect(added.every((face) => face.source.includes("/arimo/"))).toBe(true);

    added.length = 0;
    const noLatin = await createResolver({ fetchFn: failing("Arimo"), FontFaceCtor: FakeFontFace as unknown as typeof FontFace, documentFonts: fonts })("Arial");
    expect(noLatin).toBeUndefined();
    expect(added.length).toBeGreaterThan(0);
    expect(added.every((face) => face.source.includes("/lateef/"))).toBe(true);
  });

  test("never rejects: delivery failure resolves undefined for browser fallback", async () => {
    const resolver = createResolver({
      fetchFn: (async () => new Response("oops", { status: 500 })) as unknown as typeof fetch,
      FontFaceCtor: FakeFontFace as unknown as typeof FontFace,
      documentFonts: { add(_face: unknown): void {} } as unknown as FontFaceSet,
    });
    await expect(resolver("Calibri")!).resolves.toBeUndefined();
    const offline = createResolver({
      fetchFn: (async () => {
        throw new TypeError("fetch failed");
      }) as unknown as typeof fetch,
      FontFaceCtor: FakeFontFace as unknown as typeof FontFace,
      documentFonts: { add(_face: unknown): void {} } as unknown as FontFaceSet,
    });
    await expect(offline("Calibri")!).resolves.toBeUndefined();
    await expect(offline("Arial")!).resolves.toBeUndefined();
  });

  test("default export resolves undefined headlessly instead of throwing", () => {
    // Bun has fetch but no FontFace/document: the shared default must stay
    // inert outside a browser, since unit tests run headless by design.
    expect(defaultResolver("Calibri")).toBeUndefined();
    expect(defaultResolver("Aptos")).toBeUndefined();
    expect(defaultResolver("Segoe UI")).toBeUndefined();
  });
});
