import { describe, expect, test } from "bun:test";
import {
  createResolver,
  defaultResolver,
  FONT_SUBSTITUTIONS,
  fontCssUrl,
  parseFontCss,
  SUBSTITUTE_FAMILIES,
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
    expect(FONT_SUBSTITUTIONS.Arial).toBe("Liberation Sans");
    expect(FONT_SUBSTITUTIONS["Times New Roman"]).toBe("Liberation Serif");
    expect(FONT_SUBSTITUTIONS["Courier New"]).toBe("Liberation Mono");
    expect(FONT_SUBSTITUTIONS["Traditional Arabic"]).toBe("Noto Naskh Arabic");
    expect(FONT_SUBSTITUTIONS["Sakkal Majalla"]).toBe("Noto Naskh Arabic");
    expect(FONT_SUBSTITUTIONS["Simplified Arabic"]).toBe("Noto Sans Arabic");
    expect(FONT_SUBSTITUTIONS["Arabic Typesetting"]).toBe("Amiri");
  });

  test("every substitute is a deliverable Google Fonts family", () => {
    for (const substitute of Object.values(FONT_SUBSTITUTIONS)) {
      expect(SUBSTITUTE_FAMILIES, substitute).toContain(substitute);
    }
  });

  test("lookup is case-insensitive and tolerates quotes and whitespace", () => {
    expect(substituteFamily("calibri")).toBe("Carlito");
    expect(substituteFamily("  \"Calibri\"  ")).toBe("Carlito");
    expect(substituteFamily("TRADITIONAL ARABIC")).toBe("Noto Naskh Arabic");
    expect(substituteFamily("Cambria")).toBe("Caladea");
    expect(substituteFamily("Aptos")).toBeUndefined();
    expect(substituteFamily("")).toBeUndefined();
  });
});

describe("CDN font CSS", () => {
  test("requests regular, bold, and italics in one stylesheet per family", () => {
    const url = fontCssUrl(["Noto Naskh Arabic", "Carlito"]);
    expect(url.startsWith("https://fonts.googleapis.com/css2?")).toBe(true);
    expect(url).toContain("family=Noto+Naskh+Arabic:ital,wght@0,400;0,700;1,400;1,700");
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

describe("default resolver", () => {
  test("returns undefined for unknown families without touching the network", async () => {
    let fetched = 0;
    const resolver = createResolver({
      fetchFn: (async () => {
        fetched++;
        throw new Error("must not fetch");
      }) as unknown as typeof fetch,
    });
    expect(resolver("Aptos")).toBeUndefined();
    expect(fetched).toBe(0);
  });

  test("registers every subset face once per family and resolves the primary", async () => {
    const added: Array<{ family: string; source: string; descriptors: Record<string, string> | undefined }> = [];
    const loads: string[] = [];
    class FakeFontFace {
      family: string;
      constructor(family: string, source: string, descriptors?: Record<string, string>) {
        this.family = family;
        added.push({ family, source, descriptors });
      }
      async load(): Promise<this> {
        loads.push(this.family);
        return this;
      }
    }
    const fonts = { add(_face: unknown): void {} };
    let fetches = 0;
    const resolver = createResolver({
      fetchFn: (async () => {
        fetches++;
        return new Response(SAMPLE_CSS, { status: 200 });
      }) as unknown as typeof fetch,
      FontFaceCtor: FakeFontFace as unknown as typeof FontFace,
      documentFonts: fonts as unknown as FontFaceSet,
    });
    const first = await resolver("Arabic Typesetting");
    const second = await resolver("arabic typesetting");
    expect(first).toBeInstanceOf(FakeFontFace);
    expect(second).toBe(first);
    expect(fetches).toBe(1);
    expect(added).toHaveLength(2);
    expect(added[0]?.source).toContain("https://fonts.gstatic.com/");
    expect(added[0]?.descriptors).toMatchObject({ style: "normal", weight: "400" });
    expect(added[0]?.descriptors?.unicodeRange).toContain("U+0600-06FF");
    expect(loads).toEqual(["Amiri", "Amiri"]);
  });

  test("never rejects: delivery failure resolves undefined for browser fallback", async () => {
    const resolver = createResolver({
      fetchFn: (async () => new Response("oops", { status: 500 })) as unknown as typeof fetch,
      FontFaceCtor: class {
        async load(): Promise<this> {
          return this;
        }
      } as unknown as typeof FontFace,
      documentFonts: { add(_face: unknown): void {} } as unknown as FontFaceSet,
    });
    await expect(resolver("Calibri")!).resolves.toBeUndefined();
    const offline = createResolver({
      fetchFn: (async () => {
        throw new TypeError("fetch failed");
      }) as unknown as typeof fetch,
      FontFaceCtor: class {
        async load(): Promise<this> {
          return this;
        }
      } as unknown as typeof FontFace,
      documentFonts: { add(_face: unknown): void {} } as unknown as FontFaceSet,
    });
    await expect(offline("Calibri")!).resolves.toBeUndefined();
  });

  test("default export resolves undefined headlessly instead of throwing", () => {
    // Bun has fetch but no FontFace/document: the shared default must stay
    // inert outside a browser, since unit tests run headless by design.
    expect(defaultResolver("Calibri")).toBeUndefined();
    expect(defaultResolver("Aptos")).toBeUndefined();
  });
});
