import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { automaticTextColor, indexPptx, parsePptxSlide, parseTextBody } from "../src/core/pptx.js";
import { parseXml } from "../src/core/xml.js";
import type { ShapeModel, TableModel } from "../src/core/model.js";

const COVERAGE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

function shapeWithText(slide: { elements: ShapeModel[] }, text: string): ShapeModel {
  const found = slide.elements.find(
    (element) => element.kind === "shape" && element.text !== undefined
      && element.text.paragraphs.some((p) => p.runs.some((r) => r.text.includes(text))),
  );
  if (found === undefined || found.kind !== "shape") throw new Error(`expected a shape containing "${text}"`);
  return found;
}

describe("PPTX automatic text color", () => {
  test("dark solid backgrounds resolve white", () => {
    expect(automaticTextColor({ kind: "solid", color: "#3A4961" })).toBe("#FFFFFF");
  });

  test("light, missing, and non-solid backgrounds resolve black", () => {
    expect(automaticTextColor({ kind: "solid", color: "#FFFFFF" })).toBe("#000000");
    expect(automaticTextColor(undefined)).toBe("#000000");
    expect(automaticTextColor({ kind: "unsupported", source: "pattern" })).toBe("#000000");
  });

  test("a translucent background is read as black, so the ink stays black on it", () => {
    // `a:alpha` really does produce this shape (core-shared.test.ts reads one
    // off `resolveDrawingColor`). A luminance probe cannot read `rgb(...)`, so
    // the fill must fall through to black: if a "smarter" probe treated a
    // half-transparent black as dark, it would paint white text on a dark
    // slide, which is exactly the silent-fallback class §28 fixed.
    expect(automaticTextColor({ kind: "solid", color: "rgb(0 0 0 / 0.5)" })).toBe("#000000");
    // A bare hex without its hash is the same colour, not an unparseable one.
    expect(automaticTextColor({ kind: "solid", color: "3A4961" })).toBe("#FFFFFF");
  });

  test("fill-less titles resolve Automatic contrast, never chrome inheritance", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const deck = await indexPptx(pkg);
    const slide3 = await parsePptxSlide(pkg, deck, 2);
    // The slides are white, so Automatic is black, as in native PowerPoint.
    // The dark path is pinned by the `automaticTextColor` unit test above.
    expect(slide3.background).toMatchObject({ kind: "solid", color: "#FFFFFF" });
    // The title run carries only a text outline (`a:ln`) — no fill. The
    // model must carry an explicit resolved color so the renderer never
    // inherits bench chrome on theme toggles.
    const title = shapeWithText(slide3 as unknown as { elements: ShapeModel[] }, "أشكال ورسومات");
    const colors = title.text!.paragraphs.flatMap((p) => p.runs.map((r) => r.color));
    expect(colors.length).toBeGreaterThan(0);
    for (const color of colors) expect(color).toBe("#000000");
    // The outline itself still resolves (black edge from the WordArt style).
    expect(title.text!.paragraphs[0]!.runs[0]!.outline).toMatchObject({ color: "#000000" });
    const slide4 = await parsePptxSlide(pkg, deck, 3);
    const gradientTitle = shapeWithText(slide4 as unknown as { elements: ShapeModel[] }, "تدرجات لونية");
    for (const color of gradientTitle.text!.paragraphs.flatMap((p) => p.runs.map((r) => r.color))) {
      expect(color).toBe("#000000");
    }
  });

  test("authored table style text wins over automatic contrast", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const deck = await indexPptx(pkg);
    const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
    const table = slides.flatMap((slide) => slide.elements).find((element): element is TableModel => element.kind === "table");
    if (table === undefined) throw new Error("fixture structure changed: expected a gate table");
    // Header cells carry direct authored red fills, which win over the
    // firstRow style white — direct formatting keeps precedence.
    const headerColors = table.rows[0]!.cells.flatMap((cell) =>
      cell.text.paragraphs.flatMap((p) => p.runs.filter((r) => r.text !== "").map((r) => r.color)),
    );
    expect(headerColors.length).toBeGreaterThan(0);
    for (const color of headerColors) expect(color).toBe("#FF0000");
    // Body cells have no direct fill: the wholeTbl style dk1 (black) applies,
    // not bench-chrome inheritance and not automatic contrast.
    const bodyStyleCheck = table.rows[1]!.cells[0]!.text.paragraphs[0]!.runs[0]!.color;
    expect(bodyStyleCheck).toBe("#000000");
    const bodyColors = table.rows.slice(1).flatMap((row) =>
      row.cells.flatMap((cell) => cell.text.paragraphs.flatMap((p) => p.runs.filter((r) => r.text !== "").map((r) => r.color))),
    );
    expect(bodyColors.length).toBeGreaterThan(0);
    for (const color of bodyColors) expect(color).toBe("#000000");
  });
});

describe("PPTX bullets (coverage slide 5)", () => {
  test("explicit bullet colors and sizes resolve", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const deck = await indexPptx(pkg);
    const slide = await parsePptxSlide(pkg, deck, 4);
    const bullets = slide.elements.flatMap((element) =>
      element.kind === "shape" && element.text !== undefined ? element.text.paragraphs : [],
    ).filter((paragraph) => paragraph.bullet !== undefined);
    expect(bullets.length).toBeGreaterThan(5);
    // Follow-text default: no explicit color or size.
    const plain = bullets.find((p) => p.runs.some((r) => r.text.includes("افتراضية")));
    expect(plain?.bullet).toMatchObject({ char: "•" });
    expect(plain?.bullet?.color).toBeUndefined();
    expect(plain?.bullet?.sizePct).toBeUndefined();
    // Customized: red 150% markers. The Wingdings square ("q") and check mark
    // ("ü") arrive as their Unicode shapes with no symbol font, so a device
    // without Wingdings paints the shape, not the letter.
    const red = bullets.find((p) => p.runs.some((r) => r.text.includes("حمراء")));
    expect(red?.bullet).toMatchObject({ color: "#FF0000", sizePct: 150 });
    expect(bullets.some((p) => p.bullet?.typeface === "Wingdings")).toBe(false);
    const shapes = bullets.map((p) => p.bullet?.char).filter((char) => char === "❑" || char === "✓");
    expect(shapes.sort()).toEqual(["✓", "✓", "❑"]);
  });

  test("numbering restarts per type and picture bullets resolve media", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const deck = await indexPptx(pkg);
    const slide = await parsePptxSlide(pkg, deck, 4);
    const bullets = slide.elements.flatMap((element) =>
      element.kind === "shape" && element.text !== undefined ? element.text.paragraphs : [],
    ).filter((paragraph) => paragraph.bullet !== undefined);
    const chars = bullets.map((p) => p.bullet?.char);
    // A new autonum type starts a new list (alpha restarts at a), not d)/e).
    expect(chars).toContain("1.");
    expect(chars).toContain("2.");
    expect(chars).toContain("a)");
    expect(chars).toContain("b)");
    expect(chars).toContain("i.");
    expect(chars).toContain("ii.");
    const pictured = bullets.find((p) => p.bullet?.pictureMediaId !== undefined);
    expect(pictured?.bullet?.pictureMediaId).toBe("ppt/media/image10.svg");
    expect(pictured?.bullet?.pictureContentType).toBe("image/svg+xml");
  });
});

describe("PPTX WordArt (coverage slide 6)", () => {
  function wordartShape(slide: Awaited<ReturnType<typeof parsePptxSlide>>, text: string) {
    const found = slide.elements.find(
      (element): element is ShapeModel =>
        element.kind === "shape" && element.text !== undefined
        && element.text.paragraphs.some((p) => p.runs.some((r) => r.text.includes(text))),
    );
    if (found === undefined) throw new Error(`expected a WordArt shape containing "${text}"`);
    return found;
  }

  test("gradient text resolves centered radial fills, glow resolves, warp is marked", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const deck = await indexPptx(pkg);
    const slide = await parsePptxSlide(pkg, deck, 5);
    // Centered `path="circle"` gradient text (PowerPoint's built-in preset)
    // resolves to a real radial gradient, not an unsupported fallback.
    const graded = wordartShape(slide, "متدرج");
    expect(graded.text!.paragraphs[0]!.runs[0]!.gradient).toMatchObject({
      kind: "path-gradient",
      path: "circle",
      focus: { left: 0.5, top: 0.5, right: 0.5, bottom: 0.5 },
      stops: [
        { position: 0, color: "#9E9E9E" },
        { position: 0.5, color: "#C4C4C4" },
        { position: 1, color: "#E3E3E3" },
      ],
    });
    const glowing = wordartShape(slide, "متوهج");
    expect(glowing.text!.paragraphs[0]!.runs[0]!.glow).toMatchObject({ radiusPt: 18 });
    expect(glowing.text!.paragraphs[0]!.runs[0]!.glow?.color).toContain("255 106 28");
    const warped = wordartShape(slide, "مقوس");
    expect(warped.text!.textWarp).toBe("textArchUp");
  });

  test("outline-only text and whole-shape hyperlinks resolve", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const deck = await indexPptx(pkg);
    const slide = await parsePptxSlide(pkg, deck, 5);
    const hollow = wordartShape(slide, "تعبئة");
    expect(hollow.text!.paragraphs[0]!.runs[0]).toMatchObject({
      noFill: true,
      outline: { color: "#000000" },
    });
    // PowerPoint wrote this link on the shape (`p:cNvPr/a:hlinkClick`), not
    // the runs: the shape carries it, run-level links stay absent.
    const link = wordartShape(slide, "تجريبي");
    expect(link.hyperlink).toBe("https://example.com/");
    for (const run of link.text!.paragraphs[0]!.runs) {
      expect(run.hyperlink).toBeUndefined();
    }
  });
});

describe("PPTX run font styles", () => {
  async function themeOf() {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const deck = await indexPptx(pkg);
    return deck.theme;
  }

  test("underline, strike, caps, spacing, baseline, highlight", async () => {
    const theme = await themeOf();
    const body = parseTextBody(
      parseXml(
        `<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:r><a:rPr u="sng" strike="sngStrike" cap="all" spc="200" baseline="30000">` +
          `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>` +
          `<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>` +
          `</a:rPr><a:t>Styled</a:t></a:r></a:p>` +
          `</a:txBody>`,
      ),
      theme,
    );
    const run = body.paragraphs[0]!.runs[0]!;
    expect(run.underline).toBe(true);
    expect(run.underlineStyle).toBe("sng");
    expect(run.strikethrough).toBe(true);
    expect(run.caps).toBe("all");
    expect(run.letterSpacingPt).toBeCloseTo(2, 6);
    expect(run.baselinePct).toBe(30);
    expect(run.highlight).toBe("#FFFF00");
    expect(run.color).toBe("#FF0000");
  });

  test("u=none means no underline; small caps maps", async () => {
    const theme = await themeOf();
    const fixed = parseTextBody(
      parseXml(
        `<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:r><a:rPr u="none" cap="small"/><a:t>Plain</a:t></a:r></a:p>` +
          `</a:txBody>`,
      ),
      theme,
    );
    const run = fixed.paragraphs[0]!.runs[0]!;
    expect(run.underline).toBeUndefined();
    expect(run.underlineStyle).toBeUndefined();
    expect(run.caps).toBe("small");
  });

  test("noFill and grpFill mark outline-only and group fills", async () => {
    const theme = await themeOf();
    const body = parseTextBody(
      parseXml(
        `<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:r><a:rPr><a:noFill/><a:ln><a:noFill/></a:ln></a:rPr><a:t>Ghost</a:t></a:r>` +
          `<a:r><a:rPr><a:grpFill/></a:rPr><a:t>Grouped</a:t></a:r></a:p>` +
          `</a:txBody>`,
      ),
      theme,
      { autoColor: "#000000" },
    );
    expect(body.paragraphs[0]!.runs[0]!.noFill).toBe(true);
    expect(body.paragraphs[0]!.runs[1]!.groupFill).toBe(true);
  });

  test("hlinkClick resolves externals through rels", async () => {
    const theme = await themeOf();
    const body = parseTextBody(
      parseXml(
        `<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:r><a:rPr><a:hlinkClick r:id="rId9"/></a:rPr><a:t>Link</a:t></a:r></a:p>` +
          `</a:txBody>`,
      ),
      theme,
      { links: new Map([["rId9", "https://example.com/x"]]) },
    );
    expect(body.paragraphs[0]!.runs[0]!.hyperlink).toBe("https://example.com/x");
  });
});
