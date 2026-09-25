import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide, parseTable, parseTextBody } from "../src/core/pptx.js";
import { parseXml } from "../src/core/xml.js";
import { readDrawingTheme } from "../src/core/theme.js";
import type { TableModel } from "../src/core/model.js";

const COVERAGE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const PRODUCTION = `${import.meta.dir}/../../../fixtures/post-production-review.pptx`;

describe("PPTX text fidelity", () => {
  test("resolves numbered lists to display strings per level", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const deck = await indexPptx(pkg);
    const body = parseTextBody(
      parseXml(
        `<a:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:pPr lvl="0"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:rPr/><a:t>First</a:t></a:r></a:p>` +
          `<a:p><a:pPr lvl="0"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:rPr/><a:t>Second</a:t></a:r></a:p>` +
          `<a:p><a:pPr lvl="1"><a:buAutoNum type="alphaLcPeriod" startAt="3"/></a:pPr><a:r><a:rPr/><a:t>Nested</a:t></a:r></a:p>` +
          `<a:p><a:pPr lvl="0"><a:buAutoNum type="romanUcPeriod"/></a:pPr><a:r><a:rPr/><a:t>Third</a:t></a:r></a:p>` +
          `</a:txBody>`,
      ),
      deck.theme,
    );
    // Counters restart per numbering type (a new type starts a new list),
    // so the trailing roman list opens at I, not III.
    expect(body.paragraphs.map((paragraph) => paragraph.bullet?.char)).toEqual(["1.", "2.", "c.", "I."]);
    expect(body.paragraphs[0]?.level).toBe(0);
    expect(body.paragraphs[2]?.level).toBe(1);
  });

  test("keeps explicit bullets independent from numbering counters", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const deck = await indexPptx(pkg);
    const body = parseTextBody(
      parseXml(
        `<a:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:rPr/><a:t>Bullet</a:t></a:r></a:p>` +
          `<a:p><a:pPr><a:buNone/></a:pPr><a:r><a:rPr/><a:t>Plain</a:t></a:r></a:p>` +
          `</a:txBody>`,
      ),
      deck.theme,
    );
    expect(body.paragraphs[0]?.bullet).toEqual({ char: "•" });
    expect(body.paragraphs[1]?.bullet).toBeUndefined();
  });

  test("preserves paragraph indents for table and shape text alike", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const deck = await indexPptx(pkg);
    const theme = deck.theme;
    const body = parseTextBody(
      parseXml(
        `<a:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:pPr marL="342900" indent="-342900"><a:buChar char="•"/></a:pPr><a:r><a:rPr/><a:t>Item</a:t></a:r></a:p>` +
          `</a:txBody>`,
      ),
      theme,
    );
    expect(body.paragraphs[0]).toMatchObject({ leftIndent: 0.375, firstLineIndent: -0.375 });
  });
});

describe("PPTX table fidelity", () => {
  test("resolves production cell margins, borders, and direct fills", async () => {
    const pkg = await OpcPackage.open(await Bun.file(PRODUCTION).arrayBuffer());
    const deck = await indexPptx(pkg);
    const slide = await parsePptxSlide(pkg, deck, 4);
    const table = slide.elements.find((element): element is TableModel => element.kind === "table");
    if (table === undefined) throw new Error("fixture structure changed: expected a production table");
    expect(table.columns).toHaveLength(4);
    expect(table.rows).toHaveLength(6);
    expect(table.cellFills[0]).toEqual([1, 2, 3, 4].map(() => ({ kind: "solid" as const, color: "#12171F" })));
    expect(table.rows[0]?.cells[0]?.margins).toEqual({ left: 0.1, right: 0.1, top: 0.05, bottom: 0.05 });
    expect(table.rows[0]?.cells[0]?.borders).toMatchObject({
      left: { color: "#D8D3C9" },
      right: { color: "#D8D3C9" },
      top: { color: "#D8D3C9" },
      bottom: { color: "#D8D3C9" },
    });
    expect(table.rows[0]?.cells[0]?.borders?.left?.width).toBeCloseTo(12700 / 914400, 6);
  });

  test("resolves gate-fixture style borders and text so cells never inherit chrome", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const deck = await indexPptx(pkg);
    const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
    const table = slides.flatMap((slide) => slide.elements).find((element): element is TableModel => element.kind === "table");
    if (table === undefined) throw new Error("fixture structure changed: expected a gate table");
    // A cell stating no insets gets DrawingML's defaults, as PowerPoint paints it.
    expect(table.rows[0]?.cells[0]?.margins).toEqual({ left: 0.1, right: 0.1, top: 0.05, bottom: 0.05 });
    // Medium Style 2 Accent 1: white style edges and header/body text colors.
    expect(table.rows[0]?.cells[0]?.borders).toMatchObject({
      left: { color: "#FFFFFF" },
      top: { color: "#FFFFFF" },
    });
    for (const row of table.rows) {
      for (const cell of row.cells) {
        for (const paragraph of cell.text.paragraphs) {
          for (const run of paragraph.runs) {
            if (run.text !== "") expect(run.color).toBeDefined();
          }
        }
      }
    }
    expect(table.rows[0]?.cells[0]?.columnSpan).toBeUndefined();
  });

  test("resolves grid and row spans while skipping merge continuations", async () => {
    const pkg = await OpcPackage.open(await Bun.file(PRODUCTION).arrayBuffer());
    const theme = await readDrawingTheme(pkg);
    const frame = parseXml(
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Merge"/></p:nvGraphicFramePr>` +
        `<p:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="1828800"/></p:xfrm>` +
        `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>` +
        `<a:tblPr/><a:tblGrid><a:gridCol w="4572000"/><a:gridCol w="4572000"/></a:tblGrid>` +
        `<a:tr h="914400"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Wide</a:t></a:r></a:p></a:txBody><a:tcPr gridSpan="2"/></a:tc></a:tr>` +
        `<a:tr h="914400"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Left</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>` +
        `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Covered</a:t></a:r></a:p></a:txBody><a:tcPr hMerge="1"/></a:tc></a:tr>` +
        `<a:tr h="914400"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Tall</a:t></a:r></a:p></a:txBody><a:tcPr rowSpan="2"/></a:tc>` +
        `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Side</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>` +
        `</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`,
    );
    const table = parseTable(frame, theme, new Map())!;
    expect(table.rows[0]?.cells).toHaveLength(1);
    expect(table.rows[0]?.cells[0]?.columnSpan).toBe(2);
    expect(table.rows[1]?.cells).toHaveLength(1);
    expect(table.rows[2]?.cells[0]?.rowSpan).toBe(2);
  });
});
