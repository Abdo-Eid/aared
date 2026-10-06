import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, overrideColorMap, parsePptxSlide, parseTable, parseTextBody } from "../src/core/pptx.js";
import { themeForColorMap } from "../src/core/theme.js";
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
        // PowerPoint writes a merge on the cell itself: `a:tc gridSpan`, `a:tc hMerge`, `a:tc rowSpan`, `a:tc vMerge`.
        `<a:tr h="914400"><a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Wide</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></a:txBody><a:tcPr/></a:tc></a:tr>` +
        `<a:tr h="914400"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Left</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>` +
        `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Right</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>` +
        `<a:tr h="914400"><a:tc rowSpan="2"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Tall</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>` +
        `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Side</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>` +
        `<a:tr h="914400"><a:tc vMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></a:txBody><a:tcPr/></a:tc>` +
        `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Below</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>` +
        `</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`,
    );
    const table = parseTable(frame, theme, new Map())!;
    expect(table.rows[0]?.cells).toHaveLength(1);
    expect(table.rows[0]?.cells[0]?.columnSpan).toBe(2);
    expect(table.rows[1]?.cells).toHaveLength(2);
    expect(table.rows[2]?.cells[0]?.rowSpan).toBe(2);
    expect(table.rows[3]?.cells).toHaveLength(1);
    expect(table.rows[3]?.cells[0]?.text.paragraphs[0]?.runs[0]?.text).toBe("Below");
  });

  test("cuts a span that runs past the table to the columns and rows that are left", async () => {
    const pkg = await OpcPackage.open(await Bun.file(PRODUCTION).arrayBuffer());
    const theme = await readDrawingTheme(pkg);
    const cell = (attributes: string, text: string) =>
      `<a:tc ${attributes}><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>${text}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`;
    const frame = parseXml(
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Hostile"/></p:nvGraphicFramePr>` +
        `<p:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="1828800"/></p:xfrm>` +
        `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>` +
        `<a:tblPr/><a:tblGrid><a:gridCol w="4572000"/><a:gridCol w="4572000"/></a:tblGrid>` +
        `<a:tr h="914400">${cell(`gridSpan="99999999" rowSpan="99999999"`, "Huge")}${cell("", "Next")}</a:tr>` +
        `<a:tr h="914400">${cell("", "A")}${cell("", "B")}</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`,
    );
    const table = parseTable(frame, theme, new Map())!;
    expect(table.rows[0]?.cells[0]?.columnSpan).toBe(2);
    expect(table.rows[0]?.cells[0]?.rowSpan).toBe(2);
  });

  // `AddTextbox` writes a bodyPr with no insets; PowerPoint still pads it.
  test("gives a text body no bodyPr states DrawingML's default insets", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const { theme } = await indexPptx(pkg);
    const body = (bodyPr: string) => parseTextBody(parseXml(`<a:txBody>${bodyPr}<a:lstStyle/><a:p><a:r><a:t>x</a:t></a:r></a:p></a:txBody>`), theme);
    expect(body(`<a:bodyPr wrap="none"/>`).insets).toEqual({ left: 0.1, right: 0.1, top: 0.05, bottom: 0.05 });
    expect(body(`<a:bodyPr lIns="0" tIns="0"/>`).insets).toEqual({ left: 0, right: 0.1, top: 0, bottom: 0.05 });
  });

  // Slide 2's title `bodyPr` carries only `a:normAutofit`; the master says
  // `anchor="ctr"` and PowerPoint centres the text in the 1.45in box.
  test("takes the anchor from the master when the slide's own bodyPr states none", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const deck = await indexPptx(pkg);
    const slide = await parsePptxSlide(pkg, deck, 1);
    const title = slide.elements.find((element) => element.kind === "shape" && element.placeholder?.type === "title");
    if (title?.kind !== "shape" || title.text === undefined) throw new Error("fixture structure changed: expected slide 2's title");
    expect(title.text.verticalAlign).toBe("middle");
    // An authored anchor still wins, and wrap cascades the same way.
    const body = (bodyPr: string, inherited: string) =>
      parseTextBody(parseXml(`<a:txBody>${bodyPr}<a:lstStyle/><a:p><a:r><a:t>x</a:t></a:r></a:p></a:txBody>`), deck.theme, { bodies: [parseXml(`<a:txBody>${inherited}</a:txBody>`)] });
    expect(body(`<a:bodyPr anchor="b"><a:normAutofit/></a:bodyPr>`, `<a:bodyPr anchor="ctr" wrap="none"/>`)).toMatchObject({ verticalAlign: "bottom", wrapping: "none" });
  });

  test("reads a slide's or layout's colour-map override, and leaves a master mapping alone", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const { theme } = await indexPptx(pkg);
    const dark = parseXml(
      `<p:sld><p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" ` +
        `accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sld>`,
    );
    const remapped = themeForColorMap(theme, overrideColorMap(dark));
    expect(remapped.colors.bg1).toBe(theme.colors.dk1 ?? "");
    expect(remapped.colors.tx1).toBe(theme.colors.lt1 ?? "");
    expect(remapped.colors.bg2).toBe(theme.colors.dk2 ?? "");
    // `a:masterClrMapping` (and no element at all) keep the map above.
    expect(overrideColorMap(parseXml(`<p:sld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`))).toBeUndefined();
    expect(overrideColorMap(parseXml(`<p:sld/>`))).toBeUndefined();
  });

  test("draws a built-in Medium Style 2 the deck does not define, and leaves other missing styles unstyled", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const { theme } = await indexPptx(pkg);
    const frame = (styleId: string) => parseXml(
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Styled"/></p:nvGraphicFramePr>` +
        `<p:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="1828800"/></p:xfrm>` +
        `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>` +
        `<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="9144000"/></a:tblGrid>` +
        `<a:tr h="914400"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Head</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>` +
        `<a:tr h="914400"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Body</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>` +
        `</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`,
    );
    const styles = new Map();
    // Medium Style 2 - Accent 2: header solid accent2, white edges, white bold text.
    const accent2 = parseTable(frame("{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}"), theme, styles)!;
    expect(accent2.cellFills[0]?.[0]).toEqual({ kind: "solid", color: theme.colors.accent2! });
    expect(accent2.cellFills[1]?.[0]).toMatchObject({ kind: "solid" });
    expect(accent2.rows[0]?.cells[0]?.borders?.left?.color).toBe(theme.colors.lt1!);
    expect(accent2.rows[0]?.cells[0]?.borders?.bottom?.width).toBeCloseTo(38100 / 914400, 6);
    expect(styles.has("{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}")).toBe(true);
    // A GUID PowerPoint would draw from its own copy and aared has no definition for.
    const unknown = parseTable(frame("{00000000-0000-0000-0000-000000000000}"), theme, new Map())!;
    expect(unknown.cellFills[0]?.[0]).toBeUndefined();
  });

  // Ground truth: probe en-us-arabic-run (PowerPoint's own PDF).
  test("keeps Arabic words in written order around ASCII punctuation in a left-to-right en-US run only", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
    const { theme } = await indexPptx(pkg);
    const runs = (text: string, lang: string, paragraphProperties = ""): Array<{ text: string; ltrRun?: boolean }> =>
      parseTextBody(
        parseXml(`<a:txBody><a:bodyPr/><a:lstStyle/><a:p>${paragraphProperties}<a:r><a:rPr lang="${lang}"/><a:t>${text}</a:t></a:r></a:p></a:txBody>`),
        theme,
      ).paragraphs[0]!.runs.map((run) => ({ text: run.text, ...(run.ltrRun === true ? { ltrRun: true } : {}) }));

    expect(runs("واحد + اثنان", "en-US")).toEqual([{ text: "واحد " }, { text: "+", ltrRun: true }, { text: " اثنان" }]);
    expect(runs("واحد (اثنان) ثالثة", "en-US")).toEqual([
      { text: "واحد " }, { text: "(", ltrRun: true }, { text: "اثنان" }, { text: ")", ltrRun: true }, { text: " ثالثة" },
    ]);
    // Spaces alone, an Arabic language, a right-to-left paragraph, and punctuation
    // with no Arabic word beside it all follow plain Unicode bidi.
    expect(runs("واحد اثنان ثالثة", "en-US")).toEqual([{ text: "واحد اثنان ثالثة" }]);
    expect(runs("واحد + اثنان", "ar-SA")).toEqual([{ text: "واحد + اثنان" }]);
    expect(runs("واحد + اثنان", "en-US", `<a:pPr rtl="1"/>`)).toEqual([{ text: "واحد + اثنان" }]);
    expect(runs("abc + واحد", "en-US")).toEqual([{ text: "abc + واحد" }]);
  });

  test("reads the theme's own line widths for lnRef edges", async () => {
    // Two real themes, two answers: the Aptos-era Office theme states 1, 1.5
    // and 2 pt, the 2013-era one 0.5, 1 and 1.5 pt.
    const coverage = await readDrawingTheme(await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer()));
    expect(coverage.lineWidths).toEqual([12700, 19050, 25400]);
    const production = await readDrawingTheme(await OpcPackage.open(await Bun.file(PRODUCTION).arrayBuffer()));
    expect(production.lineWidths).toEqual([6350, 12700, 19050]);
  });
});
