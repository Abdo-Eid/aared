import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import { dropsTopSpacing, estimateBlockHeightPx, paginate } from "../src/core/docx/pagination.js";
import { lineFontSizePt } from "../src/core/typography.js";
import type { PageStartReason } from "../src/core/docx/pagination.js";
import type { DocumentSection, FlowNode, TextRunModel } from "../src/core/model.js";

const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

function joined(node: { runs: TextRunModel[] }): string {
  return node.runs.map((run) => run.text).join("");
}

function pageText(page: { body: FlowNode[] }): string {
  const parts: string[] = [];
  const visit = (nodes: FlowNode[]): void => {
    for (const node of nodes) {
      if (node.type === "paragraph" || node.type === "heading") parts.push(joined(node));
      else if (node.type === "list") for (const item of node.items) parts.push(joined(item.paragraph));
      else if (node.type === "table") {
        for (const row of node.rows) {
          for (const cell of row.cells) visit(cell.nodes);
        }
      }
    }
  };
  visit(page.body);
  return parts.join("\n");
}

async function productionBrief(): Promise<DocumentSection[]> {
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  return parseDocxSections(pkg, await indexDocx(pkg));
}

describe("DOCX layout pagination", () => {
  test("merged coverage fixture retains section boundaries and physical page geometry", async () => {
    const layout = paginate(await productionBrief());
    expect(layout.warnings).toEqual([]);
    // The seven wrapping examples share portrait page 5; the empty section
    // markers do not create pages of their own.
    expect(layout.pages).toHaveLength(7);
    expect(layout.pages.map((page) => page.orientation)).toEqual(["portrait", "portrait", "portrait", "landscape", "portrait", "landscape", "portrait"]);
    expect(layout.pages.map((page) => page.startReason)).toEqual([
      "document",
      "explicit-page-break",
      "explicit-page-break",
      "section-next-page",
      "section-next-page",
      "section-next-page",
      "section-next-page",
    ]);
    expect(layout.pages.map((page) => page.index)).toEqual([1, 2, 3, 4, 5, 6, 7]);
    expect(layout.pages.map((page) => page.sectionIndex)).toEqual([0, 0, 0, 1, 2, 3, 4]);
    // Real section geometry drives every page, not a fixed body padding.
    expect(layout.pages[0]?.widthPx).toBeCloseTo(793.73, 0);
    expect(layout.pages[0]?.heightPx).toBeCloseTo(1122.53, 0);
    expect(layout.pages[3]?.widthPx).toBeCloseTo(1122.53, 0);
    expect(layout.pages[3]?.heightPx).toBeCloseTo(793.73, 0);
    const margins = layout.pages[0]?.marginsPx;
    expect(margins?.top).toBeCloseTo(48, 1);
    expect(margins?.right).toBeCloseTo(48, 1);
    expect(margins?.bottom).toBeCloseTo(48, 1);
    expect(margins?.left).toBeCloseTo(75.6, 1);
  });

  test("the merged final section page takes its own first-page header", async () => {
    const layout = paginate(await productionBrief());
    const headerText = (page: { header?: FlowNode[] }): string => (page.header ?? []).map((node) => (node.type === "paragraph" || node.type === "heading" ? joined(node) : "")).join("\n");
    const finalPage = layout.pages[6]!;
    expect(headerText(finalPage)).toBe("رأس الصفحة الأولى · First-page header");
    // Printed numbers run on from the earlier sections: no restart is authored.
    expect(finalPage.number).toBe(7);
  });

  test("merged coverage fixture content remains in its authored sections", async () => {
    const layout = paginate(await productionBrief());
    const [first, second, third, fourth] = layout.pages.map(pageText);
    expect(first).toContain("Contents");
    expect(second).toContain("2. Technical Specifications");
    // The technical table stays on page 2 with its footnote reference...
    expect(second).toContain("Parameter");
    // ...while the overflow page carries the schedule and contacts...
    expect(third).toContain("3. Schedule");
    expect(third).toContain("4. Contacts");
    expect(third).not.toContain("Appendix A");
    // ...and the landscape section holds the appendix.
    expect(fourth).toContain("Appendix A — Shot List, Episode 04");
    const wrappingPage = pageText(layout.pages[4]!);
    for (const label of [
      "التفاف مربع عند الطرف الأيسر",
      "التفاف مربع عند الطرف الأيمن",
      "أعلى وأسفل الصورة",
      "صورة خلف النص",
      "التفاف محكم حول الصورة",
      "التفاف مربع في الوسط",
      "التفاف عبر الصورة",
    ]) expect(wrappingPage).toContain(label);
    // Footnotes follow their references: f1 in the page-2 spec table, f2 in
    // the page-3 schedule.
    expect(layout.pages[1]?.footnotes.map((note) => note.id)).toEqual(["f1"]);
    expect(layout.pages[2]?.footnotes.map((note) => note.id)).toEqual(["f2"]);
    // Page 6 holds both charts: the inline one as its own block, the anchored
    // one on its paragraph.
    const chartPage = layout.pages[5]!;
    expect(chartPage.body.filter((node) => node.type === "chart")).toHaveLength(1);
    expect(chartPage.body.some((node) => node.type === "paragraph" && node.floats?.some((float) => float.type === "chart"))).toBe(true);
  });

  test("explicit breaks split unconditionally, soft hints follow policy", async () => {
    const para = (text: string): FlowNode => ({ type: "paragraph", runs: [{ text }], direction: "ltr" });
    const sections: DocumentSection[] = [{
      properties: {
        breakType: "nextPage",
        page: {
          widthTwips: 12240, heightTwips: 15840, orientation: "portrait",
          margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 },
        },
      },
      root: [
        para("one"),
        { type: "pageBreak", strength: "mandatory" },
        para("two"),
        { type: "pageBreak", strength: "preferred" },
        para("three"),
      ],
    }];
    // Default ("prefer") splits on a stale hint once the page is half full —
    // a single short paragraph is not enough to trigger it.
    expect(paginate(sections).pages.map((page) => page.startReason)).toEqual(["document", "explicit-page-break"]);
    expect(paginate(sections, { softBreaks: "force" }).pages.map((page) => page.startReason)).toEqual([
      "document", "explicit-page-break", "explicit-page-break",
    ]);
    expect(paginate(sections, { softBreaks: "ignore" }).pages.map((page) => page.startReason)).toEqual([
      "document", "explicit-page-break",
    ]);
  });

  test("section transitions open fresh pages with their own geometry", async () => {
    const para = (text: string): FlowNode => ({ type: "paragraph", runs: [{ text }], direction: "ltr" });
    const geometry = (widthTwips: number, heightTwips: number, orientation: "portrait" | "landscape") => ({
      widthTwips, heightTwips, orientation,
      margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 },
    });
    const sections: DocumentSection[] = [
      { properties: { breakType: "nextPage", page: geometry(12240, 15840, "portrait") }, root: [para("front")] },
      { properties: { breakType: "continuous", page: geometry(12240, 15840, "portrait") }, root: [para("same page")] },
      { properties: { breakType: "nextPage", page: geometry(16838, 11906, "landscape") }, root: [para("appendix")] },
    ];
    const layout = paginate(sections);
    expect(layout.pages).toHaveLength(2);
    expect(layout.pages[0]?.body).toHaveLength(2);
    expect(layout.pages[0]?.startReason).toBe("document");
    expect(layout.pages[1]?.startReason).toBe("section-next-page");
    expect(layout.pages[1]?.orientation).toBe("landscape");
  });

  test("even-page sections insert a blank filler when parity is wrong", async () => {
    const para = (text: string): FlowNode => ({ type: "paragraph", runs: [{ text }], direction: "ltr" });
    const geometry = {
      widthTwips: 12240, heightTwips: 15840, orientation: "portrait" as const,
      margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 },
    };
    const sections: DocumentSection[] = [
      { properties: { breakType: "nextPage", page: geometry }, root: [para("one")] },
      { properties: { breakType: "nextPage", page: geometry }, root: [para("two")] },
      { properties: { breakType: "evenPage", page: geometry }, root: [para("three")] },
    ];
    const layout = paginate(sections);
    // Pages 1-2 are content; the even section would open on odd page 3, so a
    // blank filler holds the parity and content lands on page 4.
    expect(layout.pages).toHaveLength(4);
    expect(layout.pages[2]?.body).toHaveLength(0);
    expect(layout.pages[2]?.startReason).toBe("section-even-page");
    expect(pageText(layout.pages[3]!)).toContain("three");
  });

  test("blocks taller than an empty page are kept with a warning", async () => {
    const sections: DocumentSection[] = [{
      properties: {
        breakType: "nextPage",
        page: {
          widthTwips: 12240, heightTwips: 15840, orientation: "portrait",
          margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 },
        },
      },
      root: [{ type: "image", mediaId: "word/media/huge.png", contentType: "image/png", width: 7, height: 20 }],
    }];
    const layout = paginate(sections);
    expect(layout.pages).toHaveLength(1);
    expect(layout.warnings.map((warning) => warning.code)).toEqual(["oversize-block"]);
  });

  // Word breaks a long table between rows. aared used to move a table taller
  // than a page to a fresh page whole, which left the page before it blank
  // (a `pageBreakBefore` paragraph alone) and clipped the table's tail.
  test("a table taller than a page breaks between rows, never inside a vertical merge", () => {
    const row = (text: string, rowSpan?: number) => ({
      height: 1,
      cells: [{ nodes: [{ type: "paragraph" as const, direction: "ltr" as const, runs: [{ text }] }], ...(rowSpan === undefined ? {} : { rowSpan }) }],
    });
    // 1in rows on a 9in body: rows 8-9 are one vertical merge, so the first
    // break falls after row 7, not inside the merge.
    const rows = Array.from({ length: 20 }, (_, index) => row(`row ${index + 1}`, index === 7 ? 2 : undefined));
    const sections: DocumentSection[] = [{
      properties: {
        breakType: "nextPage",
        page: {
          widthTwips: 12240, heightTwips: 15840, orientation: "portrait",
          margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 },
        },
      },
      root: [
        { type: "paragraph", direction: "ltr", runs: [{ text: "lead" }] },
        { type: "table", direction: "ltr", columns: [6.5], rows },
      ],
    }];
    const layout = paginate(sections);
    expect(layout.warnings).toEqual([]);
    const rowsPerPage = layout.pages.map((page) => page.body.filter((node) => node.type === "table").map((node) => node.type === "table" ? node.rows.length : 0));
    // Every page carries rows (no blank page) and every row is placed once.
    expect(rowsPerPage.every((counts) => counts.length === 1 && counts[0]! > 0)).toBe(true);
    expect(rowsPerPage.flat().reduce((sum, count) => sum + count, 0)).toBe(20);
    expect(pageText(layout.pages[0]!)).toContain("lead");
    // The first page holds rows 1-7: row 8 starts a merge that row 9 closes.
    expect(pageText(layout.pages[0]!)).toContain("row 7");
    expect(pageText(layout.pages[0]!)).not.toContain("row 8");
    expect(pageText(layout.pages[1]!).startsWith("row 8")).toBe(true);
  });

  // Word's own saved layout can mark exactly where it broke a long table
  // (DECISION_LOG §45): that boundary wins over the later one aared's own
  // measurement would otherwise prefer, whenever both agree it is valid
  // (clean of a vertical merge, no `w:keepNext` chain) and it still fits.
  test("a row Word's own layout marked wins the table split over a later boundary that also fits", () => {
    const row = (text: string, pageBreakBefore?: true) => ({
      height: 1,
      cells: [{ nodes: [{ type: "paragraph" as const, direction: "ltr" as const, runs: [{ text }] }] }],
      ...(pageBreakBefore === undefined ? {} : { pageBreakBefore }),
    });
    // Same 20-row, 1in-row table as above (breaks after row 7 unmarked), but
    // row 5 (index 4) carries Word's own marker.
    const rows = Array.from({ length: 20 }, (_, index) => row(`row ${index + 1}`, index === 4 ? true : undefined));
    const sections: DocumentSection[] = [{
      properties: {
        breakType: "nextPage",
        page: {
          widthTwips: 12240, heightTwips: 15840, orientation: "portrait",
          margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 },
        },
      },
      root: [
        { type: "paragraph", direction: "ltr", runs: [{ text: "lead" }] },
        { type: "table", direction: "ltr", columns: [6.5], rows },
      ],
    }];
    const layout = paginate(sections);
    expect(pageText(layout.pages[0]!)).toContain("row 4");
    expect(pageText(layout.pages[0]!)).not.toContain("row 5");
    expect(pageText(layout.pages[1]!).startsWith("row 5")).toBe(true);
  });

  // A table whose whole first row is marked (not a split boundary Word chose
  // mid-table, but the table itself starting on a fresh page) is honoured
  // the same way a top-level `w:lastRenderedPageBreak` hint is: only once
  // the current page is already more than half full.
  test("a table marked to start a fresh page opens one once the current page is over half full", () => {
    const table = (marked: boolean): FlowNode => ({
      type: "table",
      direction: "ltr",
      columns: [6.5],
      rows: [{
        height: 1,
        cells: [{ nodes: [{ type: "paragraph", direction: "ltr", runs: [{ text: "first row" }] }] }],
        ...(marked ? { pageBreakBefore: true as const } : {}),
      }],
    });
    const page = (
      widthTwips: number, heightTwips: number, marked: boolean,
    ): DocumentSection => ({
      properties: {
        breakType: "nextPage",
        page: {
          widthTwips, heightTwips, orientation: "portrait",
          margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 },
        },
      },
      // Enough short paragraphs to push well past half the 9in body before the table.
      root: [
        ...Array.from({ length: 40 }, (_, index): FlowNode => ({ type: "paragraph", direction: "ltr", runs: [{ text: `filler ${index}` }] })),
        table(marked),
      ],
    });
    const unmarked = paginate([page(12240, 15840, false)]);
    expect(unmarked.pages).toHaveLength(1);
    const marked = paginate([page(12240, 15840, true)]);
    expect(marked.pages).toHaveLength(2);
    expect(pageText(marked.pages[1]!)).toContain("first row");
  });

  test("estimator folds margins and wraps cells in column widths", async () => {
    // A bare paragraph reports its line box plus spacing, standalone.
    const para: FlowNode = {
      type: "paragraph",
      runs: [{ text: "Hello world, this is a short line." }],
      direction: "ltr",
      fontSizePt: 12,
      spaceBeforePt: 6,
      spaceAfterPt: 6,
    };
    expect(estimateBlockHeightPx(para, 639)).toBeCloseTo(8 + 18.4 + 8, 0);
    // Empty documents paginate to nothing, never a blank page.
    expect(paginate([])).toEqual({ pages: [], warnings: [] });
  });

  test("a header or footer taller than its margin pushes the body inward, never clips it", async () => {
    const para = (sizePt: number): FlowNode => ({ type: "paragraph", runs: [{ text: "Header", fontSizePt: sizePt }], direction: "ltr" });
    const section = (header?: FlowNode[], footer?: FlowNode[]): DocumentSection => ({
      properties: {
        breakType: "nextPage",
        page: {
          widthTwips: 12240, heightTwips: 15840, orientation: "portrait",
          margins: { top: 720, right: 1440, bottom: 720, left: 1440, header: 360, footer: 360, gutter: 0 },
        },
      },
      root: [{ type: "paragraph", runs: [{ text: "body" }], direction: "ltr" }],
      ...(header === undefined ? {} : { header }),
      ...(footer === undefined ? {} : { footer }),
    });
    const tallHeader = para(30);
    const tallFooter = para(20);
    const [page] = paginate([section([tallHeader], [tallFooter])]).pages;
    const columnPx = 6.5 * 96;
    // Word grows a header down from its distance (0.25in) and starts the body below it.
    expect(page?.bodyInsetsPx.top).toBeCloseTo(24 + estimateBlockHeightPx(tallHeader, columnPx), 1);
    expect(page?.bodyInsetsPx.bottom).toBeCloseTo(24 + estimateBlockHeightPx(tallFooter, columnPx), 1);
    // The authored margins are kept as authored; only the derived insets move.
    expect(page?.marginsPx.top).toBeCloseTo(48, 1);
    expect(page?.bodyInsetsPx.top).toBeGreaterThan(48);
    // A short header stays inside its margin, and no header leaves the margin alone.
    const [short] = paginate([section([para(8)]), section()]).pages;
    expect(short?.bodyInsetsPx.top).toBeCloseTo(48, 1);
    expect(paginate([section()]).pages[0]?.bodyInsetsPx).toEqual({ top: 48, bottom: 48 });
  });

  test("space-before survives at a section's first page and is dropped where the flow merely continues", () => {
    const starts: PageStartReason[] = ["document", "section-next-page", "section-even-page", "section-odd-page", "explicit-page-break", "overflow"];
    expect(starts.map(dropsTopSpacing)).toEqual([false, false, false, false, true, true]);
  });

  test("a paragraph's line box follows its runs, not its default size", () => {
    const run = (fontSizePt?: number): TextRunModel => ({ text: "x", ...(fontSizePt === undefined ? {} : { fontSizePt }) });
    // 9pt runs under a 10.5pt style are a 9pt line (Word's table rows of 9pt text measure 38.7px, not 42.9px).
    expect(lineFontSizePt([run(9), run(9)], 10.5)).toBe(9);
    // A run that states no size shows the paragraph default through, and the tallest run wins.
    expect(lineFontSizePt([run(9), run()], 10.5)).toBe(10.5);
    expect(lineFontSizePt([run(9), run(12)], 10.5)).toBe(12);
    // An unsized tab between 8pt runs (a footer's "text, tab, page number") sets no size of its own.
    expect(lineFontSizePt([run(8), { text: "\t" }, run(8)], 10.5)).toBe(8);
    // An empty paragraph, or one with no sizes at all, falls back to the default (which may be absent).
    expect(lineFontSizePt([{ text: "" }], 10.5)).toBe(10.5);
    expect(lineFontSizePt([run()], undefined)).toBeUndefined();
  });
});
