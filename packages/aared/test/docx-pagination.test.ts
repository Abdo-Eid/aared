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
  test("bidi-coverage fixture paginates to four structural pages", async () => {
    const layout = paginate(await productionBrief());
    expect(layout.warnings).toEqual([]);
    expect(layout.pages).toHaveLength(4);
    expect(layout.pages.map((page) => page.orientation)).toEqual([
      "portrait",
      "portrait",
      "portrait",
      "landscape",
    ]);
    expect(layout.pages.map((page) => page.startReason)).toEqual([
      "document",
      "explicit-page-break",
      "overflow",
      "section-next-page",
    ]);
    expect(layout.pages.map((page) => page.index)).toEqual([1, 2, 3, 4]);
    expect(layout.pages.map((page) => page.sectionIndex)).toEqual([0, 0, 0, 1]);
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

  test("bidi-coverage fixture content maps to the Word reference structure", async () => {
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
    // Footnotes follow their references: f1 in the page-2 spec table, f2 in
    // the page-3 schedule.
    expect(layout.pages[1]?.footnotes.map((note) => note.id)).toEqual(["f1"]);
    expect(layout.pages[2]?.footnotes.map((note) => note.id)).toEqual(["f2"]);
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
        headerRefs: [], footerRefs: [],
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
      { properties: { breakType: "nextPage", page: geometry(12240, 15840, "portrait"), headerRefs: [], footerRefs: [] }, root: [para("front")] },
      { properties: { breakType: "continuous", page: geometry(12240, 15840, "portrait"), headerRefs: [], footerRefs: [] }, root: [para("same page")] },
      { properties: { breakType: "nextPage", page: geometry(16838, 11906, "landscape"), headerRefs: [], footerRefs: [] }, root: [para("appendix")] },
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
      { properties: { breakType: "nextPage", page: geometry, headerRefs: [], footerRefs: [] }, root: [para("one")] },
      { properties: { breakType: "nextPage", page: geometry, headerRefs: [], footerRefs: [] }, root: [para("two")] },
      { properties: { breakType: "evenPage", page: geometry, headerRefs: [], footerRefs: [] }, root: [para("three")] },
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
        headerRefs: [], footerRefs: [],
      },
      root: [{ type: "image", mediaId: "word/media/huge.png", contentType: "image/png", width: 7, height: 20 }],
    }];
    const layout = paginate(sections);
    expect(layout.pages).toHaveLength(1);
    expect(layout.warnings.map((warning) => warning.code)).toEqual(["oversize-block"]);
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
        headerRefs: [], footerRefs: [],
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
