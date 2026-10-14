import { describe, expect, test } from "bun:test";
import { storedZip } from "./zip.js";
import { OpcPackage } from "../src/core/opc.js";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import { headingLevelForStyle, paragraphAncestry, parseStyles } from "../src/core/docx/styles.js";
import { isOrderedFormat, markerFor, parseNumbering, startFor } from "../src/core/docx/numbering.js";
import { attr, firstChild, parseXml } from "../src/core/xml.js";
import type { DocxHeadingNode, DocxParagraphNode, DocumentSection, FlowNode, TextRunModel } from "../src/core/model.js";
import { withPart } from "./mutate.js";

const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;
const MANIFEST_PATH = `${import.meta.dir}/../../../fixtures/docx-coverage.json`;

interface CoverageManifest {
  id: string;
  blocks: Record<string, { hasText: string }>;
}

function joined(node: { runs: TextRunModel[] }): string {
  return node.runs.map((run) => run.text).join("");
}

function cellText(cell: { nodes: FlowNode[] }): string {
  return cell.nodes
    .map((node) => (node.type === "paragraph" || node.type === "heading" ? joined(node) : node.type))
    .join("//");
}

function blockText(node: FlowNode): string {
  if (node.type === "paragraph" || node.type === "heading") return joined(node);
  if (node.type === "list") return node.items.map((item) => joined(item.paragraph)).join("\n");
  if (node.type === "table") {
    return node.rows
      .map((row) => row.cells.map((cell) => cellText(cell)).join("|"))
      .join("\n");
  }
  if (node.type === "footnote") return node.nodes.map((child) => (child.type === "paragraph" || child.type === "heading" ? joined(child) : child.type)).join("//");
  if (node.type === "image") return node.mediaId;
  if (node.type === "chart") return `chart:${node.alt ?? ""}`;
  if (node.type === "textBox") return `textBox:${node.content.length}`;
  if (node.type === "pageBreak") return node.strength;
  return node.label;
}

function allRuns(nodes: FlowNode[]): TextRunModel[] {
  const out: TextRunModel[] = [];
  const visit = (list: FlowNode[]): void => {
    for (const node of list) {
      if (node.type === "paragraph" || node.type === "heading") out.push(...node.runs);
      else if (node.type === "list") for (const item of node.items) out.push(...item.paragraph.runs);
      else if (node.type === "table") for (const row of node.rows) for (const cell of row.cells) visit(cell.nodes);
      else if (node.type === "footnote") visit(node.nodes);
    }
  };
  visit(nodes);
  return out;
}
function summaries(sections: DocumentSection[]): string {
  const describe = (nodes: FlowNode[]): string => nodes.map((node) => `${node.type}:${blockText(node).slice(0, 60)}`).join("\n");
  return sections
    .map((section, index) =>
      `section ${index} (${section.properties.page.orientation}, ${section.properties.breakType}):\n${describe(section.root)}\nheader:\n${describe(section.header ?? [])}\nfooter:\n${describe(section.footer ?? [])}`)
    .join("\n");
}

function allRoot(sections: DocumentSection[]): FlowNode[] {
  return sections.flatMap((section) => section.root);
}

async function fixture(): Promise<DocumentSection[]> {
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  return parseDocxSections(pkg, await indexDocx(pkg));
}

async function findBlock(sections: DocumentSection[], id: string): Promise<FlowNode> {
  const manifest = JSON.parse(await Bun.file(MANIFEST_PATH).text()) as CoverageManifest;
  const selector = manifest.blocks[id];
  if (selector === undefined) throw new Error(`fixture ${manifest.id}: unknown block selector "${id}"`);
  const areas: FlowNode[][] = sections.flatMap((section) => [section.root, section.header ?? [], section.footer ?? []]);
  const matches = areas.flat().filter((node) => {
    if (node.type === "list") return node.items.some((item) => joined(item.paragraph).includes(selector.hasText));
    return blockText(node).includes(selector.hasText);
  });
  if (matches.length === 1) return matches[0]!;
  throw new Error(
    `fixture ${manifest.id}: selector "${id}" expected one block containing "${selector.hasText}", found ${matches.length}\n${summaries(sections)}`,
  );
}


const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>` +
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Default Extension="png" ContentType="image/png"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `</Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `</Relationships>`;

/** One transparent pixel: image parts for synthetic drawing tests. */
const PIXEL_PNG = new Uint8Array([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21,
  196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 218, 99, 96, 96, 96, 96, 0, 0, 0, 5, 0, 1, 165, 164, 203, 37, 0,
  0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
]);

function syntheticDocx(documentXml: string, relsXml?: string, extra?: Record<string, Uint8Array | string>): Uint8Array {
  const documentRels = `<?xml version="1.0" encoding="UTF-8"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relsXml ?? ""}</Relationships>`;
  return storedZip({
    "[Content_Types].xml": CONTENT_TYPES,
    "_rels/.rels": PACKAGE_RELS,
    "word/document.xml": documentXml,
    "word/_rels/document.xml.rels": documentRels,
    ...extra,
  });
}

async function parseSynthetic(documentXml: string, relsXml?: string, extra?: Record<string, Uint8Array | string>): Promise<DocumentSection[]> {
  const pkg = await OpcPackage.open(syntheticDocx(documentXml, relsXml, extra).buffer as ArrayBuffer);
  return parseDocxSections(pkg, await indexDocx(pkg));
}

const WORD_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;

describe("DOCX flow parser", () => {
  test("preserves section line-grid pitch and paragraph line controls", async () => {
    const sections = await parseSynthetic(
      `<?xml version="1.0" encoding="UTF-8"?><w:document ${WORD_NS}><w:body>` +
        `<w:p><w:pPr><w:snapToGrid w:val="0"/><w:spacing w:beforeLines="150" w:afterLines="50"/></w:pPr><w:r><w:t>grid sample</w:t></w:r></w:p>` +
        `<w:sectPr><w:docGrid w:type="lines" w:linePitch="480"/></w:sectPr>` +
        `</w:body></w:document>`,
    );
    expect(sections[0]?.properties.docGrid).toMatchObject({ type: "lines", linePitch: 480 });
    const paragraph = sections[0]?.root.find((node) => node.type === "paragraph");
    expect(paragraph).toMatchObject({ spaceBeforeLines: 1.5, spaceAfterLines: 0.5, snapToGrid: false });
  });

  test("resolves the gate fixture to structured-cloneable document sections", async () => {
    const sections = await fixture();
    // Body, landscape appendix, portrait wrapping examples, landscape charts,
    // and the final portrait lists/table section each retain their geometry.
    expect(sections).toHaveLength(6);
    expect(sections[0]?.title).toBe("1. Project Overview");
    expect(sections[1]?.title).toBe("Appendix A — Shot List, Episode 04");
    expect(sections[0]?.properties.page).toMatchObject({ widthTwips: 11906, heightTwips: 16838, orientation: "portrait" });
    expect(sections[0]?.properties.page.margins).toMatchObject({ top: 720, right: 720, bottom: 720, left: 1134 });
    expect(sections[1]?.properties.page).toMatchObject({ widthTwips: 16838, heightTwips: 11906, orientation: "landscape" });
    expect(sections[1]?.properties.page.margins).toMatchObject({ top: 900, right: 1000, bottom: 900, left: 1000 });
    const root = allRoot(sections);
    expect(root.length).toBeGreaterThan(30);
    // Footnotes resolve onto the last section in reference order; separators never arrive.
    const footnotes = root.filter((node) => node.type === "footnote");
    expect(footnotes.map((node) => (node.type === "footnote" ? node.label : "?"))).toEqual(["1", "2"]);
    const textBoxes = root.flatMap((node) => node.type === "paragraph" || node.type === "heading" ? node.floats ?? [] : [node])
      .filter((node) => node.type === "textBox");
    // Two original boxes plus the seventeen appended text-box coverage cases
    // (`test/docx-text-box.test.ts` reads those individually).
    expect(textBoxes).toHaveLength(25);
    expect(textBoxes[0]).toMatchObject({ type: "textBox", anchor: { wrap: "none" }, fill: { kind: "solid", color: "#E8F2E8" } });
    if (textBoxes[0]?.type !== "textBox") throw new Error("expected the fixture text box");
    const textBoxText = textBoxes[0].content.flatMap((node) => node.type === "paragraph" ? node.runs.map((run) => run.text) : []).join(" ");
    expect(textBoxText).toContain("صندوق ملاحظات: مراجعة التسليم");
    expect(textBoxText).toContain("English note: confirm the final delivery.");
    // Merged sections retain explicit breaks and Word's saved layout hints;
    // the two empty section-ending page-break markers are structural only.
    const breaks = root.filter((node) => node.type === "pageBreak");
    // The appended text-box pages each carry one real break of their own, and
    // the shared-effects page opens its own section.
    expect(breaks).toHaveLength(14);
    // Four appended pages were authored as real page breaks; the rest are Word's
    // own saved layout hints.
    expect(breaks.filter((node) => node.type === "pageBreak" && node.strength === "mandatory")).toHaveLength(5);
    expect(breaks.filter((node) => node.type === "pageBreak" && node.strength === "preferred")).toHaveLength(9);
    // The pull quote keeps its gold left border from `w:pBdr`.
    const quote = root.find(
      (node): node is DocxParagraphNode =>
        node.type === "paragraph" && blockText(node).includes("Cut on the movement"),
    );
    if (quote === undefined) throw new Error(`fixture structure changed: expected quote paragraph\n${summaries(sections)}`);
    expect(quote.borders?.left).toMatchObject({ color: "#B87F1E", style: "solid" });
    expect(quote.leftIndent).toBeCloseTo(0.25, 5);
    expect(root.some((node) => node.type === "image")).toBe(true);
    expect(root.some((node) => node.type === "table")).toBe(true);
    expect(root.some((node) => node.type === "heading" && node.level === 1)).toBe(true);
    expect(root.some((node) => node.type === "heading" && node.level === 2)).toBe(true);
    expect(structuredClone(sections)).toEqual(sections);
  });

  test("keeps mixed Arabic+Latin runs with script-correct typefaces", async () => {
    const sections = await fixture();
    const caption = allRoot(sections).find(
      (node): node is DocxParagraphNode =>
        node.type === "paragraph" && blockText(node).includes("Table 1 — Delivery specification"),
    );
    if (caption === undefined) throw new Error(`fixture structure changed: expected Table 1 caption\n${summaries(sections)}`);
    expect(caption.direction).toBe("ltr");
    expect(caption.runs.length).toBeGreaterThanOrEqual(3);
    const latin = caption.runs[0]!;
    const arabic = caption.runs[caption.runs.length - 1]!;
    expect(latin.text).toContain("Table 1");
    expect(latin.fontFamily).toBe("Arial");
    expect(arabic.text).toContain("الجدول");
    expect(arabic.fontFamily).toBe("Arial");
    expect(arabic.fontSizePt).toBe(9);
    // The heading mixes Cambria Latin with an Arial Arabic tail in one block.
    const heading = allRoot(sections).find((node) => node.type === "heading" && blockText(node).includes("1.1 Objectives"));
    if (heading === undefined || heading.type !== "heading") {
      throw new Error(`fixture structure changed: expected 1.1 heading\n${summaries(sections)}`);
    }
    expect(heading.direction).toBe("ltr");
    expect(heading.runs.map((run) => run.fontFamily)).toEqual(["Cambria", "Arial", "Arial"]);
  });

  test("marks explicit RTL paragraphs independently of alignment", async () => {
    const sections = await fixture();
    const arabic = await findBlock(sections, "rtl-paragraph");
    if (arabic.type !== "paragraph") throw new Error("fixture selector rtl-paragraph expected a paragraph");
    expect(arabic.direction).toBe("rtl");
    expect(arabic.runs.map((run) => run.text).join("")).toContain("صوت النهر");
    expect(arabic.runs.every((run) => run.fontFamily === "Arial")).toBe(true);
  });

  test("resolves the heading outline from paragraph styles", async () => {
    const sections = await fixture();
    const outline = allRoot(sections).filter((node) => node.type === "heading");
    expect(outline.map((node) => (node.type === "heading" ? [node.level, node.text] : []))).toEqual([
      [1, "1. Project Overview"],
      [2, "1.1 Objectives   الأهداف"],
      [2, "1.2 Deliverables   المخرجات"],
      [1, "2. Technical Specifications"],
      [1, "3. Schedule"],
      [1, "4. Contacts"],
      [1, "5. Delivery tracker / متابعة التسليم"],
      [1, "Appendix A — Shot List, Episode 04"],
    ]);
  });

  test("preserves table logical order, spans, shading, and alignment", async () => {
    const sections = await fixture();
    const first = await findBlock(sections, "table");
    if (first.type !== "table") throw new Error("fixture selector table expected a table");
    expect(first.direction).toBe("ltr");
    expect(first.columns).toHaveLength(2);
    expect(first.rows).toHaveLength(4);
    expect(first.rows[0]?.cells.map((cell) => cellText(cell))).toEqual([
      "Client//العميل",
      "Nile Heritage Foundation / مؤسسة تراث النيل",
    ]);
    expect(first.rows[0]?.cells[0]?.shading).toBe("#F4F1EC");
    expect(first.rows[0]?.cells[0]?.verticalAlign ?? "middle").toBe("middle");
    // The spec table exercises real spans: a colspan header and a rowspan body cell.
    const spec = allRoot(sections).find((node) => node.type === "table" && blockText(node).includes("Parameter / المعيار"));
    if (spec === undefined || spec.type !== "table") {
      throw new Error(`fixture structure changed: expected spec table\n${summaries(sections)}`);
    }
    expect(spec.rows[0]?.cells[1]?.columnSpan).toBe(2);
    const merged = spec.rows[2]?.cells[3];
    expect(merged?.rowSpan).toBe(2);
    expect(merged?.shading).toBe("#FBF7F0");
    expect(spec.rows[3]?.cells).toHaveLength(3);
  });

  /**
   * The delivery tracker (section 5) is a Word-authored right-to-left table
   * (`w:bidiVisual`, probe docx-bidi-visual): Word wrote the thick border
   * set on its physical right as `w:left` (the start side) and kept the
   * 0.3in cell margin set on its physical left as `w:left`.
   */
  test("reads Word's right-to-left table flag and keeps cells in logical order", async () => {
    const sections = await fixture();
    const table = await findBlock(sections, "rtl-table");
    if (table.type !== "table") throw new Error("fixture selector rtl-table expected a table");
    expect(table.direction).toBe("rtl");
    expect(table.rows[0]?.cells.map((cell) => cellText(cell))).toEqual(["الحلقة", "موقع التصوير", "المدة", "الحالة"]);
    expect(table.rows[4]?.cells[0]?.columnSpan).toBe(2);
    expect(table.borders?.left).toMatchObject({ widthPt: 3, color: "#1E3A5A" });
    expect(table.borders?.right).toMatchObject({ widthPt: 0.5, color: "#C8C8C8" });
    expect(table.cellMargins).toEqual({ top: 0, bottom: 0, left: 432 / 1440, right: 72 / 1440 });
    // `w:tblInd` (25 dxa) is measured from the start side.
    expect(table.indent).toBeCloseTo(25 / 1440, 6);
  });

  // Word keeps a run without `w:rtl` together left to right inside a
  // right-to-left paragraph: "15 Oct 2026" and "71%" as written (probe
  // run-direction). Runs with `w:rtl` follow the bidi algorithm alone.
  test("marks left-to-right runs inside a right-to-left paragraph", async () => {
    const sections = await fixture();
    const paragraph = allRoot(sections).find((node) => node.type === "paragraph" && joined(node).startsWith("آخر تحديث"));
    if (paragraph?.type !== "paragraph") throw new Error(`expected the tracker update line: ${summaries(sections)}`);
    expect(paragraph.direction).toBe("rtl");
    expect(paragraph.runs.map((run) => [run.text, run.ltrRun ?? false])).toEqual([
      ["آخر تحديث ", false],
      ["15 Oct 2026", true],
      [" بنسبة إنجاز ", false],
      ["71%", true],
    ]);
  });

  test("groups numbering into lists with resolved markers and levels", async () => {
    const sections = await fixture();
    const bullets = await findBlock(sections, "bullet-list");
    if (bullets.type !== "list") throw new Error("fixture selector bullet-list expected a list");
    expect(bullets.items).toHaveLength(7);
    expect(bullets.items.map((item) => item.marker)).toEqual(["—", "—", "◦", "—", "—", "—", "—"]);
    expect(bullets.items.map((item) => item.level)).toEqual([0, 0, 1, 0, 0, 0, 0]);
    expect(bullets.items.every((item) => item.ordered === false)).toBe(true);
    expect(bullets.items[2]?.paragraph.leftIndent).toBeCloseTo(0.625, 5);
    const ordered = await findBlock(sections, "ordered-list");
    if (ordered.type !== "list") throw new Error("fixture selector ordered-list expected a list");
    expect(ordered.items.map((item) => item.marker)).toEqual(["1.", "2.", "a)", "b)", "3."]);
    expect(ordered.items.map((item) => item.level)).toEqual([0, 0, 1, 1, 0]);
    expect(ordered.items.map((item) => item.ordered)).toEqual([true, true, true, true, true]);
  });

  test("resolves images, footnotes, headers, and footers", async () => {
    const sections = await fixture();
    const root = allRoot(sections);
    const images = root.filter((node) => node.type === "image");
    expect(images).toHaveLength(3);
    expect(images[0]).toMatchObject({ mediaId: "word/media/image1.png", contentType: "image/png" });
    expect(images[1]).toMatchObject({ contentType: "image/jpeg" });
    for (const image of images) {
      if (image.type !== "image") continue;
      expect(image.width).toBeGreaterThan(0);
      expect(image.height).toBeGreaterThan(0);
      expect(image.anchor).toBeUndefined();
      // Raster-only pictures resolve with no fallback, exactly as before the
      // SVG-first change (probe `svg-pictures`).
      expect(image.fallbackMediaId).toBeUndefined();
      expect(image.fallbackContentType).toBeUndefined();
    }
    // Both document images sit in centered paragraphs, so Word centers them.
    expect(images.slice(0, 2).map((node) => (node.type === "image" ? node.align : undefined))).toEqual(["center", "center"]);
    const footnote = await findBlock(sections, "footnote");
    if (footnote.type !== "footnote") throw new Error("fixture selector footnote expected a footnote");
    expect(footnote).toMatchObject({ id: "f1", label: "1" });
    // Footnote markers ride inside body content as superscript references
    // (this fixture's first marker sits in a spec-table cell).
    const marker = allRuns(root).find((run) => run.footnoteRef === "f1");
    expect(marker).toMatchObject({ verticalAlign: "superscript", text: "1" });
    // Each section resolves its own header/footer references. The appendix
    // (last) section keeps the pre-section output: header/footer once, with
    // PAGE/NUMPAGES marked as live fields that keep their cached text.
    const appendix = sections[1]!;
    // The logo is an inline drawing beside the wordmark, so the wordmark is
    // one paragraph whose first run carries the picture; a Heading 1 styled
    // mark follows it (a heading outside the outline).
    expect(appendix.header).toHaveLength(2);
    expect(appendix.header?.[1]).toMatchObject({ type: "heading", text: "Appendix / ملحق" });
    const logo = allRuns(appendix.header ?? []).find((run) => run.image !== undefined);
    expect(logo?.image).toMatchObject({ contentType: "image/png" });
    expect(appendix.footer).toHaveLength(1);
    const headerText = (appendix.header ?? []).map(blockText).join("\n");
    const footerText = (appendix.footer ?? []).map(blockText).join("\n");
    expect(headerText).toContain("NILE FRAME STUDIO");
    expect(footerText).toContain("Internal — Confidential");
    // Flow keeps each field's cached result once ("Page 4 of 4" as last saved).
    expect(footerText).toContain("Page 4 of 4");
    const fieldRuns = allRuns(appendix.footer ?? []).filter((run) => run.field !== undefined);
    expect(fieldRuns.map((run) => run.field).sort()).toEqual(["NUMPAGES", "PAGE"]);
    // The front section keeps its own portrait reference. (Word's even/first
    // headers held only the org label marking and were stripped with it.)
    expect(sections[0]?.properties.headerParts).toBeDefined();
    expect(sections[0]?.properties.footerParts).toBeDefined();
  });

  test("the last section's lists print the numbers Word prints, and its table keeps a cell with only an empty paragraph", async () => {
    // Authored in Word (`tools/probe/docx-lists-headers`): three items, a plain paragraph, two items that continue
    // the first list, then a list that restarts at 1 (Word writes a start override). The markers are Word's own
    // (Range.ListFormat.ListString).
    const sections = await fixture();
    const last = sections.find(section => section.root.some(node => node.type === "list" && node.items.some(item => joined(item.paragraph) === "قائمة جديدة: أولاً")))!;
    const lists = last.root.filter((node) => node.type === "list");
    expect(lists.flatMap((node) => node.items.map((item) => item.marker))).toEqual(["1.", "2.", "3.", "4.", "5.", "1.", "2."]);
    expect(lists.flatMap((node) => node.items.map((item) => joined(item.paragraph)))[5]).toBe("قائمة جديدة: أولاً");
    const table = last.root.find((node) => node.type === "table");
    if (table?.type !== "table") throw new Error(`expected the table in the last section\n${summaries(sections)}`);
    const emptyCell = table.rows[1]!.cells[0]!;
    expect(emptyCell.nodes).toHaveLength(1);
    expect(emptyCell.nodes[0]).toMatchObject({ type: "paragraph", runs: [] });
    expect(table.rows[1]!.cells[1]!.nodes.map((node) => (node.type === "paragraph" ? joined(node) : ""))).toEqual(["فارغ"]);
  });

  test("the last section has its own first-page header and shows the previous section's default one", async () => {
    // Authored in Word (`tools/probe/docx-lists-headers`): a section after the appendix that turns on a
    // different first page and writes only that header; its default header and footer are linked to the previous section's.
    const sections = await fixture();
    const appendix = sections[1]!;
    const last = sections.find(section => section.properties.headerParts?.first !== undefined)!;
    expect(last.properties.titlePage).toBe(true);
    expect(last.properties.headerParts?.first).toBeDefined();
    expect(last.properties.headerParts?.default).toBeUndefined();
    expect(last.properties.footerParts?.default).toBeUndefined();
    expect(last.firstHeader?.map(blockText).join("\n")).toBe("رأس الصفحة الأولى · First-page header");
    expect(last.header).toEqual(appendix.header);
    expect(last.footer).toEqual(appendix.footer);
    expect(last.header).not.toEqual(last.firstHeader);
    // The earlier sections do not turn it on.
    expect(sections[0]!.properties.titlePage).toBeUndefined();
    expect(appendix.properties.titlePage).toBeUndefined();
  });

  test("reads Word's charts with the shared chart reader, inline and anchored, Arabic intact", async () => {
    const sections = await fixture();
    const root = allRoot(sections);
    // Inline: the chart is the block after its text paragraph, at its 5 x 3in extent, centred like an inline picture.
    const text = await findBlock(sections, "chart-inline");
    const inline = root[root.indexOf(text) + 1];
    if (inline?.type !== "chart") throw new Error(`expected an inline chart after its paragraph\n${summaries(sections)}`);
    expect(inline).toMatchObject({ width: 5, height: 3, align: "center" });
    expect(inline.anchor).toBeUndefined();
    const bars = inline.spec.groups[0]!;
    expect(bars).toMatchObject({ kind: "bar", direction: "col", grouping: "clustered" });
    expect(bars.series.map((s) => [s.name, s.values])).toEqual([["المسجّلة", [39, 27, 8]], ["المخطّطة", [42, 40, 38]]]);
    expect(bars.series[0]!.categories).toEqual(["الحلقة الرابعة", "الحلقة الخامسة", "الحلقة السادسة"]);
    expect(inline.spec.title?.lines).toEqual(["الساعات المسجلة لكل حلقة"]);
    // Anchored: carried by its paragraph as a float, placed like a picture (square wrap, left column edge).
    const host = await findBlock(sections, "chart-anchored");
    if (host.type !== "paragraph") throw new Error("expected the anchored chart's paragraph");
    const anchored = host.floats?.find((float) => float.type === "chart");
    if (anchored?.type !== "chart") throw new Error(`expected an anchored chart float\n${summaries(sections)}`);
    expect(anchored).toMatchObject({ width: 3.5, height: 2.25 });
    expect(anchored.anchor).toMatchObject({ wrap: "square", behind: false, horizontal: { from: "column", align: "left" } });
    // Plain data end to end: it crosses the Worker boundary and JSON alike.
    expect(JSON.parse(JSON.stringify(inline))).toEqual(inline);
  });

  test("inherits paragraph and run properties through style chains", async () => {
    const styles = parseStyles(parseXml(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:cs="Arial"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>` +
        `<w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:color w:val="FF0000"/><w:sz w:val="24"/></w:rPr></w:style>` +
        `<w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/><w:rPr><w:b/></w:rPr></w:style>` +
        `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>` +
        `</w:styles>`,
    ));
    expect(styles.defaultRPr).toBeDefined();
    // The chain runs from the root base to the style itself: colour and size on Base, bold on Child.
    const chain = paragraphAncestry(styles, "Child");
    expect(chain).toHaveLength(2);
    expect(attr(firstChild(chain[0]!.rPr!, "color")!, "w:val")).toBe("FF0000");
    expect(firstChild(chain[1]!.rPr!, "b")).toBeDefined();
    expect(headingLevelForStyle("Heading1", styles)).toBe(1);
    expect(headingLevelForStyle("Child", styles)).toBeUndefined();
    // Duplicate style ids resolve last-wins, as real files repeat them.
    const dupes = parseStyles(parseXml(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:style w:type="paragraph" w:styleId="Heading1"><w:rPr><w:sz w:val="32"/></w:rPr></w:style>` +
        `<w:style w:type="paragraph" w:styleId="Heading1"><w:rPr><w:sz w:val="30"/></w:rPr></w:style>` +
        `</w:styles>`,
    ));
    expect(attr(firstChild(dupes.paragraph.get("Heading1")!.rPr!, "sz")!, "w:val")).toBe("30");
    // Cycles terminate instead of looping.
    const cyclic = parseStyles(parseXml(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="B"/></w:style>` +
        `<w:style w:type="paragraph" w:styleId="B"><w:basedOn w:val="A"/></w:style>` +
        `</w:styles>`,
    ));
    expect(paragraphAncestry(cyclic, "A")).toHaveLength(2);
  });

  test("an empty paragraph is a node of its own, sized by its paragraph mark, never a heading or a page break", async () => {
    const sections = await parseSynthetic(
      `<?xml version="1.0" encoding="UTF-8"?><w:document ${WORD_NS}><w:body>` +
        `<w:p><w:r><w:t>before</w:t></w:r></w:p>` +
        `<w:p/>` +
        `<w:p><w:pPr><w:rPr><w:sz w:val="48"/></w:rPr></w:pPr></w:p>` +
        `<w:p><w:pPr><w:bidi/><w:rPr><w:rtl/><w:sz w:val="48"/><w:szCs w:val="20"/></w:rPr></w:pPr></w:p>` +
        `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` +
        `<w:p><w:r><w:t>after</w:t></w:r></w:p>` +
        `</w:body></w:document>`,
    );
    const nodes = sections[0]!.root;
    expect(nodes.map((node) => (node.type === "paragraph" ? `p:${node.runs.length}:${node.fontSizePt ?? "-"}:${node.direction}` : node.type))).toEqual([
      "p:1:-:ltr",
      "p:0:-:ltr", // an empty paragraph keeps its line
      "p:0:24:ltr", // a 24pt paragraph mark makes it a 24pt line (Word: one 36.8px line)
      "p:0:10:rtl", // a right-to-left mark takes the complex-script size
      "pageBreak", // a page break alone is not also an empty line
      "p:1:-:ltr",
    ]);
  });

  test("bold and italic: Latin runs follow b/i, complex-script runs follow bCs/iCs (Word probe docx-cs-bold)", async () => {
    // Word 16 drew each case below: a run with only `<w:b/>` is not bold in Arabic,
    // `<w:bCs/>` alone never bolds Latin, and `w:rtl` makes even Latin letters complex-script.
    const run = (text: string, rPr: string, rtl = false): string => `<w:r><w:rPr>${rtl ? "<w:rtl/>" : ""}${rPr}</w:rPr><w:t>${text}</w:t></w:r>`;
    const cases: Array<[string, boolean | undefined, boolean | undefined]> = [
      // [rPr, latin bold, arabic bold]
      ["<w:b/>", true, undefined],
      ["<w:bCs/>", undefined, true],
      ["<w:b/><w:bCs/>", true, true],
      ['<w:b/><w:bCs w:val="0"/>', true, false],
      ['<w:b w:val="0"/><w:bCs/>', false, true],
    ];
    const sections = await parseSynthetic(
      `<?xml version="1.0" encoding="UTF-8"?><w:document ${WORD_NS}><w:body>` +
        cases.map(([rPr]) => `<w:p>${run("Latin", rPr)}${run("عربي", rPr)}${run("RTL", rPr, true)}</w:p>`).join("") +
        `<w:p>${run("Latin", "<w:i/>")}${run("عربي", "<w:i/>")}${run("Latin", "<w:iCs/>")}${run("عربي", "<w:iCs/>")}</w:p>` +
        `</w:body></w:document>`,
    );
    const paragraphs = sections[0]!.root.filter((node): node is DocxParagraphNode => node.type === "paragraph");
    cases.forEach(([rPr, latin, arabic], index) => {
      const [latinRun, arabicRun, rtlRun] = paragraphs[index]!.runs;
      expect(latinRun!.bold, `${rPr} latin`).toBe(latin);
      expect(arabicRun!.bold, `${rPr} arabic`).toBe(arabic);
      expect(rtlRun!.bold, `${rPr} Latin letters in a w:rtl run`).toBe(arabic);
    });
    const italics = paragraphs[cases.length]!.runs.map((r) => r.italic);
    expect(italics).toEqual([true, undefined, undefined, true]);
  });

  test("resolves numbering markers across formats and start overrides", async () => {
    const numbering = parseNumbering(parseXml(
      `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>` +
        `<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl></w:abstractNum>` +
        `<w:num w:numId="9"><w:abstractNumId w:val="7"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="3"/></w:lvlOverride></w:num>` +
        `</w:numbering>`,
    ));
    const levels = numbering.abstracts.get("7")!;
    const decimal = levels.get(0)!;
    const letter = levels.get(1)!;
    expect(isOrderedFormat("bullet")).toBe(false);
    expect(isOrderedFormat("decimal")).toBe(true);
    expect(startFor(numbering, "9", 0, decimal)).toBe(3);
    expect(markerFor(decimal, levels, 0, [3])).toBe("3.");
    expect(markerFor(letter, levels, 1, [3, 2])).toBe("b)");
    expect(markerFor({ numFmt: "bullet", lvlText: "●", start: 1 }, levels, 0, [1])).toBe("●");
  });

  test("numbers lists as Word does: per list definition, restarted by a start override, a heading and an empty item count", async () => {
    // The document the probe `docx-list-counters` gave Word 16; the expected markers are what Word printed
    // (Range.ListFormat.ListString) for each numbered paragraph, in order.
    const XML = `<?xml version="1.0" encoding="UTF-8"?>`;
    const MAIN_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const documentOf = (inner: string): string => `${XML}<w:document ${WORD_NS}><w:body>${inner}</w:body></w:document>`;
    const stylesOf = (inner: string): string => `${XML}<w:styles xmlns:w="${MAIN_NS}">${inner}</w:styles>`;
    const STYLES_REL = `<Relationship Id="rIdStyles" Type="${RELATIONSHIP}/styles" Target="styles.xml"/>`;
    const NUMBERING_REL = `<Relationship Id="rIdNumbering" Type="${RELATIONSHIP}/numbering" Target="numbering.xml"/>`;
    const level = (i: number, format: string, text: string): string => `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/></w:lvl>`;
    const abstracts = [0, 1, 2, 3, 4].map((id) => `<w:abstractNum w:abstractNumId="${id}">${level(0, "decimal", "%1.")}${level(1, "lowerLetter", "%2)")}</w:abstractNum>`).join("");
    const num = (id: number, abstract: number, override = ""): string => `<w:num w:numId="${id}"><w:abstractNumId w:val="${abstract}"/>${override}</w:num>`;
    const numbering = `${XML}<w:numbering xmlns:w="${MAIN_NS}">${abstracts}` +
      num(1, 0) + num(2, 0) + num(3, 1) + num(4, 1, `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>`) +
      num(5, 2) + num(6, 3) + num(7, 4) + num(8, 4) + num(9, 0) + `</w:numbering>`;
    const item = (text: string, numId: number, ilvl = 0): string =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>${text === "" ? "" : `<w:r><w:t>${text}</w:t></w:r>`}</w:p>`;
    const body =
      item("a", 1) + item("b", 1) + item("c", 2) + item("d", 2) + // same definition: one sequence
      item("e", 3) + item("f", 3) + item("g", 4) + item("h", 4) + // a start override restarts at its first use
      item("i", 5) + item("", 5) + item("j", 5) + // an empty numbered paragraph takes a number
      item("k", 6) + `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>H</w:t></w:r></w:p>` + item("l", 6) + // so does a numbered heading
      item("m", 7) + item("n", 7) + item("o", 8) + item("p", 7) +
      item("q", 9) + item("r", 9, 1) + item("s", 9, 1) + item("t", 9) + item("u", 9, 1); // continues abstract 0, deeper levels restart
    const sections = await parseSynthetic(
      documentOf(body),
      STYLES_REL + NUMBERING_REL,
      {
        "word/numbering.xml": numbering,
        "word/styles.xml": stylesOf(`<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:numPr><w:numId w:val="6"/></w:numPr><w:outlineLvl w:val="0"/></w:pPr></w:style>`),
      },
    );
    const markers = allRoot(sections).flatMap((node) => (node.type === "list" ? node.items.map((entry) => entry.marker) : node.type === "heading" ? ["(heading)"] : []));
    expect(markers).toEqual([
      "1.", "2.", "3.", "4.",
      "1.", "2.", "1.", "2.",
      "1.", "2.", "3.",
      "1.", "(heading)", "3.",
      "1.", "2.", "3.", "4.",
      "5.", "a)", "b)", "6.", "a)",
    ]);
  });

  test("shows tracked changes as final through the real section parser", async () => {
    // No Word-authored revision sample exists in this repo; this synthetic
    // package still travels the full OPC → index → section pipeline.
    const sections = await parseSynthetic(
      `<?xml version="1.0" encoding="UTF-8"?>` +
        `<w:document ${WORD_NS}><w:body>` +
        `<w:p><w:r><w:t xml:space="preserve">kept</w:t></w:r>` +
        `<w:ins w:author="a"><w:r><w:t xml:space="preserve"> inserted</w:t></w:r></w:ins>` +
        `<w:del w:author="a"><w:r><w:delText> removed</w:delText></w:r></w:del>` +
        `<w:r><w:delText>gone</w:delText><w:t xml:space="preserve"> stayed</w:t></w:r>` +
        `<w:moveFrom><w:r><w:t xml:space="preserve"> moved-away</w:t></w:r></w:moveFrom>` +
        `<w:moveTo><w:r><w:t xml:space="preserve"> moved-here</w:t></w:r></w:moveTo></w:p>` +
        `</w:body></w:document>`,
    );
    expect(sections).toHaveLength(1);
    expect(sections[0]?.properties.page.orientation).toBe("portrait");
    const root = sections[0]?.root ?? [];
    expect(root).toHaveLength(1);
    const paragraph = root[0];
    if (paragraph?.type !== "paragraph") throw new Error(`expected one paragraph, got ${summaries(sections)}`);
    expect(paragraph.runs.map((run) => run.text).join("")).toBe("kept inserted stayed moved-here");
  });

  test("keeps a live page field's cached result on the field run, not in the flow", async () => {
    const field = (instruction: string, cached: string): string =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:t>${cached}</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const sections = await parseSynthetic(
      `<?xml version="1.0" encoding="UTF-8"?>` +
        `<w:document ${WORD_NS}><w:body>` +
        `<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${field("PAGE", "21")}` +
        `<w:r><w:t xml:space="preserve"> of </w:t></w:r>${field("NUMPAGES", "42")}` +
        `<w:r><w:t xml:space="preserve"> / </w:t></w:r>` +
        `<w:fldSimple w:instr="PAGE"><w:r><w:t>7</w:t></w:r></w:fldSimple></w:p>` +
        // Only PAGE/NUMPAGES are live: another field's cached result is its content.
        `<w:p>${field("PAGEREF _Toc1 \\h", "9")}</w:p>` +
        `</w:body></w:document>`,
    );
    const [live, other] = sections[0]?.root ?? [];
    if (live?.type !== "paragraph" || other?.type !== "paragraph") throw new Error(`expected two paragraphs, got ${summaries(sections)}`);
    expect(live.runs.map((run) => run.text).join("")).toBe("Page 21 of 42 / 7");
    expect(live.runs.filter((run) => run.field !== undefined).map((run) => [run.field, run.text])).toEqual([
      ["PAGE", "21"],
      ["NUMPAGES", "42"],
      ["PAGE", "7"],
    ]);
    expect(other.runs.map((run) => run.text).join("")).toBe("9");
    expect(other.runs.every((run) => run.field === undefined)).toBe(true);
  });

  test("parses basic floats and adapts charts without a duplicate engine", async () => {
    const picture = (id: string): string =>
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:blipFill><a:blip r:embed="${id}"/></pic:blipFill></pic:pic>` +
      `</a:graphicData></a:graphic>`;
    const sections = await parseSynthetic(
      `<?xml version="1.0" encoding="UTF-8"?>` +
        `<w:document ${WORD_NS} xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body>` +
        `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" descr="inline art"/>` +
        `${picture("rIdImage")}</wp:inline></w:drawing></w:r></w:p>` +
        `<w:p><w:r><w:t xml:space="preserve">before </w:t></w:r>` +
        `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="91440" distR="182880" behindDoc="0"><wp:extent cx="914400" cy="457200"/><wp:docPr id="2" descr="anchored art"/>` +
        `<wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH>` +
        `<wp:positionV relativeFrom="paragraph"><wp:posOffset>76200</wp:posOffset></wp:positionV>` +
        `<wp:wrapSquare wrapText="bothSides"/>` +
        `${picture("rIdImage")}</wp:anchor></w:drawing></w:r>` +
        `<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>` +
        `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="3" descr="sales chart"/>` +
        `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
        `<c:chart/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>` +
        `</w:body></w:document>`,
      `<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/art.png"/>`,
      { "word/media/art.png": PIXEL_PNG },
    );
    // The anchored picture belongs to its paragraph: the text around it stays one paragraph.
    const kinds = allRoot(sections).map((node) => node.type);
    expect(kinds).toEqual(["image", "paragraph", "unsupported"]);
    const [inline, host, chart] = allRoot(sections);
    if (inline?.type !== "image" || host?.type !== "paragraph" || chart?.type !== "unsupported") {
      throw new Error(`unexpected drawing blocks: ${summaries(sections)}`);
    }
    expect(inline).toMatchObject({
      mediaId: "word/media/art.png",
      contentType: "image/png",
      width: 1,
      height: 1,
      alt: "inline art",
    });
    expect(inline.anchor).toBeUndefined();
    // The anchored drawing rides on its paragraph without splitting the text.
    expect(blockText(host)).toBe("before  after");
    expect(host.floats).toHaveLength(1);
    expect(host.floats![0]).toMatchObject({ width: 1, height: 0.5, alt: "anchored art" });
    expect(host.floats![0]!.anchor).toEqual({
      wrap: "square",
      behind: false,
      wrapText: "bothSides",
      horizontal: { from: "column", align: "right" },
      vertical: { from: "paragraph", offset: 76200 / 914400 },
      distance: { top: 0, bottom: 0, left: 0.1, right: 0.2 },
    });
    // A chart with no resolvable chart part (this document has no relationship for it) becomes the labelled fallback.
    expect(chart).toMatchObject({ source: "chart", label: "sales chart" });
  });

  test("splits sections at embedded breaks and keeps break signals as nodes", async () => {
    const sections = await parseSynthetic(
      `<?xml version="1.0" encoding="UTF-8"?>` +
        `<w:document ${WORD_NS}><w:body>` +
        `<w:p><w:r><w:t xml:space="preserve">front</w:t></w:r></w:p>` +
        `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` +
        `<w:p><w:pPr><w:pageBreakBefore/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t xml:space="preserve">middle</w:t></w:r></w:p>` +
        `<w:p><w:r><w:lastRenderedPageBreak/><w:t xml:space="preserve">tail</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t xml:space="preserve">split </w:t><w:lastRenderedPageBreak/><w:t xml:space="preserve">line</w:t></w:r></w:p>` +
        `<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>` +
        `</w:body></w:document>`,
    );
    expect(sections).toHaveLength(2);
    // Section 0 closes at the embedded break: front text, the explicit
    // break, then the middle paragraph with its own pageBreakBefore marker.
    expect(sections[0]?.root.map((node) => node.type)).toEqual(["paragraph", "pageBreak", "pageBreak", "paragraph"]);
    expect(sections[0]?.root[1]).toMatchObject({ strength: "mandatory" });
    expect(sections[0]?.properties.page.orientation).toBe("portrait");
    // Section 1: a rendered-page-break hint before any text is a preferred
    // break before its paragraph; one mid-paragraph (Word's line-level layout
    // cache) never splits the text. The trailing sectPr makes it landscape.
    expect(sections[1]?.root.map((node) => node.type)).toEqual(["pageBreak", "paragraph", "paragraph"]);
    expect(sections[1]?.root[0]).toMatchObject({ strength: "preferred" });
    expect(sections[1]?.root[2]).toMatchObject({ runs: [{ text: "split " }, { text: "line" }] });
    expect(sections[1]?.properties.page.orientation).toBe("landscape");
  });

  test("does not turn an empty section-ending pageBreakBefore into a blank page", async () => {
    const sections = await parseSynthetic(
      `<?xml version="1.0" encoding="UTF-8"?>` +
        `<w:document ${WORD_NS}><w:body>` +
        `<w:p><w:r><w:t>before</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:pageBreakBefore/><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr></w:pPr></w:p>` +
        `<w:p><w:r><w:t>after</w:t></w:r></w:p>` +
        `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>` +
        `</w:body></w:document>`,
    );
    expect(sections).toHaveLength(2);
    expect(sections[0]?.root.map((node) => node.type)).toEqual(["paragraph"]);
    expect(sections[0]?.properties.page.orientation).toBe("landscape");
    expect(sections[1]?.root.map((node) => node.type)).toEqual(["paragraph"]);
    expect(sections[1]?.properties.page.orientation).toBe("portrait");
  });

  test("resolves paragraph borders per side, skipping nil edges", async () => {
    const sections = await parseSynthetic(
      `<?xml version="1.0" encoding="UTF-8"?>` +
        `<w:document ${WORD_NS}><w:body>` +
        `<w:p><w:pPr><w:pBdr><w:left w:val="single" w:color="B87F1E" w:sz="18" w:space="14"/><w:right w:val="nil"/></w:pBdr><w:ind w:left="360"/></w:pPr>` +
        `<w:r><w:t xml:space="preserve">quoted</w:t></w:r></w:p>` +
        `</w:body></w:document>`,
    );
    expect(sections).toHaveLength(1);
    const paragraph = sections[0]?.root[0];
    if (paragraph?.type !== "paragraph") throw new Error(`expected one paragraph, got ${summaries(sections)}`);
    expect(paragraph.leftIndent).toBeCloseTo(0.25, 5);
    expect(paragraph.borders?.left).toMatchObject({ color: "#B87F1E", style: "solid" });
    expect(paragraph.borders?.left?.widthPt).toBeCloseTo(2.25, 5);
    expect(paragraph.borders?.left?.spacePt).toBe(14);
    expect(paragraph.borders?.right).toBeUndefined();
  });

  test("inherits host paragraph alignment onto inline images, never floats", async () => {
    const picture = `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic>` +
      `</a:graphicData></a:graphic>`;
    const drawing = (): string =>
      `<wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" descr="art"/>` +
      `${picture}</wp:inline>`;
    const drawingAnchor = `<wp:anchor><wp:extent cx="914400" cy="914400"/><wp:docPr id="2" descr="art"/>` +
      `<wp:positionH><wp:align>right</wp:align></wp:positionH><wp:wrapSquare wrapText="bothSides"/>${picture}</wp:anchor>`;
    const sections = await parseSynthetic(
      `<?xml version="1.0" encoding="UTF-8"?>` +
        `<w:document ${WORD_NS}><w:body>` +
        `<w:p><w:r><w:drawing>${drawing()}</w:drawing></w:r></w:p>` +
        `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>${drawing()}</w:drawing></w:r></w:p>` +
        `<w:p><w:pPr><w:bidi/><w:jc w:val="both"/></w:pPr><w:r><w:drawing>${drawing()}</w:drawing></w:r></w:p>` +
        `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>${drawingAnchor}</w:drawing></w:r></w:p>` +
        `</w:body></w:document>`,
      `<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/art.png"/>`,
      { "word/media/art.png": PIXEL_PNG },
    );
    const root = allRoot(sections);
    const images = root.filter((node) => node.type === "image");
    // The anchored picture is carried by its (otherwise empty) paragraph, not placed in the flow.
    expect(images).toHaveLength(3);
    const hosts = root.filter((node): node is Extract<FlowNode, { type: "paragraph" }> => node.type === "paragraph" && node.floats !== undefined);
    expect(hosts[0]?.floats).toHaveLength(1);
    const host = hosts[0]!;
    expect(host.runs).toEqual([]);
    expect(host.floats).toHaveLength(1);
    // A picture has no text to justify, so Word leaves it on the paragraph's
    // start side: left by default, right in a right-to-left paragraph (as the
    // Word-authored coverage page paints its picture-only paragraph), and
    // centered where the paragraph asks for it. An anchored picture is placed
    // by its anchor, never by the paragraph.
    expect(images.map((node) => (node.type === "image" ? node.align : []))).toEqual(["left", "center", "right"]);
    expect(host.floats![0]!.type === "image" || host.floats![0]!.type === "chart" ? host.floats![0]!.align : undefined).toBeUndefined();
    expect(host.floats![0]!.anchor?.horizontal.align).toBe("right");
  });

  test("rejects an already-aborted parser request with AbortError", async () => {
    const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
    const controller = new AbortController();
    controller.abort();
    await expect(indexDocx(pkg, controller.signal)).rejects.toMatchObject({
      name: "AbortError",
      message: "aared/core: request was cancelled",
    });
  });
});

describe("DOCX Word-fidelity parsing", () => {
  const XML = `<?xml version="1.0" encoding="UTF-8"?>`;
  const MAIN_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const documentOf = (body: string): string => `${XML}<w:document ${WORD_NS}><w:body>${body}</w:body></w:document>`;
  const stylesOf = (inner: string): string => `${XML}<w:styles xmlns:w="${MAIN_NS}">${inner}</w:styles>`;
  const STYLES_REL = `<Relationship Id="rIdStyles" Type="${RELATIONSHIP}/styles" Target="styles.xml"/>`;
  const NUMBERING_REL = `<Relationship Id="rIdNumbering" Type="${RELATIONSHIP}/numbering" Target="numbering.xml"/>`;
  const paragraphsOf = (sections: DocumentSection[]): DocxParagraphNode[] =>
    allRoot(sections).filter((node): node is DocxParagraphNode => node.type === "paragraph");

  test("reads left/right as start/end in bidi paragraphs, as Word paints them", async () => {
    const para = (properties: string): string => `<w:p><w:pPr>${properties}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>`;
    const sections = await parseSynthetic(documentOf(
      para(`<w:bidi/><w:jc w:val="right"/>`) +
        para(`<w:bidi/><w:jc w:val="left"/>`) +
        para(`<w:bidi/><w:jc w:val="end"/>`) +
        para(`<w:bidi/><w:jc w:val="start"/>`) +
        para(`<w:bidi/><w:jc w:val="center"/>`) +
        para(`<w:jc w:val="right"/>`) +
        para(`<w:jc w:val="both"/>`) +
        para(``),
    ));
    // Verified against a native Word export: `<w:bidi/><w:jc w:val="right"/>` paints on the left.
    expect(paragraphsOf(sections).map((node) => node.align)).toEqual([
      "left", "right", "left", "right", "center", "right", "justify", undefined,
    ]);
  });

  test("resolves tab stops through the style chain, honouring clears and leaders", async () => {
    const sections = await parseSynthetic(
      documentOf(
        `<w:p><w:pPr><w:pStyle w:val="Child"/><w:tabs><w:tab w:val="clear" w:pos="720"/><w:tab w:val="center" w:pos="4500"/></w:tabs></w:pPr>` +
          `<w:r><w:t>x</w:t><w:tab/><w:t>y</w:t></w:r></w:p>`,
      ),
      STYLES_REL,
      {
        "word/styles.xml": stylesOf(
          `<w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:tabs>` +
            `<w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:leader="dot" w:pos="9000"/>` +
            `</w:tabs></w:pPr></w:style>` +
            `<w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/></w:style>`,
        ),
      },
    );
    // The inherited left stop is cleared, the center stop is added, the leader survives, order ascends.
    expect(paragraphsOf(sections)[0]?.tabStops).toEqual([
      { pos: 4500 / 1440, align: "center" },
      { pos: 9000 / 1440, align: "right", leader: "dot" },
    ]);
  });

  test("keeps each paragraph's line rule, every level replacing the one it inherits", async () => {
    const para = (properties: string): string => `<w:p><w:pPr>${properties}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>`;
    const sections = await parseSynthetic(
      documentOf(
        para("") +
          para(`<w:pStyle w:val="Multiple"/>`) +
          para(`<w:pStyle w:val="Floor"/>`) +
          para(`<w:pStyle w:val="Floor"/><w:spacing w:line="300" w:lineRule="exact"/>`) +
          para(`<w:pStyle w:val="Floor"/><w:spacing w:line="276" w:lineRule="auto"/>`),
      ),
      STYLES_REL,
      {
        "word/styles.xml": stylesOf(
          `<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:line="280" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>` +
            `<w:style w:type="paragraph" w:styleId="Multiple"><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:style>` +
            `<w:style w:type="paragraph" w:styleId="Floor"><w:pPr><w:spacing w:line="320" w:lineRule="atLeast"/></w:pPr></w:style>`,
        ),
      },
    );
    const spacing = paragraphsOf(sections).map(({ lineSpacing, lineSpacingPt, lineSpacingAtLeast }) => ({ lineSpacing, lineSpacingPt, lineSpacingAtLeast }));
    expect(spacing).toEqual([
      // The docDefaults' exact 14pt pitch.
      { lineSpacing: undefined, lineSpacingPt: 14, lineSpacingAtLeast: undefined },
      // A style's multiple replaces the inherited exact pitch rather than sitting beside it.
      { lineSpacing: 1.5, lineSpacingPt: undefined, lineSpacingAtLeast: undefined },
      // `atLeast` is a floor, not a pitch.
      { lineSpacing: undefined, lineSpacingPt: 16, lineSpacingAtLeast: true },
      // Direct exact over an at-least style is exact again.
      { lineSpacing: undefined, lineSpacingPt: 15, lineSpacingAtLeast: undefined },
      // Direct `auto` clears both inherited point fields.
      { lineSpacing: 1.15, lineSpacingPt: undefined, lineSpacingAtLeast: undefined },
    ]);
  });

  test("composes a direct paragraph border with its style's, side by side", async () => {
    const edge = (side: string, color: string): string => `<w:${side} w:val="single" w:sz="8" w:space="1" w:color="${color}"/>`;
    const sections = await parseSynthetic(
      documentOf(`<w:p><w:pPr><w:pStyle w:val="Boxed"/><w:pBdr>${edge("left", "00FF00")}</w:pBdr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`),
      STYLES_REL,
      { "word/styles.xml": stylesOf(`<w:style w:type="paragraph" w:styleId="Boxed"><w:pPr><w:pBdr>${edge("top", "FF0000")}</w:pBdr></w:pPr></w:style>`) },
    );
    const borders = paragraphsOf(sections)[0]?.borders;
    expect(borders?.top).toMatchObject({ color: "#FF0000" });
    expect(borders?.left).toMatchObject({ color: "#00FF00" });
  });

  test("shrinks justified spaces only in Word 2013 layout (compatibility mode 15)", async () => {
    const SETTINGS_REL = `<Relationship Id="rIdSettings" Type="${RELATIONSHIP}/settings" Target="settings.xml"/>`;
    const settingsOf = (mode: string): string =>
      `${XML}<w:settings xmlns:w="${MAIN_NS}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/></w:compat></w:settings>`;
    const body = documentOf(
      `<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:bidi/><w:jc w:val="distribute"/></w:pPr><w:r><w:t>b</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t>c</w:t></w:r></w:p>`,
    );
    const shrinks = async (settings?: string): Promise<Array<true | undefined>> => {
      const sections = settings === undefined
        ? await parseSynthetic(body)
        : await parseSynthetic(body, SETTINGS_REL, { "word/settings.xml": settings });
      return paragraphsOf(sections).map((node) => node.justifyShrink);
    };
    // Justified and distributed, either direction; never a left-aligned paragraph.
    expect(await shrinks(settingsOf("15"))).toEqual([true, true, undefined]);
    // Word 2010 layout (probe docx-justify) and a document with no mode never shrink.
    expect(await shrinks(settingsOf("14"))).toEqual([undefined, undefined, undefined]);
    expect(await shrinks()).toEqual([undefined, undefined, undefined]);
    // The coverage fixture is saved by Word in mode 15: its justified paragraphs shrink.
    const justified = allRoot(await fixture()).filter((node): node is DocxParagraphNode => node.type === "paragraph" && node.align === "justify");
    expect(justified.length).toBeGreaterThan(0);
    expect(justified.every((node) => node.justifyShrink === true)).toBe(true);
  });

  test("reads caps and character spacing, with all caps outranking small caps", async () => {
    const run = (properties: string, text: string): string => `<w:r><w:rPr>${properties}</w:rPr><w:t>${text}</w:t></w:r>`;
    const sections = await parseSynthetic(documentOf(
      `<w:p>${run(`<w:caps/><w:spacing w:val="30"/>`, "a")}${run(`<w:smallCaps/>`, "b")}` +
        `${run(`<w:caps/><w:smallCaps/>`, "c")}${run(`<w:caps w:val="0"/>`, "d")}</w:p>`,
    ));
    const runs = paragraphsOf(sections)[0]!.runs;
    expect(runs.map((item) => item.caps)).toEqual(["all", "small", "all", undefined]);
    expect(runs[0]?.letterSpacingPt).toBeCloseTo(1.5, 5);
    expect(runs[1]?.letterSpacingPt).toBeUndefined();
  });

  test("resolves table width, borders and cell padding, defaulting to Word's Normal Table", async () => {
    const edge = (name: string, size: number): string => `<w:${name} w:val="single" w:sz="${size}" w:space="0" w:color="D8D3C9"/>`;
    const sections = await parseSynthetic(documentOf(
      `<w:tbl><w:tblPr><w:tblW w:w="9120" w:type="dxa"/>` +
        `<w:tblBorders>${edge("top", 4)}${edge("insideH", 2)}<w:insideV w:val="nil"/></w:tblBorders>` +
        `<w:tblCellMar><w:left w:w="10" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="4560"/><w:gridCol w:w="4560"/></w:tblGrid>` +
        `<w:tr><w:tc><w:tcPr><w:tcW w:w="4560" w:type="dxa"/>` +
        `<w:tcMar><w:top w:w="80" w:type="dxa"/></w:tcMar><w:tcBorders>${edge("bottom", 8)}</w:tcBorders></w:tcPr>` +
        `<w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    ));
    const table = allRoot(sections).find((node) => node.type === "table");
    if (table?.type !== "table") throw new Error(`expected a table: ${summaries(sections)}`);
    expect(table.width).toBeCloseTo(9120 / 1440, 5);
    expect(table.borders?.top).toMatchObject({ color: "#D8D3C9", widthPt: 0.5, style: "solid" });
    expect(table.borders?.insideH).toMatchObject({ widthPt: 0.25 });
    // `nil` is an authored "no border", distinct from an unset edge.
    expect(table.borders?.insideV).toBe("none");
    expect(table.borders?.bottom).toBeUndefined();
    // The table padding overrides one side; the rest keep the Normal Table default.
    expect(table.cellMargins).toEqual({ top: 0, bottom: 0, left: 10 / 1440, right: 108 / 1440 });
    expect(table.rows[0]?.cells[0]?.margins).toEqual({ top: 80 / 1440 });
    expect(table.rows[0]?.cells[0]?.borders?.bottom).toMatchObject({ widthPt: 1 });
    expect(table.rows[0]?.cells[1]?.margins).toBeUndefined();
  });

  test("reads a leading w:lastRenderedPageBreak on a row's first cell as its own page-break-before marker", async () => {
    const BREAK = `<w:r><w:lastRenderedPageBreak/></w:r>`;
    const cell = (text: string, lead?: string): string => `<w:tc><w:p>${lead ?? ""}<w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
    const sections = await parseSynthetic(documentOf(
      `<w:tbl><w:tblGrid><w:gridCol w:w="4560"/><w:gridCol w:w="4560"/></w:tblGrid>` +
        // Row 0: no marker.
        `<w:tr>${cell("a1")}${cell("b1")}</w:tr>` +
        // Row 1: the marker leads the FIRST cell — a real row-start signal.
        `<w:tr>${cell("a2", BREAK)}${cell("b2")}</w:tr>` +
        // Row 2: the marker leads only the SECOND cell — first cell has none,
        // so this row is not marked (the common, unambiguous case only).
        `<w:tr>${cell("a3")}${cell("b3", BREAK)}</w:tr>` +
        // Row 3: the marker sits after real text in the first cell's first
        // paragraph — mid-paragraph, dropped exactly like a plain paragraph's
        // (Notes §44/§45), never promoted to a row marker.
        `<w:tr><w:tc><w:p><w:r><w:t>a4</w:t></w:r><w:r><w:lastRenderedPageBreak/></w:r><w:r><w:t> more</w:t></w:r></w:p></w:tc>${cell("b4")}</w:tr>` +
        `</w:tbl>`,
    ));
    const table = allRoot(sections).find((node) => node.type === "table");
    if (table?.type !== "table") throw new Error(`expected a table: ${summaries(sections)}`);
    expect(table.rows.map((row) => row.pageBreakBefore)).toEqual([undefined, true, undefined, undefined]);
    // The mid-paragraph occurrence is dropped, not spliced into the text.
    expect(joined(table.rows[3]!.cells[0]!.nodes[0] as DocxParagraphNode)).toBe("a4 more");
  });

  test("reads a row's authored w:trHeight from w:trPr, not from w:tr", async () => {
    // Synthetic, because no Word-authored fixture carries a `w:trHeight` (the
    // coverage document's rows set only `w:tblHeader`). This pins the nesting
    // rule rather than a reading: ECMA-376 17.4.81 puts `w:trHeight` inside
    // `w:trPr`, and reading it off `w:tr` silently matched nothing, so every
    // authored row height was dropped. 1440 twips is one inch.
    const cell = (text: string): string => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
    const sections = await parseSynthetic(documentOf(
      `<w:tbl><w:tblGrid><w:gridCol w:w="4560"/></w:tblGrid>` +
        `<w:tr><w:trPr><w:trHeight w:val="1440" w:hRule="atLeast"/></w:trPr>${cell("tall")}</w:tr>` +
        `<w:tr>${cell("no height")}</w:tr>` +
        // `w:hRule="exact"` reads the same: the authored value is a floor, and
        // the consumer takes the larger of content and height.
        `<w:tr><w:trPr><w:trHeight w:val="720" w:hRule="exact"/></w:trPr>${cell("exact")}</w:tr>` +
        `</w:tbl>`,
    ));
    const table = allRoot(sections).find((node) => node.type === "table");
    if (table?.type !== "table") throw new Error(`expected a table: ${summaries(sections)}`);
    expect(table.rows.map((row) => row.height)).toEqual([1, undefined, 0.5]);
  });

  test("carries a list level's marker formatting and suffix onto its items", async () => {
    const item = (level: number): string =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>`;
    const sections = await parseSynthetic(
      documentOf(item(0) + item(1)),
      NUMBERING_REL,
      {
        "word/numbering.xml": `${XML}<w:numbering xmlns:w="${MAIN_NS}"><w:abstractNum w:abstractNumId="0">` +
          `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="—"/>` +
          `<w:pPr><w:ind w:left="460" w:hanging="240"/></w:pPr><w:rPr><w:b/><w:color w:val="B87F1E"/></w:rPr></w:lvl>` +
          `<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/><w:suff w:val="space"/></w:lvl>` +
          `</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
      },
    );
    const list = allRoot(sections).find((node) => node.type === "list");
    if (list?.type !== "list") throw new Error(`expected a list: ${summaries(sections)}`);
    const [first, second] = list.items;
    expect(first?.markerRun).toMatchObject({ text: "—", bold: true, color: "#B87F1E" });
    expect(first?.markerSuffix).toBeUndefined();
    // A level with no `w:rPr` follows the paragraph text; only a non-default suffix is recorded.
    expect(second?.markerRun).toBeUndefined();
    expect(second?.markerSuffix).toBe("space");
  });

  test("Word's default Symbol bullet becomes a Unicode bullet with no symbol font", async () => {
    const sections = await parseSynthetic(
      documentOf(`<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>`),
      NUMBERING_REL,
      {
        // What Word writes for its default bullet list: Symbol's private-use U+F0B7.
        "word/numbering.xml": `${XML}<w:numbering xmlns:w="${MAIN_NS}"><w:abstractNum w:abstractNumId="0">` +
          `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/>` +
          `<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>` +
          `</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
      },
    );
    const list = allRoot(sections).find((node) => node.type === "list");
    if (list?.type !== "list") throw new Error(`expected a list: ${summaries(sections)}`);
    const item = list.items[0]!;
    expect(item.marker).toBe("•");
    expect(item.markerRun?.text).toBe("•");
    expect(item.markerRun?.fontFamily).toBeUndefined();
  });

  const NUMBERED = {
    "word/numbering.xml": `${XML}<w:numbering xmlns:w="${MAIN_NS}"><w:abstractNum w:abstractNumId="0">` +
      `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>` +
      `<w:lvl w:ilvl="1"><w:start w:val="4"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2)"/></w:lvl>` +
      `</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
  };
  const numbered = (text: string, level = 0): string =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const markersIn = (nodes: FlowNode[]): string[] =>
    nodes.flatMap((node) => (node.type === "list" ? node.items.map((item) => `${item.marker} ${joined(item.paragraph)}`) : node.type === "table" ? node.rows.flatMap((row) => row.cells.flatMap((cell) => markersIn(cell.nodes))) : []));

  test("numbers one list across the body and a table cell in document order", async () => {
    const cell = `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${numbered("cell")}</w:tc></w:tr></w:tbl>`;
    const sections = await parseSynthetic(documentOf(numbered("one") + cell + numbered("three")), NUMBERING_REL, NUMBERED);
    expect(markersIn(allRoot(sections))).toEqual(["1. one", "2. cell", "3. three"]);
  });

  test("a deeper list level starts over at its own start value after a higher one", async () => {
  // The Word-verified case above proves the restart with every level starting
  // at 1. This one proves the other half: a level that starts at 4 in the file
  // restarts at 4, not at 1, and the counters survive a shallower item between
  // two deeper ones.
  const sections = await parseSynthetic(documentOf(numbered("a") + numbered("b", 1) + numbered("c", 1) + numbered("d") + numbered("e", 1)), NUMBERING_REL, NUMBERED);
  expect(markersIn(allRoot(sections))).toEqual(["1. a", "4) b", "5) c", "2. d", "4) e"]);
});

test("a content control wrapping a table cell's paragraphs keeps them", async () => {
    const cell = (inner: string): string => `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${inner}</w:tc></w:tr></w:tbl>`;
    const plain = await parseSynthetic(documentOf(cell(`<w:p><w:r><w:t>form field</w:t></w:r></w:p>`)));
    const wrapped = await parseSynthetic(documentOf(cell(`<w:sdt><w:sdtContent><w:p><w:r><w:t>form field</w:t></w:r></w:p></w:sdtContent></w:sdt>`)));
    expect(allRoot(wrapped)).toEqual(allRoot(plain));
    expect(blockText(allRoot(wrapped)[0]!)).toBe("form field");
  });

  test("a merged-down cell that also spans columns keeps the columns after it in place", async () => {
    const tc = (properties: string, text: string): string => `<w:tc><w:tcPr>${properties}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
    const grid = `<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>`;
    const rows =
      `<w:tr>${tc(`<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>`, "A")}${tc(`<w:vMerge w:val="restart"/>`, "B")}</w:tr>` +
      `<w:tr>${tc(`<w:gridSpan w:val="2"/><w:vMerge/>`, "")}${tc(`<w:vMerge/>`, "")}</w:tr>`;
    const [table] = allRoot(await parseSynthetic(documentOf(`<w:tbl>${grid}${rows}</w:tbl>`)));
    if (table?.type !== "table") throw new Error("expected a table");
    expect(table.rows[1]!.cells).toHaveLength(0);
    expect(table.rows[0]!.cells.map((cell) => [cell.columnSpan ?? 1, cell.rowSpan])).toEqual([[2, 2], [1, 2]]);
  });

  test("text after a page break in the same run lands after the break", async () => {
    const sections = await parseSynthetic(documentOf(`<w:p><w:r><w:t>before</w:t><w:br w:type="page"/><w:t>after</w:t></w:r></w:p>`));
    expect(allRoot(sections).map((node) => (node.type === "paragraph" ? joined(node) : node.type === "pageBreak" ? `break:${node.strength}` : node.type))).toEqual(["before", "break:mandatory", "after"]);
  });

  test("a section with no header reference shows the previous one's, and only a default reference counts", async () => {
    const relation = (id: string, part: string): string => `<Relationship Id="${id}" Type="${RELATIONSHIP}/header" Target="${part}"/>`;
    const header = (text: string): string => `${XML}<w:hdr xmlns:w="${MAIN_NS}"><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:hdr>`;
    const sectPr = (references: string): string => `<w:sectPr>${references}</w:sectPr>`;
    const sections = await parseSynthetic(
      documentOf(
        `<w:p><w:pPr>${sectPr(`<w:headerReference w:type="default" r:id="rIdA"/>`)}</w:pPr><w:r><w:t>one</w:t></w:r></w:p>` +
          `<w:p><w:pPr>${sectPr(`<w:headerReference w:type="first" r:id="rIdB"/>`)}</w:pPr><w:r><w:t>two</w:t></w:r></w:p>` +
          `<w:p><w:r><w:t>three</w:t></w:r></w:p>${sectPr("")}`,
      ),
      relation("rIdA", "header1.xml") + relation("rIdB", "header2.xml"),
      { "word/header1.xml": header("Default header"), "word/header2.xml": header("First-page header") },
    );
    expect(sections.map((section) => (section.header ?? []).map(blockText).join(""))).toEqual(["Default header", "Default header", "Default header"]);
  });

  test("a drawing with no wp:docPr/@id reads exactly as one with it: DOCX never uses a drawing's id", async () => {
    let removedIds = 0;
    const stripped = await withPart("docx-coverage.docx", "word/document.xml", (xml) => {
      return xml.replace(/(<wp:docPr) id="[^"]*"/g, (_, tag: string) => { removedIds++; return tag; });
    });
    expect(removedIds).toBeGreaterThan(0);
    const sections = await parseDocxSections(stripped, await indexDocx(stripped));
    const original = await fixture();
    // Only the count differs (it sits on the first section and drives the renderer's `docx-missing-id` warning).
    const { missingIds, ...first } = sections[0]!;
    expect(missingIds).toBe(removedIds);
    expect(original[0]!.missingIds).toBeUndefined();
    expect(JSON.stringify([first, ...sections.slice(1)])).toBe(JSON.stringify(original));
    expect(JSON.stringify(sections)).toContain('"type":"image"');
  });

  test("the count of drawings with no id covers headers and footnotes too, and a drawing with its id adds none", async () => {
    const drawing = (id: string): string =>
      `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr ${id} name="Picture"/>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const image = `<Relationship Id="rIdImage" Type="${RELATIONSHIP}/image" Target="media/pixel.png"/>`;
    const sections = await parseSynthetic(
      documentOf(`<w:p>${drawing('id="1"')}${drawing("")}<w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr>`),
      image + `<Relationship Id="rIdHeader" Type="${RELATIONSHIP}/header" Target="header1.xml"/><Relationship Id="rIdNotes" Type="${RELATIONSHIP}/footnotes" Target="footnotes.xml"/>`,
      {
        "word/media/pixel.png": PIXEL_PNG,
        "word/header1.xml": `${XML}<w:hdr xmlns:w="${MAIN_NS}" xmlns:r="${RELATIONSHIP}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:p>${drawing("")}</w:p></w:hdr>`,
        "word/_rels/header1.xml.rels": `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${image}</Relationships>`,
        "word/footnotes.xml": `${XML}<w:footnotes xmlns:w="${MAIN_NS}" xmlns:r="${RELATIONSHIP}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:footnote w:id="1"><w:p>${drawing("")}</w:p></w:footnote></w:footnotes>`,
        "word/_rels/footnotes.xml.rels": `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${image}</Relationships>`,
      },
    );
    // One body drawing with its id adds none; the body, header and footnote drawings without one are three.
    expect(sections[0]!.missingIds).toBe(3);
  });

  test("a hyperlink in a header and one in a footnote resolve through their own part's relationships", async () => {
    // The same relationship id names a different target in each part: the body's rels must never answer for them.
    const external = (id: string, target: string): string => `<Relationship Id="${id}" Type="${RELATIONSHIP}/hyperlink" Target="${target}" TargetMode="External"/>`;
    const rels = (...entries: string[]): string => `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.join("")}</Relationships>`;
    const link = (text: string): string => `<w:hyperlink r:id="rId9"><w:r><w:t>${text}</w:t></w:r></w:hyperlink>`;
    const sections = await parseSynthetic(
      documentOf(`<w:p>${link("body link")}<w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr>`),
      external("rId9", "https://example.com/body") +
        `<Relationship Id="rIdHeader" Type="${RELATIONSHIP}/header" Target="header1.xml"/>` +
        `<Relationship Id="rIdNotes" Type="${RELATIONSHIP}/footnotes" Target="footnotes.xml"/>`,
      {
        "word/header1.xml": `${XML}<w:hdr xmlns:w="${MAIN_NS}" xmlns:r="${RELATIONSHIP}"><w:p>${link("header link")}</w:p></w:hdr>`,
        "word/_rels/header1.xml.rels": rels(external("rId9", "https://example.com/header")),
        "word/footnotes.xml": `${XML}<w:footnotes xmlns:w="${MAIN_NS}" xmlns:r="${RELATIONSHIP}"><w:footnote w:id="1"><w:p>${link("note link")}</w:p></w:footnote></w:footnotes>`,
        "word/_rels/footnotes.xml.rels": rels(external("rId9", "https://example.com/note")),
      },
    );
    const runOf = (nodes: FlowNode[] | undefined): TextRunModel | undefined => {
      const paragraph = nodes?.find((node): node is DocxParagraphNode => node.type === "paragraph");
      return paragraph?.runs[0];
    };
    expect(runOf(sections[0]!.root)).toMatchObject({ text: "body link", hyperlink: "https://example.com/body" });
    expect(runOf(sections[0]!.header)).toMatchObject({ text: "header link", hyperlink: "https://example.com/header" });
    const note = sections[0]!.root.find((node) => node.type === "footnote");
    if (note?.type !== "footnote") throw new Error("expected the footnote");
    expect(runOf(note.nodes)).toMatchObject({ text: "note link", hyperlink: "https://example.com/note" });
  });

  test("paints a TOC's own links as plain text but keeps the Hyperlink style elsewhere", async () => {
    const link = `<w:hyperlink w:anchor="_Toc1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Entry</w:t></w:r></w:hyperlink>`;
    const sections = await parseSynthetic(
      documentOf(
        `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>${link}</w:p>` +
          `<w:p>${link}</w:p>`,
      ),
      STYLES_REL,
      {
        "word/styles.xml": stylesOf(
          `<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/></w:style>` +
            `<w:style w:type="character" w:styleId="Hyperlink"><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>`,
        ),
      },
    );
    const [toc, plain] = paragraphsOf(sections).map((node) => node.runs[0]);
    expect(toc).toMatchObject({ text: "Entry", hyperlink: "#_Toc1" });
    expect(toc?.color).toBeUndefined();
    expect(toc?.underline).toBeUndefined();
    expect(plain).toMatchObject({ hyperlink: "#_Toc1", color: "#0563C1", underline: true });
  });

  test("space in lines is 12pt a line, inherits apart from points, and wins when above zero", async () => {
    // Ground truth: Word's own PDF export of these exact cases (probe line-spacing-units). A style's
    // 1 line beats a paragraph's own 6pt; `beforeLines="0"` means "in points"; lines alone need no
    // points; Word's own save of "1 line" carries both and agrees.
    const para = (spacing: string, style = ""): string => `<w:p><w:pPr>${style}<w:spacing ${spacing}/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`;
    const lineStyle = `<w:pStyle w:val="LineStyle"/>`;
    const sections = await parseSynthetic(
      documentOf(
        para(`w:before="120"`, lineStyle) +
          para(`w:beforeLines="0" w:before="120"`, lineStyle) +
          para(`w:beforeLines="50"`) +
          para(`w:beforeLines="100" w:before="240" w:afterLines="25" w:after="0"`),
      ),
      STYLES_REL,
      { "word/styles.xml": stylesOf(`<w:style w:type="paragraph" w:styleId="LineStyle"><w:pPr><w:spacing w:beforeLines="100"/></w:pPr></w:style>`) },
    );
    expect(paragraphsOf(sections).map((node) => [node.spaceBeforePt, node.spaceAfterPt])).toEqual([
      [12, undefined], [6, undefined], [6, undefined], [12, 3],
    ]);
  });

  test("an empty theme complex-script slot resolves to the theme's Arab script font, as Word paints it", async () => {
    // Ground truth (native Word probe: minor `Arab` = Tahoma, major = Segoe UI): a `w:rtl`
    // run whose theme `cs` typeface is empty is set in the `Arab` entry; a literal `w:cs` wins.
    const theme = `${XML}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="probe">` +
      `<a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Arab" typeface="Segoe UI"/></a:majorFont>` +
      `<a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Arab" typeface="Tahoma"/></a:minorFont>` +
      `</a:fontScheme></a:themeElements></a:theme>`;
    const contentTypes = CONTENT_TYPES.replace(
      "</Types>",
      `<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`,
    );
    const arabic = "الاسم";
    const run = (rpr: string, text: string): string => `<w:r><w:rPr>${rpr}</w:rPr><w:t>${text}</w:t></w:r>`;
    const sections = await parseSynthetic(
      documentOf(
        `<w:p>${run(`<w:rFonts w:cstheme="minorBidi"/><w:rtl/>`, arabic)}` +
          `${run(`<w:rFonts w:cstheme="majorBidi"/><w:rtl/>`, arabic)}` +
          `${run(`<w:rFonts w:cs="Courier New"/><w:rtl/>`, arabic)}` +
          `${run(`<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>`, "Latin")}</w:p>`,
      ),
      `<Relationship Id="rIdTheme" Type="${RELATIONSHIP}/theme" Target="theme/theme1.xml"/>`,
      { "[Content_Types].xml": contentTypes, "word/theme/theme1.xml": theme },
    );
    const runs = paragraphsOf(sections)[0]!.runs;
    expect(runs.map((item) => item.fontFamily)).toEqual(["Tahoma", "Segoe UI", "Courier New", "Aptos"]);
    // The resolved complex-script slot is recorded too, so a font resolver can act on it.
    expect(runs[0]?.complexScriptTypeface).toBe("Tahoma");
    expect(runs[3]?.complexScriptTypeface).toBeUndefined();
  });

  test("an authored outline level makes a heading of a paragraph with no heading style", async () => {
    const sections = await parseSynthetic(documentOf(
      `<w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:r><w:t>Second level</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:outlineLvl w:val="9"/></w:pPr><w:r><w:t>Body text</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t>Plain</w:t></w:r></w:p>`,
    ));
    const [outlined, body, plain] = allRoot(sections);
    expect(outlined).toMatchObject({ type: "heading", level: 2, text: "Second level" });
    // Level 9 is Word's "body text": it names no heading.
    expect(body?.type).toBe("paragraph");
    expect(plain?.type).toBe("paragraph");
  });

  test("the heading rule: style levels, inherited outline levels, and never deleted or field text", async () => {
    const styles = stylesOf(
      `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>` +
        `<w:style w:type="paragraph" w:styleId="Custom"><w:name w:val="Chapter"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>`,
    );
    const document = documentOf(
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Named heading</w:t></w:r></w:p>` +
        // Outline level inherited through a style: a heading, though its style id says nothing.
        `<w:p><w:pPr><w:pStyle w:val="Custom"/></w:pPr><w:r><w:t>Chapter</w:t></w:r></w:p>` +
        // Level 9 is body text, however authored.
        `<w:p><w:pPr><w:outlineLvl w:val="9"/></w:pPr><w:r><w:t>Body</w:t></w:r></w:p>` +
        // A direct outline level on a paragraph that names no heading style.
        `<w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:r><w:t>Direct outline</w:t></w:r></w:p>` +
        // A paragraph with no outline level at all is body text too.
        `<w:p><w:r><w:t>Plain</w:t></w:r></w:p>` +
        // Deleted text and a field instruction never reach the title.
        `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Live</w:t></w:r>` +
        `<w:del w:id="1"><w:r><w:delText>Gone</w:delText></w:r></w:del>` +
        `<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 </w:instrText></w:r></w:p>` +
        `<w:sdt><w:sdtContent><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>In a control</w:t></w:r></w:p></w:sdtContent></w:sdt>`,
    );
    const pkg = await OpcPackage.open(syntheticDocx(document, STYLES_REL, { "word/styles.xml": styles }).buffer as ArrayBuffer);
    const index = await indexDocx(pkg);
    const sections = await parseDocxSections(pkg, index);
    const parsed = allRoot(sections).filter((node) => node.type === "heading").map((node) => (node.type === "heading" ? node.text : ""));
    // One rule, read once: the public outline (`document.headings`) is these
    // nodes, never a second derivation that could disagree.
    expect(parsed).toEqual(["Named heading", "Chapter", "Direct outline", "Live", "In a control"]);
    // The index names the one unit by the document's title, which this file doesn't set.
    expect(index.title).toBeUndefined();
  });

  test("a run inside a content control keeps its own character style", async () => {
    const sections = await parseSynthetic(
      documentOf(
        `<w:p><w:sdt><w:sdtContent><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>value</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
      ),
      STYLES_REL,
      { "word/styles.xml": stylesOf(`<w:style w:type="character" w:styleId="Strong"><w:rPr><w:b/></w:rPr></w:style>`) },
    );
    expect(paragraphsOf(sections)[0]?.runs[0]).toMatchObject({ text: "value", bold: true });
  });

  test("recognises live page fields whatever their case or switches, and keeps other simple fields' cached text", async () => {
    const sections = await parseSynthetic(documentOf(
      `<w:p><w:fldSimple w:instr=" page \\* MERGEFORMAT "><w:r><w:t>7</w:t></w:r></w:fldSimple>` +
        `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>9</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>` +
        `<w:r><w:t xml:space="preserve"> on </w:t></w:r>` +
        `<w:fldSimple w:instr=" DATE "><w:r><w:t>21 September 2026</w:t></w:r></w:fldSimple></w:p>`,
    ));
    const runs = paragraphsOf(sections)[0]!.runs;
    expect(runs.filter((run) => run.field !== undefined).map((run) => [run.field, run.text])).toEqual([
      ["PAGE", "7"],
      ["NUMPAGES", "9"],
    ]);
    // A DATE field is not live in a preview, but the date Word cached in it is real text.
    expect(joined({ runs })).toBe("79 on 21 September 2026");
  });

  test("keeps a picture on its text line, and a picture-only paragraph as a block", async () => {
    const drawing = (kind: "inline" | "anchor"): string =>
      `<w:r><w:drawing><wp:${kind}><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" descr="logo"/>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic>` +
      `</a:graphicData></a:graphic></wp:${kind}></w:drawing></w:r>`;
    const sections = await parseSynthetic(
      documentOf(
        `<w:p>${drawing("inline")}<w:r><w:t xml:space="preserve"> wordmark</w:t></w:r></w:p>` +
          `<w:p>${drawing("inline")}</w:p>`,
      ),
      `<Relationship Id="rIdImage" Type="${RELATIONSHIP}/image" Target="media/art.png"/>`,
      { "word/media/art.png": PIXEL_PNG },
    );
    const root = allRoot(sections);
    // Text plus a picture: one paragraph whose first run is the picture, sized in inches.
    const [withText, pictureOnly] = root;
    if (withText?.type !== "paragraph") throw new Error(`expected a paragraph first: ${summaries(sections)}`);
    expect(withText.runs[0]).toMatchObject({ text: "", image: { mediaId: "word/media/art.png", width: 1, height: 0.5, alt: "logo" } });
    expect(joined(withText)).toBe(" wordmark");
    // A picture alone in its paragraph stays the block image it always was.
    expect(pictureOnly?.type).toBe("image");
  });
});

/**
 * Numbered headings and table-of-contents fields (probe
 * `numbered-heading-probe`, Word 16). The fixture Word authored there reports
 * `ListString` `1.` for its Heading 1 paragraph and paints `1. Project
 * Overview`; its `TOC \h \o "1-2"` cache predates two later headings, which
 * Word's own refreshed export lists and the cached result does not.
 */
describe("DOCX numbered headings and table-of-contents fields", () => {
  const XML = `<?xml version="1.0" encoding="UTF-8"?>`;
  const MAIN_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const documentOf = (body: string): string => `${XML}<w:document ${WORD_NS}><w:body>${body}</w:body></w:document>`;
  const STYLES_REL = `<Relationship Id="rIdStyles" Type="${RELATIONSHIP}/styles" Target="styles.xml"/>`;
  const NUMBERING_REL = `<Relationship Id="rIdNumbering" Type="${RELATIONSHIP}/numbering" Target="numbering.xml"/>`;
  /** Heading 1/2 plus the TOC levels, as Word writes them. */
  const STYLES = `${XML}<w:styles xmlns:w="${MAIN_NS}">` +
    `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:pPr><w:spacing w:after="100"/></w:pPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/><w:pPr><w:ind w:left="210"/></w:pPr></w:style>` +
    `</w:styles>`;
  /**
   * The heading list Word authored: `abstractNum` 1 level 0 with `suff=space`
   * and no hanging indent, exactly the `w:ind left="0" firstLine="0"` level
   * behind `1. Project Overview`.
   */
  const NUMBERING = `${XML}<w:numbering xmlns:w="${MAIN_NS}">` +
    `<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>` +
    `<w:suff w:val="space"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl></w:abstractNum>` +
    `<w:num w:numId="7"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
  const heading = (style: string, text: string, extra = ""): string =>
    `<w:p><w:pPr><w:pStyle w:val="${style}"/>${extra}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const tocTabs = `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9576"/></w:tabs>`;
  const pageref = (anchor: string, page: string): string =>
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> PAGEREF ${anchor} \\h </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${page}</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  /** One cached entry's content: a linked entry, its tab, and its `PAGEREF`. */
  const entryRuns = (anchor: string, entry: string, page: string): string =>
    `<w:hyperlink w:anchor="${anchor}"><w:r><w:t>${entry}</w:t></w:r><w:r><w:tab/></w:r>${pageref(anchor, page)}</w:hyperlink>`;
  /** One cached TOC entry paragraph, as Word saves it. */
  const cachedEntry = (anchor: string, entry: string, page: string, style = "TOC1"): string =>
    `<w:p><w:pPr><w:pStyle w:val="${style}"/>${tocTabs}</w:pPr>${entryRuns(anchor, entry, page)}</w:p>`;
  /**
   * A complex TOC field spanning paragraphs, in the shape Word saves it (read
   * from probe `numbered-heading-probe`): the instruction and the first entry
   * share the opening paragraph, later entries are their own paragraphs, and
   * the field's `end` is a paragraph of its own. The nested `PAGEREF` fields
   * inside each entry are why field tracking has to survive a paragraph
   * boundary.
   */
  const toc = (instruction: string, cached: Array<[entry: string, page: string]>): string =>
    `<w:p><w:pPr><w:pStyle w:val="TOC1"/>${tocTabs}</w:pPr>` +
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve">${instruction}</w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    (cached.length === 0 ? "" : entryRuns("_Toc1", cached[0]![0], cached[0]![1])) +
    `</w:p>` +
    cached.slice(1).map(([entry, page], index) => cachedEntry(`_Toc${index + 2}`, entry, page, index % 2 === 0 ? "TOC1" : "TOC2")).join("") +
    `<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
  const numbered = (text: string, numId = 7): string =>
    heading("Heading1", text, `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`);
  const parse = (body: string): Promise<DocumentSection[]> =>
    parseSynthetic(
      documentOf(body),
      STYLES_REL + NUMBERING_REL,
      { "word/styles.xml": STYLES, "word/numbering.xml": NUMBERING },
    );
  const paragraphs = (sections: DocumentSection[]): DocxParagraphNode[] =>
    allRoot(sections).filter((node): node is DocxParagraphNode => node.type === "paragraph");
  const headings = (sections: DocumentSection[]): DocxHeadingNode[] =>
    allRoot(sections).filter((node): node is DocxHeadingNode => node.type === "heading");

  test("paints a numbered heading's marker and leads its outline title with it", async () => {
    const sections = await parse(numbered("Project Overview") + numbered("Deliverables"));
    const [first, second] = headings(sections);
    if (first === undefined) throw new Error(`expected the numbered headings\n${summaries(sections)}`);
    // Word reports `ListString` `1.` then `2.` for these two paragraphs and
    // paints `1. Project Overview`: the marker is resolved, not typed.
    expect(first).toMatchObject({ level: 1, marker: "1.", markerSuffix: "space", text: "1. Project Overview" });
    expect(second).toMatchObject({ marker: "2.", text: "2. Deliverables" });
    // The authored text alone stays in the runs the renderer paints, so
    // search and outline keep one text each: the runs are the heading's own,
    // `text` is what a reader sees.
    expect(first.runs.map((run) => run.text).join("")).toBe("Project Overview");
    // The level authors `w:ind left="0" firstLine="0"`: the marker sits in
    // the paragraph's own indents with no hanging one of its own.
    expect(first.leftIndent).toBeCloseTo(0, 6);
    expect(first.firstLineIndent).toBeUndefined();
  });

  test("a heading's numbering continues the same sequence as its list items", async () => {
    // Word counts per list definition, so the heading between two items takes
    // its place in one sequence: 1., [heading 2.], 3.
    const listItem = (text: string): string =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const sections = await parse(listItem("first") + numbered("Middle") + listItem("last"));
    // The heading is not a list item, so it also ends the list run: two
    // one-item lists, and the heading's own marker between them.
    const lists = allRoot(sections).filter((node) => node.type === "list");
    expect(lists.flatMap((list) => list.type === "list" ? list.items.map((item) => item.marker) : [])).toEqual(["1.", "3."]);
    expect(headings(sections)[0]).toMatchObject({ marker: "2.", text: "2. Middle" });
  });

  test("refreshes a TOC field from the document headings, including a table-cell one", async () => {
    // The cached result predates `5. Delivery tracker` and the table-cell
    // heading; Word's own refreshed export lists both.
    const body =
      toc('TOC \\h \\o "1-2"', [["Old entry", "9"]]) +
      heading("Heading1", "Project Overview") +
      heading("Heading2", "Objectives") +
      numbered("Delivery tracker") +
      `<w:tbl><w:tr><w:tc>${heading("Heading2", "Total")}</w:tc></w:tr></w:tbl>` +
      heading("Heading1", "Appendix");
    const sections = await parse(body);
    const entries = paragraphs(sections).filter((node) => node.styleId?.startsWith("TOC") === true);
    // Three entries replaced the one cached paragraph, and the table-cell
    // heading joins the document-order list Word collects.
    expect(entries.map((node) => node.runs.map((run) => run.text).join(""))).toEqual([
      "Project Overview\t",
      "Objectives\t",
      "1. Delivery tracker\t",
      "Total\t",
      "Appendix\t",
    ]);
    // Level one is `TOC1`, level two `TOC2` (entry level = heading level - \o min + 1).
    expect(entries.map((node) => node.styleId)).toEqual(["TOC1", "TOC2", "TOC1", "TOC2", "TOC1"]);
    // `\h` links each entry at its heading's outline fragment.
    // Four body headings take outline indexes 1-4; the table-cell heading is
    // outside the outline and links at its own `tocAnchor`.
    expect(entries.map((node) => node.runs[0]?.hyperlink)).toEqual([
      "#aared-heading-1", "#aared-heading-2", "#aared-heading-3", "#aared-toc-1", "#aared-heading-4",
    ]);
    // Each entry ends in a live page reference; `\o "1-2"` keeps numbers on all.
    expect(entries.map((node) => node.runs.at(-1)?.field)).toEqual([
      "PAGEREF", "PAGEREF", "PAGEREF", "PAGEREF", "PAGEREF",
    ]);
  });

  test("a refreshed entry carries the heading's numbering and keeps its cached page number", async () => {
    // `1. Delivery tracker` is a real numbered heading, so the TOC entry
    // repeats its marker (Word lists `5. Delivery tracker / متابعة التسليم`).
    const body =
      toc('TOC \\h \\o "1-2"', [["1. Delivery tracker", "3"]]) +
      numbered("Delivery tracker") +
      heading("Heading2", "Background");
    const entries = paragraphs(await parse(body)).filter((node) => node.styleId?.startsWith("TOC") === true);
    // The entry for a numbered heading repeats its marker, so the refreshed
    // entry reads `1. Delivery tracker` — what Word lists for this document.
    expect(entries[0]?.runs[0]?.text).toBe("1. ");
    expect(entries[0]?.runs.map((run) => run.text).join("")).toBe("1. Delivery tracker\t3");
    // The entry matching a cached one keeps its saved page number as the
    // flow-mode fallback; layout mode substitutes the laid-out page.
    expect(entries.find((node) => joined(node).startsWith("1. "))?.runs.at(-1)?.text).toBe("3");
    // The heading added since the TOC was last saved has no cache: its entry
    // carries no number until layout mode resolves the live one.
    expect(entries.find((node) => joined(node).startsWith("Background"))?.runs.at(-1)).toMatchObject({ field: "PAGEREF", text: "" });
  });

  test("honours the TOC level range, page-number omission, and entry separator", async () => {
    // Two fields over the same two headings, each with its own switches, so
    // one field's selection cannot mask the other's.
    const selected = toc('TOC \\o "2" \\n 1 \\p "-"', []);
    const allLevels = toc('TOC \\o "1-2" \\n', []);
    const entriesOf = async (body: string): Promise<DocxParagraphNode[]> => {
      const sections = await parse(body);
      return paragraphs(sections).filter((node) => node.styleId?.startsWith("TOC") === true);
    };
    const levelTwo = await entriesOf(heading("Heading1", "Top") + heading("Heading2", "Under") + selected);
    // `\o "2"` collects only the level-2 heading; being entry level 1 of that
    // range, `\n 1` drops its page number, and `\p "-"` would have separated
    // entry from number by a hyphen (no number to separate here).
    expect(levelTwo.map((node) => joined(node))).toEqual(["Under"]);
    expect(levelTwo[0]?.runs.some((run) => run.field === "PAGEREF")).toBe(false);
    const both = await entriesOf(heading("Heading1", "Top") + heading("Heading2", "Under") + allLevels);
    // `\n` with no range drops numbers at every level, so both entries keep none.
    expect(both.map((node) => joined(node))).toEqual(["Top", "Under"]);
    expect(both.every((node) => node.runs.some((run) => run.field === "PAGEREF") === false)).toBe(true);
    // Without `\n`, each entry carries a live page reference.
    const numberedEntries = await entriesOf(heading("Heading1", "Top") + toc('TOC \\o "1" \\p "-"', []));
    expect(numberedEntries[0]?.runs.at(-1)).toMatchObject({ field: "PAGEREF" });
    // `\p "-"` is the authored gap instead of the default tab with dot leaders.
    expect(numberedEntries[0]?.runs.at(-2)?.text).toBe("-");
  });

  test("a TOC limited to a bookmark range lists only the headings inside it", async () => {
    // `{ TOC \b Part1 }` collects entries from the bookmarked range only. The
    // bookmark is a named label: `bookmarkStart` with the name, `bookmarkEnd`
    // with the same id, either between paragraphs or inside one.
    const mark = (id: string, name: string): string => `<w:bookmarkStart w:id="${id}" w:name="${name}"/>`;
    const unmark = (id: string): string => `<w:bookmarkEnd w:id="${id}"/>`;
    // Between paragraphs, the range covers exactly the two headings inside it.
    const between = toc('TOC \\h \\o "1-2" \\b Part1', [["Old entry", "9"]]) +
      heading("Heading1", "Before") +
      mark("5", "Part1") +
      heading("Heading2", "Inside one") +
      heading("Heading2", "Inside two") +
      unmark("5") +
      heading("Heading1", "After");
    const entriesOf = async (body: string): Promise<string[]> => {
      const sections = await parse(body);
      return paragraphs(sections).filter((node) => node.styleId?.startsWith("TOC") === true).map((node) => joined(node));
    };
    expect(await entriesOf(between)).toEqual(["Inside one\t", "Inside two\t"]);
    // Written inside the headings it wraps, it covers the same two.
    const markedHeading = (text: string, open: boolean, close: boolean): string =>
      `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>` +
      (open ? mark("7", "Part1") : "") +
      `<w:r><w:t>${text}</w:t></w:r>` +
      (close ? unmark("7") : "") +
      `</w:p>`;
    const inside = toc('TOC \\h \\o "1-2" \\b Part1', [["Old entry", "9"]]) +
      heading("Heading1", "Before") +
      markedHeading("Inside one", true, false) +
      markedHeading("Inside two", false, true) +
      heading("Heading1", "After");
    expect(await entriesOf(inside)).toEqual(["Inside one\t", "Inside two\t"]);
  });

  test("several ranges of one bookmark name list the headings in any of them", async () => {
    // Word merges same-named ranges into one, so a heading inside either is
    // listed, in document order. Each range is its own id, as Word writes them.
    const body = toc('TOC \\h \\o "1-2" \\b Part1', []) +
      heading("Heading1", "Outside") +
      `<w:bookmarkStart w:id="1" w:name="Part1"/>` +
      heading("Heading2", "First range") +
      `<w:bookmarkEnd w:id="1"/>` +
      `<w:bookmarkStart w:id="2" w:name="Part1"/>` +
      heading("Heading2", "Second range") +
      `<w:bookmarkEnd w:id="2"/>` +
      `<w:bookmarkStart w:id="3" w:name="Other"/>` +
      heading("Heading2", "Other bookmark") +
      `<w:bookmarkEnd w:id="3"/>`;
    const entries = paragraphs(await parse(body))
      .filter((node) => node.styleId?.startsWith("TOC") === true)
      .map((node) => joined(node));
    // Neither the heading before the first range nor the one under a different
    // bookmark name is listed.
    expect(entries).toEqual(["First range\t", "Second range\t"]);
  });

  test("keeps the saved result when a TOC's bookmark is not in the document", async () => {
    // The refresh rule: a `\b` naming a bookmark this document does not hold
    // cannot be honoured, and listing the whole document into a field that
    // asked for a chapter would be worse than the result Word last saved.
    const body = toc('TOC \\h \\o "1-2" \\b Missing', [["Kept entry", "7"]]) +
      heading("Heading1", "Later") +
      heading("Heading2", "Also later");
    const entries = paragraphs(await parse(body)).filter((node) => node.styleId?.startsWith("TOC") === true);
    expect(entries.map((node) => joined(node))).toEqual(["Kept entry\t7"]);
    // A `\b` with no name at all is malformed, and falls back the same way.
    const nameless = toc('TOC \\h \\b', [["Kept entry", "7"]]) + heading("Heading1", "Later");
    expect(
      paragraphs(await parse(nameless)).filter((node) => node.styleId?.startsWith("TOC") === true).map((node) => joined(node)),
    ).toEqual(["Kept entry\t7"]);
  });

  test("a bookmark spanning a table lists that table's cell headings", async () => {
    // A table is one block, so a range around it covers the cell headings a
    // whole-document TOC would have listed anyway.
    const body = toc('TOC \\h \\o "1-2" \\b Part1', []) +
      `<w:bookmarkStart w:id="3" w:name="Part1"/>` +
      `<w:tbl><w:tr><w:tc>${heading("Heading2", "Cell total")}</w:tc></w:tr></w:tbl>` +
      `<w:bookmarkEnd w:id="3"/>` +
      heading("Heading1", "Outside");
    const entries = paragraphs(await parse(body)).filter((node) => node.styleId?.startsWith("TOC") === true);
    expect(entries.map((node) => joined(node))).toEqual(["Cell total\t"]);
  });

  test("keeps the saved result when a TOC's switches are not implemented", async () => {
    // `\c` collects SEQ-numbered captions and `\f` TC entries; aared does not
    // recompute either, so refreshing would silently drop entries.
    for (const instruction of ['TOC \\h \\c "tables"', 'TOC \\h \\f t', 'TOC \\h \\l "1-2"']) {
      const sections = await parse(toc(instruction, [["Kept entry", "7"]]) + heading("Heading1", "Later"));
      const tocParagraphs = paragraphs(sections).filter((node) => node.styleId?.startsWith("TOC") === true);
      expect(tocParagraphs.map((node) => joined(node)), instruction).toEqual(["Kept entry\t7"]);
    }
  });

  test("keeps the saved result when a TOC field never balances", async () => {
    // The `begin` and `separate` are here but the field's `end` is not: a
    // half-read field must never be refreshed.
    const unbalanced =
      `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve">TOC \\h \\o "1-2"</w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:t>Kept entry</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>7</w:t></w:r></w:p>`;
    const sections = await parse(unbalanced + heading("Heading1", "Later"));
    expect(paragraphs(sections).map((node) => joined(node))).toContain("Kept entry\t7");
  });

  test("reads a TOC field's instruction split across runs", async () => {
    // Word can break an instruction mid-word (`\o` and its range are separate
    // `instrText` runs); the switches still resolve and the field refreshes.
    // An instruction broken mid-word across two `instrText` runs, then a
    // separate paragraph for the field's end (as Word saves a real TOC).
    const split =
      `<w:p><w:pPr><w:pStyle w:val="TOC1"/>${tocTabs}</w:pPr>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve">TOC \\h \\o</w:instrText></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> "1-2"</w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:t>Old</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>9</w:t></w:r>` +
      `</w:p>` +
      `<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const entries = paragraphs(await parse(split + heading("Heading2", "Later"))).filter((node) => node.styleId?.startsWith("TOC") === true);
    expect(entries.map((node) => joined(node))).toEqual(["Later\t"]);
  });

  /**
   * `TOC1` spacing and a `TOC2` indent are the paragraph properties the cached
   * entries carried for free and a refreshed entry has to resolve itself.
   */
  const STYLES_LAYOUT = `${XML}<w:styles xmlns:w="${MAIN_NS}">` +
    `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="30"/></w:rPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:pPr><w:spacing w:after="100"/></w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/><w:pPr><w:spacing w:after="100"/><w:ind w:left="210"/></w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:style>` +
    `</w:styles>`;
  const entriesWith = async (styles: string, body: string): Promise<DocxParagraphNode[]> =>
    paragraphs(await parseWith(styles, body)).filter((node) => node.styleId?.startsWith("TOC") === true);
  const parseWith = (styles: string, body: string): Promise<DocumentSection[]> =>
    parseSynthetic(
      documentOf(body),
      STYLES_REL + NUMBERING_REL,
      { "word/styles.xml": styles, "word/numbering.xml": NUMBERING },
    );

  test("a refreshed entry takes its TOC style's spacing, indent and text size", async () => {
    const body =
      heading("Heading1", "Project Overview") +
      heading("Heading2", "Objectives") +
      toc('TOC \\o "1-2" \\n', []);
    const entries = await entriesWith(STYLES_LAYOUT, body);
    // `TOC1` authors `spacing after=100` (5pt), and every entry takes it.
    expect(entries.map((node) => node.spaceAfterPt)).toEqual([5, 5]);
    // `TOC2` authors `ind left=210`; a level-1 entry takes none.
    expect(entries.map((node) => node.leftIndent)).toEqual([undefined, 210 / 1440]);
    // The headings are 15pt and 12pt; the entries take the TOC style's 11pt —
    // a table of contents printed at heading size would be a loud difference.
    expect(entries.map((node) => node.fontSizePt)).toEqual([11, 11]);
    expect(entries.map((node) => node.runs[0]?.fontSizePt)).toEqual([11, 11]);
    // The entry is set in the TOC style, not the heading's: `Heading1`'s bold
    // is style-derived and Word drops it, so an entry in bold would differ on
    // every document whose headings are styled.
    expect(entries[0]?.runs[0]?.bold).toBeUndefined();
    expect(entries[1]?.runs[0]?.bold).toBeUndefined();
    // The entry's text is still the heading's, in the TOC style's typeface.
    expect(entries.map((node) => joined(node))).toEqual(["Project Overview", "Objectives"]);
  });

  test("a refreshed entry takes its direction from the TOC style, not the heading", async () => {
    // An entry for a right-to-left heading in a left-to-right TOC still paints
    // left to right: Word reads the entry paragraph's direction from its `TOCn`
    // style, and taking it from the heading mirrors the entry and its leader.
    const arabicHeading =
      `<w:p><w:pPr><w:pStyle w:val="Heading2"/><w:bidi/></w:pPr>` +
      `<w:r><w:rPr><w:rtl/></w:rPr><w:t>الإجمالي</w:t></w:r></w:p>`;
    const entries = await entriesWith(STYLES_LAYOUT, heading("Heading1", "Top") + arabicHeading + toc('TOC \\o "1-2" \\n', []));
    expect(entries.map((node) => node.direction)).toEqual(["ltr", "ltr"]);
    // A TOC style that is itself right-to-left paints every entry that way.
    // `\o "1-2"` puts this Heading 2 at entry level 2, which is `TOC2`.
    const rtlStyles = STYLES_LAYOUT.replace('<w:name w:val="toc 2"/><w:pPr>', '<w:name w:val="toc 2"/><w:pPr><w:bidi/>');
    const rtlEntries = await entriesWith(rtlStyles, arabicHeading + toc('TOC \\o "1-2" \\n', []));
    expect(rtlEntries.map((node) => node.direction)).toEqual(["rtl"]);
  });
});
