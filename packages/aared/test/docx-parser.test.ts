import { describe, expect, test } from "bun:test";
import { storedZip } from "./zip.js";
import { OpcPackage } from "../src/core/opc.js";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import { headingLevelForStyle, paragraphAncestry, parseStyles } from "../src/core/docx/styles.js";
import { isOrderedFormat, markerFor, parseNumbering, startFor } from "../src/core/docx/numbering.js";
import { parseXml } from "../src/core/xml.js";
import type { DocxParagraphNode, DocumentSection, FlowNode, TextRunModel } from "../src/core/model.js";

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
  test("resolves the gate fixture to structured-cloneable document sections", async () => {
    const sections = await fixture();
    // Portrait body closed by the embedded `sectPr`, landscape appendix by
    // the trailing body `sectPr`.
    expect(sections).toHaveLength(2);
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
    // The authored page breaks survive as mandatory break nodes (the brief's
    // one, plus one opening each of the two picture-wrapping pages); the
    // `lastRenderedPageBreak` hints Word saved ride along as soft ones (three
    // in the brief, one more on each wrapping page).
    const breaks = root.filter((node) => node.type === "pageBreak");
    expect(breaks).toHaveLength(8);
    expect(breaks.filter((node) => node.type === "pageBreak" && node.strength === "mandatory")).toHaveLength(3);
    expect(breaks.filter((node) => node.type === "pageBreak" && node.strength === "preferred")).toHaveLength(5);
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
    expect(images).toHaveLength(2);
    expect(images[0]).toMatchObject({ mediaId: "word/media/image1.png", contentType: "image/png" });
    expect(images[1]).toMatchObject({ contentType: "image/jpeg" });
    for (const image of images) {
      if (image.type !== "image") continue;
      expect(image.width).toBeGreaterThan(0);
      expect(image.height).toBeGreaterThan(0);
      expect(image.anchor).toBeUndefined();
    }
    // Both document images sit in centered paragraphs, so Word centers them.
    expect(images.map((node) => (node.type === "image" ? node.align : undefined))).toEqual(["center", "center"]);
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
    const appendix = sections[sections.length - 1]!;
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
    expect(sections[0]?.properties.headerRefs).toHaveLength(1);
    expect(sections[0]?.properties.footerRefs).toHaveLength(1);
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
    expect(paragraphAncestry(styles, "Child")).toHaveLength(2);
    expect(paragraphAncestry(styles, "Child")[0]?.rPr).toBeDefined();
    expect(paragraphAncestry(styles, "Child")[1]?.rPr).toBeDefined();
    expect(headingLevelForStyle("Heading1", styles)).toBe(1);
    expect(headingLevelForStyle("Child", styles)).toBeUndefined();
    // Duplicate style ids resolve last-wins, as real files repeat them.
    const dupes = parseStyles(parseXml(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:style w:type="paragraph" w:styleId="Heading1"><w:rPr><w:sz w:val="32"/></w:rPr></w:style>` +
        `<w:style w:type="paragraph" w:styleId="Heading1"><w:rPr><w:sz w:val="30"/></w:rPr></w:style>` +
        `</w:styles>`,
    ));
    expect(dupes.paragraph.get("Heading1")?.rPr).toBeDefined();
    // Cycles terminate instead of looping.
    const cyclic = parseStyles(parseXml(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="B"/></w:style>` +
        `<w:style w:type="paragraph" w:styleId="B"><w:basedOn w:val="A"/></w:style>` +
        `</w:styles>`,
    ));
    expect(paragraphAncestry(cyclic, "A")).toHaveLength(2);
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
    const picture = (id: string, descr: string): string =>
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:blipFill><a:blip r:embed="${id}"/></pic:blipFill></pic:pic>` +
      `</a:graphicData></a:graphic>`;
    const sections = await parseSynthetic(
      `<?xml version="1.0" encoding="UTF-8"?>` +
        `<w:document ${WORD_NS} xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body>` +
        `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" descr="inline art"/>` +
        `${picture("rIdImage", "inline art")}</wp:inline></w:drawing></w:r></w:p>` +
        `<w:p><w:r><w:t xml:space="preserve">before </w:t></w:r>` +
        `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="91440" distR="182880" behindDoc="0"><wp:extent cx="914400" cy="457200"/><wp:docPr id="2" descr="anchored art"/>` +
        `<wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH>` +
        `<wp:positionV relativeFrom="paragraph"><wp:posOffset>76200</wp:posOffset></wp:positionV>` +
        `<wp:wrapSquare wrapText="bothSides"/>` +
        `${picture("rIdImage", "anchored art")}</wp:anchor></w:drawing></w:r>` +
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
    // Charts adapt to the shared-contract fallback; DOCX ships no chart engine.
    expect(chart).toMatchObject({ source: "chart", label: "sales chart" });
  });

  test("splits sections at embedded breaks and keeps break signals as nodes", async () => {    const sections = await parseSynthetic(
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
        `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>${drawingAnchor}</w:drawing></w:r></w:p>` +
        `</w:body></w:document>`,
      `<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/art.png"/>`,
      { "word/media/art.png": PIXEL_PNG },
    );
    const root = allRoot(sections);
    const images = root.filter((node) => node.type === "image");
    // The anchored picture is carried by its (otherwise empty) paragraph, not placed in the flow.
    expect(images).toHaveLength(2);
    const host = root[2];
    if (host?.type !== "paragraph") throw new Error(`expected a paragraph carrying the float, got ${summaries(sections)}`);
    expect(host.runs).toEqual([]);
    expect(host.floats).toHaveLength(1);
    // Plain paragraph: no alignment carried. Centered paragraph: centered
    // image. Centered paragraph with an anchored picture: paragraph alignment
    // never positions it, the anchor does.
    expect(images.map((node) => (node.type === "image" ? node.align : []))).toEqual([undefined, "center"]);
    expect(host.floats![0]!.align).toBeUndefined();
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
        // (DECISION_LOG §44/§45), never promoted to a row marker.
        `<w:tr><w:tc><w:p><w:r><w:t>a4</w:t></w:r><w:r><w:lastRenderedPageBreak/></w:r><w:r><w:t> more</w:t></w:r></w:p></w:tc>${cell("b4")}</w:tr>` +
        `</w:tbl>`,
    ));
    const table = allRoot(sections).find((node) => node.type === "table");
    if (table?.type !== "table") throw new Error(`expected a table: ${summaries(sections)}`);
    expect(table.rows.map((row) => row.pageBreakBefore)).toEqual([undefined, true, undefined, undefined]);
    // The mid-paragraph occurrence is dropped, not spliced into the text.
    expect(joined(table.rows[3]!.cells[0]!.nodes[0] as DocxParagraphNode)).toBe("a4 more");
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
    expect(parsed).toEqual(["Named heading", "Chapter", "Live", "In a control"]);
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
