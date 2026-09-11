import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import { headingLevelForStyle, paragraphAncestry, parseStyles } from "../src/core/docx/styles.js";
import { isOrderedFormat, markerFor, parseNumbering, startFor } from "../src/core/docx/numbering.js";
import { parseXml } from "../src/core/xml.js";
import type { DocxParagraphNode, DocumentSection, FlowNode, TextRunModel } from "../src/core/model.js";

const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/arabic-docx-bidi-coverage.docx`;
const MANIFEST_PATH = `${import.meta.dir}/../../../fixtures/arabic-docx-bidi-coverage.json`;

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

function crc32(bytes: Uint8Array): number {
  let table = crc32Table;
  if (table === undefined) {
    table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c % 2 === 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c;
    }
    crc32Table = table;
  }
  let crc = 0xffffffff;
  for (const byte of bytes) crc = table[(crc ^ byte) & 0xff]! ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

let crc32Table: Uint32Array | undefined;

/** Minimal stored (uncompressed) ZIP writer: synthetic OPC packages for parser edge cases. */
function storedZip(files: Record<string, Uint8Array | string>): Uint8Array {
  const encoder = new TextEncoder();
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  const push = (target: Uint8Array[], view: DataView | Uint8Array): void => {
    target.push(view instanceof Uint8Array ? view : new Uint8Array(view.buffer));
  };
  for (const [name, content] of Object.entries(files)) {
    const nameBytes = encoder.encode(name);
    const data = typeof content === "string" ? encoder.encode(content) : content;
    const header = new DataView(new ArrayBuffer(30));
    header.setUint32(0, 0x04034b50, true);
    header.setUint16(4, 20, true);
    header.setUint16(6, 0x0800, true);
    header.setUint16(8, 0, true);
    header.setUint32(14, crc32(data), true);
    header.setUint32(18, data.byteLength, true);
    header.setUint32(22, data.byteLength, true);
    header.setUint16(26, nameBytes.byteLength, true);
    push(chunks, header);
    push(chunks, nameBytes);
    push(chunks, data);
    const entry = new DataView(new ArrayBuffer(46));
    entry.setUint32(0, 0x02014b50, true);
    entry.setUint16(10, 0, true);
    entry.setUint32(20, data.byteLength, true);
    entry.setUint32(24, data.byteLength, true);
    entry.setUint16(28, nameBytes.byteLength, true);
    entry.setUint32(42, offset, true);
    push(central, entry);
    push(central, nameBytes);
    offset += 30 + nameBytes.byteLength + data.byteLength;
  }
  for (const part of central) chunks.push(part);
  const centralSize = central.reduce((sum, part) => sum + part.byteLength, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, Object.keys(files).length, true);
  end.setUint16(10, Object.keys(files).length, true);
  end.setUint32(12, centralSize, true);
  end.setUint32(16, offset, true);
  push(chunks, end);
  const out = new Uint8Array(offset + centralSize + 22);
  let position = 0;
  for (const part of chunks) {
    out.set(part, position);
    position += part.byteLength;
  }
  return out;
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
    expect(sections[0]?.properties.page.margins).toMatchObject({ top: 1100, right: 1160, bottom: 1000, left: 1160 });
    expect(sections[1]?.properties.page).toMatchObject({ widthTwips: 16838, heightTwips: 11906, orientation: "landscape" });
    expect(sections[1]?.properties.page.margins).toMatchObject({ top: 900, right: 1000, bottom: 900, left: 1000 });
    const root = allRoot(sections);
    expect(root.length).toBeGreaterThan(30);
    // Footnotes resolve onto the last section in reference order; separators never arrive.
    const footnotes = root.filter((node) => node.type === "footnote");
    expect(footnotes.map((node) => (node.type === "footnote" ? node.label : "?"))).toEqual(["1", "2"]);
    // The single authored page break survives as a mandatory break node.
    const breaks = root.filter((node) => node.type === "pageBreak");
    expect(breaks).toHaveLength(1);
    expect(breaks[0]).toMatchObject({ strength: "mandatory" });
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
    expect(images[0]).toMatchObject({ mediaId: "word/media/d439a01409e90ec57db9bab183c6bb15b707c2ec.png", contentType: "image/png" });
    expect(images[1]).toMatchObject({ contentType: "image/jpeg" });
    for (const image of images) {
      if (image.type !== "image") continue;
      expect(image.width).toBeGreaterThan(0);
      expect(image.height).toBeGreaterThan(0);
      expect(image.float).toBeUndefined();
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
    // the static PAGE/NUMPAGES fiction now marked as live fields.
    const appendix = sections[sections.length - 1]!;
    expect(appendix.header).toHaveLength(2);
    expect(appendix.footer).toHaveLength(1);
    const headerText = (appendix.header ?? []).map(blockText).join("\n");
    const footerText = (appendix.footer ?? []).map(blockText).join("\n");
    expect(headerText).toContain("NILE FRAME STUDIO");
    expect(footerText).toContain("Internal — Confidential");
    expect(footerText).toContain("Page 1 of 1");
    const fieldRuns = allRuns(appendix.footer ?? []).filter((run) => run.field !== undefined);
    expect(fieldRuns.map((run) => run.field).sort()).toEqual(["NUMPAGES", "PAGE"]);
    // The front section keeps its own portrait references.
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
        `<w:r><w:drawing><wp:anchor><wp:extent cx="914400" cy="457200"/><wp:docPr id="2" descr="anchored art"/>` +
        `<wp:positionH><wp:align>right</wp:align></wp:positionH>` +
        `${picture("rIdImage", "anchored art")}</wp:anchor></w:drawing></w:r>` +
        `<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>` +
        `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="3" descr="sales chart"/>` +
        `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
        `<c:chart/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>` +
        `</w:body></w:document>`,
      `<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/art.png"/>`,
      { "word/media/art.png": PIXEL_PNG },
    );
    const kinds = allRoot(sections).map((node) => node.type);
    expect(kinds).toEqual(["image", "paragraph", "image", "paragraph", "unsupported"]);
    const [inline, before, floated, after, chart] = allRoot(sections);
    if (inline?.type !== "image" || floated?.type !== "image" || chart?.type !== "unsupported") {
      throw new Error(`unexpected drawing blocks: ${summaries(sections)}`);
    }
    expect(inline).toMatchObject({
      mediaId: "word/media/art.png",
      contentType: "image/png",
      width: 1,
      height: 1,
      alt: "inline art",
    });
    expect(inline.float).toBeUndefined();
    // The anchored drawing splits its paragraph and floats right.
    expect(blockText(before!)).toBe("before ");
    expect(blockText(after!)).toBe(" after");
    expect(floated).toMatchObject({ float: "right", width: 1, height: 0.5, alt: "anchored art" });
    // Charts adapt to the shared-contract fallback; DOCX ships no chart engine.
    expect(chart).toMatchObject({ source: "chart", label: "sales chart" });
  });

  test("splits sections at embedded breaks and keeps break signals as nodes", async () => {    const sections = await parseSynthetic(
      `<?xml version="1.0" encoding="UTF-8"?>` +
        `<w:document ${WORD_NS}><w:body>` +
        `<w:p><w:r><w:t xml:space="preserve">front</w:t></w:r></w:p>` +
        `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` +
        `<w:p><w:pPr><w:pageBreakBefore/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t xml:space="preserve">middle</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t xml:space="preserve">tail</w:t><w:lastRenderedPageBreak/></w:r></w:p>` +
        `<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>` +
        `</w:body></w:document>`,
    );
    expect(sections).toHaveLength(2);
    // Section 0 closes at the embedded break: front text, the explicit
    // break, then the middle paragraph with its own pageBreakBefore marker.
    expect(sections[0]?.root.map((node) => node.type)).toEqual(["paragraph", "pageBreak", "pageBreak", "paragraph"]);
    expect(sections[0]?.root[1]).toMatchObject({ strength: "mandatory" });
    expect(sections[0]?.properties.page.orientation).toBe("portrait");
    // Section 1 carries the soft hint after `tail`; the trailing sectPr
    // makes it landscape.
    expect(sections[1]?.root.map((node) => node.type)).toEqual(["paragraph", "pageBreak"]);
    expect(sections[1]?.root[1]).toMatchObject({ strength: "preferred" });
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
      `<wp:positionH><wp:align>right</wp:align></wp:positionH>${picture}</wp:anchor>`;
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
    const images = allRoot(sections).filter((node) => node.type === "image");
    expect(images).toHaveLength(3);
    // Plain paragraph: no alignment carried. Centered paragraph: centered
    // image. Centered paragraph with an anchored float: float wins.
    expect(images.map((node) => (node.type === "image" ? [node.align, node.float] : []))).toEqual([
      [undefined, undefined],
      ["center", undefined],
      [undefined, "right"],
    ]);
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
