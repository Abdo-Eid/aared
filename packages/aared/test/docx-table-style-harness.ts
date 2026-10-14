/** Inheritance variants of the existing Word-authored table, not new fixtures. */
import { OpcPackage } from "../src/core/opc.js";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import { FIXTURES } from "./mutate.js";
import { storedZip } from "./zip.js";

export async function inheritedTableSections(childProperties = "", directProperties = "", middleProperties = "") {
  const source = await OpcPackage.open(await Bun.file(`${FIXTURES}/docx-coverage.docx`).arrayBuffer());
  const parts: Record<string, Uint8Array> = {};
  for (const name of source.partNames()) parts[name] = await source.readPart(name);
  const decoder = new TextDecoder(), encoder = new TextEncoder();
  const styles = decoder.decode(parts["word/styles.xml"]);
  const stylePattern = /<w:style\b[^>]*w:styleId="GridTable4-Accent1"[^>]*>[\s\S]*?<\/w:style>/;
  if (!stylePattern.test(styles)) throw new Error("the Word-authored table style is missing");
  const edges = ["top", "right", "bottom", "left", "insideH", "insideV"]
    .map(side => `<w:${side} w:val="single" w:sz="8" w:color="123456"/>`).join("");
  const margins = ["top", "right", "bottom", "left"]
    .map((side, index) => `<w:${side} w:w="${(index + 1) * 144}" w:type="dxa"/>`).join("");
  const base = `<w:style w:type="table" w:styleId="InheritedTableBase"><w:tblPr>` +
    `<w:tblBorders>${edges}</w:tblBorders><w:tblCellMar>${margins}</w:tblCellMar>` +
    `</w:tblPr></w:style>`;
  const middle = `<w:style w:type="table" w:styleId="InheritedTableMiddle">` +
    `<w:basedOn w:val="InheritedTableBase"/>${middleProperties}</w:style>`;
  const child = `<w:style w:type="table" w:styleId="GridTable4-Accent1">` +
    `<w:basedOn w:val="InheritedTableMiddle"/>${childProperties}</w:style>`;
  parts["word/styles.xml"] = encoder.encode(styles.replace(stylePattern, base + middle + child));
  const document = decoder.decode(parts["word/document.xml"]);
  let found = false;
  const rewritten = document.replace(/<w:tblPr\b[^>]*>[\s\S]*?<\/w:tblPr>/g, properties => {
    if (!properties.includes('w:val="GridTable4-Accent1"')) return properties;
    found = true;
    // Keep the real table, its look and its geometry, varying only the source
    // of table-level borders/padding so direct fixture values cannot mask it.
    return properties.replace(/<w:tblBorders\b[^>]*>[\s\S]*?<\/w:tblBorders>/g, "")
      .replace(/<w:tblCellMar\b[^>]*>[\s\S]*?<\/w:tblCellMar>/g, "")
      .replace("</w:tblPr>", `${directProperties}</w:tblPr>`);
  });
  if (!found) throw new Error("the Word-authored styled table is missing");
  parts["word/document.xml"] = encoder.encode(rewritten);
  const pkg = await OpcPackage.open(storedZip(parts).slice().buffer as ArrayBuffer);
  return parseDocxSections(pkg, await indexDocx(pkg));
}
