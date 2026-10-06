import { storedZip } from "./zip.js";

// Synthetic documents for the header-selection tests (`docx-headers.test.ts`, `docx-headers.visual.test.ts`):
// no coverage deck carries first-page or even-page headers. Not a test itself.

const W = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const SIZE = `<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>`;

const headerPart = (text: string): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:hdr>`;
export const refs = (spec: Array<[type: string, id: string]>): string => spec.map(([type, id]) => `<w:headerReference w:type="${type}" r:id="${id}"/>`).join("");
/** A body paragraph that ends its page. */
export const page = (label: string, last = false): string => `<w:p><w:r><w:t>${label}</w:t></w:r>${last ? "" : `<w:r><w:br w:type="page"/></w:r>`}</w:p>`;
/** The last paragraph of a section, carrying its `sectPr`. */
export const endOfSection = (label: string, sectPr: string): string => `<w:p><w:pPr>${sectPr}</w:pPr><w:r><w:t>${label}</w:t></w:r></w:p>`;

const HEADERS = ["A-DEFAULT", "A-FIRST", "A-EVEN", "C-DEFAULT", "D-DEFAULT"];

export function build(body: string, settings: string): Uint8Array {
  const parts: Record<string, string> = {};
  HEADERS.forEach((text, index) => {
    parts[`word/header${index + 1}.xml`] = headerPart(text);
  });
  return storedZip({
    "[Content_Types].xml":
      `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>` +
      HEADERS.map((_, index) => `<Override PartName="/word/header${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`).join("") +
      `</Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels":
      `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      HEADERS.map((_, index) => `<Relationship Id="rId${index + 1}" Type="${REL}/header" Target="header${index + 1}.xml"/>`).join("") +
      `<Relationship Id="rId9" Type="${REL}/settings" Target="settings.xml"/></Relationships>`,
    "word/settings.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W}>${settings}</w:settings>`,
    "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`,
    ...parts,
  });
}


/** Four sections: A (first and even headers, 4 pages), B (inherits, 2), C (titlePg only, 2), D (restarts at 1, 3). */
export function fourSectionDocument(): Uint8Array {
  const A = `<w:sectPr>${refs([["default", "rId1"], ["first", "rId2"], ["even", "rId3"]])}${SIZE}<w:titlePg/></w:sectPr>`;
  const B = `<w:sectPr><w:type w:val="nextPage"/>${SIZE}</w:sectPr>`;
  const C = `<w:sectPr><w:type w:val="nextPage"/>${refs([["default", "rId4"]])}${SIZE}<w:titlePg/></w:sectPr>`;
  const D = `<w:sectPr><w:type w:val="nextPage"/>${refs([["default", "rId5"]])}${SIZE}<w:pgNumType w:start="1"/></w:sectPr>`;
  const body =
    page("A1") + page("A2") + page("A3") + endOfSection("A4", A) +
    page("B5") + endOfSection("B6", B) +
    page("C7") + endOfSection("C8", C) +
    page("D1") + page("D2") + page("D3", true) + D;
  return build(body, "<w:evenAndOddHeaders/>");
}
