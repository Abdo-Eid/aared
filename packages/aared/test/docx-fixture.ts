import { OpcPackage } from "../src/core/opc.js";
import { attr, firstChild } from "../src/core/xml.js";
import type { XmlElement } from "../src/core/xml.js";

/**
 * An input oracle for docx-coverage.docx, whose pages are explicitly authored
 * rather than overflow-driven. Count its body section/page breaks directly
 * from XML, independently of parseDocxSections, paginate and renderer output.
 * Empty section-marker paragraphs also carry pageBreakBefore in Word's save;
 * those flags do not add a second page. A new overflow/continuous-section case
 * needs its own native-backed expectation, not an extension of this counter.
 */
export async function authoredCoveragePages(pkg: OpcPackage): Promise<number> {
  const body = firstChild(await pkg.readXml("word/document.xml"), "body");
  if (body === undefined) throw new Error("DOCX coverage fixture has no body");
  let pages = 1;
  const runBreaks = (element: XmlElement): number => {
    if (element.local === "drawing" || element.local === "pict") return 0;
    if (element.local === "br") return attr(element, "type") === "page" ? 1 : 0;
    return element.children.reduce((sum, child) => sum + runBreaks(child), 0);
  };
  for (const paragraph of body.children.filter((element) => element.local === "p")) {
    const properties = firstChild(paragraph, "pPr");
    const section = properties && firstChild(properties, "sectPr");
    if (section !== undefined) {
      const type = firstChild(section, "type");
      if (type !== undefined && attr(type, "val") !== "nextPage") {
        throw new Error("DOCX coverage page oracle needs a native expectation for this section type");
      }
      pages++;
    } else if (properties !== undefined) {
      const before = firstChild(properties, "pageBreakBefore");
      if (before !== undefined && !["0", "false", "off"].includes(attr(before, "val") ?? "")) pages++;
    }
    // Only inline content: section properties and nested drawing stories are
    // not main-body page breaks.
    for (const inline of paragraph.children.filter((element) => element.local !== "pPr")) pages += runBreaks(inline);
  }
  return pages;
}
