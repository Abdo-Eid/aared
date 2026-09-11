/**
 * Workbook index and shared strings: the cheap open path. `indexXlsx()`
 * reads only the package relationships plus `xl/workbook.xml`, so opening a
 * 100-sheet workbook costs the same as opening a one-sheet workbook; each
 * sheet is inflated later by `parseXlsxSheet()` in `sheet.ts`.
 */

import { attr, childElements, firstChild, parseXml, textContent } from "../xml.js";
import type { XmlElement } from "../xml.js";
import type { OpcPackage } from "../opc.js";
import { readDrawingTheme } from "../theme.js";
import type { DrawingTheme } from "../theme.js";
import { throwIfAborted } from "../zip.js";
import { readXlsxStyles } from "./styles.js";
import type { XlsxStyles } from "./styles.js";

export interface XlsxSheetReference {
  /** Tab name. */
  name: string;
  /** Worksheet part, e.g. `xl/worksheets/sheet1.xml`. */
  part: string;
  state: "visible" | "hidden" | "veryHidden";
}

export interface XlsxWorkbookIndex {
  /** `xl/workbook.xml` part backing the sheets below. */
  workbookPart: string;
  sheets: XlsxSheetReference[];
  /** True when `workbookPr date1904="1"` (legacy Mac date system). */
  date1904: boolean;
  theme: DrawingTheme;
}

/** Fully resolved read context for `parseXlsxSheet()`. */
export interface XlsxWorkbook extends XlsxWorkbookIndex {
  sharedStrings: string[];
  styles: XlsxStyles;
}

function sheetState(value: string | undefined): XlsxSheetReference["state"] {
  if (value === "hidden") return "hidden";
  if (value === "veryHidden") return "veryHidden";
  return "visible";
}

async function workbookPart(pkg: OpcPackage, signal?: AbortSignal): Promise<string> {
  throwIfAborted(signal);
  const packageRelationships = await pkg.packageRelationships(signal);
  throwIfAborted(signal);
  const office = packageRelationships.find((relationship) => relationship.type.endsWith("/officeDocument"));
  if (office === undefined || office.targetMode === "External") {
    throw new Error("aared/core: XLSX has no internal officeDocument relationship");
  }
  const contentType = pkg.contentTypeOf(office.target);
  if (
    contentType !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    && contentType !== "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
  ) {
    throw new Error("aared/core: package is not an XLSX workbook");
  }
  return office.target;
}

/**
 * Index an XLSX without inflating any sheet, style, or string part. This is
 * the performance boundary: opening lists sheets; `parseXlsxSheet()` parses
 * the unit entering the viewport.
 */
export async function indexXlsx(pkg: OpcPackage, signal?: AbortSignal): Promise<XlsxWorkbookIndex> {
  throwIfAborted(signal);
  const part = await workbookPart(pkg, signal);
  throwIfAborted(signal);
  const root = parseXml(await pkg.readText(part, signal));
  throwIfAborted(signal);
  if (root.local !== "workbook") throw new Error("aared/core: xl/workbook.xml root must be <workbook>");
  const relationships = await pkg.relationships(part, signal);
  throwIfAborted(signal);
  const sheets = firstChild(root, "sheets");
  const references: XlsxSheetReference[] = [];
  if (sheets !== undefined) {
    for (const sheet of childElements(sheets, "sheet")) {
      throwIfAborted(signal);
      const id = attr(sheet, "r:id");
      const name = attr(sheet, "name");
      if (id === undefined || name === undefined) continue;
      const relationship = relationships.find((candidate) => candidate.id === id && candidate.type.endsWith("/worksheet"));
      if (relationship === undefined || relationship.targetMode === "External") continue;
      references.push({ name, part: relationship.target, state: sheetState(attr(sheet, "state")) });
    }
  }
  const properties = firstChild(root, "workbookPr");
  const date1904 = properties !== undefined && (attr(properties, "date1904") === "1" || attr(properties, "date1904") === "true");
  const theme = await readDrawingTheme(pkg, signal);
  throwIfAborted(signal);
  return { workbookPart: part, sheets: references, date1904, theme };
}

function stringItemText(item: XmlElement): string {
  // Plain `<t>` items and rich-text `<r><t>` runs concatenate; whitespace is
  // preserved because the parser keeps character data verbatim.
  let out = "";
  for (const child of item.children) {
    if (child.local === "t") out += textContent(child);
    else if (child.local === "r") {
      const run = firstChild(child, "t");
      if (run !== undefined) out += textContent(run);
    }
  }
  if (out === "") out = textContent(item);
  return out;
}

/** Read `xl/sharedStrings.xml` (absent when a workbook stores only inline values). */
export async function readSharedStrings(pkg: OpcPackage, workbookPart: string, signal?: AbortSignal): Promise<string[]> {
  throwIfAborted(signal);
  const relationships = await pkg.relationships(workbookPart, signal);
  throwIfAborted(signal);
  const part = relationships.find((relationship) => relationship.type.endsWith("/sharedStrings"))?.target
    ?? (pkg.hasPart("xl/sharedStrings.xml") ? "xl/sharedStrings.xml" : undefined);
  if (part === undefined || !pkg.hasPart(part)) return [];
  const root = parseXml(await pkg.readText(part, signal));
  throwIfAborted(signal);
  if (root.local !== "sst") throw new Error("aared/core: xl/sharedStrings.xml root must be <sst>");
  const strings = childElements(root, "si").map(stringItemText);
  throwIfAborted(signal);
  return strings;
}

/** Resolve everything `parseXlsxSheet()` needs beyond the sheet part itself. */
export async function readWorkbook(pkg: OpcPackage, signal?: AbortSignal): Promise<XlsxWorkbook> {
  throwIfAborted(signal);
  const index = await indexXlsx(pkg, signal);
  throwIfAborted(signal);
  const [sharedStrings, styles] = await Promise.all([
    readSharedStrings(pkg, index.workbookPart, signal),
    readXlsxStyles(pkg, index.workbookPart, signal),
  ]);
  throwIfAborted(signal);
  return { ...index, sharedStrings, styles };
}
