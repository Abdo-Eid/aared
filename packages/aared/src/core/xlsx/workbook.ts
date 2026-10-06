/**
 * Workbook index and shared strings: the cheap open path. `indexXlsx()`
 * reads only the package relationships plus `xl/workbook.xml`, so opening a
 * 100-sheet workbook costs the same as opening a one-sheet workbook; each
 * sheet is inflated later by `parseXlsxSheet()` in `sheet.ts`.
 */

import { attr, childElements, firstChild, ownString, textContent } from "../xml.js";
import type { XmlElement } from "../xml.js";
import type { OpcPackage } from "../opc.js";
import { readDrawingTheme } from "../theme.js";
import type { DrawingTheme } from "../theme.js";
import { damagedDocument, throwIfAborted } from "../errors.js";
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
  /** The package carries a VBA project (`xl/vbaProject.bin`). */
  hasMacros: boolean;
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

/**
 * Index an XLSX without inflating any sheet, style, or string part. This is
 * the performance boundary: opening lists sheets; `parseXlsxSheet()` parses
 * the unit entering the viewport.
 */
export async function indexXlsx(pkg: OpcPackage, signal?: AbortSignal): Promise<XlsxWorkbookIndex> {
  throwIfAborted(signal);
  const part = await pkg.officeDocument("xlsx", signal);
  throwIfAborted(signal);
  const root = await pkg.readXml(part, signal);
  if (root.local !== "workbook") throw damagedDocument("xl/workbook.xml root must be <workbook>");
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
  const hasMacros = pkg.partNames().some((name) => name.toLowerCase() === "xl/vbaproject.bin");
  return { workbookPart: part, sheets: references, date1904, theme, hasMacros };
}

/** The text of a shared-string `<si>` or an inline-string `<is>`: plain `<t>` and rich-text `<r><t>` runs concatenated. */
export function stringItemText(item: XmlElement): string {
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

/**
 * Read `xl/sharedStrings.xml` (absent when a workbook stores only inline
 * values). The `<si>` items are streamed: a million-string table never
 * exists as an element tree, only as the resulting `string[]`.
 */
export async function readSharedStrings(pkg: OpcPackage, workbookPart: string, signal?: AbortSignal): Promise<string[]> {
  throwIfAborted(signal);
  const relationships = await pkg.relationships(workbookPart, signal);
  throwIfAborted(signal);
  const part = relationships.find((relationship) => relationship.type.endsWith("/sharedStrings"))?.target
    ?? (pkg.hasPart("xl/sharedStrings.xml") ? "xl/sharedStrings.xml" : undefined);
  if (part === undefined) return [];
  const strings: string[] = [];
  const root = await pkg.readXmlIfPresent(part, signal, {
    streamChildrenOf: "sst",
    onChild: (item) => {
      // The workbook keeps these for the session: own them, or each one pins the whole part text.
      if (item.local === "si") strings.push(ownString(stringItemText(item)));
    },
  });
  if (root === undefined) return [];
  if (root.local !== "sst") throw damagedDocument("xl/sharedStrings.xml root must be <sst>");
  return strings;
}

/**
 * Resolve everything `parseXlsxSheet()` needs beyond the sheet part itself.
 * Pass the `indexXlsx()` result when the caller already has it (the worker
 * always does) so `workbook.xml` and the theme are not parsed a second time.
 */
export async function readWorkbook(pkg: OpcPackage, signal?: AbortSignal, known?: XlsxWorkbookIndex): Promise<XlsxWorkbook> {
  throwIfAborted(signal);
  const index = known ?? await indexXlsx(pkg, signal);
  throwIfAborted(signal);
  const [sharedStrings, styles] = await Promise.all([
    readSharedStrings(pkg, index.workbookPart, signal),
    readXlsxStyles(pkg, index.workbookPart, signal, index.theme),
  ]);
  throwIfAborted(signal);
  return { ...index, sharedStrings, styles };
}
