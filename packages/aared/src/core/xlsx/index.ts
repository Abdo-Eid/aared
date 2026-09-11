/**
 * Worker-side XLSX parsing surface (headless: no DOM).
 *
 * The worker orchestrator (`core/worker.ts`) owns session wiring; these
 * functions are the parse layer it calls: `indexXlsx()` for the cheap open
 * path, `readWorkbook()` for shared context, `parseXlsxSheet()` for one
 * requested sheet.
 */

export { formatCodeForId, formatNumberSection, formatNumberValue, formatSerialDate, isDateFormatCode, isDateFormatId, serialToDate } from "./numberfmt.js";
export { readXlsxStyles, resolveStyleColor } from "./styles.js";
export type { DifferentialStyle, XlsxStyles } from "./styles.js";
export { indexXlsx, readSharedStrings, readWorkbook } from "./workbook.js";
export type { XlsxSheetReference, XlsxWorkbook, XlsxWorkbookIndex } from "./workbook.js";
export { adaptChartGraphic, parseCellRange, parseCellReference, parseXlsxSheet, sliceSheetToRange } from "./sheet.js";
export { cellAddress, columnName, parseA1Range } from "./range.js";
export type { A1Range } from "./range.js";
