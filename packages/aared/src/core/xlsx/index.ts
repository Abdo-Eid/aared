/**
 * Worker-side XLSX parsing surface (headless: no DOM).
 *
 * The worker orchestrator (`core/worker.ts`) owns session wiring and imports
 * only these: `indexXlsx()` for the cheap open path, `readWorkbook()` for the
 * shared context, `parseXlsxSheet()` for one requested sheet, and
 * `sliceSheetToRange()` for a windowed read. The main thread reaches the A1
 * helpers through `range.ts` directly.
 */

export { indexXlsx, readWorkbook } from "./workbook.js";
export type { XlsxWorkbook, XlsxWorkbookIndex } from "./workbook.js";
export { parseXlsxSheet, sheetText, sliceSheetToRange } from "./sheet.js";
