import type { Format } from "../lib/format";
import type { OpenAdapter } from "./types";
import { openDocx } from "./docx";
import { openDocxRenderer } from "./docx-second";
import { openPdf } from "./pdf";
import { openPptx } from "./pptx";
import { openAiden0zPptx } from "./aiden0z";
import { openMdurbarPptx } from "./mdurbar";
import { openXlsx } from "./xlsx";
import { openSilurusPptx } from "./silurus-pptx";
import { openXlscalcXlsx } from "./xlscalc";
import { openSpreadsheetPreview } from "./spreadsheet-preview";
import { openOmniDocx, openOmniPptx, openOmniXlsx } from "./omni";

/**
 * Per-format renderer registry. `ADAPTERS` stays the legacy single-baseline
 * map (per-format default) so existing cards keep working; `BASELINES` is the
 * multi-renderer future the selector UI and the harness consume. Every entry
 * is lazily imported inside its adapter, so opening one renderer never pays
 * for — or breaks from — another.
 */
export const BASELINES: Record<Format, Record<string, OpenAdapter>> = {
  pdf: { pdf: openPdf },
  docx: {
    "docx-preview": openDocx,
    "docx-renderer": openDocxRenderer,
    omni: openOmniDocx,
  },
  pptx: {
    pptxviewjs: openPptx,
    silurus: openSilurusPptx,
    aiden0z: openAiden0zPptx,
    mdurbar: openMdurbarPptx,
    omni: openOmniPptx,
  },
  xlsx: {
    sheetjs: openXlsx,
    xlscalc: openXlscalcXlsx,
    "spreadsheet-preview": openSpreadsheetPreview,
    omni: openOmniXlsx,
  },
};

/** Legacy alias target per format — what `"baseline"` means in the UI. */
export const DEFAULT_BASELINE: Record<Format, string> = {
  pdf: "pdf",
  docx: "docx-preview",
  pptx: "pptxviewjs",
  xlsx: "sheetjs",
};

export const ADAPTERS: Record<Format, OpenAdapter> = {
  pdf: openPdf,
  docx: openDocx,
  pptx: openPptx,
  xlsx: openXlsx,
};

export type {
  BenchTheme,
  DocxOpened,
  OpenedDoc,
  OutlineEntry,
  PdfOpened,
  PptxOpened,
  XlsxOpened,
} from "./types";
export { isAbort } from "./types";
