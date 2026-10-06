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
import { openApolloDocx } from "./apollo";
import { openExtendDocx, openExtendPptx, openExtendXlsx } from "./extend";
import { openPptxBrowser } from "./pptx-browser";

/**
 * Per-format renderer registry, keyed by renderer id: what the selector UI
 * lists and the harness measures (`RENDERERS_BY_FORMAT` in
 * tools/bench-measure.ts names the same ids and must stay in step). Every
 * entry is lazily imported inside its adapter, so opening one renderer never
 * pays for, or breaks from, another.
 */
export const BASELINES: Record<Format, Record<string, OpenAdapter>> = {
  pdf: { pdf: openPdf },
  docx: {
    "docx-preview": openDocx,
    "docx-renderer": openDocxRenderer,
    apollo: openApolloDocx,
    extend: openExtendDocx,
    omni: openOmniDocx,
  },
  pptx: {
    pptxviewjs: openPptx,
    silurus: openSilurusPptx,
    aiden0z: openAiden0zPptx,
    mdurbar: openMdurbarPptx,
    "pptx-browser": openPptxBrowser,
    extend: openExtendPptx,
    omni: openOmniPptx,
  },
  xlsx: {
    sheetjs: openXlsx,
    xlscalc: openXlscalcXlsx,
    "spreadsheet-preview": openSpreadsheetPreview,
    extend: openExtendXlsx,
    omni: openOmniXlsx,
  },
};

/** The baseline compare mode starts with, per format. */
export const DEFAULT_BASELINE: Record<Format, string> = {
  pdf: "pdf",
  docx: "docx-preview",
  pptx: "pptxviewjs",
  xlsx: "sheetjs",
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
