import type { Format } from "../lib/format";
import type { OpenAdapter } from "./types";
import { openDocx } from "./docx";
import { openPdf } from "./pdf";
import { openPptx } from "./pptx";
import { openXlsx } from "./xlsx";

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
