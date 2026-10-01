import type { Format } from "./format";

export type DocOrigin = "sample" | "upload";

export interface DocMeta {
  key: string;
  name: string;
  format: Format;
  sizeBytes: number;
  /** Same-origin sample path or object URL; doubles as the download href. */
  url: string;
  bytes: Uint8Array;
  origin: DocOrigin;
}

export type Renderer = "baseline" | "aared" | BaselineRenderer;

/**
 * Baseline renderer ids. `"baseline"` is the legacy alias for the per-format
 * default (see `DEFAULT_BASELINE` in adapters); new UI should name the
 * renderer explicitly. Flat union on purpose: not every renderer handles
 * every format — `BASELINE_RENDERERS` owns the per-format allow-list.
 */
export type BaselineRenderer =
  | "pptxviewjs"
  | "sheetjs"
  | "docx-preview"
  | "silurus"
  | "omni"
  | "aiden0z"
  | "mdurbar"
  | "docx-renderer"
  | "spreadsheet-preview"
  | "xlscalc"
  | "apollo"
  | "extend"
  | "pptx-browser";

export function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}
