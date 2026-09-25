import type { Format } from "../lib/format";

export type BenchTheme = "light" | "dark";

export interface OutlineEntry {
  level: number;
  text: string;
  /** Baseline libraries expose no navigation API, so their outline scrolls
   * to the painted DOM node directly. */
  element?: Element;
  /** aared exposes real heading navigation (`goToHeading`); preferred over
   * `element` when present. */
  onSelect?: () => void;
}

interface BaseOpened {
  pageCount: number;
  /** Milliseconds from adapter entry to first meaningful paint. */
  ms: number;
  destroy(): void;
}

export interface PdfOpened extends BaseOpened {
  format: "pdf";
  /** Low-res data URLs for the rail, capped by the adapter. */
  thumbs: string[];
  scrollToPage(page: number): void;
}

export interface DocxOpened extends BaseOpened {
  format: "docx";
  outline: OutlineEntry[];
  /** Widest laid-out page in CSS px, measured before any scaling. */
  paperWidth: number;
  /** Laid-out pages in order, for current-page-aware zoom. */
  pages: Array<{ width: number; height: number }>;
}

export interface PptxOpened extends BaseOpened {
  format: "pptx";
  /** Zero-based slide index. */
  goToSlide(index: number): void | Promise<void>;
}

export interface XlsxOpened extends BaseOpened {
  format: "xlsx";
  sheets: string[];
  renderSheet(index: number): void | Promise<void>;
}

export type OpenedDoc = PdfOpened | DocxOpened | PptxOpened | XlsxOpened;

export type OpenAdapter = (
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
) => Promise<OpenedDoc>;

export function abortError(): Error {
  return new DOMException("Aborted", "AbortError");
}

/**
 * Cancellation from either side: a DOM `AbortError` (the baselines) or aared's
 * typed abort, which carries the same name. Same rule as aared's `isAbort`,
 * kept local so baseline adapters never import the code under measurement.
 */
export function isAbort(error: unknown): boolean {
  return error instanceof Error && error.name === "AbortError";
}
