/**
 * `aared/preview`: one-call auto-detect preview for embeds.
 *
 * `preview({ source, target })` detects the format in the Worker, lazily
 * `import()`s only the matching format entrypoint, opens, renders the first
 * unit, and returns a `PreviewHandle` owning navigation, zoom, events, and
 * full teardown (`destroy()` closes the Worker session and clears what was
 * painted). The handle is the embed viewer (`embed-viewer.ts`) that
 * `<aared-viewer>` and the React hooks run on too; this entry declares its
 * public types. Options set where it starts; the handle's methods change it.
 *
 * Import-time safe: no DOM or `Worker` globals are touched until `preview()`
 * or `detect()` runs (SSR-safe imports). Live in a separate subpath so
 * one-format consumers (`import { open } from "aared/xlsx"`) never include
 * this facade — bundlers emit the format chunks the viewer's dynamic
 * `import()`s resolve to, and only the detected one loads at runtime. (The
 * Worker script carries all three parsers regardless.)
 *
 * The grammar is the format entrypoints': `goTo(i)` switches units (slides,
 * sheets; a document is one unit), and per-format controls move within the
 * unit on screen (`xlsx.goToCell`, `docx.goToHeading`/`goToPage`). `search()`
 * finds places to go in the viewer's own document. To build from the file's
 * contents (range reads, media access, several units at once), use the
 * headless format entrypoints directly.
 */

import type {
  AaredError,
  Appearance,
  CellRef,
  DocumentFormat,
  FontResolver,
  MediaResolver,
  ReadOptions,
  Source,
  UnitRef,
  WorkerInit,
  ZoomValue,
} from "../core/index.js";
import type { PptxInfo, PptxRenderEventMap, PptxSearchMatch } from "../pptx/index.js";
import type { XlsxInfo, XlsxRenderEventMap, XlsxSearchMatch } from "../xlsx/index.js";
import type { DocxInfo, DocxPageInfo, DocxPagination, DocxRenderEventMap, DocxSearchMatch, HeadingInfo } from "../docx/index.js";
import { sourceBytes, WorkerSession } from "../core/session.js";
import { throwIfAborted } from "../core/errors.js";
import type { EventTargetOf, ZoomChangeEvent } from "../render-handle.js";
import { openAutoViewer } from "../auto-viewer.js";

export type { DocxPageInfo, DocxPagination, DocxSearchMatch, HeadingInfo, SoftBreakPolicy } from "../docx/index.js";
export type { PptxSearchMatch } from "../pptx/index.js";
export type { XlsxSearchMatch } from "../xlsx/index.js";

/**
 * Moving within the sheet on screen — the places of an XLSX unit. Switching
 * sheets is the facade's own `goTo(index)`. One object for the viewer's whole
 * life, acting on whichever sheet is on screen, so a reference kept across
 * `goTo()` keeps working; after `destroy()` `goToCell` rejects with `closed`
 * and the rest do nothing.
 */
export interface XlsxPreviewControls {
  /** The single selected cell, or `null`. */
  readonly selection: CellRef | null;
  /** Select one cell, scroll it into view if needed, and focus it. See `XlsxRenderHandle.goToCell`. */
  goToCell(cell: string): Promise<void>;
  /** Scroll a cell or range into view without selecting. See `XlsxRenderHandle.scrollTo`. */
  scrollTo(target: { cell?: string; range?: string; align?: "start" | "center" | "end" }): void;
  /** Select one cell and focus it without scrolling. */
  focus(cell: string): void;
}

/**
 * Moving within the document on screen — the places of DOCX's single unit:
 * its body headings (both modes) and its laid-out pages (layout mode only).
 * One object for the viewer's whole life; after `destroy()` its moves reject
 * with `closed`.
 */
export interface DocxPreviewControls {
  /** Laid-out pages: `0` and `[]` in flow mode, which has no pages. */
  readonly pageCount: number;
  readonly pages: readonly DocxPageInfo[];
  /** The body heading outline, in document order (`document.headings.list()`). */
  readonly headings: readonly HeadingInfo[];
  /** Scroll a heading into view by index or title. See `DocxRenderHandle.goToHeading`. */
  goToHeading(target: string | number): Promise<void>;
  /** Scroll a laid-out page into view (layout mode). See `DocxRenderHandle.goToPage`. */
  goToPage(index: number): Promise<void>;
  /** Navigate to a result; active search highlights remain in place. */
  goToSearchMatch(match: DocxSearchMatch): Promise<void>;
}

/**
 * Optional settings for `preview()`; `PreviewInput` adds its source and target.
 * `index`, `zoom` and `appearance` are where the viewer starts; the handle's
 * `goTo()`, `setZoom()` and `setAppearance()` change them.
 */
export interface PreviewOptions {
  /**
   * Assert a format and skip detection. A mismatch rejects with a typed
   * `unsupported-format` `AaredError` naming the detected format — the same
   * error the per-format `open()`s raise.
   */
  type?: DocumentFormat;
  /** Zero-based starting unit index (slide/sheet; a document is one unit). */
  index?: number;
  signal?: AbortSignal;
  /** Opt-in font resolution, forwarded to `open()` (see `aared-fonts`). */
  fontResolver?: FontResolver;
  /** Override worker construction (see `WorkerInit` for details). */
  worker?: WorkerInit;
  /**
   * Starting appearance, forwarded to `render()`. What it leaves out takes
   * the embed default: a workbook shows its sheet tabs unless
   * `spreadsheet.sheetTabs` is `false`.
   */
  appearance?: Appearance;
  /** Starting zoom; default per format (`fit-page` slides, `1` sheets, `fit-width` documents). */
  zoom?: ZoomValue;
  /** Override the default session-backed media resolution. */
  mediaResolver?: MediaResolver;
  /** DOCX layout (flow or physical pages), forwarded to `render()`; ignored for other formats. */
  pagination?: DocxPagination;
}

/** One object describing the file, render target and optional preview settings. */
export interface PreviewInput extends PreviewOptions {
  source: Source;
  target: HTMLElement;
}

/**
 * What `preview()` returns: one handle type per format, told apart by
 * `format`. Checking `viewer.format === "xlsx"` makes `viewer.xlsx`
 * non-null and `viewer.info` an `XlsxInfo`; the other formats' controls are
 * `null`.
 */
export type PreviewHandle = PptxPreview | XlsxPreview | DocxPreview;

/** Viewer search updates highlights; document and sheet search remain data-only. */
export interface PreviewSearchOptions extends ReadOptions {
  /** Highlight matching painted text. Defaults to true; false clears existing highlights. */
  highlight?: boolean;
}

/** Live find position: zero-based active index, or null when no match is active. */
export interface SearchNavigationState {
  readonly index: number | null;
  readonly count: number;
}

/** The current highlighted search's navigation; steps wrap at either end. */
export interface SearchNavigation extends SearchNavigationState {
  next(): Promise<void>;
  previous(): Promise<void>;
}

/** A presentation's preview: `format` `"pptx"`. A slide has no places inside it, so there is no `pptx` controls object. */
export interface PptxPreview extends PreviewBase {
  readonly format: "pptx";
  readonly info: PptxInfo;
  readonly xlsx: null;
  readonly docx: null;
  /**
   * Search the deck the viewer has open (no second parse): one result for each
   * text occurrence, a place to go with `goTo(match.slide.index)`. The same search as
   * the format's `deck.search()` (the API page's Search grammar); `signal` cancels it.
   * Highlights painted matches by default; `highlight: false` returns data only
   * and clears previous highlights. Empty query clears. The query stays active
   * across navigation and text redraws. A failure rejects and fires no `error`
   * (that event is for moves); after
   * `destroy()` it rejects with `closed`.
   */
  search(query: string, options?: PreviewSearchOptions): Promise<PptxSearchMatch[]>;
}

/** A workbook's preview: `format` `"xlsx"`, with the sheet on screen's places on `xlsx`. */
export interface XlsxPreview extends PreviewBase {
  readonly format: "xlsx";
  readonly info: XlsxInfo;
  /** Moving within the sheet on screen. See `XlsxPreviewControls`. */
  readonly xlsx: XlsxPreviewControls;
  readonly docx: null;
  /**
   * Search the sheet on screen, as Excel's Find does: each match is a cell
   * of that sheet, reached with `xlsx.goToCell(match.address)`. Another
   * sheet is searched by `goTo()`ing it first. The same rules as
   * `PptxPreview.search()`.
   */
  search(query: string, options?: PreviewSearchOptions): Promise<XlsxSearchMatch[]>;
}

/** A document's preview: `format` `"docx"`, with its headings and pages on `docx`. */
export interface DocxPreview extends PreviewBase {
  readonly format: "docx";
  readonly info: DocxInfo;
  readonly xlsx: null;
  /** Moving within the document. See `DocxPreviewControls`. */
  readonly docx: DocxPreviewControls;
  /**
   * Search the document the viewer has open: each match names the nearest
   * heading above it. `docx.goToSearchMatch(match)` navigates to its story
   * copy or nearest body heading.
   * The same rules as `PptxPreview.search()`.
   */
  search(query: string, options?: PreviewSearchOptions): Promise<DocxSearchMatch[]>;
}

/**
 * The events a preview handle fires, by name: `addEventListener` types each
 * `event.detail`. The render handle's own events are forwarded; the unit
 * events and `error` are the facade's.
 */
export interface PreviewEventMap {
  /** Active highlighted search position or total changed. */
  searchchange: CustomEvent<SearchNavigationState>;
  /** The unit changed (`goTo()`, an XLSX sheet tab). */
  indexchange: CustomEvent<{ index: number }>;
  /** PPTX, alongside `indexchange`. */
  slidechange: CustomEvent<UnitRef>;
  /** XLSX, alongside `indexchange`. */
  sheetchange: CustomEvent<UnitRef>;
  zoomchange: ZoomChangeEvent;
  /** A navigation failed (`goTo()`, or a move on `xlsx`/`docx`); `cause` is what the call rejected with. */
  error: CustomEvent<{ message: string; cause: AaredError }>;
  linkactivate: PptxRenderEventMap["linkactivate"] | XlsxRenderEventMap["linkactivate"] | DocxRenderEventMap["linkactivate"];
  /** XLSX. */
  cellactivate: XlsxRenderEventMap["cellactivate"];
  /** XLSX. */
  selectionchange: XlsxRenderEventMap["selectionchange"];
}

/** What every format's preview handle has: navigation, zoom, events, teardown. */
interface PreviewBase extends EventTargetOf<PreviewEventMap> {
  /** Counter and previous/next navigation for the active highlighted search. */
  readonly searchNavigation: SearchNavigation;
  /** The currently rendered unit index. */
  readonly index: number;
  /** The requested zoom. Carried across `goTo()`. */
  readonly zoom: ZoomValue;
  /**
   * The zoom this format opens at when no `zoom` option is given:
   * `"fit-page"` slides, `1` sheets, `"fit-width"` documents. `resetZoom()`
   * goes back to it. For a slide or a document it is not `1`: `setZoom(1)`
   * pins the natural size, which can overflow a small host.
   */
  readonly defaultZoom: ZoomValue;
  /** The scale actually painted — the number, or what the fit mode computed. */
  readonly scale: number;
  /**
   * Render another unit, replacing the current output. Stale navigations
   * lose: only the newest `goTo()` paints. Rejects with a `not-found`
   * `AaredError` outside `0 <= index < info.count`. Emits `indexchange`
   * plus the format-native sibling (`slidechange` for PPTX, `sheetchange`
   * for XLSX, both `{ index, title }`); going to the unit already shown
   * emits nothing. Zoom and appearance carry over. A failure rejects and
   * also fires `error` (`{ message, cause }`), so whoever holds the handle
   * hears it even when the caller drops the promise; the unit on screen
   * stays. After `destroy()` it rejects with `closed`.
   */
  goTo(index: number): Promise<void>;
  /**
   * The next unit: `goTo()` one step on from where the newest move is
   * heading, so three quick presses move three units, not one. At the last
   * unit it changes nothing and resolves; a document is one unit, so there it
   * always does. Rejects and fires `error` like `goTo()`; after `destroy()`
   * it rejects with `closed`.
   */
  next(): Promise<void>;
  /** The previous unit, the same way as `next()`; at the first unit it changes nothing. */
  previous(): Promise<void>;
  /** Change zoom; fires `zoomchange` only when the value changes. Kept across `goTo()`. */
  setZoom(zoom: ZoomValue): void;
  /**
   * Zoom in one step: `step` added to the zoom, where `1` is 100% (default
   * `0.1`, so each press adds 10 points: 100% to 110% to 120%). From a fit
   * mode it steps from the scale on screen (`scale`), so a slide fitted at
   * 62% goes to 72%, not 110%. The result is rounded to whole percents, so
   * repeated presses don't drift, and stops at 300%. A `step` that is not a
   * number above 0 throws `invalid-argument`. After `destroy()` it does
   * nothing.
   */
  zoomIn(step?: number): void;
  /**
   * Zoom out one step: `step` taken away (default `0.1`), so a zoom in and a
   * zoom out by the same step come back to the same size. Stops at 25%; the
   * same `step` rules as `zoomIn()`.
   */
  zoomOut(step?: number): void;
  /** Back to `defaultZoom`, the zoom the format opens at. */
  resetZoom(): void;
  /**
   * Replace the viewer appearance in place without re-opening the document:
   * forwards to the current render handle, keeping the unit index, zoom,
   * selection, and Worker session, and applies to every later `goTo()`.
   * Replaces, never merges: what it leaves out goes back to the default (a
   * workbook's sheet tabs stay on unless `spreadsheet.sheetTabs` is `false`).
   * Content stays source-faithful; only the theme marker and viewer chrome
   * follow the theme.
   */
  setAppearance(appearance: Appearance): void;
  /**
   * Full teardown: destroy the render output, end the Worker session,
   * and clear what was painted into the target. The handle is inert after.
   * The target element itself stays in the caller's document.
   */
  destroy(): void;
}

/**
 * Detect a document's format without opening it: reads content types only in
 * the Worker and keeps no session. Headless-safe (no DOM).
 */
export async function detect(
  source: Source,
  options?: { signal?: AbortSignal; worker?: WorkerInit },
): Promise<DocumentFormat> {
  const bytes = await sourceBytes("aared/preview", source, options?.signal);
  throwIfAborted(options?.signal);
  return WorkerSession.detectFormat("aared/preview", bytes, options);
}

/**
 * Preview any supported document in one call: detect (or assert `type`),
 * open, and render into `target`. The target is owned by the preview while
 * mounted — every `goTo()` replaces its contents, and `destroy()` clears
 * them. One format chunk loads at runtime: the detected one.
 */
export async function preview(
  input: PreviewInput,
): Promise<PreviewHandle> {
  const { source, target, ...options } = input;
  return openAutoViewer({ scope: "aared/preview", source, target, options });
}
