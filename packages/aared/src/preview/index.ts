/**
 * `aared/preview`: one-call auto-detect preview for embeds.
 *
 * `preview(source, target)` detects the format in the Worker, lazily
 * `import()`s only the matching format entrypoint, opens, renders the first
 * unit, and returns a `PreviewHandle` owning navigation, zoom, events, and
 * full teardown (`destroy()` closes the Worker session and clears what was
 * painted). It is a thin delegator over `aared/pptx`, `aared/xlsx`, and
 * `aared/docx` — every format route still parses in the Worker and paints
 * real DOM/SVG on the main thread; nothing here re-implements a renderer.
 *
 * Import-time safe: no DOM or `Worker` globals are touched until `preview()`
 * or `detect()` runs (SSR-safe imports). Live in a separate subpath so
 * one-format consumers (`import { open } from "aared/xlsx"`) never include
 * this facade — bundlers emit the format chunks the dynamic `import()`s
 * below resolve to, and only the detected one loads at runtime. (The Worker
 * script carries all three parsers regardless.)
 *
 * The grammar is the format entrypoints': `goTo(i)` switches units (slides,
 * sheets; a document is one unit), and per-format controls move within the
 * unit on screen (`xlsx.goToCell`, `docx.goToHeading`/`goToPage`). For
 * search, range reads, media access, or several units at once, build on the
 * headless format entrypoints directly.
 */

import type {
  Appearance,
  CellRef,
  DocumentFormat,
  DocumentInfo,
  FontResolver,
  MediaResolver,
  OpenOptions,
  Source,
  WorkerInit,
  ZoomValue,
} from "../core/index.js";
import type { PptxDocument } from "../pptx/index.js";
import type { XlsxDocument, XlsxRenderHandle } from "../xlsx/index.js";
import type { DocxDocument, DocxPageInfo, DocxPagination, DocxRenderHandle, HeadingInfo } from "../docx/index.js";
import { AaredError } from "../core/index.js";
import { sourceBytes, WorkerSession } from "../core/session.js";
import { throwIfAborted } from "../core/errors.js";
import { sameZoom } from "../render-handle.js";
import { DEFAULT_ZOOM, announceUnit, followSheet, forwardRenderEvents, openFormat } from "../format-route.js";
import type { AnyDocument, RenderHandle } from "../format-route.js";

export type { DocxPageInfo, DocxPagination, HeadingInfo, SoftBreakPolicy } from "../docx/index.js";

/**
 * Moving within the sheet on screen — the places of an XLSX unit. Switching
 * sheets is the facade's own `goTo(index)`.
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
}

/** Options for `preview()`. Everything is optional except `source`/`target`. */
export interface PreviewOptions {
  /**
   * Assert a format and skip detection. A mismatch rejects with a typed
   * `unsupported-format` `AaredError` naming the detected format — the same
   * error the per-format `open()`s raise.
   */
  type?: DocumentFormat;
  /** Zero-based initial unit index (slide/sheet; a document is one unit). */
  index?: number;
  signal?: AbortSignal;
  /** Opt-in font resolution, forwarded to `open()` (see `aared-fonts`). */
  fontResolver?: FontResolver;
  /** Override worker construction (see `WorkerInit` for details). */
  worker?: WorkerInit;
  /** Viewer appearance, forwarded to `render()`. */
  appearance?: Appearance;
  /** Starting zoom; default per format (`fit-page` slides, `1` sheets, `fit-width` documents). */
  zoom?: ZoomValue;
  /** Override the default session-backed media resolution. */
  mediaResolver?: MediaResolver;
  /** DOCX layout (flow or physical pages), forwarded to `render()`; ignored for other formats. */
  pagination?: DocxPagination;
}

/** Controller returned by `preview()`: navigation, zoom, events, teardown. */
export interface PreviewHandle extends EventTarget {
  /** The detected (or asserted) format. */
  readonly format: DocumentFormat;
  /** Document metadata resolved at `open()` time. */
  readonly info: DocumentInfo;
  /** The currently rendered unit index. */
  readonly index: number;
  /** The requested zoom. Carried across `goTo()`. */
  readonly zoom: ZoomValue;
  /** The scale actually painted — the number, or what the fit mode computed. */
  readonly scale: number;
  /**
   * Render another unit, replacing the current output. Stale navigations
   * lose: only the newest `goTo()` paints. Rejects with a `not-found`
   * `AaredError` outside `0 <= index < info.count`. Emits `indexchange`
   * plus the format-native sibling (`slidechange` for PPTX, `sheetchange`
   * for XLSX, both `{ index, title }`); going to the unit already shown
   * emits nothing. Zoom and appearance carry over.
   */
  goTo(index: number): Promise<void>;
  /** Change zoom; fires `zoomchange` only when the value changes. Kept across `goTo()`. */
  setZoom(zoom: ZoomValue): void;
  /** Moving within the sheet on screen; `null` for other formats and before the first render. */
  readonly xlsx: XlsxPreviewControls | null;
  /** Moving within the document on screen; `null` for other formats and before the first render. */
  readonly docx: DocxPreviewControls | null;
  /**
   * Replace the viewer appearance in place without re-opening the document:
   * forwards to the current render handle, keeping the unit index, zoom,
   * selection, and Worker session, and applies to every later `goTo()`.
   * Replaces, never merges. Content stays source-faithful; only the theme
   * marker and viewer chrome follow the theme.
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

/** The `open()` options a preview forwards: only what the caller set. */
function openOptionsOf(options: PreviewOptions | undefined): OpenOptions | undefined {
  const signal = options?.signal;
  const worker = options?.worker;
  const fontResolver = options?.fontResolver;
  if (fontResolver === undefined && signal === undefined && worker === undefined) return undefined;
  return {
    ...(fontResolver === undefined ? {} : { fontResolver }),
    ...(signal === undefined ? {} : { signal }),
    ...(worker === undefined ? {} : { worker }),
  };
}

function assertIndex(index: number, count: number): void {
  if (!Number.isInteger(index) || index < 0 || index >= count) {
    throw new AaredError(
      "not-found",
      `aared/preview: unit index ${String(index)} is outside this ${count}-unit document`,
    );
  }
}

class PreviewViewer extends EventTarget implements PreviewHandle {
  #destroyed = false;
  #generation = 0;
  #index: number;
  #zoom: ZoomValue | undefined;
  #rendered: RenderHandle | null = null;
  #xlsx: XlsxPreviewControls | null = null;
  #docx: DocxPreviewControls | null = null;
  #forwardController: AbortController | null = null;

  constructor(
    readonly format: DocumentFormat,
    private readonly document: AnyDocument,
    private readonly target: HTMLElement,
    private options: PreviewOptions | undefined,
  ) {
    super();
    this.#index = options?.index ?? 0;
    this.#zoom = options?.zoom;
  }

  get info(): DocumentInfo {
    return this.document.info;
  }

  get index(): number {
    return this.#index;
  }

  get zoom(): ZoomValue {
    return this.#zoom ?? DEFAULT_ZOOM[this.format];
  }

  get scale(): number {
    return this.#rendered?.scale ?? 1;
  }

  get xlsx(): XlsxPreviewControls | null {
    return this.#xlsx;
  }

  get docx(): DocxPreviewControls | null {
    return this.#docx;
  }

  /**
   * Render the initial unit; called once before the handle is returned. No
   * `indexchange` fires here: the starting state is reported by `preview()`
   * resolving (read `.index`), never by an event.
   */
  async mount(signal?: AbortSignal): Promise<void> {
    assertIndex(this.#index, this.document.info.count);
    throwIfAborted(signal);
    await this.renderUnit(this.#index);
  }

  async goTo(index: number): Promise<void> {
    if (this.#destroyed) return;
    assertIndex(index, this.document.info.count);
    if (index === this.#index && this.#rendered !== null) return;
    await this.renderUnit(index);
    if (this.#destroyed) return;
    this.#index = index;
    announceUnit(this, this.format, this.document, index);
  }

  setZoom(zoom: ZoomValue): void {
    if (this.#destroyed) return;
    this.#zoom = zoom;
    this.#rendered?.setZoom(zoom);
  }

  setAppearance(appearance: Appearance): void {
    if (this.#destroyed) return;
    if (this.options === undefined) this.options = { appearance };
    else this.options = { ...this.options, appearance };
    this.#rendered?.setAppearance(appearance);
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    this.#generation++;
    this.#forwardController?.abort();
    this.#forwardController = null;
    this.#rendered?.destroy();
    this.#rendered = null;
    this.#xlsx = null;
    this.#docx = null;
    this.document.close();
    this.target.replaceChildren();
  }

  private async renderUnit(index: number): Promise<void> {
    const generation = ++this.#generation;
    const stale = (): boolean => this.#destroyed || generation !== this.#generation;
    const { appearance, mediaResolver } = this.options ?? {};
    // Only what the caller set: each format's `render()` owns its defaults
    // (session media — PPTX routes shared layout pictures through one cache —
    // and its default zoom).
    const common = {
      target: this.target,
      ...(appearance === undefined ? {} : { appearance }),
      ...(mediaResolver === undefined ? {} : { mediaResolver }),
      ...(this.#zoom === undefined ? {} : { zoom: this.#zoom }),
    };
    let next: RenderHandle;
    if (this.format === "pptx") {
      const renderer = await import("../pptx/index.js");
      if (stale()) return;
      next = await renderer.render({ ...common, slide: (this.document as PptxDocument).slides.get(index) });
    } else if (this.format === "xlsx") {
      const renderer = await import("../xlsx/index.js");
      if (stale()) return;
      next = await renderer.render({ ...common, sheet: (this.document as XlsxDocument).sheets.get(index) });
    } else {
      const renderer = await import("../docx/index.js");
      if (stale()) return;
      const pagination = this.options?.pagination;
      next = await renderer.render({
        ...common,
        document: this.document as DocxDocument,
        ...(pagination === undefined ? {} : { pagination }),
      });
    }
    if (stale()) {
      next.destroy();
      return;
    }
    // A zoom set while this unit was rendering was announced by the old
    // handle; apply it before forwarding is attached so it isn't announced twice.
    if (this.#zoom !== undefined && !sameZoom(next.zoom, this.#zoom)) next.setZoom(this.#zoom);
    this.#rendered?.destroy();
    this.#rendered = next;
    this.#xlsx = this.format === "xlsx" ? xlsxControls(next as unknown as XlsxRenderHandle) : null;
    this.#docx = this.format === "docx" ? docxControls(next as unknown as DocxRenderHandle, this.document as DocxDocument) : null;
    this.#forwardController?.abort();
    this.#forwardController = new AbortController();
    forwardRenderEvents(next, this, this.#forwardController.signal);
    // A sheet-tab click switches sheets inside the handle: follow it, so
    // `index` stays true and the switch is announced like any `goTo()`.
    followSheet(next, this.#forwardController.signal, (sheet) => {
      if (this.#destroyed || sheet === this.#index) return;
      this.#index = sheet;
      announceUnit(this, this.format, this.document, sheet);
    });
  }
}

/** One stable controls object per rendered sheet, delegating to its handle. */
function xlsxControls(handle: XlsxRenderHandle): XlsxPreviewControls {
  return {
    get selection() {
      return handle.selection;
    },
    goToCell: (cell) => handle.goToCell(cell),
    scrollTo: (target) => handle.scrollTo(target),
    focus: (cell) => handle.focus(cell),
  };
}

/** One stable controls object per rendered document, delegating to its handle. */
function docxControls(handle: DocxRenderHandle, document: DocxDocument): DocxPreviewControls {
  return {
    pageCount: handle.pageCount,
    pages: handle.pages,
    headings: document.headings.list(),
    goToHeading: (target) => handle.goToHeading(target),
    goToPage: (index) => handle.goToPage(index),
  };
}

/**
 * Preview any supported document in one call: detect (or assert `type`),
 * open, and render into `target`. The target is owned by the preview while
 * mounted — every `goTo()` replaces its contents, and `destroy()` clears
 * them. One format chunk loads at runtime: the detected one.
 */
export async function preview(
  source: Source,
  target: HTMLElement,
  options?: PreviewOptions,
): Promise<PreviewHandle> {
  const signal = options?.signal;
  const bytes = await sourceBytes("aared/preview", source, signal);
  throwIfAborted(signal);
  // An asserted type skips detection: the per-format `open()` below raises
  // the typed `unsupported-format` mismatch itself, naming what it found.
  const format = options?.type ?? await WorkerSession.detectFormat("aared/preview", bytes, {
    ...(signal === undefined ? {} : { signal }),
    ...(options?.worker === undefined ? {} : { worker: options.worker }),
  });
  throwIfAborted(signal);
  const document = await openFormat(format, bytes, openOptionsOf(options));
  const viewer: PreviewViewer = new PreviewViewer(format, document, target, options);
  try {
    if (document.info.count === 0) {
      throw new AaredError(
        "invalid-document",
        "aared/preview: this document has no slides, sheets, or sections to show",
      );
    }
    await viewer.mount(signal);
    if (signal?.aborted === true) throw new AaredError("aborted", "aared/preview: request was cancelled");
    return viewer;
  } catch (error) {
    viewer.destroy();
    throw error;
  }
}
