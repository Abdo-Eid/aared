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
 * one-format consumers (`import { open } from "aared/xlsx"`) never pay for
 * the formats they don't use — bundlers only include the format chunks the
 * dynamic `import()`s below resolve to, and only the detected one loads at
 * runtime.
 *
 * `preview()` covers open → render → navigate → destroy. For search, range
 * reads, media access, or per-unit control, build on the headless format
 * entrypoints directly.
 */

import type {
  Appearance,
  DocumentFormat,
  DocumentInfo,
  FontResolver,
  MediaResolver,
  Source,
  ZoomValue,
} from "../core/index.js";
import type { SoftBreakPolicy } from "../core/docx/pagination.js";
import type { PptxDocument } from "../pptx/index.js";
import type { XlsxDocument } from "../xlsx/index.js";
import type { DocxDocument } from "../docx/index.js";
import { AaredError } from "../core/index.js";
import { sourceBytes, WorkerSession } from "../core/session.js";
import { throwIfAborted } from "../core/zip.js";

type AnyDocument = PptxDocument | XlsxDocument | DocxDocument;

interface RenderHandle extends EventTarget {
  destroy(): void;
  setZoom(zoom: ZoomValue): void;
  readonly zoom: ZoomValue;
}

/** DOCX-only knobs passed through to `aared/docx` rendering. */
export interface PreviewDocxOptions {
  pagination?: { mode?: "flow" | "layout"; softBreaks?: SoftBreakPolicy };
}

/** Options for `preview()`. Everything is optional except `source`/`target`. */
export interface PreviewOptions {
  /**
   * Assert a format and skip detection. A mismatch rejects with a typed
   * `unsupported-format` `AaredError` naming the detected format — the same
   * error the per-format `open()`s raise.
   */
  type?: DocumentFormat;
  /** Zero-based initial unit index (slide/sheet; DOCX renders as one unit). */
  index?: number;
  signal?: AbortSignal;
  /** Opt-in font resolution, forwarded to `open()` (see `aared-fonts`). */
  fontResolver?: FontResolver;
  /** Viewer appearance, forwarded to `render()`. */
  appearance?: Appearance;
  /** Override the default session-backed media resolution. */
  mediaResolver?: MediaResolver;
  /** DOCX-only rendering knobs (e.g. `pagination: { mode: "layout" }`). */
  docx?: PreviewDocxOptions;
}

/** Controller returned by `preview()`: navigation, zoom, events, teardown. */
export interface PreviewHandle extends EventTarget {
  /** The detected (or asserted) format. */
  readonly format: DocumentFormat;
  /** Document metadata resolved at `open()` time. */
  readonly info: DocumentInfo;
  /** The currently rendered unit index. */
  readonly index: number;
  /** The active zoom, forwarded from the underlying render handle. */
  readonly zoom: ZoomValue;
  /**
   * Render another unit, replacing the current output. Stale navigations
   * lose: only the newest `goTo()` paints. Rejects with a `not-found`
   * `AaredError` outside `0 <= index < info.count`. Emits `indexchange`
   * plus the format-native sibling (`slidechange` for PPTX,
   * `sheetchange` for XLSX), mirroring `<aared-viewer>`.
   */
  goTo(index: number): Promise<void>;
  /** Change zoom; emits `zoomchange` (forwarded from the render handle). */
  setZoom(zoom: ZoomValue): void;
  /**
   * Full teardown: destroy the render output, terminate the Worker session,
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
  options?: { signal?: AbortSignal },
): Promise<DocumentFormat> {
  const bytes = await sourceBytes("aared/preview", source, options?.signal);
  throwIfAborted(options?.signal);
  return WorkerSession.detectFormat("aared/preview", bytes, options?.signal);
}

async function openFormat(
  format: DocumentFormat,
  bytes: ArrayBuffer,
  options?: PreviewOptions,
): Promise<AnyDocument> {
  const signal = options?.signal;
  const openOptions = options?.fontResolver === undefined && signal === undefined
    ? undefined
    : {
      ...(options?.fontResolver === undefined ? {} : { fontResolver: options.fontResolver }),
      ...(signal === undefined ? {} : { signal }),
    };
  if (format === "pptx") {
    const pptx = await import("../pptx/index.js");
    return pptx.open(bytes, openOptions);
  }
  if (format === "xlsx") {
    const xlsx = await import("../xlsx/index.js");
    return xlsx.open(bytes, openOptions);
  }
  const docx = await import("../docx/index.js");
  return docx.open(bytes, openOptions);
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
  #rendered: RenderHandle | null = null;
  #forwardController: AbortController | null = null;

  constructor(
    readonly format: DocumentFormat,
    private readonly document: AnyDocument,
    private readonly target: HTMLElement,
    private readonly options: PreviewOptions | undefined,
  ) {
    super();
    this.#index = options?.index ?? 0;
  }

  get info(): DocumentInfo {
    return this.document.info;
  }

  get index(): number {
    return this.#index;
  }

  get zoom(): ZoomValue {
    return this.#rendered?.zoom ?? "fit-width";
  }

  /**
   * Render the initial unit; called once before the handle is returned. No
   * `indexchange` fires here: listeners can only attach after `preview()`
   * resolves, so the initial unit (`options.index ?? 0`, readable via
   * `.index`) is announced by return, not by event. `goTo()` navigations
   * always emit `indexchange` and the format-native sibling.
   */
  async mount(signal?: AbortSignal): Promise<void> {
    assertIndex(this.#index, this.document.info.count);
    throwIfAborted(signal);
    await this.renderUnit(this.#index);
  }

  async goTo(index: number): Promise<void> {
    this.assertUsable();
    assertIndex(index, this.document.info.count);
    if (index === this.#index && this.#rendered !== null) return;
    await this.renderUnit(index);
    if (this.#destroyed) return;
    this.#index = index;
    this.announce(index);
  }

  setZoom(zoom: ZoomValue): void {
    this.assertUsable();
    this.#rendered?.setZoom(zoom);
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    this.#generation++;
    this.#forwardController?.abort();
    this.#forwardController = null;
    this.#rendered?.destroy();
    this.#rendered = null;
    this.document.close();
    this.target.replaceChildren();
  }

  private assertUsable(): void {
    if (this.#destroyed) throw new AaredError("aborted", "aared/preview: this preview is destroyed");
  }

  /** `indexchange` always fires; format-native siblings fire when they apply. */
  private announce(index: number): void {
    this.dispatchEvent(new CustomEvent("indexchange", { detail: { index } }));
    if (this.format === "pptx") {
      this.dispatchEvent(new CustomEvent("slidechange", { detail: { slide: index } }));
    } else if (this.format === "xlsx") {
      const name = (this.document as XlsxDocument).sheets.list()[index]?.info.name;
      this.dispatchEvent(new CustomEvent("sheetchange", { detail: { sheet: name ?? index } }));
    }
  }

  private async renderUnit(index: number): Promise<void> {
    const generation = ++this.#generation;
    const stale = (): boolean => this.#destroyed || generation !== this.#generation;
    const explicitMedia = this.options?.mediaResolver;
    const appearance = this.options?.appearance;
    let next: RenderHandle;
    if (this.format === "pptx") {
      const renderer = await import("../pptx/index.js");
      if (stale()) return;
      const slide = (this.document as PptxDocument).slides.get(index);
      const mediaResolver: MediaResolver = explicitMedia ?? ((mediaId: string) => slide.media(mediaId));
      next = await renderer.render({
        slide,
        target: this.target,
        ...(appearance === undefined ? {} : { appearance }),
        mediaResolver,
      });
    } else if (this.format === "xlsx") {
      const renderer = await import("../xlsx/index.js");
      if (stale()) return;
      const sheet = (this.document as XlsxDocument).sheets.get(index);
      const mediaResolver: MediaResolver = explicitMedia ?? ((mediaId: string) => sheet.media(mediaId));
      next = await renderer.render({
        sheet,
        target: this.target,
        ...(appearance === undefined ? {} : { appearance }),
        mediaResolver,
      });
    } else {
      const renderer = await import("../docx/index.js");
      if (stale()) return;
      const document = this.document as DocxDocument;
      const pagination = this.options?.docx?.pagination;
      // Passing the document (not bare sections) defaults mediaResolver to
      // the owning session — the same rule as PPTX/XLSX handle inputs.
      next = await renderer.render({
        document,
        target: this.target,
        ...(explicitMedia === undefined ? {} : { mediaResolver: explicitMedia }),
        ...(pagination === undefined ? {} : { pagination }),
        ...(appearance === undefined ? {} : { appearance }),
      });
    }
    if (stale()) {
      next.destroy();
      return;
    }
    this.#rendered?.destroy();
    this.#rendered = next;
    this.forwardRenderEvents(next);
  }

  private forwardRenderEvents(handle: RenderHandle): void {
    this.#forwardController?.abort();
    const controller = new AbortController();
    this.#forwardController = controller;
    const signal = controller.signal;
    for (const type of ["zoomchange", "linkactivate", "cellactivate", "selectionchange", "sheetchange"]) {
      handle.addEventListener(type, (event) => {
        this.dispatchEvent(new CustomEvent(type, { detail: (event as CustomEvent).detail }));
      }, { signal });
    }
  }
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
  const format = options?.type ?? await WorkerSession.detectFormat("aared/preview", bytes, signal);
  throwIfAborted(signal);
  const document = await openFormat(format, bytes, options);
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
