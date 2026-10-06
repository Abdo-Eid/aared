/**
 * The viewer every embed layer runs on. `preview()` hands it out,
 * `<aared-viewer>` drives it from its attributes, and the React hooks mount it
 * through `preview()`, so opening, newest-wins navigation, zoom, the embed
 * appearance defaults, events and teardown are written once. Internal: not an
 * entry, so nothing here is public API; the handle's public types live in
 * `preview/index.ts`.
 *
 * It is a thin delegator over `aared/pptx`, `aared/xlsx` and `aared/docx`:
 * every format still parses in the Worker and paints real DOM/SVG on the main
 * thread; nothing here re-implements a renderer.
 */

import type { Appearance, DocumentFormat, DocumentInfo, OpenOptions, Source, ZoomValue } from "./core/index.js";
import type { PptxDocument, PptxSearchMatch } from "./pptx/index.js";
import type { XlsxDocument, XlsxRenderHandle, XlsxSearchMatch } from "./xlsx/index.js";
import type { DocxDocument, DocxPageInfo, DocxRenderHandle, DocxSearchMatch } from "./docx/index.js";
import type {
  DocxPreviewControls,
  PptxPreview,
  PreviewEventMap,
  PreviewHandle,
  PreviewOptions,
  PreviewSearchOptions,
  SearchNavigation,
  XlsxPreviewControls,
} from "./preview/index.js";
import { AaredError } from "./core/index.js";
import { sourceBytes } from "./core/session.js";
import { isAbort, throwIfAborted } from "./core/errors.js";
import { watchSystemTheme } from "./appearance.js";
import type { SearchHighlighter } from "./search-highlight.js";
import { DEFAULT_ZOOM, safeScale, sameZoom } from "./render-handle.js";
import {
  announceUnit,
  assertHasUnits,
  assertUnitIndex,
  embedAppearance,
  followSheet,
  forwardRenderEvents,
} from "./format-route.js";
import type { AnyDocument, RenderHandle, ViewerRoute } from "./format-route.js";

/** Listed as a record so a name added to `PreviewEventMap` and left out here fails to compile. */
const VIEWER_EVENT_NAMES: Readonly<Record<keyof PreviewEventMap, true>> = {
  searchchange: true,
  indexchange: true,
  slidechange: true,
  sheetchange: true,
  zoomchange: true,
  error: true,
  linkactivate: true,
  cellactivate: true,
  selectionchange: true,
};

/** Every event the viewer fires, for a layer that re-fires them on itself (the element). */
export const VIEWER_EVENTS = Object.keys(VIEWER_EVENT_NAMES) as ReadonlyArray<keyof PreviewEventMap>;

/** `zoomIn()`/`zoomOut()`: the default step (each press 10 points: 100% to 110%) and the range they stay in, 25%–300%. */
const ZOOM_STEP = 0.1;
const MIN_ZOOM = 0.25;
const MAX_ZOOM = 3;

/** What `openViewer()` takes. */
export interface OpenViewerInput {
  /** The public entry the caller used (`aared/preview`, `aared/element`): every message starts with it. */
  scope: string;
  source: Source;
  target: HTMLElement;
  options?: PreviewOptions | undefined;
  /** Fixed for a single-format entry, resolved lazily by the auto-detect layer. */
  route: ViewerRoute | ((bytes: ArrayBuffer, options: PreviewOptions | undefined) => Promise<ViewerRoute>);
  /** Bytes read before parsing starts (the element's `progress` event). */
  onProgress?: ((loaded: number, total: number | undefined) => void) | undefined;
}

/**
 * Resolve the supplied format route, open, and render the starting unit into
 * `target`; resolves with the handle once it is painted. On any failure the
 * viewer is destroyed (the target cleared, the Worker session ended) before
 * the rejection. `options.signal` cancels it.
 */
export async function openViewer({ scope, source, target, options, onProgress, route: providedRoute }: OpenViewerInput): Promise<PreviewHandle> {
  const signal = options?.signal;
  const bytes = await sourceBytes(scope, source, signal, onProgress);
  throwIfAborted(signal);
  const route = typeof providedRoute === "function" ? await providedRoute(bytes, options) : providedRoute;
  throwIfAborted(signal);
  const document = await route.open(bytes, openOptionsOf(options));
  const viewer = new PreviewViewer(scope, route.format, document, target, options, route);
  // A cancelled open tears the viewer down at once, not when its first render
  // settles: the render then paints nothing, and the target is cleared before
  // a newer open into it can paint (a late teardown would wipe that paint).
  const cancel = (): void => viewer.destroy();
  signal?.addEventListener("abort", cancel, { once: true });
  try {
    assertHasUnits(scope, document);
    await viewer.mount(signal);
    throwIfAborted(signal);
    // `format` decides the member: the viewer's controls and `info` match it.
    return viewer as unknown as PreviewHandle;
  } catch (error) {
    viewer.destroy();
    // A render the cancel cut short may have failed with `closed`; the caller asked to cancel.
    throwIfAborted(signal);
    throw error;
  } finally {
    signal?.removeEventListener("abort", cancel);
  }
}

/** The `open()` options a viewer forwards: only what the caller set. */
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

/**
 * What every format's handle has: `PreviewBase`, which `preview/index.ts`
 * keeps unexported. `search` is per format too (each returns its own
 * matches), so the class's version returns their union.
 */
type ViewerBase = Omit<PptxPreview, "format" | "info" | "xlsx" | "docx" | "search">;

/**
 * One class serves every format; `openViewer()` hands it out as the
 * `PreviewHandle` member its `format` names (`xlsx`/`docx` are non-null
 * exactly for their own format).
 */
class PreviewViewer extends EventTarget implements ViewerBase {
  #destroyed = false;
  #searchGeneration = 0;
  #highlighter: SearchHighlighter | null = null;
  #searchMatches: (PptxSearchMatch | XlsxSearchMatch | DocxSearchMatch)[] = [];
  #searchIndex: number | null = null;
  #searchMove = 0;
  #searchState = "";
  readonly searchNavigation: SearchNavigation;
  #generation = 0;
  #index: number;
  #zoom: ZoomValue | undefined;
  #rendered: RenderHandle | null = null;
  #forwardController: AbortController | null = null;
  /** Cancels the render in flight: a newer render or `destroy()` stops it before it paints. */
  #rendering: AbortController | null = null;
  /** The unit the newest `goTo()` is on its way to (`null` when none is), so `next()`/`previous()` step on from it. */
  #heading: { index: number } | null = null;
  /** Stops the OS colour-scheme listener; set only while the appearance asks for the `system` theme. */
  #stopSystemTheme: (() => void) | null = null;
  readonly xlsx: XlsxPreviewControls | null;
  readonly docx: DocxPreviewControls | null;

  constructor(
    private readonly scope: string,
    readonly format: DocumentFormat,
    private readonly document: AnyDocument,
    private readonly target: HTMLElement,
    private options: PreviewOptions | undefined,
    private readonly route: ViewerRoute,
  ) {
    super();
    this.#index = options?.index ?? 0;
    const owner = this;
    this.searchNavigation = {
      get index() { return owner.format === "docx" ? owner.#highlighter?.index ?? null : owner.#searchIndex; },
      get count() { return owner.format === "docx" ? owner.#highlighter?.count ?? 0 : owner.#searchMatches.length; },
      next: () => this.#stepSearch(1),
      previous: () => this.#stepSearch(-1),
    };
    this.#zoom = options?.zoom;
    this.#followSystemTheme();
    // Built once and handed the unit on screen at each call, so a reference
    // kept across `goTo()` keeps working (it used to be rebuilt per unit, and
    // an old one acted on a destroyed render).
    const report = (error: unknown): void => this.#fail(error);
    this.xlsx = format === "xlsx"
      ? xlsxControls(scope, () => this.#rendered as unknown as XlsxRenderHandle | null, report)
      : null;
    this.docx = format === "docx"
      ? docxControls(scope, () => this.#rendered as unknown as DocxRenderHandle | null, document as DocxDocument, report)
      : null;
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

  get defaultZoom(): ZoomValue {
    return DEFAULT_ZOOM[this.format];
  }

  get scale(): number {
    return this.#rendered?.scale ?? 1;
  }

  /**
   * Render the initial unit; called once before the handle is returned. No
   * `indexchange` fires here: the starting state is reported by the open
   * resolving (read `.index`), never by an event.
   */
  async mount(signal?: AbortSignal): Promise<void> {
    assertUnitIndex(this.scope, this.#index, this.document.info.count);
    throwIfAborted(signal);
    await this.renderUnit(this.#index);
  }

  async goTo(index: number): Promise<void> {
    // A call that can do nothing says so: resolving would tell the caller the unit changed.
    if (this.#destroyed) throw new AaredError("closed", `${this.scope}: goTo() was called after destroy()`);
    const heading = { index };
    try {
      assertUnitIndex(this.scope, index, this.document.info.count);
      // Already there, with nothing on its way: nothing to do. With a move on
      // its way the screen may already show that move's unit (a render paints
      // before it settles, and the unit it replaced is gone), so going back
      // renders this unit again: the newest call wins, and starting it
      // cancels the move in flight.
      if (index === this.#index && this.#rendered !== null && this.#rendering === null) return;
      this.#heading = heading;
      // Overtaken by a later goTo() or a destroy(): that call owns `index` and the announcement.
      if (!(await this.renderUnit(index))) return;
    } catch (error) {
      this.#fail(error);
      throw error;
    } finally {
      if (this.#heading === heading) this.#heading = null;
    }
    const changed = index !== this.#index;
    this.#index = index;
    if (this.format === "pptx" && this.#searchMatches.length > 0) {
      const current = this.#searchIndex === null ? undefined : this.#searchMatches[this.#searchIndex];
      if (current === undefined || !("slide" in current) || current.slide.index !== index) {
        const first = this.#searchMatches.findIndex((match) => "slide" in match && match.slide.index === index);
        this.#searchIndex = first < 0 ? null : first;
      }
      const local = this.#searchIndex === null ? -1 : this.#searchMatches.slice(0, this.#searchIndex).filter((match) => "slide" in match && match.slide.index === index).length;
      this.#highlighter?.activate(local);
    } else if (this.format === "xlsx" && changed) {
      this.#searchMatches = [];
      this.#searchIndex = null;
    }
    this.#announceSearch();
    // Back to the unit that was on screen (the move in flight cancelled): nothing changed to announce.
    if (changed) announceUnit(this, this.format, this.document, index);
  }

  /**
   * Search the document this viewer holds: no second parse. A workbook is
   * searched on the sheet on screen, as Excel's Find is. A failure only
   * rejects; the `error` event is for moves, and the hooks show it as
   * `navigationError`, which a failed search is not.
   */
  async search(query: string, options?: PreviewSearchOptions): Promise<PptxSearchMatch[] | XlsxSearchMatch[] | DocxSearchMatch[]> {
    if (this.#destroyed) throw new AaredError("closed", `${this.scope}: search() was called after destroy()`);
    const generation = ++this.#searchGeneration;
    throwIfAborted(options?.signal);
    const matches = this.format === "xlsx"
      ? await (this.document as XlsxDocument).sheets.get(this.#index).search(query, options)
      : await (this.document as PptxDocument | DocxDocument).search(query, options);
    throwIfAborted(options?.signal);
    if (this.#destroyed || generation !== this.#searchGeneration) return matches;
    const highlightQuery = options?.highlight === false ? "" : query;
    // Data-only searches never load DOM highlighting code. One controller owns
    // the stable target, so unit switches and virtualized cells reuse it.
    if (highlightQuery !== "" && this.#highlighter === null) {
      const { createSearchHighlighter } = await import("./search-highlight.js");
      throwIfAborted(options?.signal);
      if (this.#destroyed || generation !== this.#searchGeneration) return matches;
      this.#highlighter = createSearchHighlighter(this.target, () => this.#announceSearch());
    }
    this.#searchMove++;
    this.#searchMatches = highlightQuery === "" ? [] : matches;
    this.#searchIndex = this.#searchMatches.length === 0 ? null : 0;
    if (this.format === "pptx") {
      const current = this.#searchMatches.findIndex((match) => "slide" in match && match.slide.index === this.#index);
      this.#searchIndex = current >= 0 ? current : null;
    }
    this.#highlighter?.setQuery(highlightQuery);
    if (this.format === "xlsx" && this.#searchIndex !== null) {
      this.#activateCell(this.#searchMatches[0] as XlsxSearchMatch);
      if (this.#highlighter?.index === null) this.#searchIndex = null;
    }
    this.#announceSearch();
    return matches;
  }

  #announceSearch(): void {
    if (this.#destroyed) return;
    const { index, count } = this.searchNavigation;
    const state = `${index}:${count}`;
    if (state === this.#searchState) return;
    this.#searchState = state;
    this.dispatchEvent(new CustomEvent("searchchange", { detail: { index, count } }));
  }

  #activateCell(match: XlsxSearchMatch): void {
    const selector = match.chart === undefined
      ? `[data-aared-row="${match.row}"][data-aared-col="${match.col}"]`
      : `[data-aared-chart-name=${JSON.stringify(match.chart.name)}]`;
    this.#highlighter?.activateWithin(selector);
  }

  async #stepSearch(direction: 1 | -1): Promise<void> {
    if (this.#destroyed) throw new AaredError("closed", `${this.scope}: search navigation was called after destroy()`);
    const { count, index } = this.searchNavigation;
    if (count === 0) return;
    const next = index === null ? direction === 1 ? 0 : count - 1 : (index + direction + count) % count;
    const generation = this.#searchGeneration;
    const move = ++this.#searchMove;
    if (this.format === "docx") {
      this.#highlighter?.activate(next, true);
    } else {
      const match = this.#searchMatches[next]!;
      this.#searchIndex = next;
      if (this.format === "pptx" && "slide" in match) {
        await this.goTo(match.slide.index);
        if (generation !== this.#searchGeneration || move !== this.#searchMove || this.#destroyed) return;
        const local = this.#searchMatches.slice(0, next).filter((hit) => "slide" in hit && hit.slide.index === match.slide.index).length;
        this.#highlighter?.activate(local, true);
      } else if (this.format === "xlsx" && "address" in match) {
        await this.xlsx!.goToCell(match.address);
        if (generation !== this.#searchGeneration || move !== this.#searchMove || this.#destroyed) return;
        this.#activateCell(match);
      }
    }
    this.#announceSearch();
  }

  next(): Promise<void> {
    return this.#step("next", 1);
  }

  previous(): Promise<void> {
    return this.#step("previous", -1);
  }

  /**
   * One unit on from where the newest move is heading (else the unit on
   * screen), so quick presses add up instead of repeating one move. At either
   * end there is nowhere to go: nothing changes and nothing fires.
   */
  async #step(call: "next" | "previous", delta: 1 | -1): Promise<void> {
    if (this.#destroyed) throw new AaredError("closed", `${this.scope}: ${call}() was called after destroy()`);
    const from = this.#heading?.index ?? this.#index;
    const to = Math.min(Math.max(from + delta, 0), this.document.info.count - 1);
    if (to === from) return;
    await this.goTo(to);
  }

  /**
   * A navigation failed: fire `error` (`{ message, cause }`, where `cause` is
   * what the call rejects with). Cancellations and anything after `destroy()`
   * are not failures anyone is waiting to hear about.
   */
  #fail(cause: unknown): void {
    if (isAbort(cause) || this.#destroyed) return;
    const message = cause instanceof Error ? cause.message : String(cause);
    this.dispatchEvent(new CustomEvent("error", { detail: { message, cause } }));
  }

  setZoom(zoom: ZoomValue): void {
    if (this.#destroyed) return;
    this.#zoom = zoom;
    this.#rendered?.setZoom(zoom);
  }

  zoomIn(step: number = ZOOM_STEP): void {
    this.#zoomBy("zoomIn", step, 1);
  }

  zoomOut(step: number = ZOOM_STEP): void {
    this.#zoomBy("zoomOut", step, -1);
  }

  resetZoom(): void {
    this.setZoom(this.defaultZoom);
  }

  /**
   * One zoom step in or out: `step` added or taken away (0.1 is 10 points).
   * From a fit mode the step starts at the scale on screen, not at 100%.
   * Rounded to whole percents so repeated steps don't drift, and kept within
   * 25%–300%; a step that would move the wrong way (a zoom already past the
   * limit) does nothing.
   */
  #zoomBy(call: "zoomIn" | "zoomOut", step: number, direction: 1 | -1): void {
    if (typeof step !== "number" || !Number.isFinite(step) || step <= 0) {
      throw new AaredError("invalid-argument", `${this.scope}: ${call}() takes a step above 0, such as 0.1 for 10 points; received ${String(step)}`);
    }
    if (this.#destroyed) return;
    const zoom = this.zoom;
    const base = typeof zoom === "number" ? safeScale(zoom) : this.scale;
    const stepped = direction === 1 ? Math.min(MAX_ZOOM, base + step) : Math.max(MIN_ZOOM, base - step);
    if (direction === 1 ? stepped <= base : stepped >= base) return;
    this.setZoom(Math.round(stepped * 100) / 100);
  }

  setAppearance(appearance: Appearance): void {
    if (this.#destroyed) return;
    this.options = { ...this.options, appearance };
    this.#followSystemTheme();
    this.#rendered?.setAppearance(embedAppearance(this.format, appearance));
  }

  /**
   * The `system` theme follows the OS for as long as the appearance asks for
   * it: when the scheme flips, the same appearance is applied again and each
   * renderer resolves it afresh (the unit, zoom and Worker session stay). Any
   * other theme, or `destroy()`, drops the listener, so a host that takes the
   * theme over is never overridden by the OS.
   */
  #followSystemTheme(): void {
    this.#stopSystemTheme?.();
    this.#stopSystemTheme = this.options?.appearance?.theme === "system" && !this.#destroyed
      ? watchSystemTheme(() => {
          const appearance = this.options?.appearance;
          if (appearance !== undefined) this.#rendered?.setAppearance(embedAppearance(this.format, appearance));
        })
      : null;
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    this.#searchGeneration++;
    this.#searchMove++;
    this.#searchMatches = [];
    this.#searchIndex = null;
    this.#highlighter?.destroy();
    this.#highlighter = null;
    this.#stopSystemTheme?.();
    this.#stopSystemTheme = null;
    this.#generation++;
    this.#rendering?.abort();
    this.#rendering = null;
    this.#heading = null;
    this.#forwardController?.abort();
    this.#forwardController = null;
    this.#rendered?.destroy();
    this.#rendered = null;
    this.document.close();
    this.target.replaceChildren();
  }

  /** Paint unit `index`; false when a later render or `destroy()` overtook this one (nothing changed). */
  private async renderUnit(index: number): Promise<boolean> {
    const generation = ++this.#generation;
    const stale = (): boolean => this.#destroyed || generation !== this.#generation;
    // A render still in flight paints nothing once this one starts: both go
    // into the same target, and the older one finishing last would replace
    // this one's output (which then destroys itself as replaced).
    this.#rendering?.abort();
    const rendering = new AbortController();
    this.#rendering = rendering;
    const { mediaResolver } = this.options ?? {};
    const appearance = embedAppearance(this.format, this.options?.appearance);
    // Only what the caller set, over the embed defaults: each format's
    // `render()` owns the rest (session media — PPTX routes shared layout
    // pictures through one cache — and its default zoom).
    const common = {
      target: this.target,
      signal: rendering.signal,
      ...(appearance === undefined ? {} : { appearance }),
      ...(mediaResolver === undefined ? {} : { mediaResolver }),
      ...(this.#zoom === undefined ? {} : { zoom: this.#zoom }),
    };
    let next: RenderHandle;
    try {
      const pagination = this.options?.pagination;
      next = await this.route.render({ ...common, document: this.document, index,
        ...(pagination === undefined ? {} : { pagination }),
      });
    } catch (error) {
      // Cancelled because something newer took over (a later render, or
      // `destroy()` closing the document under it): it painted nothing, and
      // the newer call owns the outcome.
      if (stale()) return false;
      throw error;
    } finally {
      if (this.#rendering === rendering) this.#rendering = null;
    }
    if (stale()) {
      next.destroy();
      return false;
    }
    // A zoom set while this unit was rendering was announced by the old
    // handle; apply it before forwarding is attached so it isn't announced twice.
    if (this.#zoom !== undefined && !sameZoom(next.zoom, this.#zoom)) next.setZoom(this.#zoom);
    this.#rendered?.destroy();
    this.#rendered = next;
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
    return true;
  }
}

/** A move on `viewer.<format>` after `destroy()`: it says so, like `goTo()`. */
function closedControls(scope: string, format: "xlsx" | "docx", call: string): AaredError {
  return new AaredError("closed", `${scope}: viewer.${format}.${call}() was called after destroy()`);
}

const NO_PAGES: readonly DocxPageInfo[] = Object.freeze([]);

/** The viewer's one `xlsx` object: each call goes to the sheet on screen (`current()`, `null` after destroy). */
function xlsxControls(scope: string, current: () => XlsxRenderHandle | null, report: (error: unknown) => void): XlsxPreviewControls {
  return {
    get selection() {
      return current()?.selection ?? null;
    },
    goToCell: async (cell) => {
      const handle = current();
      if (handle === null) throw closedControls(scope, "xlsx", "goToCell");
      return reported(report, () => handle.goToCell(cell));
    },
    scrollTo: (target) => {
      const handle = current();
      if (handle !== null) reported(report, () => handle.scrollTo(target));
    },
    focus: (cell) => {
      const handle = current();
      if (handle !== null) reported(report, () => handle.focus(cell));
    },
  };
}

/** The viewer's one `docx` object: each call goes to the document's render (`current()`, `null` after destroy). */
function docxControls(
  scope: string,
  current: () => DocxRenderHandle | null,
  document: DocxDocument,
  report: (error: unknown) => void,
): DocxPreviewControls {
  return {
    get pageCount() {
      return current()?.pageCount ?? 0;
    },
    get pages() {
      return current()?.pages ?? NO_PAGES;
    },
    headings: document.headings.list(),
    goToSearchMatch: async (match) => {
      const handle = current();
      if (handle === null) throw closedControls(scope, "docx", "goToSearchMatch");
      return reported(report, () => handle.goToSearchMatch(match));
    },
    goToHeading: async (target) => {
      const handle = current();
      if (handle === null) throw closedControls(scope, "docx", "goToHeading");
      return reported(report, () => handle.goToHeading(target));
    },
    goToPage: async (index) => {
      const handle = current();
      if (handle === null) throw closedControls(scope, "docx", "goToPage");
      return reported(report, () => handle.goToPage(index));
    },
  };
}

/**
 * Run one navigation inside the unit; when it fails, `report` it (the
 * viewer's `error` event) and still throw or reject for the caller, the way
 * `goTo()` does, so a failure is heard even when the caller drops it.
 */
function reported<T>(report: (error: unknown) => void, run: () => T): T {
  let result: T;
  try {
    result = run();
  } catch (error) {
    report(error);
    throw error;
  }
  if (result instanceof Promise) {
    return result.catch((error: unknown) => {
      report(error);
      throw error;
    }) as T;
  }
  return result;
}
