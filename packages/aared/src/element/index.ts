/**
 * `<aared-viewer>`: the default, deliberately small UI over the embed viewer
 * that `preview()` returns. Attributes are its API (`src`, `type`, `index`,
 * `zoom`, `controls`, `theme`, `pagination`); `viewer` is the open document's
 * handle; removing the element is its teardown.
 *
 * It runs on the same viewer as `preview()` (`embed-viewer.ts`), so error
 * codes, events and newest-wins are the same by construction. The attributes
 * and the viewer are one state: setting `index` or `zoom` calls the viewer,
 * and the viewer's `indexchange`/`zoomchange` (from `el.viewer.goTo()`, a
 * sheet tab, the controls) write the attribute back. `theme` owns the
 * appearance, so `viewer` leaves out `setAppearance` (and `destroy`).
 */

import type { DocumentFormat, DocumentInfo, LoadProgress, Source, ViewerTheme, WorkerInit, ZoomValue } from "../core/index.js";
import type { DocxPagination, DocxPreview, PptxPreview, PreviewEventMap, PreviewHandle, XlsxPreview } from "../preview/index.js";
import { AaredError, isAbort } from "../core/index.js";
import { abortError } from "../core/errors.js";
import { reportBackgroundError } from "../background-error.js";
import { resolveTheme, watchSystemTheme } from "../appearance.js";
import { DEFAULT_ZOOM, safeScale, sameZoom } from "../render-handle.js";
import { CLASS_SPECIFICITY, constantSheet } from "../class-sheet.js";
import { VIEWER_EVENTS } from "../format-route.js";
import { openAutoViewer } from "../auto-viewer.js";

export type { ZoomValue } from "../core/index.js";
export type { DocumentFormat } from "../core/index.js";

/** What the `file` property takes: any in-memory `Source` (a URL goes in `src`). */
type ViewerFile = Exclude<Source, string>;

/** How a DOCX is laid out: the `pagination` attribute. */
type PaginationMode = NonNullable<DocxPagination["mode"]>;

/** What the element owns itself: the appearance (its `theme` attribute) and teardown (removing it). */
type OwnedByElement = "setAppearance" | "destroy";

/**
 * The element's `viewer`: the `preview()` handle without `setAppearance` and
 * `destroy`, which belong to the element (`theme`, and removing it).
 * Checking `format` narrows it the same way (`viewer.xlsx` for a workbook).
 */
export type AaredViewerHandle =
  | Omit<PptxPreview, OwnedByElement>
  | Omit<XlsxPreview, OwnedByElement>
  | Omit<DocxPreview, OwnedByElement>;

const ICON_OPEN = '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">';
const ENTER_FULLSCREEN_ICON = `${ICON_OPEN}<path d="M8 3H5a2 2 0 0 0-2 2v3M16 3h3a2 2 0 0 1 2 2v3M8 21H5a2 2 0 0 1-2-2v-3M16 21h3a2 2 0 0 0 2-2v-3"/></svg>`;
const EXIT_FULLSCREEN_ICON = `${ICON_OPEN}<path d="M3 8h3a2 2 0 0 0 2-2V3M21 8h-3a2 2 0 0 1-2-2V3M3 16h3a2 2 0 0 1 2 2v3M21 16h-3a2 2 0 0 0-2 2v3"/></svg>`;

function isFormat(type: string | null): type is DocumentFormat {
  return type === "pptx" || type === "xlsx" || type === "docx";
}

function sourceTypeError(type: string | null): AaredError | undefined {
  if (type === null || isFormat(type)) return undefined;
  return new AaredError(
    "unsupported-format",
    `@aared/viewer/element: unsupported type "${type}"; use "pptx", "xlsx", "docx", or omit type for detection`,
  );
}

/**
 * The element's chrome, one constructed sheet per document (`constantSheet`):
 * no `<style>` text is written, so a strict `style-src` needs no
 * `'unsafe-inline'`. Theme defaults live in private properties and each visual
 * property has one rule, so a consumer's --aared-background / --aared-page-shadow
 * win in both themes (a hard-coded dark rule used to override them). A rule
 * over a renderer default carries the class sheet's two-ID bump
 * (`:not(#aared-0#aared-1)`, `class-sheet.ts`), or the default would win.
 */
const CHROME_CSS = `
  :host { display: flex; flex-direction: column; min-width: 0; min-height: 0; --_aared-stage: #f2f2f2; --_aared-shadow: 0 1px 5px rgb(0 0 0 / 0.25); --_aared-controls-ink: inherit; }
  :host([data-aared-theme="dark"]) { --_aared-stage: #1c1c1f; --_aared-shadow: 0 1px 5px rgb(0 0 0 / 0.6); --_aared-controls-ink: #e4e4e7; }
  [part="stage"] { display: flex; flex: 1 1 auto; align-items: center; justify-content: center; min-width: 0; min-height: 0; overflow: auto; background: var(--aared-background, var(--_aared-stage)); color: var(--aared-text, inherit); }
  [part="page"] { flex: 0 0 auto; width: 100%; min-width: 0; margin: auto; box-shadow: var(--aared-page-shadow, var(--_aared-shadow)); }
  :host([data-aared-format="xlsx"]) [part="stage"] { align-items: stretch; }
  :host([data-aared-format="xlsx"]) [part="page"] { margin: 0; }
  :host([data-aared-format="pptx"]) [part="page"] { align-self: stretch; margin: 0; display: flex; flex-direction: column; box-shadow: none; }
  :host([data-aared-format="pptx"]) [part="page"] > [data-aared-element="slide"] { margin: auto; flex: none; box-shadow: var(--aared-page-shadow, var(--_aared-shadow)); }
  :host([data-aared-pagination="layout"]) [part="page"] { box-shadow: none; padding-block: 1.5rem; }
  :host([data-aared-pagination="layout"]) [part="page"] [data-aared-element="page"]${CLASS_SPECIFICITY} { box-shadow: var(--aared-page-shadow, var(--_aared-shadow)); }
  [part~="controls"] { display: flex; flex: 0 0 auto; flex-wrap: wrap; align-items: center; justify-content: center; gap: .5rem; padding: .5rem; font: 14px system-ui, sans-serif; color: var(--_aared-controls-ink); }
  :host(:fullscreen) [part~="controls"] { background: var(--aared-background, var(--_aared-stage)); }
  button { font: inherit; }
`;

/**
 * The events `<aared-viewer>` fires, by name: `addEventListener` types each
 * `event.detail`. The viewer's events plus the element's own lifecycle; the
 * page's other element events keep their usual types (this element's
 * `error`, `load`, `progress` and `selectionchange` are its own).
 */
export interface AaredViewerEventMap extends Omit<HTMLElementEventMap, "error" | "load" | "progress" | "selectionchange">, PreviewEventMap {
  /** The document opened and its first unit painted. */
  load: CustomEvent<DocumentInfo>;
  /** Bytes read before parsing starts. */
  progress: CustomEvent<LoadProgress>;
}

/** `<aared-viewer>`'s own events, typed on top of `HTMLElement`'s (no runtime code). */
export interface AaredViewer {
  addEventListener<K extends keyof AaredViewerEventMap>(
    type: K,
    listener: (this: AaredViewer, event: AaredViewerEventMap[K]) => unknown,
    options?: boolean | AddEventListenerOptions,
  ): void;
  addEventListener(type: string, listener: EventListenerOrEventListenerObject, options?: boolean | AddEventListenerOptions): void;
  removeEventListener<K extends keyof AaredViewerEventMap>(
    type: K,
    listener: (this: AaredViewer, event: AaredViewerEventMap[K]) => unknown,
    options?: boolean | EventListenerOptions,
  ): void;
  removeEventListener(type: string, listener: EventListenerOrEventListenerObject, options?: boolean | EventListenerOptions): void;
}

declare global {
  interface HTMLElementTagNameMap {
    /** Typed by `@aared/viewer/element`: `document.createElement("aared-viewer")` and `querySelector("aared-viewer")` give an `AaredViewer`. */
    "aared-viewer": AaredViewer;
  }
}

/**
 * `HTMLElement` where the page has one. A server import (SSR, a Next.js
 * client component prerendered on the server) gets a stand-in, so the module
 * loads anywhere and only registers the element in a browser.
 */
const ElementBase: typeof HTMLElement = typeof HTMLElement === "undefined"
  ? (class {} as unknown as typeof HTMLElement)
  : HTMLElement;

export class AaredViewer extends ElementBase {
  static readonly observedAttributes = [
    "src",
    "type",
    "index",
    "zoom",
    "controls",
    "theme",
    "pagination",
  ] as const;

  #shadow: ShadowRoot;
  #page: HTMLDivElement;
  #controls: HTMLDivElement;
  #file: ViewerFile | null = null;
  /** Whether `file` was set after `src`: the last one set is what opens. */
  #fileLast = false;
  #workerInit: WorkerInit | undefined = undefined;
  /** The open in flight, or the one that opened `#viewer`; aborted when a newer open or a removal replaces it. */
  #openController: AbortController | null = null;
  #viewer: PreviewHandle | null = null;
  /** Detaches the element's listeners from `#viewer`. */
  #viewerListeners: AbortController | null = null;
  /** The theme `#viewer` was last given (`system` included: the viewer follows the OS itself), so only a real change is forwarded. */
  #viewerTheme: ViewerTheme | null = null;
  /** The DOCX layout `#viewer` opened with, so only a real change re-opens it. */
  #viewerPagination: PaginationMode | null = null;
  #fullscreenButton: HTMLButtonElement | null = null;
  /** Stops the OS colour-scheme listener that keeps the element's own chrome in step. */
  #stopSystemTheme: (() => void) | null = null;
  /** Set while the element writes `index`/`zoom` to match the viewer, so the write moves nothing. */
  #syncing = false;

  constructor() {
    super();
    this.#shadow = this.attachShadow({ mode: "open" });
    // Renders inside the shadow root add their own class sheet beside this one.
    const chrome = constantSheet(this.ownerDocument, CHROME_CSS);
    if (chrome !== null) this.#shadow.adoptedStyleSheets = [chrome];
    const stage = document.createElement("div");
    stage.part.add("stage");
    this.#page = document.createElement("div");
    this.#page.part.add("page");
    stage.append(this.#page);
    this.#controls = document.createElement("div");
    this.#controls.part.add("controls");
    this.#controls.part.add("toolbar");
    this.#shadow.append(stage, this.#controls);
    // No `#updateChrome()` here: it writes `data-aared-theme` on the host,
    // and `document.createElement()` / `new AaredViewer()` throw
    // NotSupportedError if the constructor adds attributes. It runs from
    // `connectedCallback` instead.
    this.#createControls();
  }

  get src(): string | null {
    return this.getAttribute("src");
  }

  set src(value: string | null) {
    if (value === null) this.removeAttribute("src");
    else this.setAttribute("src", value);
  }

  /** The requested format (`type` attribute), or `null` to detect. What opened is `viewer.format`. */
  get type(): DocumentFormat | null {
    const raw = this.getAttribute("type");
    return isFormat(raw) ? raw : null;
  }

  set type(value: DocumentFormat | null) {
    if (value === null) this.removeAttribute("type");
    else this.setAttribute("type", value);
  }

  /**
   * The open document's viewer, or `null` before one opens (and after a
   * failure or removal): `format`, `info`, `index`, `zoom`, `defaultZoom`,
   * `scale`, `goTo()`, `setZoom()`, the toolbar methods (`next()`,
   * `previous()`, `zoomIn()`, `zoomOut()`, `resetZoom()`), `search()` over
   * the open document, the places inside a unit (`xlsx`, `docx`) and the
   * events. Every move keeps the `index` and `zoom` attributes in step.
   */
  get viewer(): AaredViewerHandle | null {
    return this.#viewer;
  }

  /** The unit to show (`index` attribute), zero-based. */
  get index(): number {
    return Number(this.getAttribute("index") ?? 0);
  }

  set index(value: number) {
    this.setAttribute("index", String(value));
  }

  /**
   * The zoom: the `zoom` attribute when set, else the viewer's own, else
   * `"fit-width"` before a document opens. Each format opens at its default
   * (`fit-page` slides, `1` sheets, `fit-width` documents).
   */
  get zoom(): ZoomValue {
    const raw = this.getAttribute("zoom");
    if (raw === "fit-width" || raw === "fit-page") return raw;
    if (raw !== null) return Number(raw);
    return this.#viewer?.zoom ?? "fit-width";
  }

  set zoom(value: ZoomValue) {
    this.setAttribute("zoom", String(value));
  }

  get controls(): boolean {
    return this.hasAttribute("controls");
  }

  set controls(value: boolean) {
    this.toggleAttribute("controls", value);
  }

  /** Viewer theme (`light` | `dark` | `system`). Fixed `light` unless the host opts in to `dark` or `system`. */
  get theme(): ViewerTheme {
    const raw = this.getAttribute("theme");
    return raw === "dark" || raw === "system" ? raw : "light";
  }

  set theme(value: ViewerTheme) {
    this.setAttribute("theme", value);
  }

  /**
   * DOCX layout (`pagination` attribute): `"flow"` (default) reflows the
   * document to the element's width; `"layout"` paints physical pages
   * (`viewer.docx.pages`, `goToPage()`). Changing it re-opens a DOCX; other
   * formats ignore it.
   */
  get pagination(): PaginationMode {
    return this.getAttribute("pagination") === "layout" ? "layout" : "flow";
  }

  set pagination(value: PaginationMode) {
    this.setAttribute("pagination", value);
  }

  get file(): ViewerFile | null {
    return this.#file;
  }

  set file(value: ViewerFile | null) {
    this.#file = value;
    this.#fileLast = value !== null;
    this.#reopen();
  }

  get worker(): WorkerInit | undefined {
    return this.#workerInit;
  }

  set worker(value: WorkerInit | undefined) {
    this.#workerInit = value;
  }

  /**
   * (Re)open whichever of `file` and `src` was set last and render unit
   * `index`. Resolves once painted and `load` has fired; rejects only when
   * it fails. Automatic loads report through `error` instead. A newer
   * `open()` wins: the one it replaced rejects with
   * `aborted` (no `error` event), which `isAbort()` matches.
   */
  async open(): Promise<void> {
    this.#close();
    const controller = new AbortController();
    this.#openController = controller;
    const signal = controller.signal;
    const theme = this.theme;
    const pagination = this.pagination;
    try {
      const type = this.getAttribute("type");
      const typeError = sourceTypeError(type);
      if (typeError !== undefined) throw typeError;
      const source = this.#source();
      if (source === null) throw new AaredError("invalid-state", "@aared/viewer/element: provide a src URL or file before calling open()");
      const viewer = await openAutoViewer({
        scope: "@aared/viewer/element",
        source,
        target: this.#page,
        options: {
          ...(isFormat(type) ? { type } : {}),
          index: this.index,
          ...(this.hasAttribute("zoom") ? { zoom: this.zoom } : {}),
          appearance: { theme },
          ...(pagination === "layout" ? { pagination: { mode: pagination } } : {}),
          ...(this.#workerInit === undefined ? {} : { worker: this.#workerInit }),
          signal,
        },
        onProgress: (loaded, total) => {
          if (controller !== this.#openController) return;
          this.dispatchEvent(new CustomEvent("progress", { detail: total === undefined ? { loaded } : { loaded, total } }));
        },
      });
      try {
        await this.#catchUp(viewer, signal);
      } catch (error) {
        viewer.destroy();
        throw error;
      }
      this.#attach(viewer, theme, pagination);
      this.#updateChrome();
      // The starting unit is reported by `load` (read `index` and `viewer`),
      // never by `indexchange`: events describe changes.
      this.dispatchEvent(new CustomEvent<DocumentInfo>("load", { detail: viewer.info }));
    } catch (error) {
      // Superseded or cancelled: no `error` event (nothing failed), but the
      // promise says so instead of resolving as if this open had finished.
      if (isAbort(error)) throw error;
      if (controller !== this.#openController) throw abortError();
      this.#updateChrome();
      throw error;
    }
  }

  connectedCallback(): void {
    this.#updateChrome();
    this.#stopSystemTheme?.();
    this.#stopSystemTheme = watchSystemTheme(this.#onSystemThemeChange);
    document.addEventListener("fullscreenchange", this.#onFullscreenChange);
    this.#syncFullscreen();
    this.#reopen();
  }

  disconnectedCallback(): void {
    document.removeEventListener("fullscreenchange", this.#onFullscreenChange);
    this.#stopSystemTheme?.();
    this.#stopSystemTheme = null;
    this.#close();
  }

  attributeChangedCallback(name: string): void {
    if (name === "controls" || name === "theme") {
      this.#updateChrome();
      return;
    }
    if (name === "index" || name === "zoom") {
      // The element's own write to match the viewer moves nothing; before a
      // document is open, the open reads both attributes.
      const viewer = this.#viewer;
      if (!this.#syncing && viewer !== null) {
        if (name === "index") this.#showIndex(viewer);
        else {
          // Removing the attribute goes back to the format's default.
          const zoom = this.#wantedZoom(viewer.format);
          if (!sameZoom(viewer.zoom, zoom)) viewer.setZoom(zoom);
        }
      }
      this.#updateChrome();
      return;
    }
    if (name === "type") {
      const viewer = this.#viewer;
      if (viewer !== null) {
        // The open document's format is known: comparing needs no download.
        const type = this.getAttribute("type");
        if (type === null || type === viewer.format) return;
        const error = sourceTypeError(type) ?? new AaredError(
          "unsupported-format",
          `@aared/viewer/element: type "${type}" does not match the open document's format "${viewer.format}"`,
        );
        this.#close();
        this.#report(error);
        return;
      }
    }
    if (name === "pagination") {
      // Only a DOCX has pages, and only a real change re-opens it.
      const viewer = this.#viewer;
      if (viewer !== null && (viewer.format !== "docx" || this.pagination === this.#viewerPagination)) return;
    }
    if (name === "src" && this.src !== null) this.#fileLast = false;
    this.#reopen();
  }

  /**
   * What a change to `src`, `file`, `type` or `pagination` (or being put on
   * the page) does: open the source again, or close quietly when neither `src`
   * nor `file` is set. Emptying the viewer is not a failure, so no event
   * fires; only calling `open()` with nothing to open is (`invalid-state`).
   */
  #reopen(): void {
    if (!this.isConnected) return;
    if (this.#source() === null) this.#close();
    else void this.open().catch((error: unknown) => {
      if (this.isConnected) this.#report(error);
    });
  }

  /** What opens: whichever of `file` and `src` was set last, else the other. */
  #source(): Source | null {
    if (this.#fileLast && this.#file !== null) return this.#file;
    return this.src ?? this.#file;
  }

  /** End what is open or opening: abort an open in flight and destroy the viewer, which clears the page. */
  #close(): void {
    this.#openController?.abort();
    this.#openController = null;
    this.#viewerListeners?.abort();
    this.#viewerListeners = null;
    this.#viewer?.destroy();
    this.#viewer = null;
    this.#viewerTheme = null;
    this.#viewerPagination = null;
    this.#page.replaceChildren();
    this.#updateChrome();
  }

  /**
   * Apply what the attributes asked for while the file loaded, before `load`
   * and before the element listens, so no change event comes ahead of `load`.
   * A unit that cannot be shown fails the open, as a bad starting `index` does.
   */
  async #catchUp(viewer: PreviewHandle, signal: AbortSignal): Promise<void> {
    if (signal.aborted) throw abortError();
    while (viewer.index !== this.index) {
      await viewer.goTo(this.index);
      if (signal.aborted) throw abortError();
    }
    const zoom = this.#wantedZoom(viewer.format);
    if (!sameZoom(viewer.zoom, zoom)) viewer.setZoom(zoom);
  }

  /** Listen to a freshly opened viewer: re-fire its events here, and keep `index` and `zoom` in step. */
  #attach(viewer: PreviewHandle, theme: ViewerTheme, pagination: PaginationMode): void {
    this.#viewer = viewer;
    this.#viewerTheme = theme;
    this.#viewerPagination = pagination;
    const listeners = new AbortController();
    this.#viewerListeners = listeners;
    for (const type of VIEWER_EVENTS) {
      viewer.addEventListener(type, (event: Event) => {
        if (type === "indexchange") this.#syncAttribute("index", String(viewer.index));
        else if (type === "zoomchange") this.#syncZoom(viewer);
        // Built-in renderer actions report here; public calls reject only.
        this.dispatchEvent(new CustomEvent(type, { detail: (event as CustomEvent).detail }));
        if (type === "indexchange" || type === "zoomchange") this.#updateChrome();
      }, { signal: listeners.signal });
    }
  }

  /**
   * Move to the `index` attribute. Always through the viewer, even when it
   * names the unit on screen: a move still on its way is then cancelled
   * (with nothing on its way the viewer does nothing). This attribute owns
   * the action and reports its failure; the attribute goes back to
   * the unit on screen, unless something newer was written meanwhile.
   */
  #showIndex(viewer: PreviewHandle): void {
    const asked = this.getAttribute("index");
    void viewer.goTo(this.index).catch((error: unknown) => {
      if (this.#viewer !== viewer || this.getAttribute("index") !== asked) return;
      this.#syncAttribute("index", String(viewer.index));
      this.#updateChrome();
      this.#report(error);
    });
  }

  /** Write `index` or `zoom` to match the viewer; the write itself moves nothing. */
  #syncAttribute(name: "index" | "zoom", value: string): void {
    if (this.getAttribute(name) === value) return;
    this.#syncing = true;
    try {
      this.setAttribute(name, value);
    } finally {
      this.#syncing = false;
    }
  }

  /** After a zoom change the attribute follows, except that an absent one at the format's default stays absent. */
  #syncZoom(viewer: PreviewHandle): void {
    if (this.hasAttribute("zoom") ? sameZoom(this.zoom, viewer.zoom) : sameZoom(viewer.zoom, DEFAULT_ZOOM[viewer.format])) return;
    this.#syncAttribute("zoom", String(viewer.zoom));
  }

  /** The zoom the attribute asks for: its value, or the format's default when it is absent. */
  #wantedZoom(format: DocumentFormat): ZoomValue {
    return this.hasAttribute("zoom") ? this.zoom : DEFAULT_ZOOM[format];
  }

  /** Report only a built-in action; public calls throw/reject to their caller. */
  #report(cause: unknown): void {
    reportBackgroundError(this, "@aared/viewer/element", cause);
  }

  /** The OS scheme flipped: the chrome (`data-aared-theme`, the stage) follows a `system` theme. The viewer follows it by itself. */
  #onSystemThemeChange = (): void => {
    if (this.theme === "system") this.#updateChrome();
  };

  /**
   * The bar is the viewer's own toolbar methods behind buttons (`previous()`,
   * `next()`, `zoomIn()`, `zoomOut()`), so it pages and steps exactly as an
   * app's own toolbar would; the `index` and `zoom` attributes follow from the
   * viewer's events. With nothing open the buttons do nothing.
   */
  #createControls(): void {
    const previous = document.createElement("button");
    previous.type = "button";
    previous.part.add("button");
    previous.textContent = "Previous";
    const move = (direction: "previous" | "next"): void => {
      const viewer = this.#viewer;
      if (viewer === null) return;
      void viewer[direction]().catch((error: unknown) => {
        if (this.#viewer === viewer) this.#report(error);
      });
    };
    previous.addEventListener("click", () => move("previous"));
    const indicator = document.createElement("span");
    indicator.part.add("label");
    indicator.dataset.aaredControls = "index";
    const next = document.createElement("button");
    next.type = "button";
    next.part.add("button");
    next.textContent = "Next";
    next.addEventListener("click", () => move("next"));
    const zoomOut = document.createElement("button");
    zoomOut.type = "button";
    zoomOut.part.add("button");
    zoomOut.textContent = "-";
    zoomOut.setAttribute("aria-label", "Zoom out");
    zoomOut.addEventListener("click", () => this.#viewer?.zoomOut());
    const zoomIn = document.createElement("button");
    zoomIn.type = "button";
    zoomIn.part.add("button");
    zoomIn.textContent = "+";
    zoomIn.setAttribute("aria-label", "Zoom in");
    zoomIn.addEventListener("click", () => this.#viewer?.zoomIn());
    const zoom = document.createElement("span");
    zoom.part.add("label");
    zoom.dataset.aaredControls = "zoom";
    this.#controls.append(previous, indicator, next, zoomOut, zoom, zoomIn);
    if (typeof this.requestFullscreen === "function") {
      const fullscreen = document.createElement("button");
      fullscreen.type = "button";
      fullscreen.part.add("button");
      fullscreen.part.add("fullscreen");
      fullscreen.addEventListener("click", () => this.#toggleFullscreen());
      this.#fullscreenButton = fullscreen;
      this.#syncFullscreen();
      this.#controls.append(fullscreen);
    }
  }

  #toggleFullscreen(): void {
    if (document.fullscreenElement === this) void document.exitFullscreen().catch(() => {});
    else void this.requestFullscreen().catch(() => {});
  }

  #onFullscreenChange = (): void => this.#syncFullscreen();

  #syncFullscreen(): void {
    const button = this.#fullscreenButton;
    if (button === null) return;
    const active = document.fullscreenElement === this;
    const label = active ? "Exit full screen" : "Full screen";
    button.innerHTML = active ? EXIT_FULLSCREEN_ICON : ENTER_FULLSCREEN_ICON;
    button.title = label;
    button.setAttribute("aria-label", label);
    // Hosts embedded without permission (e.g. an iframe lacking
    // allow="fullscreen") report the API as disabled.
    button.hidden = document.fullscreenEnabled === false;
  }

  #updateChrome(): void {
    const theme = this.theme;
    this.dataset.aaredTheme = resolveTheme(theme);
    const viewer = this.#viewer;
    if (viewer === null) delete this.dataset.aaredFormat;
    else this.dataset.aaredFormat = viewer.format;
    // Physical pages carry their own shadow, so the chrome follows what opened, not the attribute alone.
    if (viewer?.format === "docx" && this.#viewerPagination === "layout") this.dataset.aaredPagination = "layout";
    else delete this.dataset.aaredPagination;
    // `theme` owns the viewer's appearance: a real change is forwarded in
    // place, so the unit, zoom and Worker session survive (and the embed
    // defaults, XLSX sheet tabs among them, stay). `system` goes as `system`:
    // the viewer then follows the OS flips itself, so one forwarded here
    // changes nothing but the chrome above.
    if (viewer !== null && this.#viewerTheme !== theme) {
      this.#viewerTheme = theme;
      viewer.setAppearance({ theme });
    }
    this.#controls.hidden = !this.controls;
    const indicator = this.#controls.querySelector<HTMLElement>('[data-aared-controls="index"]');
    const zoomIndicator = this.#controls.querySelector<HTMLElement>('[data-aared-controls="zoom"]');
    if (indicator !== null) indicator.textContent = viewer === null ? "0 / 0" : `${viewer.index + 1} / ${viewer.info.count}`;
    if (zoomIndicator !== null) {
      const zoom = this.zoom;
      zoomIndicator.textContent = zoom === "fit-width"
        ? "Fit width"
        : zoom === "fit-page"
          ? "Fit page"
          : `${Math.round(safeScale(zoom) * 100)}%`;
    }
  }
}

if (typeof customElements !== "undefined" && !customElements.get("aared-viewer")) {
  customElements.define("aared-viewer", AaredViewer);
}
