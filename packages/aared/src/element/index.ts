/**
 * `<aared-viewer>`: the default, deliberately small UI over the headless
 * aared APIs. Attributes are its API (`src`, `type`, `index`, `zoom`,
 * `controls`, `theme`); read-only `format` and `info` report what opened;
 * `goTo()` navigates; removing the element is its teardown.
 *
 * It follows the same grammar as `preview()`: `index`/`goTo()` switch units,
 * events report changes (never the starting state — that is `load`), zoom is
 * the render handle's own (so each format's fit logic applies and
 * `zoomchange` forwards naturally), and every failure is an `AaredError`
 * with the same code `preview()` uses for the same situation.
 */

import type { Appearance, DocumentFormat, DocumentInfo, WorkerInit, ZoomValue } from "../core/index.js";
import { AaredError, isAbort } from "../core/index.js";
import { abortError } from "../core/errors.js";
import { resolveTheme } from "../appearance.js";
import { DEFAULT_ZOOM, announceUnit, assertHasUnits, followSheet, forwardRenderEvents, openFormat } from "../format-route.js";
import type { AnyDocument, RenderHandle } from "../format-route.js";
import { safeScale, sameZoom } from "../render-handle.js";
import type { PptxDocument } from "../pptx/index.js";
import type { XlsxDocument } from "../xlsx/index.js";
import type { DocxDocument } from "../docx/index.js";
import { WorkerSession, sourceBytes } from "../core/session.js";

export type { ZoomValue } from "../core/index.js";
export type { DocumentFormat } from "../core/index.js";

type ViewerSource = File | Blob | ArrayBuffer | string;

export type ViewerThemeAttribute = "light" | "dark" | "system";

/** Zoom the controls bar steps through: ×1.25 per press, within 25%–300%. */
const ZOOM_STEP = 1.25;
const MIN_ZOOM = 0.25;
const MAX_ZOOM = 3;

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
    `aared/element: unsupported type "${type}"; use "pptx", "xlsx", "docx", or omit type for detection`,
  );
}

export class AaredViewer extends HTMLElement {
  static readonly observedAttributes = [
    "src",
    "type",
    "index",
    "zoom",
    "controls",
    "theme",
  ] as const;

  #doc: AnyDocument | null = null;
  #format: DocumentFormat | null = null;
  #shadow: ShadowRoot;
  #page: HTMLDivElement;
  #controls: HTMLDivElement;
  #file: File | Blob | ArrayBuffer | null = null;
  #workerInit: WorkerInit | undefined = undefined;
  #openController: AbortController | null = null;
  #generation = 0;
  #renderGeneration = 0;
  #fullscreenButton: HTMLButtonElement | null = null;
  #slideRender: RenderHandle | null = null;
  #renderAbortController: AbortController | null = null;
  #resolvedTheme: "light" | "dark" = "light";
  #contentRenderTheme: "light" | "dark" | null = null;
  #themeMediaQuery: MediaQueryList | null = null;
  /** The unit on screen, so an `index` that already matches it renders nothing. */
  #shownIndex: number | null = null;
  /** Set while the element writes its own `index` attribute (tab follow, `goTo`), so no second render runs. */
  #writingIndex = false;

  constructor() {
    super();
    this.#shadow = this.attachShadow({ mode: "open" });
    const style = document.createElement("style");
    // Theme defaults live in private properties and each visual property has
    // one rule, so a consumer's --aared-background / --aared-page-shadow win
    // in both themes (a hard-coded dark rule used to override them).
    style.textContent = `
      :host { display: flex; flex-direction: column; min-width: 0; min-height: 0; --_aared-stage: #f2f2f2; --_aared-shadow: 0 1px 5px rgb(0 0 0 / 0.25); --_aared-controls-ink: inherit; }
      :host([data-aared-theme="dark"]) { --_aared-stage: #1c1c1f; --_aared-shadow: 0 1px 5px rgb(0 0 0 / 0.6); --_aared-controls-ink: #e4e4e7; }
      [part="stage"] { display: flex; flex: 1 1 auto; align-items: center; justify-content: center; min-width: 0; min-height: 0; overflow: auto; background: var(--aared-background, var(--_aared-stage)); color: var(--aared-text, inherit); }
      [part="page"] { flex: 0 0 auto; width: 100%; min-width: 0; margin: auto; box-shadow: var(--aared-page-shadow, var(--_aared-shadow)); }
      :host([data-aared-format="xlsx"]) [part="stage"] { align-items: stretch; }
      :host([data-aared-format="xlsx"]) [part="page"] { margin: 0; }
      :host([data-aared-format="pptx"]) [part="page"] { align-self: stretch; margin: 0; display: flex; flex-direction: column; box-shadow: none; }
      :host([data-aared-format="pptx"]) [part="page"] > [data-aared-element="slide"] { margin: auto; flex: none; box-shadow: var(--aared-page-shadow, var(--_aared-shadow)); }
      [part~="controls"] { display: flex; flex: 0 0 auto; flex-wrap: wrap; align-items: center; justify-content: center; gap: .5rem; padding: .5rem; font: 14px system-ui, sans-serif; color: var(--_aared-controls-ink); }
      :host(:fullscreen) [part~="controls"] { background: var(--aared-background, var(--_aared-stage)); }
      button { font: inherit; }
    `;
    const stage = document.createElement("div");
    stage.part.add("stage");
    this.#page = document.createElement("div");
    this.#page.part.add("page");
    stage.append(this.#page);
    this.#controls = document.createElement("div");
    this.#controls.part.add("controls");
    this.#controls.part.add("toolbar");
    this.#shadow.append(style, stage, this.#controls);
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

  /** The requested format (`type` attribute), or `null` to detect. See `format` for what opened. */
  get type(): DocumentFormat | null {
    const raw = this.getAttribute("type");
    return isFormat(raw) ? raw : null;
  }

  set type(value: DocumentFormat | null) {
    if (value === null) this.removeAttribute("type");
    else this.setAttribute("type", value);
  }

  /** The detected (or asserted) format of the open document, or `null` before one opens. */
  get format(): DocumentFormat | null {
    return this.#format;
  }

  /** What `open()` knows about the open document (also the `load` detail), or `null`. */
  get info(): DocumentInfo | null {
    return this.#doc?.info ?? null;
  }

  /** The unit to show (`index` attribute), zero-based. */
  get index(): number {
    return Number(this.getAttribute("index") ?? 0);
  }

  set index(value: number) {
    this.setAttribute("index", String(value));
  }

  /**
   * The zoom: the `zoom` attribute when set, else the render handle's own, else
   * the format's default (`fit-page` slides, `1` sheets, `fit-width` documents).
   */
  get zoom(): ZoomValue {
    const raw = this.getAttribute("zoom");
    if (raw === "fit-width" || raw === "fit-page") return raw;
    if (raw !== null) return Number(raw);
    return this.#slideRender?.zoom ?? (this.#format === null ? "fit-width" : DEFAULT_ZOOM[this.#format]);
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
  get theme(): ViewerThemeAttribute {
    const raw = this.getAttribute("theme");
    return raw === "dark" || raw === "system" ? raw : "light";
  }

  set theme(value: ViewerThemeAttribute) {
    this.setAttribute("theme", value);
  }

  get file(): File | Blob | ArrayBuffer | null {
    return this.#file;
  }

  set file(value: File | Blob | ArrayBuffer | null) {
    this.#file = value;
    if (this.isConnected) void this.open().catch(() => {});
  }

  get worker(): WorkerInit | undefined {
    return this.#workerInit;
  }

  set worker(value: WorkerInit | undefined) {
    this.#workerInit = value;
  }

  /**
   * (Re)open `file` or `src` and render unit `index`. Resolves once painted
   * and `load` has fired; rejects — after firing `error` once with the same
   * `AaredError` as `detail.cause` — when it fails. A newer `open()` wins: the
   * one it replaced rejects with `aborted` (no `error` event), which `isAbort()`
   * matches.
   */
  async open(): Promise<void> {
    const generation = this.#replaceDocument();
    const controller = this.#openController!;
    const source: ViewerSource | null = this.#file ?? this.src;
    const workerInit = this.#workerInit;
    try {
      const typeError = sourceTypeError(this.getAttribute("type"));
      if (typeError !== undefined) throw typeError;
      if (source === null) throw new AaredError("invalid-state", "aared/element: provide a src URL or file before calling open()");

      const bytes = await sourceBytes("aared/element", source, controller.signal, (loaded, total) => {
        this.#progress(loaded, total, generation);
      });
      if (!this.#current(generation, controller)) throw abortError();
      // An asserted type skips detection: the per-format `open()` raises the
      // typed `unsupported-format` mismatch itself, naming what it found.
      const format = this.type ?? await WorkerSession.detectFormat("aared/element", bytes, {
        signal: controller.signal,
        ...(workerInit === undefined ? {} : { worker: workerInit }),
      });
      if (!this.#current(generation, controller)) throw abortError();
      const doc = await openFormat(
        format,
        bytes,
        workerInit === undefined ? { signal: controller.signal } : { signal: controller.signal, worker: workerInit },
      );
      if (!this.#current(generation, controller)) {
        doc.close();
        throw abortError();
      }
      try {
        assertHasUnits("aared/element", doc);
      } catch (error) {
        doc.close();
        throw error;
      }
      this.#doc = doc;
      this.#format = format;
      const page = this.#requestedPage(doc.info.count);
      await this.#render(page, generation);
      if (!this.#current(generation, controller)) throw abortError();
      // The starting unit is reported by `load` (read `index`/`info`), never
      // by `indexchange`: events describe changes.
      this.dispatchEvent(new CustomEvent<DocumentInfo>("load", { detail: doc.info }));
      this.#updateChrome();
    } catch (error) {
      // Superseded or cancelled: no `error` event (nothing failed), but the
      // promise says so instead of resolving as if this open had finished.
      if (isAbort(error)) throw error;
      if (!this.#current(generation, controller)) throw abortError();
      this.#doc?.close();
      this.#doc = null;
      this.#format = null;
      this.#shownIndex = null;
      this.#page.replaceChildren();
      this.#updateChrome();
      this.#report(error);
      throw error;
    }
  }

  /**
   * Show unit `index`. Resolves once painted; going to the unit already shown
   * is a no-op. Rejects — after firing `error` once — with a `not-found`
   * `AaredError` outside `0 <= index < info.count`.
   */
  async goTo(index: number): Promise<void> {
    try {
      this.#assertPage(index, this.#doc?.info.count);
    } catch (error) {
      this.#report(error);
      throw error;
    }
    if (index === this.#shownIndex && this.#slideRender !== null) return;
    this.#writeIndex(index);
    if (this.#doc !== null) await this.#showPage(index, true);
  }

  connectedCallback(): void {
    this.#updateChrome();
    this.#listenSystemTheme();
    document.addEventListener("fullscreenchange", this.#onFullscreenChange);
    this.#syncFullscreen();
    if (this.#file !== null || this.src !== null) void this.open().catch(() => {});
  }

  disconnectedCallback(): void {
    document.removeEventListener("fullscreenchange", this.#onFullscreenChange);
    this.#themeMediaQuery?.removeEventListener("change", this.#onSystemThemeChange);
    this.#themeMediaQuery = null;
    this.#replaceDocument();
  }

  attributeChangedCallback(name: string): void {
    if (name === "controls" || name === "theme") {
      this.#updateChrome();
      return;
    }
    if (name === "zoom") {
      // The render handle owns zoom: its fit logic applies, and its
      // `zoomchange` (fired only on a real change) forwards to us. Removing
      // the attribute goes back to the format's default.
      const handle = this.#slideRender;
      if (handle !== null && this.#format !== null) {
        const want = this.hasAttribute("zoom") ? this.zoom : DEFAULT_ZOOM[this.#format];
        if (!sameZoom(handle.zoom, want)) handle.setZoom(want);
      }
      this.#updateChrome();
      return;
    }
    if (name === "type") {
      const requestedType = this.getAttribute("type");
      if (this.#doc !== null && this.#format !== null && requestedType !== null && requestedType !== this.#format) {
        const error = sourceTypeError(requestedType) ?? new AaredError(
          "unsupported-format",
          `aared/element: type "${requestedType}" does not match the open document's format "${this.#format}"`,
        );
        this.#replaceDocument();
        this.#report(error);
        return;
      }
    }
    if (name === "index") {
      if (this.#writingIndex) return;
      if (this.#doc !== null && this.index !== this.#shownIndex) void this.#showPage(this.index, false);
      return;
    }
    if (this.isConnected) void this.open().catch(() => {});
  }

  #replaceDocument(): number {
    this.#generation++;
    this.#openController?.abort();
    this.#openController = new AbortController();
    this.#renderGeneration++;
    this.#doc?.close();
    this.#doc = null;
    this.#format = null;
    this.#shownIndex = null;
    this.#slideRender?.destroy();
    this.#slideRender = null;
    this.#renderAbortController?.abort();
    this.#renderAbortController = null;
    this.#page.replaceChildren();
    this.#updateChrome();
    return this.#generation;
  }

  #current(generation: number, controller: AbortController): boolean {
    return generation === this.#generation && controller === this.#openController && !controller.signal.aborted;
  }

  /** Write `index` without re-rendering (the caller renders, or the unit is already on screen). */
  #writeIndex(index: number): void {
    this.#writingIndex = true;
    try {
      this.setAttribute("index", String(index));
    } finally {
      this.#writingIndex = false;
    }
  }

  /** Render `index` and announce it. `rethrow` for `goTo()`; attribute changes only report. */
  async #showPage(index: number, rethrow: boolean): Promise<void> {
    try {
      this.#assertPage(index, this.#doc?.info.count);
      const generation = this.#generation;
      const rendered = await this.#render(index, generation);
      if (!rendered || generation !== this.#generation || this.#doc === null || this.#format === null) return;
      announceUnit(this, this.#format, this.#doc, index);
      this.#updateChrome();
    } catch (error) {
      if (this.#doc === null || isAbort(error)) return;
      this.#report(error);
      if (rethrow) throw error;
    }
  }

  /**
   * The one appearance the element renders each format with — used for the
   * first paint and every later theme change alike. `setAppearance` replaces
   * rather than merges, so building it twice differently would drop knobs
   * (the XLSX tab strip used to vanish on a theme toggle).
   */
  #appearanceFor(format: DocumentFormat): Appearance {
    return format === "xlsx"
      ? { theme: this.#resolvedTheme, spreadsheet: { sheetTabs: true } }
      : { theme: this.#resolvedTheme };
  }

  async #render(index: number, generation: number): Promise<boolean> {
    const doc = this.#doc;
    const format = this.#format;
    if (doc === null || format === null) throw new AaredError("invalid-state", "aared/element: no document is open");
    const renderGeneration = ++this.#renderGeneration;
    const stale = () => generation !== this.#generation || renderGeneration !== this.#renderGeneration;
    this.#slideRender?.destroy();
    this.#slideRender = null;
    // Only what the element sets: each format's `render()` owns its defaults
    // (session media, its default zoom). The zoom attribute, when set, wins.
    const common = {
      target: this.#page,
      appearance: this.#appearanceFor(format),
      ...(this.hasAttribute("zoom") ? { zoom: this.zoom } : {}),
    };
    let rendered: RenderHandle;
    if (format === "pptx") {
      const renderer = await import("../pptx/index.js");
      if (stale()) return false;
      rendered = await renderer.render({ ...common, slide: (doc as PptxDocument).slides.get(index) });
    } else if (format === "xlsx") {
      const renderer = await import("../xlsx/index.js");
      if (stale()) return false;
      rendered = await renderer.render({ ...common, sheet: (doc as XlsxDocument).sheets.get(index) });
    } else {
      const renderer = await import("../docx/index.js");
      if (stale()) return false;
      rendered = await renderer.render({ ...common, document: doc as DocxDocument });
    }
    if (stale()) {
      rendered.destroy();
      return false;
    }
    this.#slideRender = rendered;
    this.#shownIndex = index;
    this.#renderAbortController?.abort();
    const controller = new AbortController();
    this.#renderAbortController = controller;
    forwardRenderEvents(rendered, this, controller.signal);
    // A sheet-tab click switches sheets inside the renderer: follow it on
    // `index` without re-rendering, and announce it like any navigation.
    followSheet(rendered, controller.signal, (sheet) => {
      if (this.#doc === null || this.#format === null || sheet === this.#shownIndex) return;
      this.#shownIndex = sheet;
      this.#writeIndex(sheet);
      announceUnit(this, this.#format, this.#doc, sheet);
      this.#updateChrome();
    });
    // A zoom attribute changed while this unit rendered: the old handle was
    // already gone, so apply it now — the forwarded event reports it once.
    if (this.hasAttribute("zoom") && !sameZoom(rendered.zoom, this.zoom)) rendered.setZoom(this.zoom);
    this.#contentRenderTheme = this.#resolvedTheme;
    return true;
  }

  #requestedPage(total: number): number {
    const page = this.index;
    this.#assertPage(page, total);
    return page;
  }

  #assertPage(page: number, total?: number): void {
    if (!Number.isInteger(page) || page < 0) {
      throw new AaredError("invalid-argument", `aared/element: index must be a non-negative integer; received ${String(page)}`);
    }
    if (total !== undefined && page >= total) {
      throw new AaredError("not-found", `aared/element: index ${page} is outside this ${total}-unit document`);
    }
  }

  #progress(loaded: number, total: number | undefined, generation: number): void {
    if (generation !== this.#generation) return;
    this.dispatchEvent(new CustomEvent("progress", {
      detail: total === undefined ? { loaded } : { loaded, total },
    }));
  }

  /** Fire `error` once; `cause` is the `AaredError` itself (the one `goTo()`/`open()` reject with). */
  #report(cause: unknown): void {
    const message = cause instanceof Error ? cause.message : String(cause);
    this.dispatchEvent(new CustomEvent("error", { detail: { message, cause } }));
  }

  #onSystemThemeChange = (): void => {
    if (this.theme === "system") this.#updateChrome();
  };

  #listenSystemTheme(): void {
    this.#themeMediaQuery?.removeEventListener("change", this.#onSystemThemeChange);
    if (typeof matchMedia !== "undefined") {
      this.#themeMediaQuery = matchMedia("(prefers-color-scheme: dark)");
      this.#themeMediaQuery.addEventListener("change", this.#onSystemThemeChange);
    }
  }

  #createControls(): void {
    const previous = document.createElement("button");
    previous.type = "button";
    previous.part.add("button");
    previous.textContent = "Previous";
    previous.addEventListener("click", () => void this.goTo(Math.max(0, this.index - 1)).catch(() => {}));
    const indicator = document.createElement("span");
    indicator.part.add("label");
    indicator.dataset.aaredControls = "index";
    const next = document.createElement("button");
    next.type = "button";
    next.part.add("button");
    next.textContent = "Next";
    next.addEventListener("click", () => {
      const total = this.#doc?.info.count ?? 0;
      void this.goTo(Math.min(Math.max(total - 1, 0), this.index + 1)).catch(() => {});
    });
    const zoomOut = document.createElement("button");
    zoomOut.type = "button";
    zoomOut.part.add("button");
    zoomOut.textContent = "-";
    zoomOut.setAttribute("aria-label", "Zoom out");
    zoomOut.addEventListener("click", () => this.#stepZoom(1 / ZOOM_STEP));
    const zoomIn = document.createElement("button");
    zoomIn.type = "button";
    zoomIn.part.add("button");
    zoomIn.textContent = "+";
    zoomIn.setAttribute("aria-label", "Zoom in");
    zoomIn.addEventListener("click", () => this.#stepZoom(ZOOM_STEP));
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

  /**
   * +/−: step from the numeric zoom, or from the scale a fit mode actually
   * painted (so "Fit page" at 62% steps to 78%, not to 125%), clamped to
   * 25%–300% and rounded so repeated steps don't drift.
   */
  #stepZoom(factor: number): void {
    const zoom = this.zoom;
    const base = typeof zoom === "number" ? safeScale(zoom) : (this.#slideRender?.scale ?? 1);
    const next = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, base * factor));
    this.zoom = Math.round(next * 100) / 100;
  }

  #updateChrome(): void {
    const resolved = resolveTheme(this.theme);
    const prevTheme = this.#resolvedTheme;
    this.#resolvedTheme = resolved;
    this.dataset.aaredTheme = resolved;
    if (this.#format === null) delete this.dataset.aaredFormat;
    else this.dataset.aaredFormat = this.#format;
    // When the resolved theme changes and a document is open, forward it in
    // place — with the same appearance the first paint used — so the unit
    // index, zoom, and Worker session survive.
    if (resolved !== prevTheme && this.#format !== null && this.#slideRender !== null && this.#contentRenderTheme !== null) {
      this.#slideRender.setAppearance(this.#appearanceFor(this.#format));
      this.#contentRenderTheme = resolved;
    }
    this.#controls.hidden = !this.controls;
    const total = this.#doc?.info.count ?? 0;
    const indicator = this.#controls.querySelector<HTMLElement>('[data-aared-controls="index"]');
    const zoomIndicator = this.#controls.querySelector<HTMLElement>('[data-aared-controls="zoom"]');
    if (indicator !== null) indicator.textContent = total === 0 ? "0 / 0" : `${this.index + 1} / ${total}`;
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

if (!customElements.get("aared-viewer")) {
  customElements.define("aared-viewer", AaredViewer);
}
