/** The default, deliberately small UI for the headless aared APIs. */

import type { DocumentFormat, DocumentInfo, ZoomValue } from "../core/index.js";
import type { PptxDocument } from "../pptx/index.js";
import type { XlsxDocument } from "../xlsx/index.js";
import type { DocxDocument } from "../docx/index.js";
import { WorkerSession } from "../core/session.js";

export type { ZoomValue } from "../core/index.js";
export type { DocumentFormat } from "../core/index.js";

type ViewerSource = File | Blob | ArrayBuffer | string;
type AnyDocument = PptxDocument | XlsxDocument | DocxDocument;

interface RenderHandle extends EventTarget {
  destroy(): void;
}

export type ViewerThemeAttribute = "light" | "dark" | "system";

function abortError(error: unknown): boolean {
  return error instanceof DOMException && error.name === "AbortError";
}

function sourceTypeError(type: string | null): Error | undefined {
  if (type === null || type === "pptx" || type === "xlsx" || type === "docx") return undefined;
  return new Error(`<aared-viewer>: unsupported type "${type}"; use "pptx", "xlsx", "docx", or omit type for detection`);
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
  #stage: HTMLDivElement;
  #page: HTMLDivElement;
  #controls: HTMLDivElement;
  #file: File | Blob | ArrayBuffer | null = null;
  #openController: AbortController | null = null;
  #generation = 0;
  #renderGeneration = 0;
  #slideSize: { width: number; height: number } | null = null;
  #resizeObserver: ResizeObserver | null = null;
  #slideRender: RenderHandle | null = null;
  #renderAbortController: AbortController | null = null;
  #resolvedTheme: "light" | "dark" = "light";
  #contentRenderTheme: "light" | "dark" | null = null;
  #themeMediaQuery: MediaQueryList | null = null;

  constructor() {
    super();
    this.#shadow = this.attachShadow({ mode: "open" });
    const style = document.createElement("style");
    style.textContent = `
      :host { display: block; min-width: 0; min-height: 0; }
      :host([data-aared-theme="dark"]) [part="stage"] { background: #1c1c1f; }
      :host([data-aared-theme="dark"]) [part="page"] { box-shadow: 0 1px 5px rgb(0 0 0 / 0.6); }
      :host([data-aared-theme="dark"]) [part="controls"] { color: #e4e4e7; }
      [part="stage"] { display: flex; align-items: center; justify-content: center; min-width: 0; min-height: 0; height: 100%; overflow: auto; background: var(--aared-background, #f2f2f2); color: var(--aared-text, inherit); }
      [part="page"] { flex: 0 0 auto; width: 100%; min-width: 0; margin: auto; box-shadow: var(--aared-page-shadow, 0 1px 5px rgb(0 0 0 / 0.25)); }
      [part="controls"] { display: flex; align-items: center; justify-content: center; gap: .5rem; padding: .5rem; font: 14px system-ui, sans-serif; }
      button { font: inherit; }
    `;
    this.#stage = document.createElement("div");
    this.#stage.part.add("stage");
    this.#page = document.createElement("div");
    this.#page.part.add("page");
    this.#stage.append(this.#page);
    this.#controls = document.createElement("div");
    this.#controls.part.add("controls");
    this.#controls.part.add("toolbar");
    this.#shadow.append(style, this.#stage, this.#controls);
    if (typeof ResizeObserver !== "undefined") {
      this.#resizeObserver = new ResizeObserver(() => this.#updateChrome());
    }
    this.#createControls();
    this.#updateChrome();
  }

  get src(): string | null {
    return this.getAttribute("src");
  }

  set src(value: string | null) {
    if (value === null) this.removeAttribute("src");
    else this.setAttribute("src", value);
  }

  get type(): DocumentFormat | null {
    return this.getAttribute("type") as DocumentFormat | null;
  }

  get index(): number {
    return Number(this.getAttribute("index") ?? 0);
  }

  set index(value: number) {
    this.setAttribute("index", String(value));
  }

  get zoom(): ZoomValue {
    const raw = this.getAttribute("zoom");
    if (raw === "fit-width" || raw === "fit-page") return raw;
    return raw === null ? "fit-width" : Number(raw);
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

  /** Viewer chrome theme (`light` | `dark` | `system`). Content stays source-faithful. */
  get theme(): ViewerThemeAttribute {
    const raw = this.getAttribute("theme");
    return raw === "light" || raw === "dark" ? raw : "system";
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

  async open(): Promise<void> {
    const generation = this.#replaceDocument();
    const controller = this.#openController!;
    const source = this.#file ?? this.src;
    try {
      const typeError = sourceTypeError(this.getAttribute("type"));
      if (typeError !== undefined) throw typeError;
      if (source === null) throw new Error("<aared-viewer>: provide a src URL or file before calling open()");

      const bytes = await this.#readSource(source, controller.signal, generation);
      if (!this.#current(generation, controller)) return;
      const requestedType = this.getAttribute("type") as DocumentFormat | null;
      const format = requestedType ?? await WorkerSession.detectFormat("aared/element", bytes, controller.signal);
      if (!this.#current(generation, controller)) return;
      if (requestedType !== null && requestedType !== format) {
        throw new Error(`<aared-viewer>: explicit type "${requestedType}" does not match detected format "${format}"`);
      }
      const doc = await this.#openFormat(format, bytes, controller.signal);
      if (!this.#current(generation, controller)) {
        doc.close();
        return;
      }
      if (doc.info.count === 0) {
        doc.close();
        throw new Error("<aared-viewer>: this document has no slides, sheets, or sections to show");
      }
      this.#doc = doc;
      this.#format = format;
      const page = this.#requestedPage(doc.info.count);
      const rendered = await this.#render(page, generation);
      if (!this.#current(generation, controller)) return;
      this.dispatchEvent(new CustomEvent<DocumentInfo>("load", { detail: doc.info }));
      if (rendered) this.#announcePage(page);
      this.#updateChrome();
    } catch (error) {
      if (!this.#current(generation, controller) || abortError(error)) return;
      this.#doc?.close();
      this.#doc = null;
      this.#format = null;
      this.#slideSize = null;
      this.#page.replaceChildren();
      this.#updateChrome();
      this.#report(error);
      throw error;
    }
  }

  goTo(index: number): void {
    try {
      this.#assertPage(index, this.#doc?.info.count);
    } catch (error) {
      this.#report(error);
      throw error;
    }
    if (this.index !== index) {
      this.index = index;
      return;
    }
    if (this.#doc !== null) void this.#showPage(index);
  }

  connectedCallback(): void {
    this.#resizeObserver?.observe(this.#stage);
    this.#listenSystemTheme();
    if (this.#file !== null || this.src !== null) void this.open().catch(() => {});
  }

  disconnectedCallback(): void {
    this.#resizeObserver?.disconnect();
    this.#themeMediaQuery?.removeEventListener("change", this.#onSystemThemeChange);
    this.#themeMediaQuery = null;
    this.#replaceDocument();
  }

  attributeChangedCallback(name: string): void {
    if (name === "zoom" || name === "controls" || name === "theme") {
      this.#updateChrome();
      return;
    }
    if (name === "type") {
      const requestedType = this.getAttribute("type") as DocumentFormat | null;
      if (this.#doc !== null && this.#format !== null && requestedType !== null && requestedType !== this.#format) {
        const error = new Error(`<aared-viewer>: explicit type "${requestedType}" does not match detected format "${this.#format}"`);
        this.#replaceDocument();
        this.#report(error);
        return;
      }
    }
    if (name === "index") {
      if (this.#doc !== null) void this.#showPage(this.index);
      return;
    }
    if (this.isConnected) void this.open().catch(() => {});
  }

  async #openFormat(format: DocumentFormat, bytes: ArrayBuffer, signal: AbortSignal): Promise<AnyDocument> {
    if (format === "pptx") {
      const pptx = await import("../pptx/index.js");
      return pptx.open(bytes, { signal });
    }
    if (format === "xlsx") {
      const xlsx = await import("../xlsx/index.js");
      return xlsx.open(bytes, { signal });
    }
    const docx = await import("../docx/index.js");
    return docx.open(bytes, { signal });
  }

  #replaceDocument(): number {
    this.#generation++;
    this.#openController?.abort();
    this.#openController = new AbortController();
    this.#renderGeneration++;
    this.#doc?.close();
    this.#doc = null;
    this.#format = null;
    this.#slideRender?.destroy();
    this.#slideRender = null;
    this.#renderAbortController?.abort();
    this.#renderAbortController = null;
    this.#slideSize = null;
    this.#page.replaceChildren();
    this.#updateChrome();
    return this.#generation;
  }

  #current(generation: number, controller: AbortController): boolean {
    return generation === this.#generation && controller === this.#openController && !controller.signal.aborted;
  }

  async #readSource(source: ViewerSource, signal: AbortSignal, generation: number): Promise<ArrayBuffer> {
    if (typeof source !== "string") {
      const total = source instanceof ArrayBuffer ? source.byteLength : source.size;
      this.#progress(0, total, generation);
      const bytes = source instanceof ArrayBuffer ? source.slice(0) : await source.arrayBuffer();
      this.#progress(bytes.byteLength, total, generation);
      return bytes;
    }
    const response = await fetch(source, { signal });
    if (!response.ok) throw new Error(`<aared-viewer>: failed to fetch "${source}": HTTP ${response.status}`);
    const totalHeader = response.headers.get("content-length");
    const total = totalHeader === null ? undefined : Number(totalHeader);
    this.#progress(0, Number.isFinite(total) && total! >= 0 ? total : undefined, generation);
    if (response.body === null) {
      const bytes = await response.arrayBuffer();
      this.#progress(bytes.byteLength, bytes.byteLength, generation);
      return bytes;
    }
    const reader = response.body.getReader();
    const chunks: Uint8Array[] = [];
    let loaded = 0;
    while (true) {
      const chunk = await reader.read();
      if (chunk.done) break;
      chunks.push(chunk.value);
      loaded += chunk.value.byteLength;
      this.#progress(loaded, total, generation);
    }
    const bytes = new Uint8Array(loaded);
    let offset = 0;
    for (const chunk of chunks) {
      bytes.set(chunk, offset);
      offset += chunk.byteLength;
    }
    return bytes.buffer;
  }

  async #showPage(index: number): Promise<void> {
    try {
      this.#assertPage(index, this.#doc?.info.count);
      const generation = this.#generation;
      const rendered = await this.#render(index, generation);
      if (!rendered || generation !== this.#generation) return;
      this.#announcePage(index);
      this.#updateChrome();
    } catch (error) {
      if (this.#doc !== null && !abortError(error)) this.#report(error);
    }
  }

  /** `indexchange` always fires; format-native siblings fire when they apply. */
  #announcePage(index: number): void {
    this.dispatchEvent(new CustomEvent("indexchange", { detail: { index } }));
    if (this.#format === "pptx") {
      this.dispatchEvent(new CustomEvent("slidechange", { detail: { slide: index } }));
    } else if (this.#format === "xlsx") {
      const name = this.#doc !== null && this.#format === "xlsx"
        ? (this.#doc as XlsxDocument).sheets.list()[index]?.info.name
        : undefined;
      this.dispatchEvent(new CustomEvent("sheetchange", { detail: { sheet: name ?? index } }));
    }
  }

  async #render(index: number, generation: number): Promise<boolean> {
    const doc = this.#doc;
    const format = this.#format;
    if (doc === null || format === null) throw new Error("<aared-viewer>: no document is open");
    const renderGeneration = ++this.#renderGeneration;
    const stale = () => generation !== this.#generation || renderGeneration !== this.#renderGeneration;
    this.#slideRender?.destroy();
    this.#slideRender = null;
    if (format === "pptx") {
      const renderer = await import("../pptx/index.js");
      if (stale()) return false;
      const deck = doc as PptxDocument;
      const slide = deck.slides.get(index);
      this.#slideSize = { ...slide.info.size };
      const mediaResolver = (mediaId: string) => deck.media.get(mediaId);
      const rendered = await renderer.render({ slide, target: this.#page, mediaResolver });
      if (stale()) {
        rendered.destroy();
        return false;
      }
      this.#slideRender = rendered;
      this.#forwardRenderEvents(rendered);
    } else if (format === "xlsx") {
      const renderer = await import("../xlsx/index.js");
      if (stale()) return false;
      const workbook = doc as XlsxDocument;
      const sheet = workbook.sheets.get(index);
      this.#slideSize = null;
      const mediaResolver = (mediaId: string) => workbook.media.get(mediaId);
      const rendered = await renderer.render({ sheet, target: this.#page, mediaResolver, appearance: { theme: this.#resolvedTheme } });
      if (stale()) {
        rendered.destroy();
        return false;
      }
      this.#slideRender = rendered;
      this.#forwardRenderEvents(rendered);
    } else {
      const renderer = await import("../docx/index.js");
      if (stale()) return false;
      const document = doc as DocxDocument;
      this.#slideSize = null;
      const mediaResolver = (mediaId: string) => document.media.get(mediaId);
      this.#slideRender = await renderer.render({ document, target: this.#page, mediaResolver, appearance: { theme: this.#resolvedTheme } });
      this.#forwardRenderEvents(this.#slideRender!);
    }
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
      throw new RangeError(`<aared-viewer>: index must be a non-negative integer; received ${String(page)}`);
    }
    if (total !== undefined && page >= total) {
      throw new RangeError(`<aared-viewer>: index ${page} is outside this ${total}-unit document`);
    }
  }

  #progress(loaded: number, total: number | undefined, generation: number): void {
    if (generation !== this.#generation) return;
    this.dispatchEvent(new CustomEvent("progress", {
      detail: total === undefined ? { loaded } : { loaded, total },
    }));
  }

  #report(cause: unknown): void {
    const message = cause instanceof Error ? cause.message : String(cause);
    this.dispatchEvent(new CustomEvent("error", { detail: { message, cause } }));
  }

  #forwardRenderEvents(handle: RenderHandle): void {
    this.#renderAbortController?.abort();
    const controller = new AbortController();
    this.#renderAbortController = controller;
    const signal = controller.signal;
    const forward = (type: string) => {
      handle.addEventListener(type, (event) => {
        this.dispatchEvent(new CustomEvent(type, { detail: (event as CustomEvent).detail }));
      }, { signal });
    };
    for (const type of ["zoomchange", "linkactivate", "cellactivate", "selectionchange", "sheetchange"]) {
      forward(type);
    }
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
    previous.textContent = "Previous";
    previous.addEventListener("click", () => this.goTo(Math.max(0, this.index - 1)));
    const indicator = document.createElement("span");
    indicator.dataset.aaredControls = "index";
    const next = document.createElement("button");
    next.type = "button";
    next.textContent = "Next";
    next.addEventListener("click", () => {
      const total = this.#doc?.info.count ?? 0;
      this.goTo(Math.min(Math.max(total - 1, 0), this.index + 1));
    });
    const zoomOut = document.createElement("button");
    zoomOut.type = "button";
    zoomOut.textContent = "-";
    zoomOut.setAttribute("aria-label", "Zoom out");
    zoomOut.addEventListener("click", () => this.#setNumericZoom(1 / 1.25));
    const zoomIn = document.createElement("button");
    zoomIn.type = "button";
    zoomIn.textContent = "+";
    zoomIn.setAttribute("aria-label", "Zoom in");
    zoomIn.addEventListener("click", () => this.#setNumericZoom(1.25));
    const zoom = document.createElement("span");
    zoom.dataset.aaredControls = "zoom";
    this.#controls.append(previous, indicator, next, zoomOut, zoom, zoomIn);
  }

  #setNumericZoom(factor: number): void {
    const current = typeof this.zoom === "number" && Number.isFinite(this.zoom) ? this.zoom : 1;
    this.zoom = Math.min(3, Math.max(0.25, current * factor));
  }

  #updateChrome(): void {
    const theme = this.theme;
    const resolved = theme === "system"
      ? (typeof matchMedia !== "undefined" && matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light")
      : theme;
    const prevTheme = this.#resolvedTheme;
    this.#resolvedTheme = resolved;
    this.dataset.aaredTheme = resolved;
    // When the resolved theme changes and a document is open, re-render
    // the current unit so adaptive content adapts to the new chrome.
    if (resolved !== prevTheme && this.#doc !== null && this.#contentRenderTheme !== null) {
      void this.#render(this.index, this.#generation).catch(() => {});
    }
    const zoom = this.zoom;
    const numericZoom = typeof zoom === "number" && Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
    if (typeof zoom === "number") {
      this.#page.style.width = `${numericZoom * 100}%`;
      this.#page.style.maxWidth = "none";
    } else if (zoom === "fit-page" && this.#slideSize !== null && this.#stage.clientHeight > 0) {
      const width = Math.min(this.#stage.clientWidth, this.#stage.clientHeight * (this.#slideSize.width / this.#slideSize.height));
      this.#page.style.width = `${width}px`;
      this.#page.style.maxWidth = "100%";
    } else {
      this.#page.style.width = "100%";
      this.#page.style.maxWidth = "none";
    }
    this.#controls.hidden = !this.controls;
    const total = this.#doc?.info.count ?? 0;
    const page = this.index;
    const indicator = this.#controls.querySelector<HTMLElement>('[data-aared-controls="index"]');
    const zoomIndicator = this.#controls.querySelector<HTMLElement>('[data-aared-controls="zoom"]');
    if (indicator !== null) indicator.textContent = total === 0 ? "0 / 0" : `${page + 1} / ${total}`;
    if (zoomIndicator !== null) zoomIndicator.textContent = zoom === "fit-width" ? "Fit width" : zoom === "fit-page" ? "Fit page" : `${Math.round(numericZoom * 100)}%`;
  }
}

if (!customElements.get("aared-viewer")) {
  customElements.define("aared-viewer", AaredViewer);
}
