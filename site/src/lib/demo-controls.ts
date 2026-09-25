import type { ZoomValue } from "aared/core";

/**
 * Docs plumbing (not shown as a snippet): the same paging and zoom bar
 * `<aared-viewer controls>` draws, for the demos that drive aared
 * themselves. It talks to a small controller: a `PreviewHandle` is one as-is
 * (React hook, one-call `preview()`); the full-control demos wrap their own
 * `render()` handles in one (`slideController`, `sheetController`).
 */
export interface DemoController extends EventTarget {
  readonly index: number;
  readonly zoom: ZoomValue;
  readonly scale: number;
  goTo(index: number): Promise<void>;
  setZoom(zoom: ZoomValue): void;
}

const ZOOM_STEP = 1.25;
const MIN_ZOOM = 0.25;
const MAX_ZOOM = 3;

function zoomLabel(zoom: ZoomValue): string {
  if (zoom === "fit-width") return "Fit width";
  if (zoom === "fit-page") return "Fit page";
  return `${Math.round(zoom * 100)}%`;
}

/** Add the bar right after `stage`, for `count` units; returns its teardown. */
export function mountDemoControls(stage: HTMLElement, controller: DemoController, count: number): () => void {
  const bar = document.createElement("div");
  bar.className = "demo-controls";
  bar.setAttribute("role", "toolbar");
  bar.setAttribute("aria-label", "Paging and zoom");

  const button = (text: string, label: string | null, onClick: () => void): HTMLButtonElement => {
    const element = document.createElement("button");
    element.type = "button";
    element.textContent = text;
    if (label !== null) element.setAttribute("aria-label", label);
    element.addEventListener("click", onClick);
    return element;
  };
  const text = (): HTMLSpanElement => {
    const element = document.createElement("span");
    element.className = "demo-controls-label";
    return element;
  };

  const previous = button("Previous", null, () => void controller.goTo(Math.max(0, controller.index - 1)).catch(() => {}));
  const position = text();
  const next = button("Next", null, () => void controller.goTo(Math.min(count - 1, controller.index + 1)).catch(() => {}));
  // +/− step from the numeric zoom, or from the scale a fit mode painted.
  const step = (factor: number): void => {
    const base = typeof controller.zoom === "number" ? controller.zoom : controller.scale;
    controller.setZoom(Math.round(Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, base * factor)) * 100) / 100);
  };
  const zoomOut = button("-", "Zoom out", () => step(1 / ZOOM_STEP));
  const zoom = text();
  const zoomIn = button("+", "Zoom in", () => step(ZOOM_STEP));

  const sync = (): void => {
    position.textContent = `${controller.index + 1} / ${count}`;
    previous.disabled = controller.index <= 0;
    next.disabled = controller.index >= count - 1;
    zoom.textContent = zoomLabel(controller.zoom);
  };
  // A document is one unit: paging has nothing to move between.
  if (count > 1) bar.append(previous, position, next);
  bar.append(zoomOut, zoom, zoomIn);
  controller.addEventListener("indexchange", sync);
  controller.addEventListener("zoomchange", sync);
  sync();
  stage.after(bar);

  return () => {
    controller.removeEventListener("indexchange", sync);
    controller.removeEventListener("zoomchange", sync);
    bar.remove();
  };
}

/** A render handle that zooms: every format's `render()` result is one. */
interface ZoomingHandle extends EventTarget {
  readonly zoom: ZoomValue;
  readonly scale: number;
  setZoom(zoom: ZoomValue): void;
  destroy(): void;
}

/**
 * A controller over per-unit `render()` calls (PPTX: one handle per slide):
 * going to a unit renders it with the current zoom, then drops the old handle.
 */
export class UnitController<H extends ZoomingHandle> extends EventTarget implements DemoController {
  #index = 0;
  #handle: H;
  #listening: AbortController | null = null;

  constructor(first: H, private readonly renderUnit: (index: number, zoom: ZoomValue) => Promise<H>) {
    super();
    this.#handle = first;
    this.#listen();
  }

  get handle(): H {
    return this.#handle;
  }
  get index(): number {
    return this.#index;
  }
  get zoom(): ZoomValue {
    return this.#handle.zoom;
  }
  get scale(): number {
    return this.#handle.scale;
  }

  async goTo(index: number): Promise<void> {
    if (index === this.#index) return;
    const next = await this.renderUnit(index, this.#handle.zoom);
    this.#handle.destroy();
    this.#handle = next;
    this.#index = index;
    this.#listen();
    this.dispatchEvent(new Event("indexchange"));
  }

  setZoom(zoom: ZoomValue): void {
    this.#handle.setZoom(zoom);
  }

  destroy(): void {
    this.#listening?.abort();
    this.#handle.destroy();
  }

  #listen(): void {
    this.#listening?.abort();
    this.#listening = new AbortController();
    this.#handle.addEventListener("zoomchange", () => this.dispatchEvent(new Event("zoomchange")), { signal: this.#listening.signal });
  }
}
