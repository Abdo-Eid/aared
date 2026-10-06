/**
 * The controller skeleton every format render handle shares. Renderer-side.
 * A handle is an `EventTarget` that owns a zoom value (announced as
 * `zoomchange` when it changes) and a destroy: the painter attaches what
 * actually applies the zoom and tears the render down once it has painted,
 * and a call made before that only records the zoom to apply on first paint.
 * After `destroy()` the handle is inert: every method is a no-op that fires
 * nothing.
 *
 * The fit helpers live here too, because every format fits the same way:
 * measure the room its container offers with the content set aside (so a fit
 * can never read back its own result), then scale the content's natural size
 * into it.
 */

import type { DocumentFormat, RenderWarning, ZoomValue } from "./core/index.js";
import { abortError, throwIfAborted } from "./core/errors.js";

/**
 * Each format's starting zoom when the caller names none: a slide opens whole,
 * a sheet at 100% like Excel, a document to the width of its container. Read by
 * the three renderers and by the embed layers that mirror them.
 */
export const DEFAULT_ZOOM: Readonly<Record<DocumentFormat, ZoomValue>> = {
  pptx: "fit-page",
  xlsx: 1,
  docx: "fit-width",
};

/**
 * The warning every format gives when its file left drawings without an id (`code` is
 * `<format>-missing-id`): none when `count` is 0, else one entry. Word, Excel and PowerPoint
 * each call such a file damaged and offer to repair it, losing nothing visible; `outcome` says
 * what aared did. Plain ASCII, like every message that ships.
 */
export function missingIdWarning(code: string, count: number, noun: string, where: string, outcome: string): RenderWarning | undefined {
  if (count <= 0) return undefined;
  const many = count !== 1;
  return { code, message: `This file is damaged: ${count} ${noun}${many ? "s" : ""} ${where} ${many ? "have" : "has"} no id. ${outcome}` };
}

/** A numeric zoom the DOM can use: finite and positive, else 1. A `0` or `NaN` would otherwise collapse the content. */
export function safeScale(zoom: number): number {
  return zoom > 0 && Number.isFinite(zoom) ? zoom : 1;
}

/** Whether two zoom values paint the same: equal fit modes, or equal usable scales. */
export function sameZoom(a: ZoomValue, b: ZoomValue): boolean {
  if (typeof a === "number" && typeof b === "number") return safeScale(a) === safeScale(b);
  return a === b;
}

/** Room a container offers its content, in the container's own CSS pixels. */
export interface FitRoom {
  width: number;
  height: number;
}

/**
 * The room `container` offers `content`. Width is measured with the content in
 * place (a scrollbar it causes is real) and keeps fractional pixels (a 640.5px
 * box fits by exactly 640.5); padding, borders and a vertical scrollbar are
 * excluded. Height is measured with the content set aside: a container that
 * only grows to fit what it holds then reports 0 — it offers no height to fit
 * — instead of reading back the content's own height, which would make a fit
 * feed back into itself. A width-only fit (`withHeight: false`) skips that
 * measurement and its forced reflow, and reports height 0.
 */
export function measureFitRoom(container: HTMLElement, content: HTMLElement, withHeight = true): FitRoom {
  const style = getComputedStyle(container);
  const px = (value: string): number => Number.parseFloat(value) || 0;
  const frame = container.offsetWidth - container.clientWidth;
  const width = container.getBoundingClientRect().width - frame - px(style.paddingLeft) - px(style.paddingRight);
  if (!withHeight) return { width: Math.max(0, width), height: 0 };
  // Setting the content aside can shrink a page whose height is its content,
  // and the browser clamps every scroll position during that layout; nothing
  // paints before the content comes back, so putting the positions back
  // leaves the reader where they were (a Next click at the page's bottom used
  // to jump it up).
  const scrolled = scrollPositions(container);
  const display = content.style.display;
  content.style.display = "none";
  const height = container.clientHeight - px(style.paddingTop) - px(style.paddingBottom);
  content.style.display = display;
  for (const [element, top, left] of scrolled) {
    if (element.scrollTop !== top) element.scrollTop = top;
    if (element.scrollLeft !== left) element.scrollLeft = left;
  }
  return { width: Math.max(0, width), height: Math.max(0, height) };
}

/** Every ancestor that is scrolled (the document's scrolling element included), with where it is. */
function scrollPositions(from: HTMLElement): Array<[Element, number, number]> {
  const out: Array<[Element, number, number]> = [];
  for (let node: Element | null = from; node !== null; node = node.parentElement ?? ((node.getRootNode() as ShadowRoot).host ?? null)) {
    if (node.scrollTop !== 0 || node.scrollLeft !== 0) out.push([node, node.scrollTop, node.scrollLeft]);
  }
  const page = from.ownerDocument.scrollingElement;
  if (page !== null && !out.some(([element]) => element === page) && (page.scrollTop !== 0 || page.scrollLeft !== 0)) {
    out.push([page, page.scrollTop, page.scrollLeft]);
  }
  return out;
}

/**
 * The scale a fit mode gives `natural` content in `room`: `fit-width` matches
 * widths; `fit-page` also keeps the whole height in view, falling back to
 * `fit-width` when the room offers no height. `undefined` when there is no
 * width to fit yet (not laid out), so the caller keeps its last scale.
 */
export function fitScale(mode: "fit-width" | "fit-page", room: FitRoom, natural: FitRoom): number | undefined {
  if (room.width <= 0 || natural.width <= 0) return undefined;
  const byWidth = room.width / natural.width;
  if (mode === "fit-width" || room.height <= 0 || natural.height <= 0) return byWidth;
  return Math.min(byWidth, room.height / natural.height);
}

/**
 * An `EventTarget` whose `addEventListener`/`removeEventListener` know the
 * event each name carries (`M`: name to event type), so a listener reads
 * `event.detail` without a cast. Any other name still takes a plain
 * listener. Each handle's public interface extends it with its own map.
 */
export interface EventTargetOf<M> extends EventTarget {
  addEventListener<K extends keyof M & string>(type: K, listener: (event: M[K]) => void, options?: boolean | AddEventListenerOptions): void;
  addEventListener(type: string, listener: EventListenerOrEventListenerObject | null, options?: boolean | AddEventListenerOptions): void;
  removeEventListener<K extends keyof M & string>(type: K, listener: (event: M[K]) => void, options?: boolean | EventListenerOptions): void;
  removeEventListener(type: string, listener: EventListenerOrEventListenerObject | null, options?: boolean | EventListenerOptions): void;
}

/** What every render handle fires when its zoom changes. */
export type ZoomChangeEvent = CustomEvent<{ zoom: ZoomValue }>;

/**
 * Call `onReplaced` once `root` is no longer a child of `container`: another
 * render into the same target (or the host) put something else there. Every
 * format destroys its handle then, so a render that was replaced stops
 * observing, refitting and holding media URLs without the caller tracking it.
 * The handle's own swaps inside the container (the XLSX tab bar coming and
 * going) leave `root` in place and do nothing. `MutationObserver` reports
 * after the replacing call returns, so the old handle reads `destroyed` from
 * the next microtask on. Returns the disconnect for the handle's own destroy.
 */
export function onReplaced(container: Node, root: Node, replaced: () => void): () => void {
  const observer = new MutationObserver(() => {
    if (root.parentNode !== container) replaced();
  });
  observer.observe(container, { childList: true });
  return () => observer.disconnect();
}

/**
 * What a render waits for after it has painted (its font check), unless its
 * `signal` aborts first. Nobody holds a render that has not been handed back
 * yet, so a cancelled one tears itself down here: the handle is destroyed,
 * the nodes it put in `target` are taken out (only those still there, so a
 * newer render's are left alone), and the call rejects `aborted`. That makes
 * the rule one line for the caller, the way `fetch()` behaves: an aborted
 * render never hands back a handle. Once `render()` has resolved, the signal
 * is ignored and `destroy()` is the teardown. Used by the formats that paint
 * before they wait (PPTX, XLSX); DOCX waits first and resolves as it paints.
 */
export async function settleRender(
  render: { destroy(): void },
  target: Node,
  painted: readonly Node[],
  work: Promise<void>,
  signal: AbortSignal | undefined,
): Promise<void> {
  if (signal === undefined) return work;
  let cancel = (): void => {};
  const cancelled = new Promise<never>((_resolve, reject) => {
    cancel = () => reject(abortError());
  });
  signal.addEventListener("abort", cancel, { once: true });
  try {
    if (signal.aborted) cancel();
    await Promise.race([work, cancelled]);
    // Aborted while the wait was already ending: still before the handle is handed back.
    throwIfAborted(signal);
  } catch (error) {
    render.destroy();
    for (const node of painted) if (node.parentNode === target) target.removeChild(node);
    throw error;
  } finally {
    signal.removeEventListener("abort", cancel);
  }
}

export interface RenderHooks<TAppearance> {
  /** Re-applies the current zoom to the painted content; returns the scale it applied. */
  applyZoom: () => number;
  /** Disconnects observers and revokes what the render created. */
  destroy: () => void;
  applyAppearance?: (appearance: TAppearance | undefined) => void;
}

export abstract class RenderHandleBase<TAppearance> extends EventTarget {
  #zoom: ZoomValue;
  #scale = 1;
  #destroyed = false;
  #warnings: readonly RenderWarning[] = [];
  #applyZoom: () => number = () => this.#scale;
  #destroyHook: () => void = () => {};
  /** What `setAppearance` runs once the painter has attached it; a no-op before that. */
  protected applyAppearance: (appearance: TAppearance | undefined) => void = () => {};

  /**
   * `initialZoom` is applied on first paint without an event: a render's
   * starting zoom is its state, not a change. Formats pass their `DEFAULT_ZOOM`
   * when the caller gave none.
   */
  constructor(initialZoom: ZoomValue) {
    super();
    this.#zoom = initialZoom;
  }

  /** The requested zoom: a fit mode or an explicit scale. */
  get zoom(): ZoomValue {
    return this.#zoom;
  }

  /** The scale actually painted: the explicit scale, or what the fit mode computed. */
  get scale(): number {
    return this.#scale;
  }

  /**
   * Non-fatal things this render noticed about its own output (DOCX
   * pagination edge cases, a font this render couldn't actually paint with —
   * see `RenderWarning`). Empty unless the painter set some.
   */
  get warnings(): readonly RenderWarning[] {
    return this.#warnings;
  }

  /** @internal Set once by the painter once it has finished. */
  setWarnings(warnings: readonly RenderWarning[]): void {
    this.#warnings = warnings;
  }

  /** Change zoom. Always re-applies; fires `zoomchange` only when the value actually changes. */
  setZoom(zoom: ZoomValue): void {
    if (this.#destroyed) return;
    const changed = !sameZoom(this.#zoom, zoom);
    this.#zoom = zoom;
    this.refit();
    if (changed) this.dispatchEvent(new CustomEvent("zoomchange", { detail: { zoom } }));
  }

  /** @internal Re-apply the current zoom — painters call it when the room changes. */
  refit(): void {
    if (this.#destroyed) return;
    this.#scale = this.#applyZoom();
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyHook();
    this.#destroyed = true;
  }

  /** @internal Wired by the painter; `setZoom()` before this applies on first paint. */
  attach(hooks: RenderHooks<TAppearance>): void {
    this.#applyZoom = hooks.applyZoom;
    this.#destroyHook = hooks.destroy;
    if (hooks.applyAppearance !== undefined) this.applyAppearance = hooks.applyAppearance;
  }

  /** @internal Called by a painter tearing itself down (e.g. its DOM was replaced), so everything after is inert. */
  markDestroyed(): void {
    this.#destroyed = true;
  }

  get destroyed(): boolean {
    return this.#destroyed;
  }
}
