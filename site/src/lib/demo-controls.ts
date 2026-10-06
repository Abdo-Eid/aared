import type { ZoomValue } from "aared/core";

/**
 * Docs plumbing for the playground demos: the paging and zoom bar under each
 * stage, built the way an app builds its own. Nothing here is aared API.
 * `viewerToolbar()` is the Start page's "Build your own toolbar" recipe over a
 * viewer (`preview()`, the hooks); the full-control demos wire the same bar
 * to their own render handles with the helpers below.
 */

/** What `viewerToolbar()` uses of a viewer: `preview()`'s handle and the hooks' `viewer` both have it. */
interface ToolbarViewer extends EventTarget {
  readonly index: number;
  readonly zoom: ZoomValue;
  readonly info: { readonly count: number };
  next(): Promise<void>;
  previous(): Promise<void>;
  zoomIn(step?: number): void;
  zoomOut(step?: number): void;
}

/** The viewer's toolbar methods behind buttons; returns the bar's teardown. */
export function viewerToolbar(stage: HTMLElement, viewer: ToolbarViewer): () => void {
  const position = label();
  const zoom = label();
  // A failed move also fires the viewer's `error` event; the rejection needs nothing more.
  const previous = button("Previous", () => void viewer.previous().catch(() => {}));
  const next = button("Next", () => void viewer.next().catch(() => {}));
  const update = (): void => {
    position.textContent = `${viewer.index + 1} / ${viewer.info.count}`;
    previous.disabled = viewer.index <= 0;
    next.disabled = viewer.index >= viewer.info.count - 1;
    zoom.textContent = zoomLabel(viewer.zoom);
  };
  viewer.addEventListener("indexchange", update);
  viewer.addEventListener("zoomchange", update);
  update();
  // A document is one unit: paging has nothing to move between.
  const paging = viewer.info.count > 1 ? [previous, position, next] : [];
  const unmount = mountBar(stage, ...paging, button("-", () => viewer.zoomOut(), "Zoom out"), zoom, button("+", () => viewer.zoomIn(), "Zoom in"));
  return () => {
    viewer.removeEventListener("indexchange", update);
    viewer.removeEventListener("zoomchange", update);
    unmount();
  };
}

/** Put a bar holding `items` right under `stage` (styled in demos.css, joined to it); returns its removal. */
export function mountBar(stage: HTMLElement, ...items: HTMLElement[]): () => void {
  const bar = document.createElement("div");
  bar.className = "demo-controls";
  bar.setAttribute("role", "toolbar");
  bar.setAttribute("aria-label", "Paging and zoom");
  bar.append(...items);
  stage.after(bar);
  return () => bar.remove();
}

export function button(text: string, onClick: () => void, ariaLabel?: string): HTMLButtonElement {
  const element = document.createElement("button");
  element.type = "button";
  element.textContent = text;
  if (ariaLabel !== undefined) element.setAttribute("aria-label", ariaLabel);
  element.addEventListener("click", onClick);
  return element;
}

export function label(): HTMLSpanElement {
  const element = document.createElement("span");
  element.className = "demo-controls-label";
  return element;
}

export function zoomLabel(zoom: ZoomValue): string {
  if (zoom === "fit-width") return "Fit width";
  if (zoom === "fit-page") return "Fit page";
  return `${Math.round(zoom * 100)}%`;
}

/**
 * A zoom step on a render handle, which has no `zoomIn()`: what the viewer's
 * own does, by hand. Add `delta` (0.1 is 10 points) to the zoom, or to the
 * painted scale in a fit mode, within 25%–300%, rounded to whole percents.
 */
export function stepZoom(handle: { readonly zoom: ZoomValue; readonly scale: number; setZoom(zoom: ZoomValue): void }, delta: number): void {
  const base = typeof handle.zoom === "number" ? handle.zoom : handle.scale;
  handle.setZoom(Math.round(Math.min(3, Math.max(0.25, base + delta)) * 100) / 100);
}
