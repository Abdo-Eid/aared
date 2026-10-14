/** Word shape frames: browser flow inside an independently transformed paint surface. */
import { cssClass } from "../class-sheet.js";
import { resourceId } from "../resource-id.js";
import { FRAME_PAINT } from "../render-defaults.js";
import type { DocxTextBoxNode } from "../core/model.js";
import { geometryForPreset } from "../core/drawingml-shape.js";
import { PX_PER_INCH, ptToPx } from "../core/units.js";
import { loadFillPainter, loadedFillPainter } from "../fill-loader.js";
import type { MediaUrls } from "../media-url.js";
import type { AnchorFrame } from "./float-layout.js";

const FRAME = cssClass("position:relative;box-sizing:border-box;vertical-align:top;min-width:0");
const SURFACE = cssClass("position:relative;box-sizing:border-box;width:100%;height:100%;display:flex;flex-direction:column;justify-content:flex-start;overflow:clip;min-width:0;padding:0.05in 0.1in");
const CONTENT = cssClass("flex:none;min-width:0;display:flow-root;position:relative");
const GRAPHIC = cssClass("overflow:visible;pointer-events:none");
const NS = "http://www.w3.org/2000/svg";
interface State {
  model: DocxTextBoxNode;
  frame: HTMLElement;
  surface: HTMLElement;
  content: HTMLElement;
  contentFrame: AnchorFrame;
  urls: MediaUrls | undefined;
  alive: () => boolean;
  size?: { width: number; height: number };
  changed: Array<() => void>;
}
const states = new WeakMap<HTMLElement, State>();

/** Build without reading layout. The caller paints ordinary DOCX blocks into `content`. */
export function createTextBox(model: DocxTextBoxNode, anchored: boolean, urls: MediaUrls | undefined, alive: () => boolean): { frame: HTMLElement; content: HTMLElement; contentFrame: AnchorFrame } {
  const frame = document.createElement("div");
  frame.setAttribute("class", FRAME);
  frame.setAttribute("data-aared-element", "text-box");
  frame.setAttribute("data-aared-text-box-frame", "true");
  frame.setAttribute("data-aared-direction", model.direction);
  frame.style.display = anchored ? "block" : "inline-block";
  frame.style.direction = model.direction;
  frame.style.width = `${model.width * PX_PER_INCH}px`;
  frame.style.height = model.autofit === "resizeShape" ? "auto" : `${model.height * PX_PER_INCH}px`;
  frame.style.maxWidth = anchored ? "none" : "100%";
  if (model.wrapping === "none") {
    frame.style.width = "max-content";
    frame.style.minWidth = anchored ? `${model.width * PX_PER_INCH}px` : `min(100%, ${model.width * PX_PER_INCH}px)`;
  }
  const surface = document.createElement("div");
  surface.setAttribute("class", SURFACE);
  surface.setAttribute("data-aared-element", "text-box-surface");
  if (model.autofit === "resizeShape") surface.style.height = "auto";
  if (model.verticalAlign !== undefined) surface.style.justifyContent = model.verticalAlign === "middle" ? "safe center" : model.verticalAlign === "bottom" ? "safe flex-end" : "flex-start";
  const transform = model.transform;
  if (transform !== undefined) surface.style.transform = `rotate(${transform.rotation ?? 0}deg) scale(${transform.flipHorizontal ? -1 : 1}, ${transform.flipVertical ? -1 : 1})`;
  const rect = model.geometry?.textRect;
  const vb = model.geometry?.viewBox;
  const fractions = rect === undefined || vb === undefined ? { left: 0, top: 0, right: 0, bottom: 0 } : {
    left: rect.left / vb.width, right: 1 - rect.right / vb.width,
    top: rect.top / vb.height, bottom: 1 - rect.bottom / vb.height,
  };
  // Vertical percentage padding uses the containing width in CSS. Use a
  // separate content margin for geometry's vertical rectangle, settled below.
  for (const side of ["Left", "Right", "Top", "Bottom"] as const) {
    const key = side.toLowerCase() as "left" | "right" | "top" | "bottom";
    if (model.insets?.[key] === undefined && fractions[key] === 0) continue;
    const inset = (model.insets?.[key] ?? (key === "left" || key === "right" ? 0.1 : 0.05)) * PX_PER_INCH;
    surface.style[`padding${side}`] = side === "Left" || side === "Right" ? `calc(${fractions[key] * 100}% + ${inset}px)` : `${inset}px`;
  }
  const content = document.createElement("div");
  content.setAttribute("class", CONTENT);
  content.setAttribute("data-aared-element", "text-box-content");
  surface.append(content);
  frame.append(surface);
  if (model.fill?.kind === "solid") {
    frame.setAttribute("data-aared-shade", model.fill.color);
  }
  const contentFrame: AnchorFrame = { columnPx: Math.max(0, model.width * PX_PER_INCH - (model.insets?.left ?? 0.1) * PX_PER_INCH - (model.insets?.right ?? 0.1) * PX_PER_INCH), columnLeftPx: 0, resized: [] };
  const state: State = { model, frame, surface, content, contentFrame, urls, alive, changed: [] };
  states.set(frame, state);
  return { frame, content, contentFrame };
}

/** No-wrap must override paragraphs' explicit pre-wrap, without changing nested boxes. */
export function finishTextBox(frame: HTMLElement): void {
  const state = states.get(frame)!;
  if (state.model.wrapping !== "none") return;
  for (const line of state.content.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6')) {
    if (line.closest('[data-aared-text-box-frame]') === frame) line.style.whiteSpace = "pre";
  }
}

export function onTextBoxSize(frame: HTMLElement, changed: () => void): void {
  states.get(frame)?.changed.push(changed);
}

export function textBoxSize(frame: HTMLElement, fallback: { width: number; height: number }): { width: number; height: number } {
  return states.get(frame)?.size ?? fallback;
}

function paint(state: State, width: number, height: number): void {
  const { model, surface } = state;
  let geometry = model.geometry;
  // A resized round rectangle keeps circular corners, rather than stretching
  // its old short-side radius with the authored viewBox. Recover its pinned
  // adjustment from the resolved path; no preset catalog is needed in Word.
  if (geometry?.preset === "roundRect") {
    const radius = Number(/^M\s+([\d.eE+-]+)/.exec(geometry.path)?.[1]);
    if (Number.isFinite(radius)) geometry = geometryForPreset("roundRect", { width, height }, {
      adj: radius / Math.min(geometry.viewBox.width, geometry.viewBox.height) * 100000,
    }) ?? geometry;
  }
  const vb = geometry?.viewBox ?? { width, height };
  const svg = document.createElementNS(NS, "svg");
  svg.setAttribute("class", `${FRAME_PAINT} ${GRAPHIC}`);
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("data-aared-element", "text-box-graphic");
  svg.setAttribute("viewBox", `0 0 ${vb.width} ${vb.height}`);
  svg.setAttribute("preserveAspectRatio", "none");
  const outline = geometry?.path ?? `M 0 0 H ${vb.width} V ${vb.height} H 0 Z`;
  let fill = model.fill?.kind === "solid" ? model.fill.color : "none";
  const fills = loadedFillPainter();
  state.frame.removeAttribute("data-aared-fill-fallback");
  if (model.fill !== undefined && model.fill.kind !== "solid") {
    const painted = fills?.fillPaint(svg, resourceId("word-fill"), model.fill, {
      width: vb.width, height: vb.height, unitsPerPxX: vb.width / width, unitsPerPxY: vb.height / height, slideX: 0, slideY: 0,
    }, { urls: state.urls, ...(model.transform === undefined ? {} : { transform: model.transform }), outline });
    if (painted === undefined) {
      const reason = fills?.fillFallbackReason(model.fill, state.urls) ?? "painter-unavailable";
      svg.setAttribute("data-aared-fill-fallback", reason);
      state.frame.setAttribute("data-aared-fill-fallback", reason);
    }
    else fill = painted;
  }
  if (geometry?.fallback !== undefined && geometry.fallback !== "missing") svg.setAttribute("data-aared-geometry-fallback", geometry.fallback);
  const path = (d: string, paint: string, stroke: boolean): void => {
    const node = document.createElementNS(NS, "path");
    node.setAttribute("d", d);
    node.setAttribute("fill", paint);
    // Non-scaling stroke keeps the authored point width with a stretched viewBox.
    if (stroke && model.border !== undefined) {
      const border = model.border;
      const px = ptToPx(border.widthPt);
      node.setAttribute("stroke", border.color);
      node.setAttribute("stroke-width", String(px));
      node.setAttribute("vector-effect", "non-scaling-stroke");
      if (border.style === "dashed" || border.style === "dotted") node.setAttribute("stroke-dasharray", border.style === "dotted" ? `${px} ${px * 2}` : `${px * 4} ${px * 3}`);
    }
    svg.append(node);
  };
  if (geometry?.fillOnly !== undefined) path(geometry.fillOnly, fill, false);
  path(outline, fill, true);
  if (geometry?.strokeOnly !== undefined) path(geometry.strokeOnly, "none", true);
  surface.querySelector(':scope > [data-aared-element="text-box-graphic"]')?.remove();
  surface.prepend(svg);
}

/** Batched reads, then writes. A skipped lazy page retains its last valid paint. */
export function settleTextBoxes(root: HTMLElement, alive: () => boolean, observe = true): { dispose: () => void; ready: Promise<void> } {
  const boxes = Array.from(root.querySelectorAll<HTMLElement>('[data-aared-text-box-frame]')).map((frame) => states.get(frame)!).filter(Boolean);
  let disposed = false;
  const run = (): void => {
    if (disposed || !alive()) return;
    const sizes = boxes.map((state) => {
      if (state.frame.checkVisibility?.({ contentVisibilityAuto: true }) === false) return { state, width: 0, height: 0, contentWidth: 0, intrinsicWidth: 0 };
      // Computed border-box lengths are unzoomed and retain fractional pixels;
      // a transformed inner surface's bounding rect is deliberately never read.
      const style = getComputedStyle(state.frame);
      return { state, width: Number.parseFloat(style.width), height: Number.parseFloat(style.height), contentWidth: state.content.clientWidth, intrinsicWidth: state.content.scrollWidth };
    });
    for (const { state, width, height, contentWidth, intrinsicWidth } of sizes) {
      if (!(width > 0 && height > 0) || !state.alive()) continue;
      if (state.model.wrapping === "none") {
        const rect = state.model.geometry?.textRect;
        const vb = state.model.geometry?.viewBox;
        const share = rect === undefined || vb === undefined ? 0 : Math.max(0, Math.min(0.99, (rect.left + vb.width - rect.right) / vb.width));
        const insets = ((state.model.insets?.left ?? 0.1) + (state.model.insets?.right ?? 0.1)) * PX_PER_INCH;
        const wanted = Math.max(state.model.width * PX_PER_INCH, (intrinsicWidth + insets) / (1 - share));
        // Intrinsic sizing treats percentage padding as zero. Resolve that
        // geometry share once in the root's batched settle, including no-wrap.
        if (wanted > width + 0.5) state.frame.style.width = `${wanted}px`;
      }
      const changed = state.size?.width !== width || state.size.height !== height;
      const contentChanged = state.contentFrame.columnPx !== contentWidth;
      state.contentFrame.columnPx = contentWidth;
      if (contentChanged) for (const notify of state.contentFrame.resized ?? []) notify();
      if (!changed) continue;
      state.size = { width, height };
      const rect = state.model.geometry?.textRect;
      const vb = state.model.geometry?.viewBox;
      if (rect !== undefined && vb !== undefined) {
        // Autoheight solves H = content + insets + geometry's share of H.
        const share = Math.max(0, Math.min(0.99, rect.top / vb.height + 1 - rect.bottom / vb.height));
        const margins = (Number.parseFloat(state.content.style.marginTop) || 0) + (Number.parseFloat(state.content.style.marginBottom) || 0);
        const finalHeight = state.model.autofit === "resizeShape" ? (height - margins) / (1 - share) : height;
        state.content.style.marginTop = `${rect.top / vb.height * finalHeight}px`;
        state.content.style.marginBottom = `${(1 - rect.bottom / vb.height) * finalHeight}px`;
      }
      paint(state, width, height);
      for (const notify of state.changed) notify();
    }
  };
  run();
  // Geometry insets and nested frame updates settle together before pagination.
  run();
  run();
  const needsFills = boxes.some(({ model }) => model.fill !== undefined && model.fill.kind !== "solid");
  const ready = needsFills ? loadFillPainter().then(() => {
    if (disposed || !alive()) return;
    for (const state of boxes) if (state.size !== undefined && state.alive()) paint(state, state.size.width, state.size.height);
  }, () => undefined) : Promise.resolve();
  const observer = observe && boxes.length > 0 && typeof ResizeObserver === "function" ? new ResizeObserver(run) : undefined;
  for (const state of boxes) { observer?.observe(state.frame); observer?.observe(state.content); }
  return { ready, dispose: () => { disposed = true; observer?.disconnect(); } };
}
