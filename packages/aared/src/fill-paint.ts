/**
 * DrawingML fills as SVG paint servers. One painter serves every fill box —
 * a shape's geometry, a table cell, the slide background, a chart's bars and
 * areas — so they can never disagree about what a fill looks like. It lives at
 * the package root, not under `pptx/`, because charts in sheets and Word
 * documents paint through it too. Each call adds whatever the
 * fill needs to the SVG's `<defs>` and returns the value for a `fill`
 * attribute, or `undefined` when the fill cannot be painted faithfully, so
 * the caller can mark it instead of inventing a colour.
 *
 * Boxes are described in their own SVG user units plus how many of those
 * units make one CSS pixel: pattern cells and picture tiles are sized in real
 * pixels (as PowerPoint sizes them), whatever the box's viewBox is.
 */
import type {
  LinearGradientFillModel,
  LinearGradientStopModel,
  PatternFillModel,
  PathGradientFillModel,
  PictureFillModel,
  ShapeFillModel,
  TransformModel,
} from "./core/model.js";
import { PX_PER_INCH } from "./core/units.js";
import { mixResolvedColorsLinear, mixResolvedColorsSrgb } from "./core/color.js";
import type { MediaUrls } from "./media-url.js";
import { PATTERN_PRESETS } from "./pattern-presets.js";
import { svgElement } from "./svg-element.js";

/** The area one fill paints, in the user units of the SVG it paints into. */
export interface FillBox {
  width: number;
  height: number;
  /** User units per CSS pixel (at 100 % zoom) along each axis. */
  unitsPerPxX: number;
  unitsPerPxY: number;
  /** The box's top-left in slide CSS pixels: patterns are anchored to the slide origin. */
  slideX: number;
  slideY: number;
  /**
   * The box's top-left in the SVG's user units, when the SVG holds many boxes
   * (a chart's shapes) rather than one that starts at its origin. Absent is 0.
   * Gradients and picture fills follow it; patterns tile from the origin and
   * stay put.
   */
  originX?: number;
  originY?: number;
  /** Resolve a scaled gradient against this paint box rather than the path's bounding box (group fill). */
  scaledToUserSpace?: boolean;
}

export interface FillPaintOptions {
  /** Media URLs for picture fills; without them a picture fill cannot paint. */
  urls: MediaUrls | undefined;
  /** The painted element's rotation/flips, for gradients that do not rotate with it. */
  transform?: TransformModel;
  /**
   * The outline a `path="shape"` gradient follows, as SVG path data in the
   * box's user units; boxes without one (table cells, the background) use
   * their own rectangle.
   */
  outline?: string;
}

/** Marks an SVG whose picture fill is still loading; removed once it paints. */
export const FILL_LOADING_ATTRIBUTE = "data-aared-fill-loading";

function numberText(value: number): string {
  return `${Math.round(value * 1_000_000) / 1_000_000}`;
}

function defsOf(svg: SVGSVGElement): SVGDefsElement {
  const existing = svg.querySelector(":scope > defs");
  if (existing instanceof SVGDefsElement) return existing;
  const defs = svgElement("defs");
  svg.prepend(defs);
  return defs;
}

/** Intermediate stops per segment: enough that straight sRGB runs between them stay within ~1-2 levels of PowerPoint's curve. */
const SEGMENT_SAMPLES = 16;

/** Nested outline copies for a `path="shape"` gradient: steps of 1/64 stay within a few colour levels. */
const SHAPE_GRADIENT_LEVELS = 64;

function smoothstep(value: number): number {
  return value * value * (3 - 2 * value);
}

/**
 * How PowerPoint blends between a gradient's stops, which depends on the
 * stops themselves. Measured against PowerPoint renders (Notes §28),
 * each within a few colour levels:
 *
 * - two stops: linear light, the weight eased by smoothstep — a
 *   blue-to-orange midpoint is a warm grey, not the olive of an sRGB blend;
 * - three or more, first and last colour different: plain sRGB, straight;
 * - three or more that return to their first colour (A→B→A): linear light,
 *   straight. (One linear black→white→black probe eased instead, so this
 *   case is the least certain.)
 */
type GradientBlend = "eased-linear-light" | "linear-light" | "srgb";

/** A fill's own blend wins (`srgb`: Excel's cell gradients, measured straight); else PowerPoint's rule above. */
function gradientBlend(stops: readonly LinearGradientStopModel[], blend?: "srgb"): GradientBlend {
  if (blend === "srgb") return "srgb";
  if (stops.length <= 2) return "eased-linear-light";
  return stops[0]!.color === stops[stops.length - 1]!.color ? "linear-light" : "srgb";
}

function mixStops(blend: GradientBlend, from: string, to: string, local: number): string {
  switch (blend) {
    case "eased-linear-light":
      return mixResolvedColorsLinear(from, to, smoothstep(local));
    case "linear-light":
      return mixResolvedColorsLinear(from, to, local);
    case "srgb":
      return mixResolvedColorsSrgb(from, to, local);
  }
}

/** The colour Office paints at weight `t` (0-1) of a gradient. */
export function gradientColorAt(stops: readonly LinearGradientStopModel[], t: number, fillBlend?: "srgb"): string {
  const first = stops[0]!;
  if (t <= first.position) return first.color;
  const blend = gradientBlend(stops, fillBlend);
  for (let index = 1; index < stops.length; index++) {
    const from = stops[index - 1]!;
    const to = stops[index]!;
    if (t <= to.position) {
      const span = to.position - from.position;
      return mixStops(blend, from.color, to.color, span > 0 ? (t - from.position) / span : 1);
    }
  }
  return stops[stops.length - 1]!.color;
}

/**
 * Gradient stops as PowerPoint blends them, for renderers that can only
 * interpolate linearly in sRGB (SVG, CSS). A plain-sRGB gradient needs only
 * its authored stops; the linear-light blends expand each segment into
 * intermediate stops carrying PowerPoint's colours.
 */
export function officeGradientStops(stops: readonly LinearGradientStopModel[], fillBlend?: "srgb"): LinearGradientStopModel[] {
  const blend = gradientBlend(stops, fillBlend);
  if (blend === "srgb") return [...stops];
  const out: LinearGradientStopModel[] = stops.length === 0 ? [] : [stops[0]!];
  for (let index = 1; index < stops.length; index++) {
    const from = stops[index - 1]!;
    const to = stops[index]!;
    for (let sample = 1; sample <= SEGMENT_SAMPLES; sample++) {
      const local = sample / SEGMENT_SAMPLES;
      out.push({
        position: from.position + (to.position - from.position) * local,
        color: sample === SEGMENT_SAMPLES ? to.color : mixStops(blend, from.color, to.color, local),
      });
    }
  }
  return out;
}

function appendStops(gradient: SVGElement, stops: readonly LinearGradientStopModel[], blend?: "srgb"): void {
  for (const stop of officeGradientStops(stops, blend)) {
    gradient.append(svgElement("stop", { offset: `${numberText(stop.position * 100)}%`, "stop-color": stop.color }));
  }
}

function gradientAngle(fill: LinearGradientFillModel, transform: TransformModel | undefined): number {
  if (fill.rotatesWithShape) return fill.angle;
  let angle = fill.angle - (transform?.rotation ?? 0);
  if (transform?.flipHorizontal === true) angle = 180 - angle;
  if (transform?.flipVertical === true) angle = -angle;
  return angle;
}

function linearGradient(id: string, fill: LinearGradientFillModel, box: FillBox, transform: TransformModel | undefined): SVGElement {
  const angle = (gradientAngle(fill, transform) * Math.PI) / 180;
  const horizontal = Math.cos(angle);
  const vertical = Math.sin(angle);
  const extent = fill.scaled
    ? (Math.abs(horizontal) + Math.abs(vertical)) / 2
    : (Math.abs(horizontal) * box.width + Math.abs(vertical) * box.height) / 2;
  const centerX = fill.scaled ? 0.5 : box.width / 2;
  const centerY = fill.scaled ? 0.5 : box.height / 2;
  const gradient = svgElement("linearGradient", {
    id,
    x1: centerX - horizontal * extent,
    y1: centerY - vertical * extent,
    x2: centerX + horizontal * extent,
    y2: centerY + vertical * extent,
  });
  if (!fill.scaled || box.scaledToUserSpace) gradient.setAttribute("gradientUnits", "userSpaceOnUse");
  if (fill.scaled && box.scaledToUserSpace) {
    for (const key of ["x1", "x2"] as const) gradient.setAttribute(key, numberText(Number(gradient.getAttribute(key)) * box.width));
    for (const key of ["y1", "y2"] as const) gradient.setAttribute(key, numberText(Number(gradient.getAttribute(key)) * box.height));
  }
  appendStops(gradient, fill.stops, fill.blend);
  return gradient;
}

/** A path gradient's extent and focus rectangles in the box's user units. */
function pathGradientBoxes(fill: PathGradientFillModel, box: FillBox) {
  const { width: w, height: h } = box;
  const x0 = fill.extent.left * w;
  const x1 = w - fill.extent.right * w;
  const y0 = fill.extent.top * h;
  const y1 = h - fill.extent.bottom * h;
  return {
    x0, x1, y0, y1,
    fx0: x0 + fill.focus.left * w,
    fx1: x1 - fill.focus.right * w,
    fy0: y0 + fill.focus.top * h,
    fy1: y1 - fill.focus.bottom * h,
  };
}

function userSpaceLinear(id: string, stops: readonly LinearGradientStopModel[], x1: number, y1: number, x2: number, y2: number, blend?: "srgb"): SVGElement {
  const gradient = svgElement("linearGradient", { id, gradientUnits: "userSpaceOnUse", x1, y1, x2, y2 });
  appendStops(gradient, stops, blend);
  return gradient;
}

/**
 * A true circle of radius `radiusPx` CSS pixels around (cx, cy), drawn in
 * user units that may scale x and y differently: the gradient is authored in
 * x-units and squashed back to a circle on y.
 */
function userSpaceCircle(id: string, stops: readonly LinearGradientStopModel[], cx: number, cy: number, radiusPx: number, box: FillBox, blend?: "srgb"): SVGElement {
  const ratio = box.unitsPerPxY / box.unitsPerPxX;
  const gradient = svgElement("radialGradient", { id, gradientUnits: "userSpaceOnUse", cx, cy, r: radiusPx * box.unitsPerPxX });
  if (Math.abs(ratio - 1) > 1e-9) {
    gradient.setAttribute("gradientTransform", `translate(0 ${numberText(cy)}) scale(1 ${numberText(ratio)}) translate(0 ${numberText(-cy)})`);
  }
  appendStops(gradient, stops, blend);
  return gradient;
}

/**
 * `path="circle"`: a true circle in slide units, radius half the extent's
 * diagonal, measured from the focus rectangle. A point focus is one radial
 * gradient. A focus with area splits the box into nine regions — the focus
 * itself in the first stop, four side bands grading straight out, and four
 * corner quarters grading radially from the focus corners — all sharing the
 * same radius, so neighbouring regions agree on every shared edge.
 */
function circleGradient(id: string, fill: PathGradientFillModel, box: FillBox): SVGElement {
  const { x0, x1, y0, y1, fx0, fx1, fy0, fy1 } = pathGradientBoxes(fill, box);
  const radiusPx = Math.hypot((x1 - x0) / box.unitsPerPxX, (y1 - y0) / box.unitsPerPxY) / 2;
  const epsilon = 1e-6 * Math.max(box.width, box.height);
  if (Math.abs(fx1 - fx0) <= epsilon && Math.abs(fy1 - fy0) <= epsilon) {
    return userSpaceCircle(id, fill.stops, (fx0 + fx1) / 2, (fy0 + fy1) / 2, radiusPx, box, fill.blend);
  }
  const { width: w, height: h } = box;
  const clampX = (value: number): number => Math.min(Math.max(value, 0), w);
  const clampY = (value: number): number => Math.min(Math.max(value, 0), h);
  const [left, right, top, bottom] = [clampX(fx0), clampX(fx1), clampY(fy0), clampY(fy1)];
  const rx = radiusPx * box.unitsPerPxX;
  const ry = radiusPx * box.unitsPerPxY;
  const pattern = svgElement("pattern", { id, patternUnits: "userSpaceOnUse", x: 0, y: 0, width: w, height: h });
  pattern.append(svgElement("rect", { x: 0, y: 0, width: w, height: h, fill: fill.stops[fill.stops.length - 1]!.color }));
  // The four corners, four sides and the focus, each as [x0, y0, x1, y1, paint server].
  const regions: Array<[rx0: number, ry0: number, rx1: number, ry1: number, server: () => SVGElement | string]> = [
    [0, 0, left, top, () => userSpaceCircle(`${id}-tl`, fill.stops, fx0, fy0, radiusPx, box, fill.blend)],
    [right, 0, w, top, () => userSpaceCircle(`${id}-tr`, fill.stops, fx1, fy0, radiusPx, box, fill.blend)],
    [0, bottom, left, h, () => userSpaceCircle(`${id}-bl`, fill.stops, fx0, fy1, radiusPx, box, fill.blend)],
    [right, bottom, w, h, () => userSpaceCircle(`${id}-br`, fill.stops, fx1, fy1, radiusPx, box, fill.blend)],
    [0, top, left, bottom, () => userSpaceLinear(`${id}-l`, fill.stops, fx0, 0, fx0 - rx, 0, fill.blend)],
    [right, top, w, bottom, () => userSpaceLinear(`${id}-r`, fill.stops, fx1, 0, fx1 + rx, 0, fill.blend)],
    [left, 0, right, top, () => userSpaceLinear(`${id}-t`, fill.stops, 0, fy0, 0, fy0 - ry, fill.blend)],
    [left, bottom, right, h, () => userSpaceLinear(`${id}-b`, fill.stops, 0, fy1, 0, fy1 + ry, fill.blend)],
    [left, top, right, bottom, () => fill.stops[0]!.color],
  ];
  for (const [rx0, ry0, rx1, ry1, server] of regions) {
    if (rx1 - rx0 <= 0 || ry1 - ry0 <= 0) continue;
    const paint = server();
    let fillValue: string;
    if (typeof paint === "string") fillValue = paint;
    else {
      pattern.append(paint);
      fillValue = `url(#${paint.getAttribute("id")})`;
    }
    // Axis-aligned regions whose neighbours agree on every shared edge:
    // hard edges avoid pale anti-aliasing seams between them.
    pattern.append(svgElement("rect", { x: rx0, y: ry0, width: rx1 - rx0, height: ry1 - ry0, fill: fillValue, "shape-rendering": "crispEdges" }));
  }
  return pattern;
}

/**
 * `path="shape"`: colour follows the fill's own outline. A point sits at
 * weight t when it lies on the outline scaled by t toward the focus centre,
 * so the fill is nested scaled copies of the outline, largest first, each in
 * the colour for its scale. Measured on a star: PowerPoint's weights match
 * the fraction along the centre-to-outline ray within 0.05.
 */
function shapeGradient(id: string, fill: PathGradientFillModel, box: FillBox, outline: string): SVGElement {
  const { fx0, fx1, fy0, fy1 } = pathGradientBoxes(fill, box);
  const cx = (fx0 + fx1) / 2;
  const cy = (fy0 + fy1) / 2;
  const pattern = svgElement("pattern", { id, patternUnits: "userSpaceOnUse", x: 0, y: 0, width: box.width, height: box.height });
  for (let level = SHAPE_GRADIENT_LEVELS; level >= 1; level--) {
    const scale = level / SHAPE_GRADIENT_LEVELS;
    pattern.append(svgElement("path", {
      d: outline,
      fill: gradientColorAt(fill.stops, scale, fill.blend),
      transform: `translate(${numberText(cx)} ${numberText(cy)}) scale(${numberText(scale)}) translate(${numberText(-cx)} ${numberText(-cy)})`,
    }));
  }
  return pattern;
}

/**
 * `path="rect"`: the focus rectangle in the first stop, then four sides that
 * each run linearly from the focus edge to the matching extent edge. Their
 * borders are the diagonals from focus corner to extent corner, where both
 * neighbours reach the same stop, so the sides meet without a visible step.
 * Outside the extent the last stop continues.
 */
function rectGradient(id: string, fill: PathGradientFillModel, box: FillBox): SVGElement {
  const { width: w, height: h } = box;
  const { x0, x1, y0, y1, fx0, fx1, fy0, fy1 } = pathGradientBoxes(fill, box);
  const first = fill.stops[0]!.color;
  const last = fill.stops[fill.stops.length - 1]!.color;
  const pattern = svgElement("pattern", { id, patternUnits: "userSpaceOnUse", x: 0, y: 0, width: w, height: h });
  pattern.append(svgElement("rect", { x: 0, y: 0, width: w, height: h, fill: last }));
  const sides: Array<[name: string, points: number[], from: [number, number], to: [number, number]]> = [
    ["left", [x0, y0, fx0, fy0, fx0, fy1, x0, y1], [fx0, 0], [x0, 0]],
    ["right", [fx1, fy0, x1, y0, x1, y1, fx1, fy1], [fx1, 0], [x1, 0]],
    ["top", [x0, y0, x1, y0, fx1, fy0, fx0, fy0], [0, fy0], [0, y0]],
    ["bottom", [fx0, fy1, fx1, fy1, x1, y1, x0, y1], [0, fy1], [0, y1]],
  ];
  for (const [name, points, from, to] of sides) {
    const run = Math.hypot(to[0] - from[0], to[1] - from[1]);
    if (run <= 0) continue;
    const gradient = svgElement("linearGradient", {
      id: `${id}-${name}`,
      gradientUnits: "userSpaceOnUse",
      x1: from[0],
      y1: from[1],
      x2: to[0],
      y2: to[1],
    });
    appendStops(gradient, fill.stops, fill.blend);
    pattern.append(gradient);
    // crispEdges: the neighbouring sides agree along the diagonal, so hard
    // edges show no step while anti-aliased ones leave a pale hairline seam.
    pattern.append(svgElement("polygon", { points: points.map(numberText).join(" "), fill: `url(#${id}-${name})`, "shape-rendering": "crispEdges" }));
  }
  if (fx1 > fx0 && fy1 > fy0) {
    pattern.append(svgElement("rect", { x: fx0, y: fy0, width: fx1 - fx0, height: fy1 - fy0, fill: first, "shape-rendering": "crispEdges" }));
  }
  return pattern;
}

/**
 * One preset bitmap as runs of foreground cells over a background tile. A
 * cell is one CSS pixel (the 8×8 tile is 1/12 inch) and the tile grid is
 * anchored to the slide origin, both as PowerPoint draws it.
 */
function patternFill(id: string, fill: PatternFillModel, box: FillBox): SVGElement | undefined {
  const bitmap = PATTERN_PRESETS[fill.preset];
  if (bitmap === undefined) return undefined;
  const cellX = box.unitsPerPxX;
  const cellY = box.unitsPerPxY;
  const pattern = svgElement("pattern", {
    id,
    patternUnits: "userSpaceOnUse",
    x: -box.slideX * box.unitsPerPxX,
    y: -box.slideY * box.unitsPerPxY,
    width: cellX * 8,
    height: cellY * 8,
  });
  pattern.append(svgElement("rect", { x: 0, y: 0, width: cellX * 8, height: cellY * 8, fill: fill.background }));
  for (let row = 0; row < 8; row++) {
    const bits = Number.parseInt(bitmap.slice(row * 2, row * 2 + 2), 16);
    for (let column = 0; column < 8;) {
      if ((bits & (0x80 >> column)) === 0) {
        column++;
        continue;
      }
      let end = column;
      while (end < 8 && (bits & (0x80 >> end)) !== 0) end++;
      // Anti-aliased on purpose: cells sit on whole pixels at 100 %, so they
      // stay sharp there, while a shrunk slide blends them into an even
      // texture as PowerPoint does. `crispEdges` would snap some cells to a
      // device pixel and drop others, banding the pattern (moiré).
      pattern.append(svgElement("rect", {
        x: column * cellX,
        y: row * cellY,
        width: (end - column) * cellX,
        height: cellY,
        fill: fill.foreground,
      }));
      column = end;
    }
  }
  return pattern;
}

/** Where the full (uncropped) image lands along one axis, given the target span and the crop fractions. */
function uncroppedSpan(start: number, end: number, cropBefore: number, cropAfter: number): [number, number] {
  const visible = 1 - cropBefore - cropAfter;
  const size = visible > 0 ? (end - start) / visible : end - start;
  return [start - cropBefore * size, size];
}

/**
 * A picture fill. Stretched pictures place the image (uncropped by
 * `crop`) over the `stretch` rectangle and let the pattern clip the rest.
 * Tiled pictures need the image's natural size, so the tile is built once
 * the image has decoded: its size is natural pixels at 96 DPI times the
 * tile scale, aligned to the box and mirrored per `flip`.
 */
function pictureFill(svg: SVGSVGElement, id: string, fill: PictureFillModel, box: FillBox, urls: MediaUrls): SVGElement {
  const pattern = svgElement("pattern", { id, patternUnits: "userSpaceOnUse", x: 0, y: 0, width: box.width, height: box.height });
  const crop = fill.crop ?? { left: 0, top: 0, right: 0, bottom: 0 };
  svg.setAttribute(FILL_LOADING_ATTRIBUTE, "picture");
  const settle = (failed: boolean) => {
    svg.removeAttribute(FILL_LOADING_ATTRIBUTE);
    if (failed) svg.dataset.aaredFillFallback = "media-unavailable";
  };
  // The SVG source paints first; Office's raster fallback covers a read or
  // decode failure, and the fill is marked only when neither paints.
  void urls.image({
    mediaId: fill.mediaId,
    contentType: fill.contentType,
    ...(fill.fallbackMediaId === undefined || fill.fallbackContentType === undefined ? {} : {
      fallbackMediaId: fill.fallbackMediaId,
      fallbackContentType: fill.fallbackContentType,
    }),
  }).then(async (resolved) => {
    if (resolved === undefined || urls.closed) return settle(resolved === undefined && !urls.closed);
    const sources = [resolved.url, ...(resolved.fallbackUrl === undefined ? [] : [resolved.fallbackUrl])];
    if (fill.tile === undefined) {
      const stretch = fill.stretch ?? { left: 0, top: 0, right: 0, bottom: 0 };
      const [x, width] = uncroppedSpan(stretch.left * box.width, box.width - stretch.right * box.width, crop.left, crop.right);
      const [y, height] = uncroppedSpan(stretch.top * box.height, box.height - stretch.bottom * box.height, crop.top, crop.bottom);
      const picture = svgElement("image", { href: sources[0]!, x, y, width, height, preserveAspectRatio: "none" });
      // At most one decode retry: a broken fallback marks the fill, never a loop.
      if (sources[1] !== undefined) {
        const fallback = sources[1];
        picture.addEventListener("error", () => {
          const retry = picture.getAttribute("href") !== fallback;
          if (retry) picture.setAttribute("href", fallback);
          else settle(true);
        });
      }
      pattern.append(picture);
      return settle(false);
    }
    let picked: string | undefined;
    let decodedWidth = 0;
    let decodedHeight = 0;
    for (const src of sources) {
      const image = new Image();
      image.src = src;
      try {
        await image.decode();
      } catch {
        continue;
      }
      if (urls.closed) return;
      picked = src;
      decodedWidth = image.naturalWidth;
      decodedHeight = image.naturalHeight;
      break;
    }
    if (picked === undefined) return settle(true);
    const url = picked;
    const tile = fill.tile;
    const naturalWidth = decodedWidth * (1 - crop.left - crop.right);
    const naturalHeight = decodedHeight * (1 - crop.top - crop.bottom);
    const tileWidth = naturalWidth * tile.scaleX * box.unitsPerPxX;
    const tileHeight = naturalHeight * tile.scaleY * box.unitsPerPxY;
    if (!(tileWidth > 0 && tileHeight > 0)) return settle(true);
    const horizontal = tile.align.endsWith("r") ? box.width - tileWidth : tile.align === "t" || tile.align === "ctr" || tile.align === "b" ? (box.width - tileWidth) / 2 : 0;
    const vertical = tile.align.startsWith("b") ? box.height - tileHeight : tile.align === "l" || tile.align === "ctr" || tile.align === "r" ? (box.height - tileHeight) / 2 : 0;
    const mirrorX = tile.flip === "x" || tile.flip === "xy";
    const mirrorY = tile.flip === "y" || tile.flip === "xy";
    pattern.setAttribute("x", numberText(horizontal + tile.offsetX * PX_PER_INCH * box.unitsPerPxX));
    pattern.setAttribute("y", numberText(vertical + tile.offsetY * PX_PER_INCH * box.unitsPerPxY));
    pattern.setAttribute("width", numberText(tileWidth * (mirrorX ? 2 : 1)));
    pattern.setAttribute("height", numberText(tileHeight * (mirrorY ? 2 : 1)));
    const [imageX, imageWidth] = uncroppedSpan(0, tileWidth, crop.left, crop.right);
    const [imageY, imageHeight] = uncroppedSpan(0, tileHeight, crop.top, crop.bottom);
    for (const [flipX, flipY] of [[false, false], [true, false], [false, true], [true, true]] as const) {
      if ((flipX && !mirrorX) || (flipY && !mirrorY)) continue;
      // Each copy is clipped to its own tile cell so a crop never bleeds into the next tile.
      const cell = svgElement("svg", {
        x: flipX ? tileWidth : 0,
        y: flipY ? tileHeight : 0,
        width: tileWidth,
        height: tileHeight,
        viewBox: `0 0 ${numberText(tileWidth)} ${numberText(tileHeight)}`,
        preserveAspectRatio: "none",
      });
      const picture = svgElement("image", { href: url, x: imageX, y: imageY, width: imageWidth, height: imageHeight, preserveAspectRatio: "none" });
      if (flipX || flipY) {
        picture.setAttribute("transform", `translate(${flipX ? numberText(tileWidth) : 0} ${flipY ? numberText(tileHeight) : 0}) scale(${flipX ? -1 : 1} ${flipY ? -1 : 1})`);
      }
      cell.append(picture);
      pattern.append(cell);
    }
    settle(false);
  }).catch(() => {
    if (!urls.closed) settle(true);
  });
  return pattern;
}

/**
 * Add `fill`'s paint server to `svg` under `id` and return the `fill`
 * attribute value; `undefined` means this fill cannot be painted (an
 * `unsupported` source, an unknown pattern preset, a picture with no media
 * resolver) and the caller should mark it.
 */
export function fillPaint(svg: SVGSVGElement, id: string, fill: ShapeFillModel, box: FillBox, options: FillPaintOptions): string | undefined {
  let server: SVGElement | undefined;
  switch (fill.kind) {
    case "solid":
      return fill.color;
    case "linear-gradient":
      server = linearGradient(id, fill, box, options.transform);
      break;
    case "path-gradient":
      server = fill.path === "circle"
        ? circleGradient(id, fill, box)
        : fill.path === "rect"
          ? rectGradient(id, fill, box)
          : shapeGradient(id, fill, box, options.outline ?? `M 0 0 H ${numberText(box.width)} V ${numberText(box.height)} H 0 Z`);
      break;
    case "pattern":
      server = patternFill(id, fill, box);
      break;
    case "picture":
      server = options.urls === undefined ? undefined : pictureFill(svg, id, fill, box, options.urls);
      break;
    case "unsupported":
      return undefined;
  }
  if (server === undefined) return undefined;
  placeServer(server, box, fill);
  defsOf(svg).append(server);
  return `url(#${id})`;
}

/**
 * Move a paint server onto a box that does not start at the SVG's origin. The
 * servers are built in the box's own coordinates (0,0 to width,height). A
 * pattern tile grid is anchored to the origin by design, and a scaled linear
 * gradient already follows the painted shape's bounding box.
 */
function placeServer(server: SVGElement, box: FillBox, fill: ShapeFillModel): void {
  const x = box.originX ?? 0;
  const y = box.originY ?? 0;
  if ((x === 0 && y === 0) || fill.kind === "pattern" || (fill.kind === "linear-gradient" && fill.scaled && !box.scaledToUserSpace)) return;
  const name = server.localName === "pattern" ? "patternTransform" : "gradientTransform";
  const move = `translate(${numberText(x)} ${numberText(y)})`;
  const existing = server.getAttribute(name);
  server.setAttribute(name, existing === null ? move : `${move} ${existing}`);
}

/** The marker value for a fill `fillPaint` could not paint. */
export function fillFallbackReason(fill: ShapeFillModel, urls: MediaUrls | undefined): string {
  if (fill.kind === "unsupported") return fill.source;
  if (fill.kind === "picture" && urls === undefined) return "missing-resolver";
  return fill.kind;
}
