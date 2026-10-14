/**
 * DrawingML fills -> `ShapeFillModel`, in the Worker. One reader for every
 * part that states a fill (a PPTX shape, table cell or background, a chart
 * area, a chart series), so a gradient, pattern or picture means the same
 * thing wherever it sits and the page paints it with one painter
 * (`src/fill-paint.ts`). Headless: no DOM, no format parser imports.
 */

import { attr, firstChild, numberAttr, onOff } from "./xml.js";
import { resolveImageSource } from "./image-source.js";
import type { XmlElement } from "./xml.js";
import { emuToInches } from "./units.js";
import { resolveDrawingColor } from "./theme.js";
import type { DrawingTheme } from "./theme.js";
import type {
  LinearGradientStopModel,
  PictureFillModel,
  PictureTileModel,
  RelativeRectModel,
  ShapeFillModel,
} from "./model.js";

function resolveGradientStops(gradient: XmlElement, theme: DrawingTheme): LinearGradientStopModel[] | undefined {
  const stopList = firstChild(gradient, "gsLst");
  if (stopList === undefined) return undefined;
  const stops: LinearGradientStopModel[] = [];
  for (const stop of stopList.children) {
    if (stop.local !== "gs") continue;
    const position = numberAttr(stop, "pos", "a:gs");
    const color = resolveDrawingColor(stop, theme);
    if (position === undefined || position < 0 || position > 100000 || color === undefined) return undefined;
    stops.push({ position: position / 100000, color });
  }
  return stops.length > 0 ? stops : undefined;
}

/** A `CT_RelativeRect` (`l`/`t`/`r`/`b` in 1/1000 %, default 0) as fractions; undefined when the element is absent. */
function relativeRect(element: XmlElement | undefined): RelativeRectModel | undefined {
  if (element === undefined) return undefined;
  const side = (name: string): number => (numberAttr(element, name, `a:${element.local}`) ?? 0) / 100000;
  return { left: side("l"), top: side("t"), right: side("r"), bottom: side("b") };
}

const NO_INSET: RelativeRectModel = { left: 0, top: 0, right: 0, bottom: 0 };

function isZeroRect(rect: RelativeRectModel | undefined): boolean {
  return rect === undefined || (rect.left === 0 && rect.top === 0 && rect.right === 0 && rect.bottom === 0);
}

export function resolveGradientFill(gradient: XmlElement, theme: DrawingTheme): ShapeFillModel {
  const path = firstChild(gradient, "path");
  if (path !== undefined) {
    const stops = resolveGradientStops(gradient, theme);
    if (stops === undefined) return { kind: "unsupported", source: "gradient-path" };
    const kind = attr(path, "path");
    if (kind !== "circle" && kind !== "rect" && kind !== "shape") return { kind: "unsupported", source: "gradient-path" };
    // Absent rectangles take the schema default: zero insets.
    return {
      kind: "path-gradient",
      path: kind,
      focus: relativeRect(firstChild(path, "fillToRect")) ?? { ...NO_INSET },
      extent: relativeRect(firstChild(gradient, "tileRect")) ?? { ...NO_INSET },
      stops,
    };
  }
  const linear = firstChild(gradient, "lin");
  if (linear === undefined) return { kind: "unsupported", source: "gradient-linear" };
  const stops = resolveGradientStops(gradient, theme);
  if (stops === undefined) return { kind: "unsupported", source: "gradient-linear" };
  const drawingMlAngle = numberAttr(linear, "ang", "a:lin") ?? 0;
  return {
    kind: "linear-gradient",
    angle: ((drawingMlAngle / 60000) % 360 + 360) % 360,
    // An absent `scaled` is false: PowerPoint draws slide 4's `<a:lin
    // ang="2700000"/>` with bands square to a true 45° across a 1.7:1 box
    // (Notes §28), not squashed to the box's diagonal.
    scaled: onOff(linear, "scaled") ?? false,
    rotatesWithShape: onOff(gradient, "rotWithShape") ?? true,
    stops,
  };
}

/**
 * The part a fill belongs to, so a picture fill's `r:embed` resolves against
 * that part's relationships: the slide, the layout or master a background or
 * shared graphic comes from, or a chart part.
 */
export interface FillOwner {
  relationships: ReadonlyArray<{ id: string; type: string; target: string; targetMode: string }>;
  mediaContentType: (target: string) => string;
  /** Inherited from the layout or master, so many slides share the picture. */
  inherited: boolean;
  /** Resolved parent group paint; null is an explicitly unfilled group. */
  groupFill?: ShapeFillModel | null;
}

/**
 * `a:pattFill`: a preset bitmap in two colours. The bitmaps themselves live
 * with the renderer; an unknown preset name is marked there, not here.
 * `fgClr`/`bgClr` default to black on white, as PowerPoint writes them.
 */
export function resolvePatternFill(pattern: XmlElement, theme: DrawingTheme): ShapeFillModel {
  const preset = attr(pattern, "prst");
  if (preset === undefined) return { kind: "unsupported", source: "pattern" };
  return {
    kind: "pattern",
    preset,
    foreground: resolveDrawingColor(firstChild(pattern, "fgClr"), theme) ?? "#000000",
    background: resolveDrawingColor(firstChild(pattern, "bgClr"), theme) ?? "#FFFFFF",
  };
}

const TILE_ALIGNMENTS = new Set(["tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"]);
const TILE_FLIPS = new Set(["none", "x", "y", "xy"]);

/**
 * `a:blipFill` as a fill: stretched (optionally into `a:stretch/a:fillRect`)
 * or tiled (`a:tile`: `sx`/`sy` in 1/1000 %, `tx`/`ty` in EMU). A picture the
 * owner's relationships cannot resolve stays `unsupported`.
 */
export function resolvePictureFill(blipFill: XmlElement, owner: FillOwner | undefined): ShapeFillModel {
  const unsupported = { kind: "unsupported", source: "image" } as const;
  if (owner === undefined) return unsupported;
  // An embedded SVG (`asvg:svgBlip`) paints first; the main `blip` stays as
  // its raster fallback, so an SVG Office cannot decode still shows Office's
  // own PNG of it.
  const source = resolveImageSource(firstChild(blipFill, "blip"), {
    relationships: owner.relationships,
    mediaContentType: owner.mediaContentType,
  });
  if (source === undefined) return unsupported;
  const out: PictureFillModel = {
    kind: "picture",
    mediaId: source.mediaId,
    contentType: source.contentType,
    ...(source.fallbackMediaId === undefined || source.fallbackContentType === undefined ? {} : {
      fallbackMediaId: source.fallbackMediaId,
      fallbackContentType: source.fallbackContentType,
    }),
    ...(owner.inherited ? { shared: true as const } : {}),
  };
  const crop = relativeRect(firstChild(blipFill, "srcRect"));
  if (!isZeroRect(crop)) out.crop = crop!;
  const tile = firstChild(blipFill, "tile");
  if (tile !== undefined) {
    const align = attr(tile, "algn") ?? "tl";
    const flip = attr(tile, "flip") ?? "none";
    out.tile = {
      scaleX: (numberAttr(tile, "sx", "a:tile") ?? 100000) / 100000,
      scaleY: (numberAttr(tile, "sy", "a:tile") ?? 100000) / 100000,
      offsetX: emuToInches(numberAttr(tile, "tx", "a:tile") ?? 0),
      offsetY: emuToInches(numberAttr(tile, "ty", "a:tile") ?? 0),
      align: (TILE_ALIGNMENTS.has(align) ? align : "tl") as PictureTileModel["align"],
      flip: (TILE_FLIPS.has(flip) ? flip : "none") as PictureTileModel["flip"],
    };
    return out;
  }
  const stretch = firstChild(blipFill, "stretch");
  const fillRect = relativeRect(stretch === undefined ? undefined : firstChild(stretch, "fillRect"));
  if (!isZeroRect(fillRect)) out.stretch = fillRect!;
  return out;
}

/** The `EG_FillProperties` child authored directly on `container`, if any. */
const FILL_PROPERTIES = new Set(["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"]);

/**
 * The one fill child `container` authors itself, and nothing else.
 *
 * Direct children only, on purpose: a container's own fill sits beside its
 * `a:ln`, its effects and its `a:extLst`, and a deep search finds a colour in
 * one of those and reports it as the shape's fill. Word gives every inserted
 * DrawingML shape a black outline, so `<a:ln><a:solidFill><a:prstClr
 * val="black"/>` is under almost every shape; a hidden outline adds
 * `a:extLst/a:ext/a14:hiddenLine` with the same payload. Both read as the shape's
 * own fill under a subtree search, which paints a text box solid black or in its
 * outline colour.
 */
export function directFillProperty(container: XmlElement | undefined): XmlElement | undefined {
  if (container === undefined) return undefined;
  for (const child of container.children) {
    if (FILL_PROPERTIES.has(child.local)) return child;
  }
  return undefined;
}

/**
 * The fill authored directly on `container` (`p:spPr`, `p:bgPr`, `a:tcPr`):
 * `null` for an explicit `a:noFill`, `undefined` when no fill is authored so
 * the caller's style fallback applies.
 */
export function resolveDirectFill(container: XmlElement, theme: DrawingTheme, owner: FillOwner | undefined): ShapeFillModel | null | undefined {
  const fill = directFillProperty(container);
  if (fill === undefined) return undefined;
  switch (fill.local) {
    case "noFill":
      return null;
    case "gradFill":
      return resolveGradientFill(fill, theme);
    case "blipFill":
      return resolvePictureFill(fill, owner);
    case "pattFill":
      return resolvePatternFill(fill, theme);
    case "grpFill":
      return owner?.groupFill ?? (owner?.groupFill === null ? null : { kind: "unsupported", source: "group" });
    case "solidFill":
      break;
    default:
      return undefined;
  }
  const color = resolveDrawingColor(fill, theme);
  return color === undefined ? undefined : { kind: "solid", color };
}
