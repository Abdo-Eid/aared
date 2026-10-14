import type { LocalViewBoxModel, ShapeGeometryModel, TransformModel } from "./model.js";
import { attr, firstChild, numberAttr, onOff, ownString } from "./xml.js";
import type { XmlElement } from "./xml.js";
import { encodeGeometry } from "./drawingml-program.js";

export type GeometryFallback = NonNullable<ShapeGeometryModel["fallback"]>;

export interface DrawingMlShapeSize {
  width?: number;
  height?: number;
}

function localViewBox(size: DrawingMlShapeSize | undefined): LocalViewBoxModel {
  // SVG viewBoxes must be non-zero. A zero-height DrawingML line still paints
  // through its stroke, so retain its horizontal extent with a one-unit axis.
  return {
    width: Math.max(size?.width ?? 100000, 1),
    height: Math.max(size?.height ?? 100000, 1),
  };
}

function pin(values: Readonly<Record<string, number>>, name: string, defaultValue: number, maximum: number): number {
  return Math.min(Math.max(values[name] ?? defaultValue, 0), maximum);
}

/**
 * Preset paths follow the ECMA-376 `presetShapeDefinitions.xml` formulas
 * (Part 1, fifth edition — OfficeOpenXML-DrawingMLGeometries.zip). Guide
 * bases matter: several insets/skews are `ss` (short-side) based with
 * `maxAdj` clamps, not width-based, and the winding matches Office's.
 */
function pathForPreset(preset: string, viewBox: LocalViewBoxModel, adjustments: Readonly<Record<string, number>>): string | undefined {
  const { width, height } = viewBox;
  const shortSide = Math.min(width, height);
  switch (preset) {
    case "rect":
      return `M 0 0 H ${width} V ${height} H 0 Z`;
    case "roundRect": {
      const radius = (shortSide * pin(adjustments, "adj", 16667, 50000)) / 100000;
      return `M ${radius} 0 H ${width - radius} A ${radius} ${radius} 0 0 1 ${width} ${radius} V ${height - radius} A ${radius} ${radius} 0 0 1 ${width - radius} ${height} H ${radius} A ${radius} ${radius} 0 0 1 0 ${height - radius} V ${radius} A ${radius} ${radius} 0 0 1 ${radius} 0 Z`;
    }
    case "ellipse":
      // Two half-arcs between the horizontal extremes: robust at every
      // aspect, unlike a single near-closed arc that paints as zero-area.
      return `M 0 ${height / 2} A ${width / 2} ${height / 2} 0 0 1 ${width} ${height / 2} A ${width / 2} ${height / 2} 0 0 1 0 ${height / 2} Z`;
    case "line":
    case "straightConnector1":
      // A straight connector is a line: `M l t L r b` in the preset catalog.
      return `M 0 0 L ${width} ${height}`;
    case "triangle": {
      const apex = (width * pin(adjustments, "adj", 50000, 100000)) / 100000;
      return `M 0 ${height} L ${apex} 0 L ${width} ${height} Z`;
    }
    case "rtTriangle":
      // ECMA winding: vertical left edge, right angle at the bottom-left,
      // hypotenuse top-left to bottom-right.
      return `M 0 ${height} L 0 0 L ${width} ${height} Z`;
    case "diamond":
      return `M ${width / 2} 0 L ${width} ${height / 2} L ${width / 2} ${height} L 0 ${height / 2} Z`;
    case "parallelogram": {
      const skew = (shortSide * pin(adjustments, "adj", 25000, (100000 * width) / shortSide)) / 100000;
      return `M 0 ${height} L ${skew} 0 L ${width} 0 L ${width - skew} ${height} Z`;
    }
    case "hexagon": {
      const inset = (shortSide * pin(adjustments, "adj", 25000, (50000 * width) / shortSide)) / 100000;
      return `M 0 ${height / 2} L ${inset} 0 L ${width - inset} 0 L ${width} ${height / 2} L ${width - inset} ${height} L ${inset} ${height} Z`;
    }
    case "octagon": {
      const inset = (shortSide * pin(adjustments, "adj", 29289, 50000)) / 100000;
      return `M 0 ${inset} L ${inset} 0 L ${width - inset} 0 L ${width} ${inset} L ${width} ${height - inset} L ${width - inset} ${height} L ${inset} ${height} L 0 ${height - inset} Z`;
    }
    case "star5": {
      // ECMA star5: a 5-point star stretched over an hf/vf-enlarged ellipse
      // whose center sits below the box middle so the top point touches y=0.
      const a = pin(adjustments, "adj", 19098, 50000);
      const hf = (adjustments.hf ?? 105146) / 100000;
      const vf = (adjustments.vf ?? 110557) / 100000;
      const swd2 = (width / 2) * hf;
      const shd2 = (height / 2) * vf;
      const svc = (height / 2) * vf;
      const cos = (angle: number): number => Math.cos((angle * Math.PI) / 180);
      const sin = (angle: number): number => Math.sin((angle * Math.PI) / 180);
      const dx1 = swd2 * cos(18);
      const dx2 = swd2 * cos(306);
      const dy1 = shd2 * sin(18);
      const dy2 = shd2 * sin(306);
      const x1 = width / 2 - dx1;
      const x2 = width / 2 - dx2;
      const x3 = width / 2 + dx2;
      const x4 = width / 2 + dx1;
      const y1 = svc - dy1;
      const y2 = svc - dy2;
      const iwd2 = (swd2 * a) / 50000;
      const ihd2 = (shd2 * a) / 50000;
      const sdx1 = iwd2 * cos(342);
      const sdx2 = iwd2 * cos(54);
      const sdy1 = ihd2 * sin(54);
      const sdy2 = ihd2 * sin(342);
      const sx1 = width / 2 - sdx1;
      const sx2 = width / 2 - sdx2;
      const sx3 = width / 2 + sdx2;
      const sx4 = width / 2 + sdx1;
      const sy1 = svc - sdy1;
      const sy2 = svc - sdy2;
      const sy3 = svc + ihd2;
      return [
        `M ${x1} ${y1}`,
        `L ${sx2} ${sy1}`,
        `L ${width / 2} 0`,
        `L ${sx3} ${sy1}`,
        `L ${x4} ${y1}`,
        `L ${sx4} ${sy2}`,
        `L ${x3} ${y2}`,
        `L ${width / 2} ${sy3}`,
        `L ${x2} ${y2}`,
        `L ${sx1} ${sy2}`,
        "Z",
      ].join(" ");
    }
    default:
      return undefined;
  }
}

/**
 * ECMA preset `a:rect` formulas, using the same pinned adjustments as the
 * outline. Measured against PowerPoint's own laid-out text edges (probe
 * `pptx-text-rectangle`): every preset here matched within 0.01pt of the value
 * these formulas produce, after Office's authored `bodyPr` insets are added.
 */
function textRectForPreset(preset: string, { width: w, height: h }: LocalViewBoxModel, adjustments: Readonly<Record<string, number>>): ShapeGeometryModel["textRect"] {
  const ss = Math.min(w, h);
  const inset = (x: number, y = x): NonNullable<ShapeGeometryModel["textRect"]> => ({ left: x, top: y, right: w - x, bottom: h - y });
  const cos = (degrees: number): number => Math.cos((degrees * Math.PI) / 180);
  const sin = (degrees: number): number => Math.sin((degrees * Math.PI) / 180);
  switch (preset) {
    case "rect":
    case "straightConnector1": return inset(0);
    case "ellipse": return inset(w * (1 - Math.SQRT1_2) / 2, h * (1 - Math.SQRT1_2) / 2);
    case "roundRect":
      // The official guide uses the rounded constant 29289/100000, rather
      // than the exact 1 - cos(45deg), for the corner's text inset.
      return inset(ss * pin(adjustments, "adj", 16667, 50000) / 100000 * 29289 / 100000);
    case "diamond": return inset(w / 4, h / 4);
    case "triangle": {
      // `x1 = w * a / 200000`, `x3 = x1 + wd2`, `t = vc`.
      const x1 = w * pin(adjustments, "adj", 50000, 100000) / 200000;
      return { left: x1, top: h / 2, right: x1 + w / 2, bottom: h };
    }
    // `l = wd12` (a twelfth of the width), `t = h * 7 / 12`, `r = w * 7 / 12`,
    // `b = h * 11 / 12`.
    case "rtTriangle": return { left: w / 12, top: h * 7 / 12, right: w * 7 / 12, bottom: h * 11 / 12 };
    case "parallelogram": {
      // `maxAdj = 100000 w / ss`, `q2 = (1 + 5 a / maxAdj) / 12`.
      const max = 100000 * w / ss;
      const ratio = (1 + 5 * pin(adjustments, "adj", 25000, max) / max) / 12;
      return inset(w * ratio, h * ratio);
    }
    case "hexagon": {
      // `q1 = maxAdj * -1 / 2`, `q2 = a + q1` (clamped at 0), `q3 = q2 ? 4 : 2`,
      // `q4 = q2 ? 3 : 2`, `q5 = q2 ? q1 : 0`, `q6 = (a + q5) / q1`,
      // `q8 = q3 - q6 * q4`. An earlier version compared `a + q1` against 0 and
      // substituted `a` for `q5`, which is a different expression; it read
      // 16.67pt where PowerPoint lays the column at 36.67pt.
      const max = 50000 * w / ss;
      const a = pin(adjustments, "adj", 25000, max);
      const q1 = (-max / 2);
      const q2 = Math.max(a + q1, 0);
      const q3 = q2 > 0 ? 4 : 2;
      const q4 = q2 > 0 ? 3 : 2;
      const q5 = q2 > 0 ? q1 : 0;
      const q8 = q3 - (a + q5) / q1 * q4;
      return inset(w * q8 / 24, h * q8 / 24);
    }
    case "octagon": return inset(ss * pin(adjustments, "adj", 29289, 50000) / 200000);
    case "star5": {
      // ECMA star5 `a:rect l="sx1" t="sy1" r="sx4" b="sy3"`: `swd2 = wd2*hf/100000`,
      // `shd2 = hd2*vf/100000`, `svc = vc*vf/100000`, `iwd2 = swd2*a/50000`,
      // `ihd2 = shd2*a/50000`, `sdx1 = iwd2*cos(342deg)`, `sdy1 = ihd2*sin(54deg)`,
      // `sx1 = hc-sdx1`, `sx4 = hc+sdx1`, `sy1 = svc-sdy1`, `sy3 = svc+ihd2`.
      // `cos`/`sin` take degrees: radius multiplies the trig result, never the angle.
      const a = pin(adjustments, "adj", 19098, 50000);
      const hf = (adjustments.hf ?? 105146) / 100000;
      const vf = (adjustments.vf ?? 110557) / 100000;
      const swd2 = (w / 2) * hf;
      const shd2 = (h / 2) * vf;
      const svc = (h / 2) * vf;
      const iwd2 = (swd2 * a) / 50000;
      const ihd2 = (shd2 * a) / 50000;
      const sdx1 = iwd2 * cos(20520000 / 60000);
      const sdy1 = ihd2 * sin(3240000 / 60000);
      return {
        left: w / 2 - sdx1,
        top: svc - sdy1,
        right: w / 2 + sdx1,
        bottom: svc + ihd2,
      };
    }
    default: return undefined;
  }
}

/** Literal `a:avLst` overrides; formula guides remain the geometry program's concern. */
export function drawingGeometryAdjustments(prstGeom: XmlElement | undefined): Record<string, number> {
  const out: Record<string, number> = {};
  const list = prstGeom === undefined ? undefined : firstChild(prstGeom, "avLst");
  for (const guide of list?.children ?? []) {
    if (guide.local !== "gd") continue;
    const name = attr(guide, "name");
    const formula = attr(guide, "fmla");
    const match = formula === undefined ? null : /^val\s+([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*$/.exec(formula);
    const value = match === null ? NaN : Number(match[1]);
    if (name !== undefined && Number.isFinite(value)) out[ownString(name)] = value;
  }
  return out;
}

/** Rotation and flips of the owning shape, independent of its placement/extents. */
export function drawingShapeTransform(xfrm: XmlElement | undefined): TransformModel | undefined {
  if (xfrm === undefined) return undefined;
  const rotation = numberAttr(xfrm, "rot", "a:xfrm");
  const flipHorizontal = onOff(xfrm, "flipH");
  const flipVertical = onOff(xfrm, "flipV");
  if (rotation === undefined && flipHorizontal === undefined && flipVertical === undefined) return undefined;
  return {
    ...(rotation === undefined ? {} : { rotation: rotation / 60000 }),
    ...(flipHorizontal === undefined ? {} : { flipHorizontal }),
    ...(flipVertical === undefined ? {} : { flipVertical }),
  };
}

/** Resolve common preset geometries to paths in their authored local dimensions. */
export function geometryForPreset(
  preset: string | undefined,
  size: DrawingMlShapeSize | undefined,
  adjustments: Readonly<Record<string, number>> = {},
): ShapeGeometryModel | undefined {
  if (preset === undefined) return undefined;
  const viewBox = localViewBox(size);
  const path = pathForPreset(preset, viewBox, adjustments);
  const textRect = textRectForPreset(preset, viewBox, adjustments);
  return path === undefined ? undefined : { preset, path, viewBox, ...(textRect === undefined ? {} : { textRect }) };
}

/** A deliberate bounds-preserving rectangle for geometry that cannot be represented yet. */
function fallbackGeometry(size: DrawingMlShapeSize | undefined, fallback: GeometryFallback): ShapeGeometryModel {
  const viewBox = localViewBox(size);
  return { path: `M 0 0 H ${viewBox.width} V ${viewBox.height} H 0 Z`, viewBox, fallback };
}

/**
 * Geometry the fast path above cannot resolve: a visible bounds fallback that
 * also carries what the session needs to finish the job on the main thread —
 * the encoded `custGeom`, or (for a preset) just its authored adjustments,
 * since the catalog supplies the definition by name. The worker never holds
 * the preset catalog or the evaluator; if the session cannot resolve it, the
 * fallback marker is exactly what painted before.
 */
export function unresolvedGeometry(
  size: DrawingMlShapeSize | undefined,
  preset: string | undefined,
  adjustments: Readonly<Record<string, number>>,
  custGeom: XmlElement | undefined,
): ShapeGeometryModel {
  if (custGeom !== undefined) {
    const program = encodeGeometry(custGeom);
    return { ...fallbackGeometry(size, "custom"), ...(program === undefined ? {} : { deferred: { program } }) };
  }
  if (preset === undefined) return fallbackGeometry(size, "missing");
  return {
    ...fallbackGeometry(size, "unsupported"),
    preset,
    deferred: Object.keys(adjustments).length === 0 ? {} : { adjustments: { ...adjustments } },
  };
}
