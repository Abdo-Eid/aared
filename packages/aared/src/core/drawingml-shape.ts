import type { LocalViewBoxModel, ShapeGeometryModel } from "./model.js";
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

/** Resolve common preset geometries to paths in their authored local dimensions. */
export function geometryForPreset(
  preset: string | undefined,
  size: DrawingMlShapeSize | undefined,
  adjustments: Readonly<Record<string, number>> = {},
): ShapeGeometryModel | undefined {
  if (preset === undefined) return undefined;
  const viewBox = localViewBox(size);
  const path = pathForPreset(preset, viewBox, adjustments);
  return path === undefined ? undefined : { preset, path, viewBox };
}

/** A deliberate bounds-preserving rectangle for geometry that cannot be represented yet. */
export function fallbackGeometry(size: DrawingMlShapeSize | undefined, fallback: GeometryFallback): ShapeGeometryModel {
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
