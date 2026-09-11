import type { LocalViewBoxModel, ShapeGeometryModel } from "./model.js";

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

function adjustment(values: Readonly<Record<string, number>>, name: string, defaultValue: number, maximum: number): number {
  return Math.min(Math.max(values[name] ?? defaultValue, 0), maximum);
}

function pathForPreset(preset: string, viewBox: LocalViewBoxModel, adjustments: Readonly<Record<string, number>>): string | undefined {
  const { width, height } = viewBox;
  const shortSide = Math.min(width, height);
  switch (preset) {
    case "rect":
      return `M 0 0 H ${width} V ${height} H 0 Z`;
    case "roundRect": {
      const radius = (shortSide * adjustment(adjustments, "adj", 16667, 50000)) / 100000;
      return `M ${radius} 0 H ${width - radius} A ${radius} ${radius} 0 0 1 ${width} ${radius} V ${height - radius} A ${radius} ${radius} 0 0 1 ${width - radius} ${height} H ${radius} A ${radius} ${radius} 0 0 1 0 ${height - radius} V ${radius} A ${radius} ${radius} 0 0 1 ${radius} 0 Z`;
    }
    case "ellipse":
      return `M ${width / 2} 0 A ${width / 2} ${height / 2} 0 1 1 ${width / 2 - 0.01} 0 Z`;
    case "line":
      return `M 0 0 L ${width} ${height}`;
    case "triangle": {
      const apex = (width * adjustment(adjustments, "adj", 50000, 100000)) / 100000;
      return `M ${apex} 0 L ${width} ${height} H 0 Z`;
    }
    case "rtTriangle":
      return `M 0 0 H ${width} V ${height} Z`;
    case "diamond":
      return `M ${width / 2} 0 L ${width} ${height / 2} L ${width / 2} ${height} L 0 ${height / 2} Z`;
    case "parallelogram": {
      const skew = (width * adjustment(adjustments, "adj", 25000, 50000)) / 100000;
      return `M ${skew} 0 H ${width} L ${width - skew} ${height} H 0 Z`;
    }
    case "hexagon": {
      const inset = (width * adjustment(adjustments, "adj", 25000, 50000)) / 100000;
      return `M ${inset} 0 H ${width - inset} L ${width} ${height / 2} L ${width - inset} ${height} H ${inset} L 0 ${height / 2} Z`;
    }
    case "octagon": {
      const inset = (shortSide * adjustment(adjustments, "adj", 29289, 50000)) / 100000;
      return `M ${inset} 0 H ${width - inset} L ${width} ${inset} V ${height - inset} L ${width - inset} ${height} H ${inset} L 0 ${height - inset} V ${inset} Z`;
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
