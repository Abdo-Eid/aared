/**
 * Paints the chart primitives `core/chart.ts` emits into an SVG. Renderer-side
 * and shared by PPTX and XLSX charts: parsing was already one module, and the
 * painter must be too, or the same primitive draws differently per format.
 */

import type { ChartGraphicPrimitive } from "./core/model.js";
import { cssFontFamily } from "./core/typography.js";

const SVG_NAMESPACE = "http://www.w3.org/2000/svg";

/**
 * Append one primitive to `svg`. Unset styles take the safe defaults: an
 * unfilled rect or path stays unfilled (SVG's own default fill is black), text
 * without a fill is black, and a line without a stroke follows `currentColor`.
 */
export function appendChartPrimitive(svg: SVGSVGElement, primitive: ChartGraphicPrimitive): void {
  const node = document.createElementNS(SVG_NAMESPACE, primitive.kind);
  const set = (name: string, value: string | number | undefined): void => {
    if (value !== undefined) node.setAttribute(name, String(value));
  };
  switch (primitive.kind) {
    case "rect":
      set("x", primitive.x);
      set("y", primitive.y);
      set("width", primitive.width);
      set("height", primitive.height);
      break;
    case "line":
      set("x1", primitive.x1);
      set("y1", primitive.y1);
      set("x2", primitive.x2);
      set("y2", primitive.y2);
      break;
    case "path":
      set("d", primitive.path);
      break;
    case "text":
      set("x", primitive.x);
      set("y", primitive.y);
      node.textContent = primitive.text;
      set("font-size", primitive.fontSize);
      // The same stack every run paints with: the family, a delivered
      // substitute's alias, then a generic. A bare family the reader lacks
      // (Aptos, the default theme's) fell to the SVG default serif.
      set("font-family", primitive.fontFamily === undefined ? undefined : cssFontFamily(primitive.fontFamily));
      set("font-weight", primitive.fontWeight);
      set("text-anchor", primitive.textAnchor);
      break;
  }
  if (primitive.kind === "text") set("fill", primitive.fill ?? "#000000");
  else if (primitive.kind !== "line") set("fill", primitive.fill ?? "none");
  if (primitive.stroke !== undefined) {
    set("stroke", primitive.stroke.color);
    set("stroke-width", primitive.stroke.width);
  } else if (primitive.kind === "line") {
    set("stroke", "currentColor");
  }
  set("opacity", primitive.opacity);
  svg.append(node);
}
