/**
 * The Worker half of the chart engine, shared by PPTX, XLSX and DOCX: a
 * DrawingML `c:chartSpace` part (or an Office 2016 `cx:chartSpace`) plus the
 * theme -> a resolved `ChartSpecModel` for the unit model. Reading only:
 * layout and painting happen on the page (`core/chart/draw.ts`, loaded when a
 * chart first shows), so the Worker carries the readers and nothing else of
 * the engine (Notes §71). Never emits markup and never touches the DOM.
 */

import { firstChild } from "./xml.js";
import type { XmlElement } from "./xml.js";
import type { OpcPackage } from "./opc.js";
import type { DrawingTheme } from "./theme.js";
import type { FillOwner } from "./fill.js";
import { AaredError } from "./errors.js";
import { ChartTooLarge, SUPPORTED_GROUP_TAGS, readChart } from "./chart/read.js";
import { isChartEx, readChartEx } from "./chart/read-chartex.js";
import type { ChartexData } from "./chart/read-chartex.js";
import type { ChartSpecModel } from "./chart/spec.js";

export { chartexFormulas } from "./chart/read-chartex.js";
export type { ChartexData, ChartexGrid } from "./chart/read-chartex.js";

/**
 * The chart part as the owner of its picture fills: a chart's `r:embed`
 * resolves against the chart part's own relationships, whichever format holds
 * the chart. Pass the result to `readChartModel`.
 */
export async function chartFillOwner(pkg: OpcPackage, chartPart: string, signal?: AbortSignal): Promise<FillOwner> {
  return {
    relationships: await pkg.relationships(chartPart, signal),
    mediaContentType: (target) => pkg.mediaContentType(target),
    inherited: false,
  };
}

/**
 * Read a chart part. `theme` must already carry the colour-map names
 * (`themeForColorMap`). `data` holds the cell ranges an XLSX chartex part's
 * formulas point at (the part carries no cache of its own there; see
 * `chartexFormulas`). A chart the engine cannot draw (a type it does not
 * draw, more points than it draws, or no cached values) comes back with the
 * reason instead, phrased for an `unsupported` report. `owner` (from
 * `chartFillOwner`) lets a picture fill name its media; without it a picture
 * fill reads as `unsupported`.
 */
export function readChartModel(
  chartSpace: XmlElement,
  theme: DrawingTheme,
  options: { data?: ChartexData; owner?: FillOwner } = {},
): { spec: ChartSpecModel } | { unsupportedReason: string } {
  const space = chartSpace.local === "chartSpace" ? chartSpace : firstChild(chartSpace, "chartSpace") ?? chartSpace;
  let spec: ChartSpecModel | string | undefined;
  try {
    spec = isChartEx(space) ? readChartEx(space, theme, options.data, options.owner) : readChart(space, theme, options.owner);
  } catch (error) {
    if (error instanceof ChartTooLarge) return { unsupportedReason: error.message };
    // A malformed number in a fill (the shared readers are strict, as for a shape) costs
    // this chart its drawing, marked, and not the slide, sheet or document around it.
    if (error instanceof AaredError && error.code === "invalid-document") {
      return { unsupportedReason: `this chart has a damaged fill and is not previewed (${error.message.replace(/^aared\/core: /, "").replace(/\. The file may be damaged.*$/, "")})` };
    }
    throw error;
  }
  if (typeof spec === "string") return { unsupportedReason: spec };
  // A chart with series but no cached points draws nothing: say so instead of an empty frame.
  if (spec !== undefined && spec.groups.some((g) => g.series.some((s) => s.values.some((v) => v !== null)))) return { spec };
  return { unsupportedReason: unsupportedReason(space) };
}

function unsupportedReason(chartSpace: XmlElement): string {
  const chart = chartSpace.local === "chart" ? chartSpace : firstChild(chartSpace, "chart");
  const plotArea = chart === undefined ? undefined : firstChild(chart, "plotArea");
  if (plotArea === undefined) return "this chart has no plot area to preview";
  const types = plotArea.children.map((child) => child.local).filter((local) => local.endsWith("Chart"));
  const missing = types.filter((type) => !SUPPORTED_GROUP_TAGS.includes(type));
  if (missing.length > 0) return `chart type ${missing.join(", ")} is not previewed`;
  if (types.length === 0) return isChartEx(chartSpace) ? "this chart has no cached values to preview" : "chart type unknown is not previewed";
  return "this chart has no cached values to preview";
}
