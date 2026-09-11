/**
 * Shared chart-to-primitives module for PPTX, XLSX, and DOCX.
 *
 * Parses DrawingML chart parts (`c:chartSpace`) into typed
 * `ChartGraphicModel` primitives only — rect/line/path/text in local
 * unitless coordinates. Never emits raw SVG/HTML markup, never touches
 * the DOM, so it runs headless in the Worker and is reusable across
 * formats. PPTX wires `graphicFrame` charts here; XLSX/DOCX will reuse
 * the same entry points for their own drawing layers.
 */

import { attr, childElements, firstChild, textContent } from "./xml.js";
import type { XmlElement } from "./xml.js";
import { resolveDrawingColor } from "./theme.js";
import type { DrawingTheme } from "./theme.js";
import type { ChartGraphicModel, ChartGraphicPrimitive } from "./model.js";

const VIEW_WIDTH = 1000;
const VIEW_HEIGHT = 600;
const PLOT_LEFT = 90;
const PLOT_RIGHT = 30;
const PLOT_TOP = 30;
const PLOT_BOTTOM = 90;
const LEGEND_HEIGHT = 44;

const FALLBACK_PALETTE = ["#2E7D74", "#C9D6D3", "#156082", "#E8A33D", "#73A0B4", "#5B6470", "#888888"];

interface ChartSeries {
  name: string;
  fill?: string;
  categories: string[];
  values: number[];
}

function seriesName(ser: XmlElement): string {
  const tx = firstChild(ser, "tx");
  if (tx === undefined) return "";
  const findV = (element: XmlElement): string | undefined => {
    if (element.local === "v") {
      const text = textContent(element).trim();
      return text === "" ? undefined : text;
    }
    for (const child of element.children) {
      const found = findV(child);
      if (found !== undefined) return found;
    }
    return undefined;
  };
  // Series names surface as a descendant <c:v> inside strRef/strCache.
  return findV(tx) ?? "";
}

function pointsIn(container: XmlElement | undefined): Array<{ index: number; text: string }> {
  if (container === undefined) return [];
  const points: Array<{ index: number; text: string }> = [];
  const visit = (element: XmlElement): void => {
    if (element.local === "pt") {
      const rawIndex = attr(element, "idx");
      const index = rawIndex === undefined ? points.length : Number(rawIndex);
      const v = firstChild(element, "v");
      points.push({ index: Number.isFinite(index) ? index : points.length, text: v === undefined ? "" : textContent(v).trim() });
      return;
    }
    for (const child of element.children) visit(child);
  };
  visit(container);
  return [...points].sort((a, b) => a.index - b.index);
}

function categoriesOf(ser: XmlElement): string[] {
  const cat = firstChild(ser, "cat");
  if (cat === undefined) return [];
  return pointsIn(cat).map((point) => point.text);
}

function valuesOf(ser: XmlElement): number[] {
  const val = firstChild(ser, "val");
  if (val === undefined) return [];
  const out: number[] = [];
  for (const point of pointsIn(val)) {
    const numeric = Number(point.text);
    out.push(Number.isFinite(numeric) ? numeric : 0);
  }
  return out;
}

function seriesFill(ser: XmlElement, theme: DrawingTheme, fallback: string): string {
  const spPr = firstChild(ser, "spPr");
  const color = resolveDrawingColor(spPr === undefined ? undefined : firstChild(spPr, "solidFill"), theme);
  if (color !== undefined) return color;
  // Per-point fills (pie slices) are resolved by the pie path; series-level
  // callers use the palette fallback.
  return fallback;
}

function pointFills(ser: XmlElement, theme: DrawingTheme): Map<number, string> {
  const fills = new Map<number, string>();
  for (const dPt of childElements(ser, "dPt")) {
    const indexRaw = firstChild(dPt, "idx");
    const index = indexRaw === undefined ? undefined : Number(attr(indexRaw, "val") ?? NaN);
    const spPr = firstChild(dPt, "spPr");
    const color = resolveDrawingColor(spPr === undefined ? undefined : firstChild(spPr, "solidFill"), theme);
    if (index !== undefined && Number.isFinite(index) && color !== undefined) fills.set(index, color);
  }
  return fills;
}

function collectSeries(chartType: XmlElement, theme: DrawingTheme): ChartSeries[] {
  return childElements(chartType, "ser").map((ser, seriesIndex) => {
    const fallback = FALLBACK_PALETTE[seriesIndex % FALLBACK_PALETTE.length]!;
    return {
      name: seriesName(ser),
      fill: seriesFill(ser, theme, fallback),
      categories: categoriesOf(ser),
      values: valuesOf(ser),
    };
  });
}

function axisFontFamily(theme: DrawingTheme): string {
  return theme.fonts.minor.latin || theme.fonts.minor.complexScript || "Arial";
}

function escapeText(text: string): string {
  return text;
}

function barPrimitives(
  series: ChartSeries[],
  direction: "col" | "bar",
  theme: DrawingTheme,
  hasLegend: boolean,
): ChartGraphicPrimitive[] {
  const primitives: ChartGraphicPrimitive[] = [];
  const fontFamily = axisFontFamily(theme);
  const legendReserve = hasLegend ? LEGEND_HEIGHT : 0;
  const plotWidth = VIEW_WIDTH - PLOT_LEFT - PLOT_RIGHT;
  const plotHeight = VIEW_HEIGHT - PLOT_TOP - PLOT_BOTTOM - legendReserve;
  const categoryCount = Math.max(0, ...series.map((entry) => Math.max(entry.categories.length, entry.values.length)));
  if (categoryCount === 0 || series.length === 0) return primitives;
  const categories = series[0]!.categories.length === categoryCount
    ? series[0]!.categories
    : Array.from({ length: categoryCount }, (_, index) => series[0]?.categories[index] ?? `Cat ${index + 1}`);
  const allValues = series.flatMap((entry) => entry.values);
  const maxValue = Math.max(0, ...allValues);
  const minValue = Math.min(0, ...allValues);
  const span = maxValue - minValue === 0 ? 1 : maxValue - minValue;
  const yFor = (value: number): number => PLOT_TOP + plotHeight - ((value - minValue) / span) * plotHeight;
  const xFor = (value: number): number => PLOT_LEFT + ((value - minValue) / span) * plotWidth;
  const baselineY = yFor(0);
  const baselineX = xFor(0);

  // Axes and gridlines.
  const gridColor = "#E6E2DA";
  const axisColor = "#888888";
  if (direction === "col") {
    for (let tick = 0; tick <= 4; tick++) {
      const value = minValue + (span * tick) / 4;
      const y = yFor(value);
      primitives.push({ kind: "line", x1: PLOT_LEFT, y1: y, x2: PLOT_LEFT + plotWidth, y2: y, stroke: { color: tick === 0 ? axisColor : gridColor, width: 2 } });
      primitives.push({
        kind: "text", x: PLOT_LEFT - 10, y: y + 7, text: escapeText(String(Math.round(value * 100) / 100)),
        fontSize: 20, fontFamily, textAnchor: "end", fill: "#5B6470",
      });
    }
    primitives.push({ kind: "line", x1: PLOT_LEFT, y1: PLOT_TOP, x2: PLOT_LEFT, y2: PLOT_TOP + plotHeight, stroke: { color: axisColor, width: 2 } });
    primitives.push({ kind: "line", x1: PLOT_LEFT, y1: baselineY, x2: PLOT_LEFT + plotWidth, y2: baselineY, stroke: { color: axisColor, width: 2 } });
  } else {
    for (let tick = 0; tick <= 4; tick++) {
      const value = minValue + (span * tick) / 4;
      const x = xFor(value);
      primitives.push({ kind: "line", x1: x, y1: PLOT_TOP, x2: x, y2: PLOT_TOP + plotHeight, stroke: { color: tick === 0 ? axisColor : gridColor, width: 2 } });
      primitives.push({
        kind: "text", x, y: PLOT_TOP + plotHeight + 24, text: escapeText(String(Math.round(value * 100) / 100)),
        fontSize: 20, fontFamily, textAnchor: "middle", fill: "#5B6470",
      });
    }
    primitives.push({ kind: "line", x1: PLOT_LEFT, y1: PLOT_TOP + plotHeight, x2: PLOT_LEFT + plotWidth, y2: PLOT_TOP + plotHeight, stroke: { color: axisColor, width: 2 } });
    primitives.push({ kind: "line", x1: baselineX, y1: PLOT_TOP, x2: baselineX, y2: PLOT_TOP + plotHeight, stroke: { color: axisColor, width: 2 } });
  }

  if (direction === "col") {
    const groupWidth = plotWidth / categoryCount;
    const barWidth = (groupWidth * 0.7) / series.length;
    series.forEach((entry, seriesIndex) => {
      entry.values.forEach((value, categoryIndex) => {
        const groupX = PLOT_LEFT + groupWidth * categoryIndex;
        const x = groupX + (groupWidth - barWidth * series.length) / 2 + barWidth * seriesIndex;
        const y = value >= 0 ? yFor(value) : baselineY;
        const height = Math.abs(yFor(value) - baselineY);
        if (height <= 0) return;
        primitives.push({ kind: "rect", x, y, width: barWidth, height, fill: entry.fill ?? FALLBACK_PALETTE[seriesIndex % FALLBACK_PALETTE.length]! });
        primitives.push({
          kind: "text", x: x + barWidth / 2, y: (value >= 0 ? y - 6 : y + height + 20),
          text: escapeText(String(value)), fontSize: 18, fontFamily, textAnchor: "middle", fill: "#5B6470",
        });
      });
    });
    categories.forEach((category, index) => {
      primitives.push({
        kind: "text", x: PLOT_LEFT + groupWidth * index + groupWidth / 2, y: PLOT_TOP + plotHeight + 28,
        text: escapeText(category), fontSize: 20, fontFamily, textAnchor: "middle", fill: "#5B6470",
      });
    });
  } else {
    const groupHeight = plotHeight / categoryCount;
    const barHeight = (groupHeight * 0.7) / series.length;
    series.forEach((entry, seriesIndex) => {
      entry.values.forEach((value, categoryIndex) => {
        const groupY = PLOT_TOP + groupHeight * categoryIndex;
        const y = groupY + (groupHeight - barHeight * series.length) / 2 + barHeight * seriesIndex;
        const x = value >= 0 ? baselineX : xFor(value);
        const width = Math.abs(xFor(value) - baselineX);
        if (width <= 0) return;
        primitives.push({ kind: "rect", x, y, width, height: barHeight, fill: entry.fill ?? FALLBACK_PALETTE[seriesIndex % FALLBACK_PALETTE.length]! });
        primitives.push({
          kind: "text", x: (value >= 0 ? x + width + 6 : x - 6), y: y + barHeight / 2 + 6,
          text: escapeText(String(value)), fontSize: 18, fontFamily,
          textAnchor: value >= 0 ? "start" : "end", fill: "#5B6470",
        });
      });
    });
    categories.forEach((category, index) => {
      primitives.push({
        kind: "text", x: PLOT_LEFT - 10, y: PLOT_TOP + groupHeight * index + groupHeight / 2 + 6,
        text: escapeText(category), fontSize: 20, fontFamily, textAnchor: "end", fill: "#5B6470",
      });
    });
  }

  if (hasLegend) {
    const legendY = VIEW_HEIGHT - legendReserve + 20;
    let cursor = PLOT_LEFT;
    series.forEach((entry, index) => {
      const color = entry.fill ?? FALLBACK_PALETTE[index % FALLBACK_PALETTE.length]!;
      primitives.push({ kind: "rect", x: cursor, y: legendY - 14, width: 22, height: 14, fill: color });
      cursor += 28;
      const label = entry.name === "" ? `Series ${index + 1}` : entry.name;
      primitives.push({ kind: "text", x: cursor, y: legendY, text: escapeText(label), fontSize: 20, fontFamily, textAnchor: "start", fill: "#5B6470" });
      cursor += Math.max(60, label.length * 11 + 24);
    });
  }
  return primitives;
}

function linePrimitives(series: ChartSeries[], theme: DrawingTheme, hasLegend: boolean): ChartGraphicPrimitive[] {
  const primitives: ChartGraphicPrimitive[] = [];
  const fontFamily = axisFontFamily(theme);
  const legendReserve = hasLegend ? LEGEND_HEIGHT : 0;
  const plotWidth = VIEW_WIDTH - PLOT_LEFT - PLOT_RIGHT;
  const plotHeight = VIEW_HEIGHT - PLOT_TOP - PLOT_BOTTOM - legendReserve;
  const categoryCount = Math.max(0, ...series.map((entry) => Math.max(entry.categories.length, entry.values.length)));
  if (categoryCount === 0 || series.length === 0) return primitives;
  const categories = series[0]!.categories.length === categoryCount
    ? series[0]!.categories
    : Array.from({ length: categoryCount }, (_, index) => series[0]?.categories[index] ?? `Cat ${index + 1}`);
  const allValues = series.flatMap((entry) => entry.values);
  const maxValue = Math.max(0, ...allValues);
  const minValue = Math.min(0, ...allValues);
  const span = maxValue - minValue === 0 ? 1 : maxValue - minValue;
  const xFor = (index: number): number => categoryCount === 1 ? PLOT_LEFT + plotWidth / 2 : PLOT_LEFT + (plotWidth * index) / (categoryCount - 1);
  const yFor = (value: number): number => PLOT_TOP + plotHeight - ((value - minValue) / span) * plotHeight;

  for (let tick = 0; tick <= 4; tick++) {
    const value = minValue + (span * tick) / 4;
    const y = yFor(value);
    primitives.push({ kind: "line", x1: PLOT_LEFT, y1: y, x2: PLOT_LEFT + plotWidth, y2: y, stroke: { color: tick === 0 ? "#888888" : "#E6E2DA", width: 2 } });
    primitives.push({
      kind: "text", x: PLOT_LEFT - 10, y: y + 7, text: escapeText(String(Math.round(value * 100) / 100)),
      fontSize: 20, fontFamily, textAnchor: "end", fill: "#5B6470",
    });
  }
  primitives.push({ kind: "line", x1: PLOT_LEFT, y1: PLOT_TOP + plotHeight, x2: PLOT_LEFT + plotWidth, y2: PLOT_TOP + plotHeight, stroke: { color: "#888888", width: 2 } });

  series.forEach((entry, seriesIndex) => {
    const color = entry.fill ?? FALLBACK_PALETTE[seriesIndex % FALLBACK_PALETTE.length]!;
    const points = entry.values.map((value, index) => ({ x: xFor(index), y: yFor(value), value }));
    if (points.length === 0) return;
    const path = points.map((point, index) => `${index === 0 ? "M" : "L"} ${round(point.x)} ${round(point.y)}`).join(" ");
    primitives.push({ kind: "path", path, fill: "none", stroke: { color, width: 4 } });
    for (const point of points) {
      primitives.push({ kind: "rect", x: point.x - 5, y: point.y - 5, width: 10, height: 10, fill: color });
      primitives.push({
        kind: "text", x: point.x, y: point.y - 10, text: escapeText(String(point.value)),
        fontSize: 18, fontFamily, textAnchor: "middle", fill: "#5B6470",
      });
    }
  });

  categories.forEach((category, index) => {
    primitives.push({
      kind: "text", x: xFor(index), y: PLOT_TOP + plotHeight + 28,
      text: escapeText(category), fontSize: 20, fontFamily, textAnchor: "middle", fill: "#5B6470",
    });
  });

  if (hasLegend) {
    const legendY = VIEW_HEIGHT - legendReserve + 20;
    let cursor = PLOT_LEFT;
    series.forEach((entry, index) => {
      const color = entry.fill ?? FALLBACK_PALETTE[index % FALLBACK_PALETTE.length]!;
      primitives.push({ kind: "line", x1: cursor, y1: legendY - 5, x2: cursor + 26, y2: legendY - 5, stroke: { color, width: 4 } });
      cursor += 32;
      const label = entry.name === "" ? `Series ${index + 1}` : entry.name;
      primitives.push({ kind: "text", x: cursor, y: legendY, text: escapeText(label), fontSize: 20, fontFamily, textAnchor: "start", fill: "#5B6470" });
      cursor += Math.max(60, label.length * 11 + 24);
    });
  }
  return primitives;
}

function round(value: number): number {
  return Math.round(value * 100) / 100;
}

function piePrimitives(chartType: XmlElement, theme: DrawingTheme, hasLegend: boolean): ChartGraphicPrimitive[] {
  const primitives: ChartGraphicPrimitive[] = [];
  const fontFamily = axisFontFamily(theme);
  const serElements = childElements(chartType, "ser");
  const series = collectSeries(chartType, theme);
  if (series.length === 0 || serElements.length === 0) return primitives;
  const first = series[0]!;
  const firstSer = serElements[0]!;
  const total = first.values.reduce((sum, value) => sum + Math.max(value, 0), 0);
  if (total <= 0 || first.values.length === 0) return primitives;
  const fills = pointFills(firstSer, theme);
  const centerX = VIEW_WIDTH / 2;
  const centerY = (VIEW_HEIGHT - (hasLegend ? LEGEND_HEIGHT : 0)) / 2 + 10;
  const radius = Math.min(VIEW_WIDTH, VIEW_HEIGHT - (hasLegend ? LEGEND_HEIGHT : 0)) / 2 - 60;
  let angle = -90;
  first.values.forEach((value, index) => {
    const clamped = Math.max(value, 0);
    if (clamped <= 0) return;
    const sweep = (clamped / total) * 360;
    const start = (angle * Math.PI) / 180;
    const end = ((angle + sweep) * Math.PI) / 180;
    const x1 = centerX + radius * Math.cos(start);
    const y1 = centerY + radius * Math.sin(start);
    const x2 = centerX + radius * Math.cos(end);
    const y2 = centerY + radius * Math.sin(end);
    const largeArc = sweep > 180 ? 1 : 0;
    const fill = fills.get(index) ?? FALLBACK_PALETTE[index % FALLBACK_PALETTE.length]!;
    primitives.push({
      kind: "path",
      path: `M ${round(centerX)} ${round(centerY)} L ${round(x1)} ${round(y1)} A ${round(radius)} ${round(radius)} 0 ${largeArc} 1 ${round(x2)} ${round(y2)} Z`,
      fill,
      stroke: { color: "#FFFFFF", width: 2 },
    });
    const mid = ((angle + sweep / 2) * Math.PI) / 180;
    const labelX = centerX + radius * 0.65 * Math.cos(mid);
    const labelY = centerY + radius * 0.65 * Math.sin(mid);
    const category = first.categories[index] ?? "";
    const percent = `${round((clamped / total) * 1000) / 10}%`;
    primitives.push({
      kind: "text", x: round(labelX), y: round(labelY),
      text: escapeText(category === "" ? percent : `${category} ${percent}`),
      fontSize: 20, fontFamily, textAnchor: "middle", fill: "#FFFFFF",
    });
    angle += sweep;
  });

  if (hasLegend) {
    const legendY = VIEW_HEIGHT - LEGEND_HEIGHT + 22;
    let cursor = PLOT_LEFT;
    first.values.forEach((_, index) => {
      const fill = fills.get(index) ?? FALLBACK_PALETTE[index % FALLBACK_PALETTE.length]!;
      const label = first.categories[index] ?? `Slice ${index + 1}`;
      primitives.push({ kind: "rect", x: cursor, y: legendY - 14, width: 22, height: 14, fill });
      cursor += 28;
      primitives.push({ kind: "text", x: cursor, y: legendY, text: escapeText(label), fontSize: 20, fontFamily, textAnchor: "start", fill: "#5B6470" });
      cursor += Math.max(60, label.length * 11 + 24);
    });
  }
  return primitives;
}

/** Parse a DrawingML `c:chartSpace` element into typed SVG primitives. */
export function parseChartGraphic(chartSpace: XmlElement, theme: DrawingTheme): ChartGraphicModel | undefined {
  const chart = firstChild(chartSpace, "chart");
  const plotArea = chart === undefined ? undefined : firstChild(chart, "plotArea");
  if (plotArea === undefined) return undefined;
  const hasLegend = chart === undefined ? false : firstChild(chart, "legend") !== undefined;
  const barChart = firstChild(plotArea, "barChart");
  const lineChart = firstChild(plotArea, "lineChart");
  const pieChart = firstChild(plotArea, "pieChart");
  if (barChart !== undefined) {
    const direction = attr(firstChild(barChart, "barDir") ?? barChart, "val") === "bar" ? "bar" : "col";
    const series = collectSeries(barChart, theme);
    if (series.length === 0) return undefined;
    return { viewBox: { width: VIEW_WIDTH, height: VIEW_HEIGHT }, primitives: barPrimitives(series, direction, theme, hasLegend) };
  }
  if (lineChart !== undefined) {
    const series = collectSeries(lineChart, theme);
    if (series.length === 0) return undefined;
    return { viewBox: { width: VIEW_WIDTH, height: VIEW_HEIGHT }, primitives: linePrimitives(series, theme, hasLegend) };
  }
  if (pieChart !== undefined) {
    const primitives = piePrimitives(pieChart, theme, hasLegend);
    if (primitives.length === 0) return undefined;
    return { viewBox: { width: VIEW_WIDTH, height: VIEW_HEIGHT }, primitives };
  }
  return undefined;
}
