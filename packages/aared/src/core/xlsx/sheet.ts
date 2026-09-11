/**
 * Worksheet parser: SpreadsheetML sheet parts into resolved `SheetModel`s.
 *
 * View-only throughout: formulas are never evaluated (the file's cached
 * `<v>` is formatted and shown), conditional formatting is a one-time
 * snapshot computed from cached values, and charts go through a narrow
 * adapter that emits the same `ChartGraphicPrimitive` vocabulary the shared
 * chart contract uses — pending `core/chart.ts`, which this branch does not
 * yet contain. See the report note, not a duplicate engine.
 */

import { attr, childElements, firstChild, parseXml, textContent } from "../xml.js";
import type { XmlElement } from "../xml.js";
import type { OpcPackage } from "../opc.js";
import { throwIfAborted } from "../zip.js";
import { AaredError } from "../errors.js";
import { formatNumberValue } from "./numberfmt.js";
import type { XlsxStyles } from "./styles.js";
import type { XlsxWorkbook } from "./workbook.js";
import type {
  BoundsModel,
  CellModel,
  CellRange,
  ChartGraphicModel,
  ChartGraphicPrimitive,
  RowModel,
  SheetChartModel,
  SheetDrawingModel,
  SheetModel,
  SheetPictureModel,
  SheetRangeModel,
  SheetUnsupportedModel,
  XlsxCellType,
} from "../model.js";

/** EMU per inch; shared with DrawingML geometry. */
const EMUS_PER_INCH = 914400;
/** Excel column-width units (characters) to pixels at the default font. */
const PIXELS_PER_COLUMN_CHAR = 7;
const COLUMN_PADDING_PIXELS = 5;
const PIXELS_PER_INCH = 96;

function columnCharsToInches(width: number): number {
  return (width * PIXELS_PER_COLUMN_CHAR + COLUMN_PADDING_PIXELS) / PIXELS_PER_INCH;
}

function pointsToInches(points: number): number {
  return points / 72;
}

function numberAttribute(element: XmlElement, name: string, context: string): number | undefined {
  const value = attr(element, name);
  if (value === undefined) return undefined;
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) throw new Error(`aared/core: invalid ${name}="${value}" on ${context}`);
  return parsed;
}

function booleanAttribute(element: XmlElement, name: string): boolean {
  const value = attr(element, name);
  return value === "1" || value === "true";
}

/** Parse an A1 cell reference into zero-based coordinates. */
export function parseCellReference(reference: string): { row: number; col: number } {
  const match = /^([A-Za-z]{1,3})(\d{1,7})$/.exec(reference);
  if (match === null) throw new Error(`aared/core: invalid cell reference "${reference}"`);
  let col = 0;
  for (const char of match[1]!.toUpperCase()) col = col * 26 + (char.charCodeAt(0) - 64);
  const row = Number(match[2]!);
  if (row < 1 || row > 1048576 || col < 1 || col > 16384) {
    throw new Error(`aared/core: cell reference "${reference}" is outside the worksheet grid`);
  }
  return { row: row - 1, col: col - 1 };
}

/** Parse an A1 range (`A1:H1`) into an inclusive zero-based `CellRange`. */
export function parseCellRange(reference: string): CellRange {
  const parts = reference.split(":");
  const start = parseCellReference(parts[0]!);
  const end = parts.length === 1 ? start : parseCellReference(parts[1]!);
  return {
    start: { row: Math.min(start.row, end.row), col: Math.min(start.col, end.col) },
    end: { row: Math.max(start.row, end.row), col: Math.max(start.col, end.col) },
  };
}

/** Parse a `dimension` ref (`A1:H34`); single cells yield a 1x1 extent. */
function parseDimension(reference: string | undefined): { cols: number; rows: number } | undefined {
  if (reference === undefined) return undefined;
  try {
    const range = parseCellRange(reference);
    return { cols: range.end.col + 1, rows: range.end.row + 1 };
  } catch {
    return undefined;
  }
}

function inlineStringText(cell: XmlElement): string {
  const inline = firstChild(cell, "is");
  if (inline === undefined) return "";
  let out = "";
  for (const child of inline.children) {
    if (child.local === "t") out += textContent(child);
    else if (child.local === "r") {
      const run = firstChild(child, "t");
      if (run !== undefined) out += textContent(run);
    }
  }
  return out;
}

interface CachedValue {
  value: string | number | boolean | null;
  formatted: string;
  type: XlsxCellType;
}

function cachedValue(
  cell: XmlElement,
  sharedStrings: string[],
  style: { numberFormat?: string },
  date1904: boolean,
): CachedValue {
  const kind = attr(cell, "t") ?? "n";
  const raw = firstChild(cell, "v") === undefined ? undefined : textContent(firstChild(cell, "v")!);
  switch (kind) {
    case "s": {
      const index = raw === undefined ? Number.NaN : Number(raw);
      const text = Number.isInteger(index) && index >= 0 && index < sharedStrings.length ? sharedStrings[index]! : "";
      return { value: text, formatted: text, type: "shared" };
    }
    case "inlineStr": {
      const text = inlineStringText(cell);
      return { value: text, formatted: text, type: "inline" };
    }
    case "b": {
      const flag = raw === "1" || raw === "true";
      return { value: flag, formatted: flag ? "TRUE" : "FALSE", type: "boolean" };
    }
    case "e": {
      const text = raw ?? "";
      return { value: text, formatted: text, type: "error" };
    }
    case "str": {
      const text = raw ?? "";
      return { value: text, formatted: text, type: "formula-string" };
    }
    default: {
      if (raw === undefined || raw === "") return { value: null, formatted: "", type: "empty" };
      const numeric = Number(raw);
      if (!Number.isFinite(numeric)) return { value: raw, formatted: raw, type: "error" };
      const code = style.numberFormat ?? "General";
      return {
        value: numeric,
        formatted: formatNumberValue(numeric, code, date1904),
        type: "number",
      };
    }
  }
}

interface HyperlinkTarget {
  target: string;
  display?: string;
}

function hyperlinkTargets(sheet: XmlElement, relationships: { id: string; target: string }[]): Map<string, HyperlinkTarget> {
  const out = new Map<string, HyperlinkTarget>();
  const links = firstChild(sheet, "hyperlinks");
  if (links === undefined) return out;
  for (const link of childElements(links, "hyperlink")) {
    const ref = attr(link, "ref");
    if (ref === undefined) continue;
    const id = attr(link, "r:id");
    const display = attr(link, "display");
    const location = attr(link, "location");
    const target = id === undefined
      ? location === undefined ? undefined : `#${location}`
      : relationships.find((relationship) => relationship.id === id)?.target;
    if (target === undefined) continue;
    try {
      const range = parseCellRange(ref);
      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.col; col <= range.end.col; col++) {
          out.set(`${row}:${col}`, { target, ...(display === undefined ? {} : { display }) });
        }
      }
    } catch {
      continue;
    }
  }
  return out;
}

interface ConditionalRule {
  priority: number;
  dxfId?: number;
  operator?: string;
  formulas: string[];
  colorScale?: { values: Array<{ type: string; value: number }>; colors: string[] };
  dataBar?: { min: { type: string; value: number }; max: { type: string; value: number }; color: string };
}

function rgbColor(element: XmlElement | undefined): string | undefined {
  const rgb = element === undefined ? undefined : attr(element, "rgb");
  if (rgb === undefined || rgb.length !== 8) return undefined;
  return `#${rgb.slice(2).toUpperCase()}`;
}

function conditionalRules(sheet: XmlElement): Array<{ ranges: CellRange[]; rule: ConditionalRule }> {
  const out: Array<{ ranges: CellRange[]; rule: ConditionalRule }> = [];
  for (const formatting of childElements(sheet, "conditionalFormatting")) {
    const refs = attr(formatting, "sqref")?.split(/\s+/).filter((part) => part !== "") ?? [];
    const ranges: CellRange[] = [];
    for (const ref of refs) {
      try {
        ranges.push(parseCellRange(ref));
      } catch {
        continue;
      }
    }
    if (ranges.length === 0) continue;
    for (const element of childElements(formatting, "cfRule")) {
      const priority = numberAttribute(element, "priority", "c:cfRule") ?? 0;
      const dxfId = numberAttribute(element, "dxfId", "c:cfRule");
      const operator = attr(element, "operator");
      const rule: ConditionalRule = {
        priority,
        ...(dxfId === undefined ? {} : { dxfId }),
        ...(operator === undefined ? {} : { operator }),
        formulas: childElements(element, "formula").map((formula) => textContent(formula)),
      };
      const scale = firstChild(element, "colorScale");
      if (scale !== undefined) {
        const values = childElements(scale, "cfvo").map((entry) => ({
          type: attr(entry, "type") ?? "",
          value: Number(attr(entry, "val") ?? Number.NaN),
        }));
        const colors = childElements(scale, "color").map((entry) => rgbColor(entry)).filter((color): color is string => color !== undefined);
        rule.colorScale = { values, colors };
      }
      const bar = firstChild(element, "dataBar");
      if (bar !== undefined) {
        const stops = childElements(bar, "cfvo").map((entry) => ({
          type: attr(entry, "type") ?? "",
          value: Number(attr(entry, "val") ?? Number.NaN),
        }));
        const color = rgbColor(firstChild(bar, "color"));
        if (stops.length >= 2 && color !== undefined) {
          rule.dataBar = { min: stops[0]!, max: stops[1]!, color };
        }
      }
      out.push({ ranges, rule });
    }
  }
  // Excel evaluates lower priority numbers first; the first matching rule wins.
  out.sort((a, b) => a.rule.priority - b.rule.priority);
  return out;
}

function unquote(formula: string): string {
  return formula.length >= 2 && formula.startsWith('"') && formula.endsWith('"') ? formula.slice(1, -1) : formula;
}

function matchesCellIs(cell: CellModel, operator: string | undefined, formulas: string[]): boolean {
  if (operator === undefined) return false;
  const value = cell.value;
  const first = formulas[0] === undefined ? "" : unquote(formulas[0]);
  const second = formulas[1] === undefined ? "" : unquote(formulas[1]);
  const numeric = typeof value === "number" ? value : Number.NaN;
  const text = cell.formatted;
  switch (operator) {
    case "equal": return typeof value === "number" && first !== "" && !Number.isNaN(Number(first)) ? value === Number(first) : text === first;
    case "notEqual": return typeof value === "number" && first !== "" && !Number.isNaN(Number(first)) ? value !== Number(first) : text !== first;
    case "greaterThan": return numeric > Number(first);
    case "greaterThanOrEqual": return numeric >= Number(first);
    case "lessThan": return numeric < Number(first);
    case "lessThanOrEqual": return numeric <= Number(first);
    case "between": return numeric >= Number(first) && numeric <= Number(second);
    case "notBetween": return numeric < Number(first) || numeric > Number(second);
    case "containsText": return first !== "" && text.includes(first);
    case "notContainsText": return first !== "" && !text.includes(first);
    case "beginsWith": return first !== "" && text.startsWith(first);
    case "endsWith": return first !== "" && text.endsWith(first);
    default: return false;
  }
}

function lerpColor(low: string, high: string, ratio: number): string {
  const clamped = Math.min(Math.max(ratio, 0), 1);
  const channels = (hex: string): [number, number, number] => [
    Number.parseInt(hex.slice(1, 3), 16),
    Number.parseInt(hex.slice(3, 5), 16),
    Number.parseInt(hex.slice(5, 7), 16),
  ];
  const [r1, g1, b1] = channels(low);
  const [r2, g2, b2] = channels(high);
  const mix = (a: number, b: number): string => Math.round(a + (b - a) * clamped).toString(16).padStart(2, "0").toUpperCase();
  return `#${mix(r1, r2)}${mix(g1, g2)}${mix(b1, b2)}`;
}

function scaleBound(bound: { type: string; value: number }, values: number[], edge: "min" | "max"): number | undefined {
  if (bound.type === "num" && Number.isFinite(bound.value)) return bound.value;
  if (bound.type === "min" || bound.type === "max") {
    if (values.length === 0) return undefined;
    return edge === "min" ? Math.min(...values) : Math.max(...values);
  }
  return undefined;
}

/** Apply cached conditional-formatting results onto parsed cells (view-only snapshot). */
function applyConditionalFormatting(
  cells: Map<string, CellModel>,
  formats: Array<{ ranges: CellRange[]; rule: ConditionalRule }>,
  styles: XlsxStyles,
  unsupported: SheetUnsupportedModel[],
): void {
  let reportedUnsupported = false;
  const reportOnce = (detail: string): void => {
    if (reportedUnsupported) return;
    reportedUnsupported = true;
    unsupported.push({ kind: "conditional-format", detail });
  };
  for (const { ranges, rule } of formats) {
    if (rule.colorScale !== undefined) {
      const { values, colors } = rule.colorScale;
      if (values.length < 2 || colors.length < 2) {
        reportOnce("a color-scale rule needs numeric bounds; it is shown without its gradient");
        continue;
      }
      const numericValues: number[] = [];
      for (const range of ranges) {
        for (let row = range.start.row; row <= range.end.row; row++) {
          for (let col = range.start.col; col <= range.end.col; col++) {
            const cell = cells.get(`${row}:${col}`);
            if (cell !== undefined && typeof cell.value === "number") numericValues.push(cell.value);
          }
        }
      }
      const low = scaleBound(values[0]!, numericValues, "min");
      const high = scaleBound(values[values.length - 1]!, numericValues, "max");
      if (low === undefined || high === undefined) {
        reportOnce("a color-scale rule uses percentile bounds; it is shown without its gradient");
        continue;
      }
      const span = high - low === 0 ? 1 : high - low;
      const from = colors[0]!;
      const to = colors[colors.length - 1]!;
      for (const range of ranges) {
        for (let row = range.start.row; row <= range.end.row; row++) {
          for (let col = range.start.col; col <= range.end.col; col++) {
            const cell = cells.get(`${row}:${col}`);
            if (cell === undefined || typeof cell.value !== "number" || cell.conditional?.fill !== undefined) continue;
            cell.conditional = { ...(cell.conditional ?? {}), fill: lerpColor(from, to, (cell.value - low) / span) };
          }
        }
      }
      continue;
    }
    if (rule.dataBar !== undefined) {
      const numericValues: number[] = [];
      for (const range of ranges) {
        for (let row = range.start.row; row <= range.end.row; row++) {
          for (let col = range.start.col; col <= range.end.col; col++) {
            const cell = cells.get(`${row}:${col}`);
            if (cell !== undefined && typeof cell.value === "number") numericValues.push(cell.value);
          }
        }
      }
      const low = scaleBound(rule.dataBar.min, numericValues, "min");
      const high = scaleBound(rule.dataBar.max, numericValues, "max");
      if (low === undefined || high === undefined) {
        reportOnce("a data-bar rule uses unsupported bounds; its bars are not drawn");
        continue;
      }
      const span = high - low === 0 ? 1 : high - low;
      for (const range of ranges) {
        for (let row = range.start.row; row <= range.end.row; row++) {
          for (let col = range.start.col; col <= range.end.col; col++) {
            const cell = cells.get(`${row}:${col}`);
            if (cell === undefined || typeof cell.value !== "number" || cell.conditional?.dataBar !== undefined) continue;
            cell.conditional = {
              ...(cell.conditional ?? {}),
              dataBar: { ratio: Math.min(Math.max((cell.value - low) / span, 0), 1), color: rule.dataBar.color },
            };
          }
        }
      }
      continue;
    }
    if (rule.dxfId === undefined) {
      reportOnce("a conditional rule without a differential style is shown without highlighting");
      continue;
    }
    const differential = styles.differential(rule.dxfId);
    if (differential.fill === undefined && differential.font === undefined) continue;
    for (const range of ranges) {
      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.col; col <= range.end.col; col++) {
          const cell = cells.get(`${row}:${col}`);
          if (cell === undefined || cell.conditional?.fill !== undefined) continue;
          if (!matchesCellIs(cell, rule.operator, rule.formulas)) continue;
          cell.conditional = {
            ...(cell.conditional ?? {}),
            ...(differential.fill === undefined ? {} : { fill: differential.fill }),
            ...(differential.font?.color === undefined ? {} : { color: differential.font.color }),
            ...(differential.font?.bold === undefined ? {} : { bold: differential.font.bold }),
          };
        }
      }
    }
  }
}

interface DrawingAnchor {
  name: string;
  from: { col: number; row: number; colOff: number; rowOff: number };
  to?: { col: number; row: number; colOff: number; rowOff: number };
  ext?: { cx: number; cy: number };
}

function anchorFrom(element: XmlElement, context: string): DrawingAnchor["from"] {
  const column = firstChild(element, "col");
  const row = firstChild(element, "row");
  const colOff = firstChild(element, "colOff");
  const rowOff = firstChild(element, "rowOff");
  const colIndex = column === undefined ? 0 : Number(textContent(column));
  const rowIndex = row === undefined ? 0 : Number(textContent(row));
  if (!Number.isInteger(colIndex) || !Number.isInteger(rowIndex) || colIndex < 0 || rowIndex < 0) {
    throw new Error(`aared/core: invalid drawing anchor in ${context}`);
  }
  return {
    col: colIndex,
    row: rowIndex,
    colOff: colOff === undefined ? 0 : Number(textContent(colOff)) || 0,
    rowOff: rowOff === undefined ? 0 : Number(textContent(rowOff)) || 0,
  };
}

const CHART_PALETTE = ["#2E7D74", "#B24830", "#E8A33D", "#156082", "#73A0B4", "#8E6FB8"];

function seriesColor(series: XmlElement, fallback: string): string {
  const properties = firstChild(series, "spPr");
  const fill = properties === undefined ? undefined : firstChild(properties, "solidFill");
  const srgb = fill === undefined ? undefined : firstChild(fill, "srgbClr");
  const value = srgb === undefined ? undefined : attr(srgb, "val");
  return value !== undefined && /^[0-9a-fA-F]{6}$/.test(value) ? `#${value.toUpperCase()}` : fallback;
}

function cachePoints(container: XmlElement | undefined, kind: "str" | "num"): string[] {
  if (container === undefined) return [];
  const reference = firstChild(container, kind === "str" ? "strRef" : "numRef");
  const literal = firstChild(container, kind === "str" ? "strLit" : "numLit");
  const cache = reference === undefined ? undefined : firstChild(reference, kind === "str" ? "strCache" : "numCache");
  const source = cache ?? literal;
  if (source === undefined) return [];
  return childElements(source, "pt").map((point) => textContent(firstChild(point, "v") ?? point));
}

function seriesTitle(series: XmlElement): string {
  const tx = firstChild(series, "tx");
  if (tx === undefined) return "";
  const direct = firstChild(tx, "v");
  if (direct !== undefined) return textContent(direct);
  const reference = firstChild(tx, "strRef");
  const cache = reference === undefined ? undefined : firstChild(reference, "strCache");
  const point = cache === undefined ? undefined : firstChild(cache, "pt");
  if (point !== undefined) return textContent(firstChild(point, "v") ?? point);
  const rich = firstChild(tx, "rich");
  if (rich !== undefined) return textContent(rich);
  return "";
}

/**
 * Narrow spreadsheet chart adapter. Covers bar/column, line, and pie charts
 * from their cached values — the common preview cases — and reports any
 * other chart type through `unsupportedReason` instead of guessing. When the
 * shared `core/chart.ts` lands, this adapter should delegate to it; the
 * primitive vocabulary is already identical.
 */
export function adaptChartGraphic(chart: XmlElement): { graphic?: ChartGraphicModel; unsupportedReason?: string } {
  const space = chart.local === "chartSpace" ? chart : firstChild(chart, "chartSpace");
  const root = space ?? chart;
  const chartElement = firstChild(root, "chart") ?? (root.local === "chart" ? root : undefined);
  if (chartElement === undefined) return { unsupportedReason: "this chart has no plot area to preview" };
  const titleElement = firstChild(firstChild(chartElement, "title") ?? chartElement, "tx");
  const title = titleElement === undefined ? "" : textContent(titleElement);
  const plotArea = firstChild(chartElement, "plotArea");
  if (plotArea === undefined) return { unsupportedReason: "this chart has no plot area to preview" };
  const barChart = firstChild(plotArea, "barChart");
  const lineChart = firstChild(plotArea, "lineChart");
  const pieChart = firstChild(plotArea, "pieChart") ?? firstChild(plotArea, "doughnutChart");
  if (barChart !== undefined) return adaptCartesian(barChart, title, attr(barChart, "barDir") === "bar" ? "bar" : "column");
  if (lineChart !== undefined) return adaptCartesian(lineChart, title, "line");
  if (pieChart !== undefined) return adaptPie(pieChart, title);
  const present = plotArea.children.map((child) => child.local).join(", ");
  return { unsupportedReason: `chart type ${present === "" ? "unknown" : present} is not previewed` };
}

interface ChartSeries {
  name: string;
  categories: string[];
  values: number[];
  color: string;
}

function chartSeries(plot: XmlElement): ChartSeries[] {
  return childElements(plot, "ser").map((series, index) => {
    const values = cachePoints(firstChild(series, "val"), "num").map(Number).filter((value) => Number.isFinite(value));
    let categories = cachePoints(firstChild(series, "cat"), "str");
    if (categories.length === 0) {
      categories = cachePoints(firstChild(series, "cat"), "num");
    }
    return {
      name: seriesTitle(series),
      categories,
      values,
      color: seriesColor(series, CHART_PALETTE[index % CHART_PALETTE.length]!),
    };
  });
}

function chartText(
  primitives: ChartGraphicPrimitive[],
  x: number,
  y: number,
  text: string,
  options?: { fontSize?: number; anchor?: "start" | "middle" | "end"; bold?: boolean },
): void {
  if (text === "") return;
  primitives.push({
    kind: "text",
    x: Math.round(x * 10) / 10,
    y: Math.round(y * 10) / 10,
    text,
    fontSize: options?.fontSize ?? 10,
    textAnchor: options?.anchor ?? "start",
    ...(options?.bold === true ? { fontWeight: "bold" as const } : {}),
  });
}

function adaptCartesian(plot: XmlElement, title: string, mode: "column" | "bar" | "line"): { graphic: ChartGraphicModel } {
  const viewBox = { width: 400, height: 300 };
  const margin = { left: 44, right: 10, top: title === "" ? 10 : 28, bottom: 30 };
  const plotWidth = viewBox.width - margin.left - margin.right;
  const plotHeight = viewBox.height - margin.top - margin.bottom;
  const series = chartSeries(plot).filter((entry) => entry.values.length > 0);
  const categories = series[0]?.categories ?? [];
  const count = Math.max(...series.map((entry) => entry.values.length), 0);
  const flat = series.flatMap((entry) => entry.values);
  const max = Math.max(0, ...flat);
  const min = Math.min(0, ...flat);
  const span = max - min === 0 ? 1 : max - min;
  const primitives: ChartGraphicPrimitive[] = [];
  if (title !== "") chartText(primitives, viewBox.width / 2, 16, title, { fontSize: 12, anchor: "middle", bold: true });
  const baseline = margin.top + ((max - 0) / span) * plotHeight;
  const valueToY = (value: number): number => margin.top + ((max - value) / span) * plotHeight;
  // Value axis with three ticks; category slots share the plot width evenly.
  for (const tick of [max, (max + min) / 2, min]) {
    const y = valueToY(tick);
    primitives.push({ kind: "line", x1: margin.left, y1: y, x2: viewBox.width - margin.right, y2: y, stroke: { color: "#878787", width: 0.75 } });
    chartText(primitives, margin.left - 4, y + 3, String(Math.round(tick * 10) / 10), { fontSize: 9, anchor: "end" });
  }
  const slot = count === 0 ? plotWidth : plotWidth / count;
  series.forEach((entry, seriesIndex) => {
    const group = mode === "column" || mode === "bar" ? slot / Math.max(series.length, 1) : 0;
    entry.values.forEach((value, index) => {
      if (mode === "line") return;
      if (mode === "column") {
        const x = margin.left + index * slot + seriesIndex * group + group * 0.15;
        const width = group * 0.7;
        const y = value >= 0 ? valueToY(value) : baseline;
        const height = Math.abs(valueToY(value) - baseline);
        primitives.push({ kind: "rect", x: Math.round(x * 10) / 10, y: Math.round(y * 10) / 10, width: Math.round(width * 10) / 10, height: Math.round(height * 10) / 10, fill: entry.color });
      } else {
        const bandHeight = plotHeight / Math.max(count, 1);
        const y = margin.top + index * bandHeight + seriesIndex * (bandHeight / Math.max(series.length, 1)) + bandHeight / Math.max(series.length, 1) * 0.15;
        const height = (bandHeight / Math.max(series.length, 1)) * 0.7;
        const zeroX = margin.left + ((0 - min) / span) * plotWidth;
        const valueX = margin.left + ((value - min) / span) * plotWidth;
        primitives.push({
          kind: "rect",
          x: Math.round(Math.min(zeroX, valueX) * 10) / 10,
          y: Math.round(y * 10) / 10,
          width: Math.round(Math.abs(valueX - zeroX) * 10) / 10,
          height: Math.round(height * 10) / 10,
          fill: entry.color,
        });
      }
    });
  });
  if (mode === "line") {
    for (const entry of series) {
      const points = entry.values.map((value, index) => ({
        x: margin.left + (index + 0.5) * slot,
        y: valueToY(value),
      }));
      const path = points.map((point, index) => `${index === 0 ? "M" : "L"} ${Math.round(point.x * 10) / 10} ${Math.round(point.y * 10) / 10}`).join(" ");
      if (path !== "") primitives.push({ kind: "path", path, stroke: { color: entry.color, width: 2 } });
      for (const point of points) {
        primitives.push({ kind: "rect", x: Math.round((point.x - 2) * 10) / 10, y: Math.round((point.y - 2) * 10) / 10, width: 4, height: 4, fill: entry.color });
      }
    }
  }
  for (let index = 0; index < count; index++) {
    const label = categories[index] ?? `Item ${index + 1}`;
    if (mode === "bar") {
      const bandHeight = plotHeight / Math.max(count, 1);
      chartText(primitives, margin.left - 4, margin.top + index * bandHeight + bandHeight / 2 + 3, label, { fontSize: 9, anchor: "end" });
    } else {
      chartText(primitives, margin.left + index * slot + slot / 2, viewBox.height - 12, label, { fontSize: 9, anchor: "middle" });
    }
  }
  if (series.length > 1) {
    series.forEach((entry, seriesIndex) => {
      const x = margin.left + seriesIndex * 110;
      primitives.push({ kind: "rect", x, y: viewBox.height - 10, width: 10, height: 10, fill: entry.color });
      chartText(primitives, x + 13, viewBox.height - 2, entry.name === "" ? `Series ${seriesIndex + 1}` : entry.name, { fontSize: 9 });
    });
  }
  return { graphic: { viewBox, primitives } };
}

function adaptPie(plot: XmlElement, title: string): { graphic: ChartGraphicModel } {
  const viewBox = { width: 400, height: 300 };
  const primitives: ChartGraphicPrimitive[] = [];
  if (title !== "") chartText(primitives, viewBox.width / 2, 16, title, { fontSize: 12, anchor: "middle", bold: true });
  const series = chartSeries(plot).find((entry) => entry.values.length > 0);
  const values = series?.values ?? [];
  const labels = series?.categories ?? [];
  const total = values.reduce((sum, value) => sum + Math.max(value, 0), 0);
  const centerX = viewBox.width / 2;
  const centerY = (viewBox.height + (title === "" ? 0 : 20)) / 2;
  const radius = 100;
  if (total <= 0) {
    chartText(primitives, centerX, centerY, "No data", { anchor: "middle" });
    return { graphic: { viewBox, primitives } };
  }
  let angle = -Math.PI / 2;
  values.forEach((value, index) => {
    const share = Math.max(value, 0) / total;
    if (share <= 0) return;
    const start = angle;
    angle += share * Math.PI * 2;
    const largeArc = share > 0.5 ? 1 : 0;
    const x1 = centerX + radius * Math.cos(start);
    const y1 = centerY + radius * Math.sin(start);
    const x2 = centerX + radius * Math.cos(angle);
    const y2 = centerY + radius * Math.sin(angle);
    const color = CHART_PALETTE[index % CHART_PALETTE.length]!;
    primitives.push({
      kind: "path",
      path: `M ${round(centerX)} ${round(centerY)} L ${round(x1)} ${round(y1)} A ${radius} ${radius} 0 ${largeArc} 1 ${round(x2)} ${round(y2)} Z`,
      fill: color,
      stroke: { color: "#FFFFFF", width: 1 },
    });
    const mid = (start + angle) / 2;
    const label = labels[index] ?? "";
    const percent = `${Math.round(share * 1000) / 10}%`;
    chartText(primitives, centerX + (radius + 14) * Math.cos(mid), centerY + (radius + 14) * Math.sin(mid), label === "" ? percent : `${label} ${percent}`, { fontSize: 9, anchor: "middle" });
  });
  return { graphic: { viewBox, primitives } };
}

function round(value: number): number {
  return Math.round(value * 10) / 10;
}

async function parseDrawings(
  pkg: OpcPackage,
  sheetPart: string,
  columnEmu: number[],
  rowEmu: number[],
  signal?: AbortSignal,
): Promise<{ drawings: SheetDrawingModel[]; unsupported: SheetUnsupportedModel[] }> {
  throwIfAborted(signal);
  const drawings: SheetDrawingModel[] = [];
  const unsupported: SheetUnsupportedModel[] = [];
  const relationships = await pkg.relationships(sheetPart, signal);
  throwIfAborted(signal);
  const drawingRel = relationships.find((relationship) => relationship.type.endsWith("/drawing"));
  if (drawingRel === undefined) return { drawings, unsupported };
  if (!pkg.hasPart(drawingRel.target)) return { drawings, unsupported };
  const drawing = parseXml(await pkg.readText(drawingRel.target, signal));
  throwIfAborted(signal);
  const drawingRels = await pkg.relationships(drawingRel.target, signal);
  throwIfAborted(signal);
  const position = (index: number, offsets: number[]): number => {
    let total = 0;
    for (let i = 0; i < index && i < offsets.length; i++) total += offsets[i]!;
    return total;
  };
  for (const anchor of drawing.children) {
    throwIfAborted(signal);
    if (anchor.local !== "twoCellAnchor" && anchor.local !== "oneCellAnchor" && anchor.local !== "absoluteAnchor") continue;
    const fromElement = firstChild(anchor, "from");
    const toElement = firstChild(anchor, "to");
    const extElement = firstChild(anchor, "ext");
    const posElement = firstChild(anchor, "pos");
    let bounds: BoundsModel | undefined;
    let range: CellRange = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };
    try {
      if (fromElement !== undefined && toElement !== undefined) {
        const from = anchorFrom(fromElement, "xdr:from");
        const to = anchorFrom(toElement, "xdr:to");
        const x = (position(from.col, columnEmu) + from.colOff) / EMUS_PER_INCH;
        const y = (position(from.row, rowEmu) + from.rowOff) / EMUS_PER_INCH;
        const x2 = (position(to.col, columnEmu) + to.colOff) / EMUS_PER_INCH;
        const y2 = (position(to.row, rowEmu) + to.rowOff) / EMUS_PER_INCH;
        bounds = { x, y, width: Math.max(x2 - x, 0), height: Math.max(y2 - y, 0) };
        range = {
          start: { row: Math.min(from.row, to.row), col: Math.min(from.col, to.col) },
          end: { row: Math.max(from.row, to.row), col: Math.max(from.col, to.col) },
        };
      } else if (fromElement !== undefined && extElement !== undefined) {
        const from = anchorFrom(fromElement, "xdr:from");
        const cx = numberAttribute(extElement, "cx", "a:ext") ?? 0;
        const cy = numberAttribute(extElement, "cy", "a:ext") ?? 0;
        bounds = {
          x: (position(from.col, columnEmu) + from.colOff) / EMUS_PER_INCH,
          y: (position(from.row, rowEmu) + from.rowOff) / EMUS_PER_INCH,
          width: cx / EMUS_PER_INCH,
          height: cy / EMUS_PER_INCH,
        };
        range = { start: { row: from.row, col: from.col }, end: { row: from.row, col: from.col } };
      } else if (posElement !== undefined && extElement !== undefined) {
        const x = (numberAttribute(posElement, "x", "xdr:pos") ?? 0) / EMUS_PER_INCH;
        const y = (numberAttribute(posElement, "y", "xdr:pos") ?? 0) / EMUS_PER_INCH;
        bounds = {
          x,
          y,
          width: (numberAttribute(extElement, "cx", "a:ext") ?? 0) / EMUS_PER_INCH,
          height: (numberAttribute(extElement, "cy", "a:ext") ?? 0) / EMUS_PER_INCH,
        };
      }
    } catch {
      continue;
    }
    if (bounds === undefined) continue;
    const picture = firstChild(anchor, "pic");
    const frame = firstChild(anchor, "graphicFrame");
    if (picture !== undefined) {
      const blipFill = firstChild(picture, "blipFill");
      const blip = blipFill === undefined ? undefined : firstChild(blipFill, "blip");
      const embed = blip === undefined ? undefined : attr(blip, "r:embed");
      const target = embed === undefined ? undefined : drawingRels.find((relationship) => relationship.id === embed)?.target;
      const nameElement = firstChild(firstChild(picture, "nvPicPr") ?? picture, "cNvPr");
      const name = nameElement === undefined ? "Picture" : (attr(nameElement, "name") ?? "Picture");
      if (target === undefined || !pkg.hasPart(target)) {
        unsupported.push({ kind: "drawing", detail: `an image ("${name}") points at a missing part; its frame is kept`, range });
        continue;
      }
      const pictureModel: SheetPictureModel = {
        kind: "picture",
        name,
        mediaId: target,
        contentType: pkg.contentTypeOf(target) ?? "application/octet-stream",
        bounds,
        anchor: range,
      };
      drawings.push(pictureModel);
      continue;
    }
    if (frame !== undefined) {
      const graphic = firstChild(frame, "graphic");
      const data = graphic === undefined ? undefined : firstChild(graphic, "graphicData");
      const chartRef = data === undefined ? undefined : firstChild(data, "chart");
      const chartId = chartRef === undefined ? undefined : attr(chartRef, "r:id");
      const nameElement = firstChild(firstChild(frame, "nvGraphicFramePr") ?? frame, "cNvPr");
      const name = nameElement === undefined ? "Chart" : (attr(nameElement, "name") ?? "Chart");
      const target = chartId === undefined ? undefined : drawingRels.find((relationship) => relationship.id === chartId)?.target;
      if (target === undefined || !pkg.hasPart(target)) {
        unsupported.push({ kind: "drawing", detail: `a chart ("${name}") points at a missing part; its frame is kept`, range });
        continue;
      }
      const chartRoot = parseXml(await pkg.readText(target, signal));
      throwIfAborted(signal);
      const adapted = adaptChartGraphic(chartRoot);
      const chartModel: SheetChartModel = {
        kind: "chart",
        name,
        bounds,
        anchor: range,
        sourceChartId: target,
        ...(adapted.graphic === undefined ? {} : { graphic: adapted.graphic }),
        ...(adapted.unsupportedReason === undefined ? {} : { unsupportedReason: adapted.unsupportedReason }),
      };
      drawings.push(chartModel);
      if (adapted.unsupportedReason !== undefined) {
        unsupported.push({ kind: "drawing", detail: `chart "${name}": ${adapted.unsupportedReason}`, range });
      }
      continue;
    }
    unsupported.push({ kind: "drawing", detail: "a drawing shape is not previewed; its anchor is skipped", range });
  }
  throwIfAborted(signal);
  return { drawings, unsupported };
}

/** Inflate and parse exactly one worksheet previously listed by `indexXlsx()`. */
export async function parseXlsxSheet(
  pkg: OpcPackage,
  workbook: XlsxWorkbook,
  index: number,
  signal?: AbortSignal,
): Promise<SheetModel> {
  throwIfAborted(signal);
  const reference = workbook.sheets[index];
  if (reference === undefined) {
    throw new AaredError("not-found", `aared/core: sheet index ${index} is outside this ${workbook.sheets.length}-sheet workbook`);
  }
  const root = parseXml(await pkg.readText(reference.part, signal));
  throwIfAborted(signal);
  if (root.local !== "worksheet") throw new Error(`aared/core: "${reference.part}" root must be <worksheet>`);
  const unsupported: SheetUnsupportedModel[] = [];

  const dimensionElement = firstChild(root, "dimension");
  const dimension = parseDimension(dimensionElement === undefined ? undefined : attr(dimensionElement, "ref"));
  const views = firstChild(root, "sheetViews");
  const view = views === undefined ? undefined : firstChild(views, "sheetView");
  const rightToLeft = view !== undefined && booleanAttribute(view, "rightToLeft");
  // `showGridLines` defaults to true when the attribute is absent; the
  // viewer maps gridline mode "source" to this authored value.
  const showGridlines = view === undefined || attr(view, "showGridLines") !== "0";
  const pane = view === undefined ? undefined : firstChild(view, "pane");
  const frozen = {
    rows: pane === undefined ? 0 : Math.max(Math.trunc(numberAttribute(pane, "ySplit", "x:pane") ?? 0), 0),
    cols: pane === undefined ? 0 : Math.max(Math.trunc(numberAttribute(pane, "xSplit", "x:pane") ?? 0), 0),
  };

  const format = firstChild(root, "sheetFormatPr");
  const defaultRowPoints = format === undefined ? 15 : (numberAttribute(format, "defaultRowHeight", "sheetFormatPr") ?? 15);
  const defaultRowHeight = pointsToInches(defaultRowPoints);
  const defaultColChars = format === undefined ? 8.43 : (numberAttribute(format, "defaultColWidth", "sheetFormatPr") ?? 8.43);

  const columnWidths = new Map<number, number>();
  const columnHidden = new Set<number>();
  const cols = firstChild(root, "cols");
  if (cols !== undefined) {
    for (const col of childElements(cols, "col")) {
      throwIfAborted(signal);
      const min = numberAttribute(col, "min", "col") ?? 1;
      const max = numberAttribute(col, "max", "col") ?? min;
      const width = attr(col, "width");
      const hidden = booleanAttribute(col, "hidden");
      for (let i = min; i <= max; i++) {
        if (width !== undefined) {
          const chars = Number(width);
          if (Number.isFinite(chars)) columnWidths.set(i - 1, columnCharsToInches(chars));
        }
        if (hidden) columnHidden.add(i - 1);
      }
    }
  }

  const sheetRels = await pkg.relationships(reference.part, signal);
  throwIfAborted(signal);
  const hyperlinks = hyperlinkTargets(root, sheetRels.map((relationship) => ({ id: relationship.id, target: relationship.target })));

  const cells = new Map<string, CellModel>();
  const rowHeights = new Map<number, number>();
  const hiddenRows = new Set<number>();
  const data = firstChild(root, "sheetData");
  if (data !== undefined) {
    for (const row of childElements(data, "row")) {
      throwIfAborted(signal);
      const rowIndex = (numberAttribute(row, "r", "row") ?? 0) - 1;
      if (rowIndex < 0) continue;
      if (booleanAttribute(row, "hidden")) hiddenRows.add(rowIndex);
      const heightPoints = numberAttribute(row, "ht", "row");
      if (heightPoints !== undefined) rowHeights.set(rowIndex, pointsToInches(heightPoints));
      for (const cell of childElements(row, "c")) {
        throwIfAborted(signal);
        const ref = attr(cell, "r");
        if (ref === undefined) continue;
        let position: { row: number; col: number };
        try {
          position = parseCellReference(ref);
        } catch {
          continue;
        }
        const styleId = numberAttribute(cell, "s", "c") ?? 0;
        const formula = firstChild(cell, "f");
        const style = workbook.styles.cellStyle(styleId);
        const cached = cachedValue(cell, workbook.sharedStrings, style, workbook.date1904);
        const model: CellModel = {
          value: cached.value,
          formatted: cached.formatted,
          type: cached.type,
          ...(formula === undefined || textContent(formula) === "" ? {} : { formula: textContent(formula) }),
        };
        if (Object.keys(style).length > 0) model.style = style;
        const hyperlink = hyperlinks.get(`${position.row}:${position.col}`);
        if (hyperlink !== undefined) model.hyperlink = { ...hyperlink };
        cells.set(`${position.row}:${position.col}`, model);
      }
    }
  }

  const merges: CellRange[] = [];
  const mergeCells = firstChild(root, "mergeCells");
  if (mergeCells !== undefined) {
    for (const merge of childElements(mergeCells, "mergeCell")) {
      throwIfAborted(signal);
      const ref = attr(merge, "ref");
      if (ref === undefined) continue;
      try {
        merges.push(parseCellRange(ref));
      } catch {
        continue;
      }
    }
  }

  const formats = conditionalRules(root);
  if (formats.length > 0) applyConditionalFormatting(cells, formats, workbook.styles, unsupported);

  // Extent: authored dimension wins; otherwise derive from content so a
  // missing dimension still yields a bounded, renderable grid.
  let extent = dimension;
  if (extent === undefined) {
    let cols = 0;
    let rows = 0;
    for (const key of cells.keys()) {
      const [row, col] = key.split(":").map(Number);
      rows = Math.max(rows, (row ?? -1) + 1);
      cols = Math.max(cols, (col ?? -1) + 1);
    }
    for (const merge of merges) {
      rows = Math.max(rows, merge.end.row + 1);
      cols = Math.max(cols, merge.end.col + 1);
    }
    extent = { cols, rows };
  }

  const widths: number[] = [];
  for (let col = 0; col < extent.cols; col++) {
    widths.push(columnWidths.get(col) ?? columnCharsToInches(defaultColChars));
  }
  const hiddenColumns = [...columnHidden].filter((col) => col < extent.cols).sort((a, b) => a - b);
  const hidden = [...hiddenRows].filter((row) => row < extent.rows).sort((a, b) => a - b);

  const rows: RowModel[] = [];
  const byRow = new Map<number, Array<{ col: number; cell: CellModel }>>();
  for (const [key, cell] of cells) {
    const [row, col] = key.split(":").map(Number);
    const list = byRow.get(row!) ?? [];
    list.push({ col: col!, cell });
    byRow.set(row!, list);
  }
  const orderedRows = [...byRow.keys()].sort((a, b) => a - b);
  for (const rowIndex of orderedRows) {
    if (rowIndex >= extent.rows) continue;
    const entries = (byRow.get(rowIndex) ?? []).filter((entry) => entry.col < extent.cols).sort((a, b) => a.col - b.col);
    if (entries.length === 0) continue;
    const height = rowHeights.get(rowIndex);
    rows.push({
      index: rowIndex,
      cells: entries,
      ...(height === undefined ? {} : { height }),
      ...(hiddenRows.has(rowIndex) ? { hidden: true as const } : {}),
    });
  }

  // Drawing bounds need absolute EMU offsets: column widths plus row heights.
  const columnEmu = widths.map((width) => width * EMUS_PER_INCH);
  const rowEmu: number[] = [];
  for (let row = 0; row < extent.rows; row++) {
    rowEmu.push((rowHeights.get(row) ?? defaultRowHeight) * EMUS_PER_INCH);
  }
  const { drawings, unsupported: drawingUnsupported } = await parseDrawings(pkg, reference.part, columnEmu, rowEmu, signal);
  throwIfAborted(signal);
  unsupported.push(...drawingUnsupported);

  const validations = firstChild(root, "dataValidations");
  if (validations !== undefined && (numberAttribute(validations, "count", "dataValidations") ?? childElements(validations, "dataValidation").length) > 0) {
    unsupported.push({ kind: "data-validation", detail: "drop-down and input validation are not interactive in previews" });
  }
  if (sheetRels.some((relationship) => relationship.type.endsWith("/comments"))) {
    unsupported.push({ kind: "comment", detail: "cell comments are not shown in previews" });
  }
  if (sheetRels.some((relationship) => relationship.type.endsWith("/pivotTable"))) {
    unsupported.push({ kind: "pivot", detail: "pivot tables are shown as their cached values, never recalculated" });
  }
  if (pkg.partNames().some((part) => part.toLowerCase() === "xl/vbaproject.bin")) {
    unsupported.push({ kind: "macro", detail: "this workbook contains macros; code never runs in previews" });
  }
  const extensionList = firstChild(root, "extLst");
  if (extensionList !== undefined && textContent(extensionList).includes("sparkline")) {
    unsupported.push({ kind: "sparkline", detail: "sparklines are not drawn in previews" });
  }
  if (sheetRels.some((relationship) => relationship.type.endsWith("/oleObject") || relationship.type.endsWith("/control"))) {
    unsupported.push({ kind: "other", detail: "embedded objects are not shown in previews" });
  }

  throwIfAborted(signal);
  return {
    name: reference.name,
    part: reference.part,
    rightToLeft,
    showGridlines,
    extent: { ...extent },
    rows,
    merges,
    frozen: { ...frozen },
    columnWidths: widths,
    hiddenColumns,
    hiddenRows: hidden,
    defaultRowHeight,
    drawings,
    unsupported,
  };
}

/**
 * Window a parsed sheet to a numeric range for `SheetHandle.read(range)`.
 * Sheet-level metadata (name, extent, frozen panes, column widths,
 * drawings) travels intact — it is cheap and renderers need it for
 * geometry — while `rows` is sliced to the window and `merges` clipped
 * to their intersection with it. Rows absent from the window read as
 * empty, exactly like sparse authored rows.
 */
export function sliceSheetToRange(sheet: SheetModel, range: CellRange): SheetRangeModel {
  const rows = sheet.rows.filter((row) => row.index >= range.start.row && row.index <= range.end.row)
    .map((row) => ({
      ...row,
      cells: row.cells.filter(({ col }) => col >= range.start.col && col <= range.end.col),
    }))
    .filter((row) => row.cells.length > 0);
  const merges: CellRange[] = [];
  for (const merge of sheet.merges) {
    const start = {
      row: Math.max(merge.start.row, range.start.row),
      col: Math.max(merge.start.col, range.start.col),
    };
    const end = {
      row: Math.min(merge.end.row, range.end.row),
      col: Math.min(merge.end.col, range.end.col),
    };
    if (start.row <= end.row && start.col <= end.col) merges.push({ start, end });
  }
  return {
    ...sheet,
    frozen: { ...sheet.frozen },
    rows,
    merges,
    range: {
      start: { ...range.start },
      end: { ...range.end },
    },
  };
}
