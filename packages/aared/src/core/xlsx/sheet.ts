/**
 * Worksheet parser: SpreadsheetML sheet parts into resolved `SheetModel`s.
 *
 * View-only throughout: formulas are never evaluated (the file's cached
 * `<v>` is formatted and shown), conditional formatting is a one-time
 * snapshot computed from cached values, and charts are drawn by the shared
 * `core/chart.ts` module, exactly as PPTX charts are.
 */

import { attr, childElements, firstChild, numberAttr, textContent } from "../xml.js";
import type { XmlElement } from "../xml.js";
import type { OpcPackage } from "../opc.js";
import { mixColors } from "../color.js";
import { AaredError, throwIfAborted } from "../errors.js";
import { PX_PER_INCH, emuToInches, inchesToEmu, ptToInches } from "../units.js";
import { MAX_COLUMNS } from "./limits.js";
import { parseCell, spanRange } from "./range.js";
import { formatNumberValue } from "./numberfmt.js";
import type { XlsxStyles } from "./styles.js";
import type { XlsxWorkbook } from "./workbook.js";
import { chartUnsupportedReason, parseChartGraphic } from "../chart.js";
import type { DrawingTheme } from "../theme.js";
import type {
  BoundsModel,
  CellModel,
  CellRange,
  ChartGraphicModel,
  RowModel,
  SheetChartModel,
  SheetDrawingModel,
  SheetModel,
  SheetPictureModel,
  SheetRangeModel,
  SheetUnsupportedModel,
  XlsxCellType,
} from "../model.js";

/** Excel column-width units (characters) to pixels at the default font. */
const PIXELS_PER_COLUMN_CHAR = 7;
const COLUMN_PADDING_PIXELS = 5;

function columnCharsToInches(width: number): number {
  return (width * PIXELS_PER_COLUMN_CHAR + COLUMN_PADDING_PIXELS) / PX_PER_INCH;
}

function booleanAttribute(element: XmlElement, name: string): boolean {
  const value = attr(element, name);
  return value === "1" || value === "true";
}

/** Parse an A1 cell reference into zero-based coordinates. */
export function parseCellReference(reference: string): { row: number; col: number } {
  const cell = parseCell(reference);
  if (cell === undefined) throw new Error(`aared/core: invalid cell reference "${reference}"`);
  return cell;
}

/** Parse an A1 range (`A1:H1`) into an inclusive zero-based `CellRange`. */
export function parseCellRange(reference: string): CellRange {
  const parts = reference.split(":");
  const start = parseCellReference(parts[0]!);
  return spanRange(start, parts.length === 1 ? start : parseCellReference(parts[1]!));
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
  const v = firstChild(cell, "v");
  const raw = v === undefined ? undefined : textContent(v);
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

/** One parsed cell with its column, as `RowModel.cells` stores it. */
type GridEntry = RowModel["cells"][number];

/**
 * Parsed cells grouped by row, each row's cells sorted by column. This is
 * both the shape `SheetModel.rows` needs and the lookup structure for
 * hyperlinks and conditional formats: no per-cell `"row:col"` string keys,
 * and range visits cost the cells that exist, not the range's area (a rule
 * on `A1:A1048576` touches only authored cells).
 */
interface CellGrid {
  byRow: Map<number, GridEntry[]>;
  /** Keys of `byRow`, ascending. */
  rowOrder: number[];
}

/** First index in `sorted` whose value is >= `target`. */
function lowerBound<T>(sorted: readonly T[], target: number, key: (item: T) => number): number {
  let low = 0;
  let high = sorted.length;
  while (low < high) {
    const mid = (low + high) >>> 1;
    if (key(sorted[mid]!) < target) low = mid + 1;
    else high = mid;
  }
  return low;
}

const entryCol = (entry: GridEntry): number => entry.col;
const identity = (value: number): number => value;
const rowIndexOf = (row: RowModel): number => row.index;

/** Visit every parsed cell inside `range`, row-major, columns ascending. */
function forEachCellInRange(grid: CellGrid, range: CellRange, visit: (cell: CellModel) => void): void {
  const { start, end } = range;
  const visitRow = (entries: GridEntry[]): void => {
    for (let i = lowerBound(entries, start.col, entryCol); i < entries.length && entries[i]!.col <= end.col; i++) {
      visit(entries[i]!.cell);
    }
  };
  const { byRow, rowOrder } = grid;
  if (end.row - start.row < rowOrder.length) {
    for (let row = start.row; row <= end.row; row++) {
      const entries = byRow.get(row);
      if (entries !== undefined) visitRow(entries);
    }
    return;
  }
  for (let i = lowerBound(rowOrder, start.row, identity); i < rowOrder.length && rowOrder[i]! <= end.row; i++) {
    visitRow(byRow.get(rowOrder[i]!)!);
  }
}

/** Resolve `<hyperlinks>` onto the parsed cells; a later link over the same cell wins. */
function applyHyperlinks(sheet: XmlElement, relationships: readonly { id: string; target: string }[], grid: CellGrid): void {
  const links = firstChild(sheet, "hyperlinks");
  if (links === undefined) return;
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
    let range: CellRange;
    try {
      range = parseCellRange(ref);
    } catch {
      continue;
    }
    // One object per link, shared by the cells it covers (read-only model).
    const hyperlink = { target, ...(display === undefined ? {} : { display }) };
    forEachCellInRange(grid, range, (cell) => {
      cell.hyperlink = hyperlink;
    });
  }
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
      const priority = numberAttr(element, "priority", "c:cfRule") ?? 0;
      const dxfId = numberAttr(element, "dxfId", "c:cfRule");
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

/** Min and max of `values` without spreading them into `Math.min` (which overflows the stack on big ranges). */
function extremes(values: readonly number[]): { min: number; max: number } | undefined {
  if (values.length === 0) return undefined;
  let min = values[0]!;
  let max = min;
  for (const value of values) {
    if (value < min) min = value;
    if (value > max) max = value;
  }
  return { min, max };
}

function scaleBound(bound: { type: string; value: number }, values: readonly number[], edge: "min" | "max"): number | undefined {
  if (bound.type === "num" && Number.isFinite(bound.value)) return bound.value;
  if (bound.type === "min" || bound.type === "max") {
    const found = extremes(values);
    return found === undefined ? undefined : found[edge];
  }
  return undefined;
}

function numericValuesIn(grid: CellGrid, ranges: readonly CellRange[]): number[] {
  const values: number[] = [];
  for (const range of ranges) {
    forEachCellInRange(grid, range, (cell) => {
      if (typeof cell.value === "number") values.push(cell.value);
    });
  }
  return values;
}

/** Apply cached conditional-formatting results onto parsed cells (view-only snapshot). */
function applyConditionalFormatting(
  grid: CellGrid,
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
      const numericValues = numericValuesIn(grid, ranges);
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
        forEachCellInRange(grid, range, (cell) => {
          if (typeof cell.value !== "number" || cell.conditional?.fill !== undefined) return;
          cell.conditional = { ...(cell.conditional ?? {}), fill: mixColors(from, to, (cell.value - low) / span) };
        });
      }
      continue;
    }
    if (rule.dataBar !== undefined) {
      const bar = rule.dataBar;
      const numericValues = numericValuesIn(grid, ranges);
      const low = scaleBound(bar.min, numericValues, "min");
      const high = scaleBound(bar.max, numericValues, "max");
      if (low === undefined || high === undefined) {
        reportOnce("a data-bar rule uses unsupported bounds; its bars are not drawn");
        continue;
      }
      const span = high - low === 0 ? 1 : high - low;
      for (const range of ranges) {
        forEachCellInRange(grid, range, (cell) => {
          if (typeof cell.value !== "number" || cell.conditional?.dataBar !== undefined) return;
          cell.conditional = {
            ...(cell.conditional ?? {}),
            dataBar: { ratio: Math.min(Math.max((cell.value - low) / span, 0), 1), color: bar.color },
          };
        });
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
      forEachCellInRange(grid, range, (cell) => {
        if (cell.conditional?.fill !== undefined) return;
        if (!matchesCellIs(cell, rule.operator, rule.formulas)) return;
        cell.conditional = {
          ...(cell.conditional ?? {}),
          ...(differential.fill === undefined ? {} : { fill: differential.fill }),
          ...(differential.font?.color === undefined ? {} : { color: differential.font.color }),
          ...(differential.font?.bold === undefined ? {} : { bold: differential.font.bold }),
        };
      });
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

/**
 * Spreadsheet charts go through the shared chart module (`core/chart.ts`),
 * so XLSX and PPTX draw the same chart the same way. Any chart the shared
 * module cannot draw is reported through `unsupportedReason` instead of
 * guessed at; the drawing keeps its frame either way.
 */
export function adaptChartGraphic(chart: XmlElement, theme: DrawingTheme): { graphic?: ChartGraphicModel; unsupportedReason?: string } {
  const space = chart.local === "chartSpace" ? chart : firstChild(chart, "chartSpace");
  const graphic = parseChartGraphic(space ?? chart, theme);
  // A chart with series but no cached points draws nothing: say so instead of an empty frame.
  if (graphic !== undefined && graphic.primitives.length > 0) return { graphic };
  return { unsupportedReason: chartUnsupportedReason(space ?? chart) };
}

/**
 * Absolute EMU offset of a grid line: the sum of the first `index` track
 * sizes, clamped to the sheet extent (anchors past it stop growing, exactly
 * as a precomputed per-track array would). Summed in track order so the
 * floating-point result is stable; only called per drawing anchor, so the
 * sheet never materializes a row-height array the size of its extent.
 */
type GridOffset = (index: number) => number;

async function parseDrawings(
  pkg: OpcPackage,
  sheetPart: string,
  columnOffset: GridOffset,
  rowOffset: GridOffset,
  theme: DrawingTheme,
  signal?: AbortSignal,
): Promise<{ drawings: SheetDrawingModel[]; unsupported: SheetUnsupportedModel[] }> {
  throwIfAborted(signal);
  const drawings: SheetDrawingModel[] = [];
  const unsupported: SheetUnsupportedModel[] = [];
  const relationships = await pkg.relationships(sheetPart, signal);
  throwIfAborted(signal);
  const drawingRel = relationships.find((relationship) => relationship.type.endsWith("/drawing"));
  if (drawingRel === undefined) return { drawings, unsupported };
  const drawing = await pkg.readXmlIfPresent(drawingRel.target, signal);
  if (drawing === undefined) return { drawings, unsupported };
  const drawingRels = await pkg.relationships(drawingRel.target, signal);
  throwIfAborted(signal);
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
        const x = emuToInches(columnOffset(from.col) + from.colOff);
        const y = emuToInches(rowOffset(from.row) + from.rowOff);
        const x2 = emuToInches(columnOffset(to.col) + to.colOff);
        const y2 = emuToInches(rowOffset(to.row) + to.rowOff);
        bounds = { x, y, width: Math.max(x2 - x, 0), height: Math.max(y2 - y, 0) };
        range = {
          start: { row: Math.min(from.row, to.row), col: Math.min(from.col, to.col) },
          end: { row: Math.max(from.row, to.row), col: Math.max(from.col, to.col) },
        };
      } else if (fromElement !== undefined && extElement !== undefined) {
        const from = anchorFrom(fromElement, "xdr:from");
        const cx = numberAttr(extElement, "cx", "a:ext") ?? 0;
        const cy = numberAttr(extElement, "cy", "a:ext") ?? 0;
        bounds = {
          x: emuToInches(columnOffset(from.col) + from.colOff),
          y: emuToInches(rowOffset(from.row) + from.rowOff),
          width: emuToInches(cx),
          height: emuToInches(cy),
        };
        range = { start: { row: from.row, col: from.col }, end: { row: from.row, col: from.col } };
      } else if (posElement !== undefined && extElement !== undefined) {
        const x = emuToInches(numberAttr(posElement, "x", "xdr:pos") ?? 0);
        const y = emuToInches(numberAttr(posElement, "y", "xdr:pos") ?? 0);
        bounds = {
          x,
          y,
          width: emuToInches(numberAttr(extElement, "cx", "a:ext") ?? 0),
          height: emuToInches(numberAttr(extElement, "cy", "a:ext") ?? 0),
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
        unsupported.push({ kind: "drawing", detail: `an image ("${name}") points at a missing part and is not previewed`, range });
        continue;
      }
      const pictureModel: SheetPictureModel = {
        kind: "picture",
        name,
        mediaId: target,
        contentType: pkg.mediaContentType(target),
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
        unsupported.push({ kind: "drawing", detail: `a chart ("${name}") points at a missing part and is not previewed`, range });
        continue;
      }
      const chartRoot = await pkg.readXml(target, signal);
      const adapted = adaptChartGraphic(chartRoot, theme);
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

function hasOwnKeys(value: object): boolean {
  for (const _ in value) return true;
  return false;
}

/**
 * Inflate and parse exactly one worksheet previously listed by `indexXlsx()`.
 *
 * `<sheetData>` is streamed row by row (`parseXml`'s streaming mode): each
 * `<row>` element tree is built, turned into cell models, and dropped, so
 * peak memory tracks the resulting model rather than the sheet's XML tree.
 */
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
  const { styles, sharedStrings, date1904 } = workbook;
  const byRow = new Map<number, GridEntry[]>();
  /** Rows whose cells arrived out of column order (or repeated): sorted and de-duplicated once below. */
  const unsortedRows = new Set<number>();
  const rowHeights = new Map<number, number>();
  const hiddenRows = new Set<number>();
  const readRow = (row: XmlElement): void => {
    if (row.local !== "row") return;
    throwIfAborted(signal);
    const rowIndex = (numberAttr(row, "r", "row") ?? 0) - 1;
    if (rowIndex < 0) return;
    if (booleanAttribute(row, "hidden")) hiddenRows.add(rowIndex);
    const heightPoints = numberAttr(row, "ht", "row");
    if (heightPoints !== undefined) rowHeights.set(rowIndex, ptToInches(heightPoints));
    for (const cell of row.children) {
      if (cell.local !== "c") continue;
      const ref = attr(cell, "r");
      if (ref === undefined) continue;
      let position: { row: number; col: number };
      try {
        position = parseCellReference(ref);
      } catch {
        continue;
      }
      const style = styles.cellStyle(numberAttr(cell, "s", "c") ?? 0);
      const cached = cachedValue(cell, sharedStrings, style, date1904);
      const formulaElement = firstChild(cell, "f");
      const formula = formulaElement === undefined ? "" : textContent(formulaElement);
      const model: CellModel = {
        value: cached.value,
        formatted: cached.formatted,
        type: cached.type,
        ...(formula === "" ? {} : { formula }),
      };
      if (hasOwnKeys(style)) model.style = style;
      // A cell's own reference decides its row, even inside another <row>.
      let entries = byRow.get(position.row);
      if (entries === undefined) {
        entries = [];
        byRow.set(position.row, entries);
      } else if (entries[entries.length - 1]!.col >= position.col) {
        unsortedRows.add(position.row);
      }
      entries.push({ col: position.col, cell: model });
    }
  };
  const root = await pkg.readXml(reference.part, signal, { streamChildrenOf: "sheetData", onChild: readRow });
  if (root.local !== "worksheet") throw new Error(`aared/core: "${reference.part}" root must be <worksheet>`);
  for (const row of unsortedRows) {
    // Stable sort, then keep the last cell written for a repeated reference.
    const sorted = byRow.get(row)!.sort((a, b) => a.col - b.col);
    byRow.set(row, sorted.filter((entry, i) => sorted[i + 1]?.col !== entry.col));
  }
  const grid: CellGrid = { byRow, rowOrder: [...byRow.keys()].sort((a, b) => a - b) };
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
    rows: pane === undefined ? 0 : Math.max(Math.trunc(numberAttr(pane, "ySplit", "x:pane") ?? 0), 0),
    cols: pane === undefined ? 0 : Math.max(Math.trunc(numberAttr(pane, "xSplit", "x:pane") ?? 0), 0),
  };

  const format = firstChild(root, "sheetFormatPr");
  const defaultRowPoints = format === undefined ? 15 : (numberAttr(format, "defaultRowHeight", "sheetFormatPr") ?? 15);
  const defaultRowHeight = ptToInches(defaultRowPoints);
  const defaultColChars = format === undefined ? 8.43 : (numberAttr(format, "defaultColWidth", "sheetFormatPr") ?? 8.43);

  const columnWidths = new Map<number, number>();
  const columnHidden = new Set<number>();
  const cols = firstChild(root, "cols");
  if (cols !== undefined) {
    for (const col of childElements(cols, "col")) {
      throwIfAborted(signal);
      const min = numberAttr(col, "min", "col") ?? 1;
      // Clamp to the grid: a corrupt `max` must not spin through billions of columns.
      const max = Math.min(numberAttr(col, "max", "col") ?? min, MAX_COLUMNS);
      const width = attr(col, "width");
      const chars = width === undefined ? Number.NaN : Number(width);
      const inches = Number.isFinite(chars) ? columnCharsToInches(chars) : undefined;
      const hidden = booleanAttribute(col, "hidden");
      for (let i = min; i <= max; i++) {
        if (inches !== undefined) columnWidths.set(i - 1, inches);
        if (hidden) columnHidden.add(i - 1);
      }
    }
  }

  const sheetRels = await pkg.relationships(reference.part, signal);
  throwIfAborted(signal);
  applyHyperlinks(root, sheetRels, grid);

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
  if (formats.length > 0) applyConditionalFormatting(grid, formats, styles, unsupported);

  // Extent: authored dimension wins; otherwise derive from content so a
  // missing dimension still yields a bounded, renderable grid.
  let extent = dimension;
  if (extent === undefined) {
    let cols = 0;
    let rows = 0;
    for (const [row, entries] of byRow) {
      rows = Math.max(rows, row + 1);
      cols = Math.max(cols, entries[entries.length - 1]!.col + 1);
    }
    for (const merge of merges) {
      rows = Math.max(rows, merge.end.row + 1);
      cols = Math.max(cols, merge.end.col + 1);
    }
    extent = { cols, rows };
  }
  const { rows: extentRows, cols: extentCols } = extent;

  const defaultColumnWidth = columnCharsToInches(defaultColChars);
  const widths: number[] = [];
  for (let col = 0; col < extentCols; col++) {
    widths.push(columnWidths.get(col) ?? defaultColumnWidth);
  }
  const hiddenColumns = [...columnHidden].filter((col) => col < extentCols).sort((a, b) => a - b);
  const hidden = [...hiddenRows].filter((row) => row < extentRows).sort((a, b) => a - b);

  const rows: RowModel[] = [];
  for (const rowIndex of grid.rowOrder) {
    if (rowIndex >= extentRows) break;
    let entries = byRow.get(rowIndex)!;
    // Columns are sorted: trimming past the extent is a single cut.
    if (entries[entries.length - 1]!.col >= extentCols) entries = entries.slice(0, lowerBound(entries, extentCols, entryCol));
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
  const columnOffset: GridOffset = (index) => {
    let total = 0;
    for (let i = 0; i < index && i < extentCols; i++) total += inchesToEmu(widths[i]!);
    return total;
  };
  const rowOffset: GridOffset = (index) => {
    let total = 0;
    for (let i = 0; i < index && i < extentRows; i++) total += inchesToEmu(rowHeights.get(i) ?? defaultRowHeight);
    return total;
  };
  const { drawings, unsupported: drawingUnsupported } = await parseDrawings(pkg, reference.part, columnOffset, rowOffset, workbook.theme, signal);
  throwIfAborted(signal);
  unsupported.push(...drawingUnsupported);

  const validations = firstChild(root, "dataValidations");
  if (validations !== undefined && (numberAttr(validations, "count", "dataValidations") ?? childElements(validations, "dataValidation").length) > 0) {
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
  // Rows are sorted by index and cells by column: binary-search both edges.
  const rows: RowModel[] = [];
  const all = sheet.rows;
  for (let i = lowerBound(all, range.start.row, rowIndexOf); i < all.length && all[i]!.index <= range.end.row; i++) {
    const row = all[i]!;
    const cells = row.cells;
    const from = lowerBound(cells, range.start.col, entryCol);
    let to = from;
    while (to < cells.length && cells[to]!.col <= range.end.col) to++;
    if (to > from) rows.push({ ...row, cells: cells.slice(from, to) });
  }
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
