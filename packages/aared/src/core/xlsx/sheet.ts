/**
 * Worksheet parser: SpreadsheetML sheet parts into resolved `SheetModel`s.
 *
 * View-only throughout: formulas are never evaluated (the file's cached
 * `<v>` is formatted and shown), conditional formatting is a one-time
 * snapshot computed from cached values, and charts are read by the shared
 * `core/chart.ts` module, exactly as PPTX charts are.
 */

import { attr, childElements, firstChild, numberAttr, onOff, textContent } from "../xml.js";
import type { XmlElement } from "../xml.js";
import type { OpcPackage } from "../opc.js";
import { mixColors } from "../color.js";
import { AaredError, damagedDocument, throwIfAborted } from "../errors.js";
import { PX_PER_INCH, emuToInches, inchesToEmu, ptToInches } from "../units.js";
import { MAX_COLUMNS, MAX_MERGED_CELLS, MAX_MERGED_CELLS_PER_SHEET, MAX_ROWS } from "./limits.js";
import { parseCell, parseRef } from "./range.js";
import { formatNumberValue } from "./numberfmt.js";
import type { XlsxStyles } from "./styles.js";
import { stringItemText } from "./workbook.js";
import type { XlsxWorkbook } from "./workbook.js";
import { chartFillOwner, chartexFormulas, readChartModel } from "../chart.js";
import { chartText } from "../chart/text.js";
import type { SheetText } from "../rpc.js";
import { resolveChartexData } from "./chart-data.js";
import { themeForColorMap } from "../theme.js";
import type { DrawingTheme } from "../theme.js";
import type {
  BoundsModel,
  CellModel,
  CellRange,
  RowModel,
  SheetChartModel,
  SheetDrawingModel,
  SheetModel,
  SheetPictureModel,
  SheetRangeModel,
  SheetUnsupportedModel,
  XlsxCellType,
} from "../model.js";

/** The widest digit of Excel's default font (Calibri 11, Aptos Narrow 11) at 96 dpi. */
const PIXELS_PER_COLUMN_CHAR = 7;
/** The width Excel gives a column no `<col>` or `defaultColWidth` sizes (8.43 characters plus padding). */
const DEFAULT_COLUMN_PIXELS = 64;

/**
 * A stored column width (`col/@width`, `defaultColWidth`) in inches. Excel
 * stores it with the cell padding already included (ECMA-376 18.3.1.13:
 * `Truncate((chars * 7 + 5) / 7 * 256) / 256`), so the pixels are the stored
 * number times the digit width, truncated: width 26 is 182 px, 16 is 112 px
 * (read back from Excel; DECISION_LOG §72). Adding the padding again made every
 * column 5 px too wide.
 */
export function columnCharsToInches(width: number): number {
  return Math.trunc(((256 * width + Math.trunc(128 / PIXELS_PER_COLUMN_CHAR)) / 256) * PIXELS_PER_COLUMN_CHAR) / PX_PER_INCH;
}

function mergedCellCount(range: CellRange): number {
  return (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
}

/**
 * Cut every merge to the sheet's extent (cells past it are never shown) and
 * refuse the ones still past the budgets: the page indexes each covered cell.
 */
function boundMerges(merges: readonly CellRange[], extentRows: number, extentCols: number, skip: SkipDamaged): CellRange[] {
  const kept: CellRange[] = [];
  let total = 0;
  for (const merge of merges) {
    const cut: CellRange = {
      start: merge.start,
      end: {
        row: Math.min(merge.end.row, Math.max(extentRows - 1, merge.start.row)),
        col: Math.min(merge.end.col, Math.max(extentCols - 1, merge.start.col)),
      },
    };
    const cells = mergedCellCount(cut);
    if (cells > MAX_MERGED_CELLS || total + cells > MAX_MERGED_CELLS_PER_SHEET) {
      skip("merged range");
      continue;
    }
    total += cells;
    kept.push(cut);
  }
  return kept;
}

/**
 * EMU from the grid's edge to the start of track `index`: every track before
 * it at `fallback` inches, corrected by the ones `authored` sizes otherwise.
 */
function gridOffset(index: number, fallback: number, authored: ReadonlyMap<number, number>): number {
  let inches = index * fallback;
  for (const [track, size] of authored) if (track < index) inches += size - fallback;
  return inchesToEmu(inches);
}

/**
 * The track whose span holds `emu`, at or after `from` (a drawing's first
 * track) and at most `limit - 1`. Offsets only grow, so a binary search: a
 * drawing claiming to be a million rows tall costs twenty lookups.
 */
function trackAt(emu: number, from: number, limit: number, offset: (index: number) => number): number {
  let lo = Math.min(Math.max(from, 0), limit - 1);
  let hi = limit - 1;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (offset(mid + 1) < emu) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

function booleanAttribute(element: XmlElement, name: string): boolean {
  return onOff(element, name) === true;
}

/** Parse a `dimension` ref (`A1:H34`); single cells yield a 1x1 extent. */
function parseDimension(reference: string | undefined): { cols: number; rows: number } | undefined {
  const range = reference === undefined ? undefined : parseRef(reference);
  return range === undefined ? undefined : { cols: range.end.col + 1, rows: range.end.row + 1 };
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
      const inline = firstChild(cell, "is");
      const text = inline === undefined ? "" : stringItemText(inline);
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

/** Records one item a damaged sheet lost, by what it was ("cell", "hyperlink", ...). */
type SkipDamaged = (what: string) => void;

/**
 * What a damaged sheet lost, one `damaged` entry per kind of item with its
 * count, so a broken file adds a line per kind rather than one per cell.
 */
function damagedEntries(skipped: ReadonlyMap<string, number>): SheetUnsupportedModel[] {
  return [...skipped].map(([what, count]) => ({
    kind: "damaged" as const,
    detail: `${count} ${what}${count === 1 ? " could not be read and was" : "s could not be read and were"} skipped; the file may be damaged`,
  }));
}

/** Resolve `<hyperlinks>` onto the parsed cells; a later link over the same cell wins. */
function applyHyperlinks(sheet: XmlElement, relationships: readonly { id: string; target: string }[], grid: CellGrid, skip: SkipDamaged): void {
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
    const range = parseRef(ref);
    if (range === undefined) {
      skip("hyperlink");
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
  /** `cfRule/@type`: `cellIs`, `colorScale`, `duplicateValues`, `expression`, ... */
  type?: string;
  /** Lower-priority rules no longer apply to a cell this rule matched. */
  stopIfTrue: boolean;
  dxfId?: number;
  operator?: string;
  /** `cfRule/@text`: what a `containsText`, `notContainsText`, `beginsWith` or `endsWith` rule looks for. */
  text?: string;
  formulas: string[];
  colorScale?: { values: Array<{ type: string; value: number }>; colors: string[] };
  dataBar?: { min: { type: string; value: number }; max: { type: string; value: number }; color: string };
}

function rgbColor(element: XmlElement | undefined): string | undefined {
  const rgb = element === undefined ? undefined : attr(element, "rgb");
  if (rgb === undefined || rgb.length !== 8) return undefined;
  return `#${rgb.slice(2).toUpperCase()}`;
}

function conditionalRules(sheet: XmlElement, skip: SkipDamaged): Array<{ ranges: CellRange[]; rule: ConditionalRule }> {
  const out: Array<{ ranges: CellRange[]; rule: ConditionalRule }> = [];
  for (const formatting of childElements(sheet, "conditionalFormatting")) {
    const refs = attr(formatting, "sqref")?.split(/\s+/).filter((part) => part !== "") ?? [];
    const ranges: CellRange[] = [];
    for (const ref of refs) {
      const range = parseRef(ref);
      if (range === undefined) skip("conditional-format range");
      else ranges.push(range);
    }
    if (ranges.length === 0) continue;
    for (const element of childElements(formatting, "cfRule")) {
      const priority = numberAttr(element, "priority", "cfRule") ?? 0;
      const dxfId = numberAttr(element, "dxfId", "cfRule");
      const operator = attr(element, "operator");
      const type = attr(element, "type");
      const text = attr(element, "text");
      const rule: ConditionalRule = {
        priority,
        stopIfTrue: booleanAttribute(element, "stopIfTrue"),
        ...(type === undefined ? {} : { type }),
        ...(dxfId === undefined ? {} : { dxfId }),
        ...(operator === undefined ? {} : { operator }),
        ...(text === undefined ? {} : { text }),
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
    default: return false;
  }
}

/**
 * The four text rules. Excel stores what to look for in `cfRule/@text` (the
 * `<formula>` only repeats it as a worksheet formula) and compares without
 * regard to case, as its own SEARCH and `=` do.
 */
function matchesText(cell: CellModel, type: string, wanted: string | undefined): boolean {
  if (wanted === undefined || wanted === "") return false;
  const text = cell.formatted.toLowerCase();
  const needle = wanted.toLowerCase();
  switch (type) {
    case "containsText": return text.includes(needle);
    case "notContainsText": return !text.includes(needle);
    case "beginsWith": return text.startsWith(needle);
    default: return text.endsWith(needle);
  }
}

/**
 * The number a `cfvo` bound stands for, over the range's values in ascending
 * order. `percent` is a share of the min-to-max span; `percentile` is Excel's
 * PERCENTILE (linear between ranks). A `formula` bound is not evaluated
 * (view-only), so it resolves to nothing.
 */
function scaleBound(bound: { type: string; value: number }, sorted: readonly number[]): number | undefined {
  if (bound.type === "num") return Number.isFinite(bound.value) ? bound.value : undefined;
  if (sorted.length === 0) return undefined;
  const min = sorted[0]!;
  const max = sorted[sorted.length - 1]!;
  switch (bound.type) {
    case "min": return min;
    case "max": return max;
    case "percent": return Number.isFinite(bound.value) ? min + ((max - min) * bound.value) / 100 : undefined;
    case "percentile": {
      if (!Number.isFinite(bound.value)) return undefined;
      const rank = ((sorted.length - 1) * Math.min(Math.max(bound.value, 0), 100)) / 100;
      const below = Math.floor(rank);
      const above = Math.ceil(rank);
      return sorted[below]! + (sorted[above]! - sorted[below]!) * (rank - below);
    }
    default: return undefined;
  }
}

/** The colour a value takes on a 2- or 3-stop scale: the end colours past the ends, a blend between the two stops it falls between. */
function scaleColor(value: number, stops: readonly number[], colors: readonly string[]): string {
  const last = stops.length - 1;
  if (value <= stops[0]!) return colors[0]!;
  if (value >= stops[last]!) return colors[last]!;
  let segment = 0;
  while (segment < last - 1 && value >= stops[segment + 1]!) segment++;
  const span = stops[segment + 1]! - stops[segment]!;
  return mixColors(colors[segment]!, colors[segment + 1]!, span === 0 ? 0 : (value - stops[segment]!) / span);
}

/** What Excel compares for Duplicate/Unique Values: numbers by value, text without regard to case. */
function duplicateKey(cell: CellModel): string | undefined {
  const value = cell.value;
  if (value === null || value === undefined) return undefined;
  return typeof value === "string" ? `s${value.toLowerCase()}` : `${typeof value}${String(value)}`;
}

/** Rule types the viewer evaluates from cached values; anything else is reported, not guessed. */
const EVALUATED_RULES = new Set(["cellIs", "containsText", "notContainsText", "beginsWith", "endsWith", "duplicateValues", "uniqueValues"]);

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
  // Cells a `stopIfTrue` rule has matched: no later (lower-priority) rule touches them.
  const stopped = new WeakSet<CellModel>();
  for (const { ranges, rule } of formats) {
    if (rule.colorScale !== undefined) {
      const { values, colors } = rule.colorScale;
      if (values.length < 2 || values.length !== colors.length) {
        reportOnce("a color-scale rule is malformed; it is shown without its gradient");
        continue;
      }
      const sorted = numericValuesIn(grid, ranges).sort((a, b) => a - b);
      const stops = values.map((bound) => scaleBound(bound, sorted));
      if (stops.some((stop) => stop === undefined)) {
        reportOnce("a color-scale rule uses a formula bound; it is shown without its gradient");
        continue;
      }
      for (const range of ranges) {
        forEachCellInRange(grid, range, (cell) => {
          if (typeof cell.value !== "number" || stopped.has(cell) || cell.conditional?.fill !== undefined) return;
          cell.conditional = { ...(cell.conditional ?? {}), fill: scaleColor(cell.value, stops as number[], colors) };
          if (rule.stopIfTrue) stopped.add(cell);
        });
      }
      continue;
    }
    if (rule.dataBar !== undefined) {
      const bar = rule.dataBar;
      const sorted = numericValuesIn(grid, ranges).sort((a, b) => a - b);
      const low = scaleBound(bar.min, sorted);
      const high = scaleBound(bar.max, sorted);
      if (low === undefined || high === undefined) {
        reportOnce("a data-bar rule uses unsupported bounds; its bars are not drawn");
        continue;
      }
      const span = high - low === 0 ? 1 : high - low;
      for (const range of ranges) {
        forEachCellInRange(grid, range, (cell) => {
          if (typeof cell.value !== "number" || stopped.has(cell) || cell.conditional?.dataBar !== undefined) return;
          cell.conditional = {
            ...(cell.conditional ?? {}),
            dataBar: { ratio: Math.min(Math.max((cell.value - low) / span, 0), 1), color: bar.color },
          };
          if (rule.stopIfTrue) stopped.add(cell);
        });
      }
      continue;
    }
    if (rule.type !== undefined && !EVALUATED_RULES.has(rule.type)) {
      reportOnce(`a "${rule.type}" conditional rule is not evaluated; its highlighting is not shown`);
      continue;
    }
    if (rule.dxfId === undefined) {
      reportOnce("a conditional rule without a differential style is shown without highlighting");
      continue;
    }
    const differential = styles.differential(rule.dxfId);
    if (differential.fill === undefined && differential.font === undefined) continue;
    // Duplicate/Unique Values compare a cell with every other cell the rule covers.
    let occurrences: Map<string, number> | undefined;
    if (rule.type === "duplicateValues" || rule.type === "uniqueValues") {
      occurrences = new Map();
      for (const range of ranges) {
        forEachCellInRange(grid, range, (cell) => {
          const key = duplicateKey(cell);
          if (key !== undefined) occurrences!.set(key, (occurrences!.get(key) ?? 0) + 1);
        });
      }
    }
    const textRule = rule.type === "containsText" || rule.type === "notContainsText" || rule.type === "beginsWith" || rule.type === "endsWith";
    const matches = (cell: CellModel): boolean => {
      if (textRule) return matchesText(cell, rule.type!, rule.text);
      if (occurrences === undefined) return matchesCellIs(cell, rule.operator, rule.formulas);
      const key = duplicateKey(cell);
      if (key === undefined) return false;
      return rule.type === "duplicateValues" ? occurrences.get(key)! > 1 : occurrences.get(key) === 1;
    };
    for (const range of ranges) {
      forEachCellInRange(grid, range, (cell) => {
        if (stopped.has(cell) || cell.conditional?.fill !== undefined) return;
        if (!matches(cell)) return;
        cell.conditional = {
          ...(cell.conditional ?? {}),
          ...(differential.fill === undefined ? {} : { fill: differential.fill }),
          ...(differential.font?.color === undefined ? {} : { color: differential.font.color }),
          ...(differential.font?.bold === undefined ? {} : { bold: differential.font.bold }),
        };
        if (rule.stopIfTrue) stopped.add(cell);
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
    throw damagedDocument(`invalid drawing anchor in ${context}`);
  }
  return {
    col: colIndex,
    row: rowIndex,
    colOff: colOff === undefined ? 0 : Number(textContent(colOff)) || 0,
    rowOff: rowOff === undefined ? 0 : Number(textContent(rowOff)) || 0,
  };
}

/**
 * Absolute EMU offset of a grid line: the sum of the first `index` track
 * sizes, clamped to the sheet extent (anchors past it stop growing, exactly
 * as a precomputed per-track array would). Summed in track order so the
 * floating-point result is stable; only called per drawing anchor, so the
 * sheet never materializes a row-height array the size of its extent.
 */
type GridOffset = (index: number) => number;

/** The `xdr:graphicFrame` of an Office 2016 chart: inside `mc:AlternateContent/mc:Choice`, holding a `cx:chart`. */
function chartExFrame(anchor: XmlElement): XmlElement | undefined {
  for (const alternate of anchor.children) {
    if (alternate.local !== "AlternateContent") continue;
    for (const choice of alternate.children) {
      if (choice.local !== "Choice") continue;
      const frame = firstChild(choice, "graphicFrame");
      const graphic = frame === undefined ? undefined : firstChild(frame, "graphic");
      const data = graphic === undefined ? undefined : firstChild(graphic, "graphicData");
      if (frame !== undefined && data !== undefined && firstChild(data, "chart") !== undefined) return frame;
    }
  }
  return undefined;
}

async function parseDrawings(
  pkg: OpcPackage,
  workbook: XlsxWorkbook,
  sheetPart: string,
  columnOffset: GridOffset,
  rowOffset: GridOffset,
  skip: SkipDamaged,
  signal?: AbortSignal,
): Promise<{ drawings: SheetDrawingModel[]; unsupported: SheetUnsupportedModel[]; missingIds: number }> {
  throwIfAborted(signal);
  const drawings: SheetDrawingModel[] = [];
  const unsupported: SheetUnsupportedModel[] = [];
  /** Drawings kept whose `xdr:cNvPr` states no id: Excel calls that damage; nothing here uses a drawing's id. */
  let missingIds = 0;
  const lacksId = (cNvPr: XmlElement | undefined): boolean => cNvPr === undefined || attr(cNvPr, "id") === undefined;
  const relationships = await pkg.relationships(sheetPart, signal);
  throwIfAborted(signal);
  const drawingRel = relationships.find((relationship) => relationship.type.endsWith("/drawing"));
  if (drawingRel === undefined) return { drawings, unsupported, missingIds };
  const drawing = await pkg.readXmlIfPresent(drawingRel.target, signal);
  if (drawing === undefined) return { drawings, unsupported, missingIds };
  const drawingRels = await pkg.relationships(drawingRel.target, signal);
  let chartTheme: DrawingTheme | undefined;
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
      skip("drawing anchor");
      continue;
    }
    if (bounds === undefined) continue;
    const picture = firstChild(anchor, "pic");
    // An Office 2016 chart sits in `mc:AlternateContent`: the frame in the Choice, a "not available" text box as the Fallback (never painted).
    const frame = firstChild(anchor, "graphicFrame") ?? chartExFrame(anchor);
    if (picture !== undefined) {
      const blipFill = firstChild(picture, "blipFill");
      const blip = blipFill === undefined ? undefined : firstChild(blipFill, "blip");
      const embed = blip === undefined ? undefined : attr(blip, "r:embed");
      const target = embed === undefined ? undefined : drawingRels.find((relationship) => relationship.id === embed)?.target;
      const nameElement = firstChild(firstChild(picture, "nvPicPr") ?? picture, "cNvPr");
      const name = nameElement === undefined ? "Picture" : (attr(nameElement, "name") ?? "Picture");
      if (target === undefined || !pkg.hasPart(target)) {
        unsupported.push({ kind: "damaged", detail: `an image ("${name}") points at a missing part and is not previewed; the file may be damaged`, range });
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
      if (lacksId(nameElement)) missingIds++;
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
        unsupported.push({ kind: "damaged", detail: `a chart ("${name}") points at a missing part and is not previewed; the file may be damaged`, range });
        continue;
      }
      const chartRoot = await pkg.readXml(target, signal);
      // Excel has no colour map: chart DrawingML uses the default one.
      chartTheme ??= themeForColorMap(workbook.theme, undefined);
      // The same reader as PPTX, so a chart reads the same in both formats; one it
      // cannot draw keeps its frame and says why.
      // An Office 2016 chart carries no values of its own in a workbook: they sit in the cells its formulas name.
      const formulas = chartexFormulas(chartRoot);
      const chartData = formulas.length === 0 ? undefined : await resolveChartexData(pkg, workbook, formulas, signal);
      const read = readChartModel(chartRoot, chartTheme, { owner: await chartFillOwner(pkg, target, signal), ...(chartData === undefined ? {} : { data: chartData }) });
      const alt = nameElement === undefined ? undefined : attr(nameElement, "descr") || attr(nameElement, "title") || undefined;
      const chartModel: SheetChartModel = { kind: "chart", name, ...(alt === undefined ? {} : { alt }), bounds, anchor: range, sourceChartId: target, ...read };
      drawings.push(chartModel);
      if (lacksId(nameElement)) missingIds++;
      if ("unsupportedReason" in read) {
        unsupported.push({ kind: "drawing", detail: `chart "${name}": ${read.unsupportedReason}`, range });
      }
      continue;
    }
    unsupported.push({ kind: "drawing", detail: "a drawing shape is not previewed; its anchor is skipped", range });
  }
  throwIfAborted(signal);
  return { drawings, unsupported, missingIds };
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
  // Items skipped because the file itself is broken, counted by kind (`damagedEntries`).
  const skipped = new Map<string, number>();
  const skip: SkipDamaged = (what) => {
    skipped.set(what, (skipped.get(what) ?? 0) + 1);
  };
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
      const position = parseCell(ref);
      if (position === undefined) {
        skip("cell");
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
  if (root.local !== "worksheet") throw damagedDocument(`"${reference.part}" root must be <worksheet>`);
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
  const defaultColChars = format === undefined ? undefined : numberAttr(format, "defaultColWidth", "sheetFormatPr");

  const columnWidths = new Map<number, number>();
  const columnHidden = new Set<number>();
  const cols = firstChild(root, "cols");
  if (cols !== undefined) {
    // Excel's own <col> ranges never overlap, so together they cover at most the
    // grid once. A file that spends more than that (thousands of full-width
    // ranges) is corrupt: stop reading widths instead of looping per overlap.
    let columnsCovered = 0;
    for (const col of childElements(cols, "col")) {
      throwIfAborted(signal);
      const min = numberAttr(col, "min", "col") ?? 1;
      // Clamp to the grid: a corrupt `max` must not spin through billions of columns.
      const max = Math.min(numberAttr(col, "max", "col") ?? min, MAX_COLUMNS);
      if (!Number.isInteger(min) || min < 1 || min > max) {
        skip("column range");
        continue;
      }
      columnsCovered += max - min + 1;
      if (columnsCovered > MAX_COLUMNS) {
        skip("column range");
        break;
      }
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
  applyHyperlinks(root, sheetRels, grid, skip);

  const merges: CellRange[] = [];
  const mergeCells = firstChild(root, "mergeCells");
  if (mergeCells !== undefined) {
    for (const merge of childElements(mergeCells, "mergeCell")) {
      throwIfAborted(signal);
      const ref = attr(merge, "ref");
      if (ref === undefined) continue;
      const range = parseRef(ref);
      if (range === undefined) skip("merged range");
      else merges.push(range);
    }
  }

  const formats = conditionalRules(root, skip);
  if (formats.length > 0) applyConditionalFormatting(grid, formats, styles, unsupported);

  // Extent: the authored dimension, grown to the cells that exist (some
  // writers leave `<dimension ref="A1"/>` whatever the sheet holds, and Excel
  // does not trust it either). Merges count as content too.
  let extent = { cols: dimension?.cols ?? 0, rows: dimension?.rows ?? 0 };
  for (const [row, entries] of byRow) {
    extent.rows = Math.max(extent.rows, row + 1);
    extent.cols = Math.max(extent.cols, entries[entries.length - 1]!.col + 1);
  }
  for (const merge of merges) {
    // A merge over the whole grid must not make an empty sheet a million rows tall.
    if (mergedCellCount(merge) > MAX_MERGED_CELLS) continue;
    extent.rows = Math.max(extent.rows, merge.end.row + 1);
    extent.cols = Math.max(extent.cols, merge.end.col + 1);
  }
  const { rows: extentRows, cols: extentCols } = extent;
  const boundedMerges = boundMerges(merges, extentRows, extentCols, skip);

  const defaultColumnWidth = defaultColChars === undefined ? DEFAULT_COLUMN_PIXELS / PX_PER_INCH : columnCharsToInches(defaultColChars);

  // A row with no cells but its own height is still a row of the grid (and of
  // the drawings' offsets): it travels with no cells. One at the default
  // height says nothing, so it is left out.
  const heightOnly = [...rowHeights].filter(([row, height]) => !byRow.has(row) && height !== defaultRowHeight).map(([row]) => row);
  const rowOrder = heightOnly.length === 0 ? grid.rowOrder : [...grid.rowOrder, ...heightOnly].sort((a, b) => a - b);
  const rows: RowModel[] = [];
  for (const rowIndex of rowOrder) {
    if (rowIndex >= extentRows) break;
    let entries = byRow.get(rowIndex) ?? [];
    // Columns are sorted: trimming past the extent is a single cut.
    if (entries.length > 0 && entries[entries.length - 1]!.col >= extentCols) entries = entries.slice(0, lowerBound(entries, extentCols, entryCol));
    const height = rowHeights.get(rowIndex);
    if (entries.length === 0 && height === undefined) continue;
    rows.push({
      index: rowIndex,
      cells: entries,
      ...(height === undefined ? {} : { height }),
    });
  }

  // Drawing bounds need absolute EMU offsets: column widths plus row heights,
  // past the used range too (Excel's grid never ends, and a chart often sits
  // beside or below the data). Default tracks are counted at once and only the
  // authored ones walked, so an anchor far down the sheet costs nothing.
  const columnOffset: GridOffset = (index) => gridOffset(Math.min(index, MAX_COLUMNS), defaultColumnWidth, columnWidths);
  const rowOffset: GridOffset = (index) => gridOffset(Math.min(index, MAX_ROWS), defaultRowHeight, rowHeights);
  const { drawings, unsupported: drawingUnsupported, missingIds } = await parseDrawings(pkg, workbook, reference.part, columnOffset, rowOffset, skip, signal);
  throwIfAborted(signal);
  unsupported.push(...drawingUnsupported, ...damagedEntries(skipped));

  // The grid reaches every drawing: a chart placed past the data still sits on cells.
  for (const drawing of drawings) {
    const right = inchesToEmu(drawing.bounds.x + drawing.bounds.width);
    const bottom = inchesToEmu(drawing.bounds.y + drawing.bounds.height);
    extent = {
      cols: Math.max(extent.cols, trackAt(right, drawing.anchor.start.col, MAX_COLUMNS, columnOffset) + 1),
      rows: Math.max(extent.rows, trackAt(bottom, drawing.anchor.start.row, MAX_ROWS, rowOffset) + 1),
    };
  }
  const widths: number[] = [];
  for (let col = 0; col < extent.cols; col++) {
    widths.push(columnWidths.get(col) ?? defaultColumnWidth);
  }
  const hiddenColumns = [...columnHidden].filter((col) => col < extent.cols).sort((a, b) => a - b);
  const hidden = [...hiddenRows].filter((row) => row < extent.rows).sort((a, b) => a - b);

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
  if (workbook.hasMacros) {
    unsupported.push({ kind: "macro", detail: "this workbook contains macros; code never runs in previews" });
  }
  // Sparklines live in a worksheet extension (`x14:sparklineGroups`), not in cells.
  const extensionList = firstChild(root, "extLst");
  if (extensionList !== undefined && childElements(extensionList, "ext").some((ext) => firstChild(ext, "sparklineGroups") !== undefined)) {
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
    extent,
    rows,
    merges: boundedMerges,
    frozen,
    columnWidths: widths,
    hiddenColumns,
    hiddenRows: hidden,
    defaultRowHeight,
    defaultColumnWidth,
    drawings,
    unsupported,
    ...(missingIds === 0 ? {} : { missingIds }),
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
/**
 * A sheet's searchable text, for search's text read in the Worker: each
 * non-empty cell's formatted text (a string cached value when the formatted
 * text is empty; formulas are never evaluated), in row-major order, then each
 * chart's words at the cell its top-left corner is anchored to.
 */
export function sheetText(sheet: SheetModel): SheetText {
  const out: SheetText = { rows: [], cols: [], text: [], chart: [] };
  const add = (row: number, col: number, text: string, chart: string | null): void => {
    out.rows.push(row);
    out.cols.push(col);
    out.text.push(text);
    out.chart.push(chart);
  };
  for (const row of sheet.rows) {
    for (const { col, cell } of row.cells) {
      const text = cell.formatted !== "" ? cell.formatted : typeof cell.value === "string" ? cell.value : "";
      if (text !== "") add(row.index, col, text, null);
    }
  }
  for (const drawing of sheet.drawings) {
    if (drawing.kind !== "chart" || drawing.spec === undefined) continue;
    for (const text of chartText(drawing.spec)) add(drawing.anchor.start.row, drawing.anchor.start.col, text, drawing.name);
  }
  return out;
}

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
    if (to > from || cells.length === 0) rows.push({ ...row, cells: cells.slice(from, to) });
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
