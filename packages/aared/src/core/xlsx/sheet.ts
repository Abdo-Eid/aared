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
import { AaredError, damagedDocument, isAbort, throwIfAborted } from "../errors.js";
import { PX_PER_INCH, emuToInches, inchesToEmu, ptToInches } from "../units.js";
import { MAX_COLUMNS, MAX_MERGED_CELLS, MAX_MERGED_CELLS_PER_SHEET, MAX_ROWS } from "./limits.js";
import { parseCell, parseRef } from "./range.js";
import { formatNumberValue, isDateFormatCode, serialToDate } from "./numberfmt.js";
import {
  cellResultFor,
  compareScalars,
  compileFormula,
  compileCriteria,
  criteriaMatches,
  dateToSerial,
  evaluate,
  isTruthyForCF,
  makeFormulaBudget,
  referencedSheets,
  type CompiledFormula,
  type FormulaBudget,
  type FormulaCellSource,
  type RangeScan,
  type ScanResult,
} from "./formula.js";
import type { DifferentialStyle, XlsxStyles } from "./styles.js";
import { stringItemText } from "./workbook.js";
import type { XlsxWorkbook } from "./workbook.js";
import { chartFillOwner, chartexFormulas, readChartModel } from "../chart.js";
import { chartText } from "../chart/text.js";
import type { SheetText } from "../rpc.js";
import { resolveChartexData } from "./chart-data.js";
import { resolveImageSource } from "../image-source.js";
import { readDrawingEffects } from "../drawingml-effects.js";
import { drawingShapeTransform } from "../drawingml-shape.js";
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
 * (read back from Excel; Notes §72). Adding the padding again made every
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

/** Visit every parsed cell inside `range` with its position, row-major, columns ascending. */
function forEachCellWithPos(grid: CellGrid, range: CellRange, visit: (row: number, col: number, cell: CellModel) => void): void {
  const { start, end } = range;
  const { byRow, rowOrder } = grid;
  const visitRow = (row: number, entries: GridEntry[]): void => {
    for (let i = lowerBound(entries, start.col, entryCol); i < entries.length && entries[i]!.col <= end.col; i++) {
      visit(row, entries[i]!.col, entries[i]!.cell);
    }
  };
  if (end.row - start.row < rowOrder.length) {
    for (let row = start.row; row <= end.row; row++) {
      const entries = byRow.get(row);
      if (entries !== undefined) visitRow(row, entries);
    }
    return;
  }
  for (let i = lowerBound(rowOrder, start.row, identity); i < rowOrder.length && rowOrder[i]! <= end.row; i++) {
    visitRow(rowOrder[i]!, byRow.get(rowOrder[i]!)!);
  }
}

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

/** One `cfvo` bound: `value` is its number, `text` its `val` as written (a `formula` bound is a formula, not a number). */
interface ConditionalBound {
  type: string;
  value: number;
  text: string;
}

interface ConditionalDataBar {
  min: ConditionalBound;
  max: ConditionalBound;
  color: string;
  minLength: number;
  maxLength: number;
  gradient: boolean;
  showValue: boolean;
  direction: "context" | "leftToRight" | "rightToLeft";
  border: boolean;
  borderColor?: string;
  axisPosition: "automatic" | "middle" | "none";
  axisColor?: string;
  negativeFillColor?: string;
  negativeBorderColor?: string;
}

interface ConditionalRule {
  priority: number;
  /** `cfRule/@type`: `cellIs`, `colorScale`, `duplicateValues`, `expression`, ... */
  type?: string;
  /** Lower-priority rules no longer apply to a cell this rule matched. */
  stopIfTrue: boolean;
  dxfId?: number;
  /** Inline differential (x14 rules carry `<dxf>`, not a `dxfId`). */
  inline?: DifferentialStyle;
  operator?: string;
  /** `cfRule/@text`: what a `containsText`, `notContainsText`, `beginsWith` or `endsWith` rule looks for. */
  text?: string;
  formulas: string[];
  /** `top10`: `rank` is a count, or a percentage when `percent` is set; `bottom` takes the smallest. */
  rank?: number;
  percent: boolean;
  bottom: boolean;
  /** `aboveAverage`: `aboveAverage` defaults to true (ECMA-376 18.3.1.10); `equalAverage` admits cells equal to the mean. */
  aboveAverage: boolean;
  equalAverage: boolean;
  /** `cfRule/@stdDev`, present when the rule compares against a number of standard deviations. */
  stdDev?: number;
  /** `cfRule/@timePeriod`: ten values (today, yesterday, tomorrow, last7Days, thisWeek, lastWeek, nextWeek, thisMonth, lastMonth, nextMonth). */
  timePeriod?: string;
  colorScale?: { values: ConditionalBound[]; colors: string[] };
  dataBar?: ConditionalDataBar;
  iconSet?: ConditionalIconSet;
}

/** One `cfIcon` override inside a custom icon set. */
interface ConditionalCustomIcon {
  family: string;
  icon: number;
}

interface ConditionalIconSet {
  /** Family as written (`3Arrows`, `4TrafficLights`, `5Boxes`, ...; `NoIcons` means no icon). */
  family: string;
  showValue: boolean;
  reverse: boolean;
  custom: boolean;
  /** Thresholds as written; the first is ignored when assigning tiers (Excel's own default is 0 there). */
  thresholds: ConditionalBound[];
  customIcons: ConditionalCustomIcon[];
}

/** A local calendar date (year, month 1-12, day 1-31) that `timePeriod` rules are evaluated against. */
export interface XlsxToday {
  year: number;
  month: number;
  day: number;
}

/** The reader's local calendar date, captured once per parse so every sheet agrees. */
function localToday(): XlsxToday {
  const now = new Date();
  return { year: now.getFullYear(), month: now.getMonth() + 1, day: now.getDate() };
}

/** Icon families by tier count: 10 with three icons, 5 with four, 5 with five. */
const ICON_TIERS: Record<string, number> = {
  "3Arrows": 3,
  "3ArrowsGray": 3,
  "3Flags": 3,
  "3TrafficLights1": 3,
  "3TrafficLights2": 3,
  "3Signs": 3,
  "3Symbols": 3,
  "3Symbols2": 3,
  "3Stars": 3,
  "3Triangles": 3,
  "4Arrows": 4,
  "4ArrowsGray": 4,
  "4RedToBlack": 4,
  "4Rating": 4,
  "4TrafficLights": 4,
  "5Arrows": 5,
  "5ArrowsGray": 5,
  "5Rating": 5,
  "5Quarters": 5,
  "5Boxes": 5,
};

function rgbColor(element: XmlElement | undefined): string | undefined {
  const rgb = element === undefined ? undefined : attr(element, "rgb");
  if (rgb === undefined || rgb.length !== 8) return undefined;
  return `#${rgb.slice(2).toUpperCase()}`;
}

function conditionalBound(entry: XmlElement): ConditionalBound {
  const value = attr(entry, "val") ?? (firstChild(entry, "f") === undefined ? undefined : textContent(firstChild(entry, "f")!));
  return { type: attr(entry, "type") ?? "", value: value === undefined ? Number.NaN : Number(value), text: value ?? "" };
}

/** x14 stores data-bar appearance in a worksheet-level extension, linked by the base rule's x14:id. */
function conditionalDataBarExtensions(sheet: XmlElement): Map<string, XmlElement> {
  const out = new Map<string, XmlElement>();
  const extensions = firstChild(sheet, "extLst");
  if (extensions === undefined) return out;
  for (const extension of childElements(extensions, "ext")) {
    const group = firstChild(extension, "conditionalFormattings");
    const formats = group === undefined
      ? childElements(extension, "conditionalFormatting")
      : childElements(group, "conditionalFormatting");
    for (const formatting of formats) {
      for (const rule of childElements(formatting, "cfRule")) {
        if (attr(rule, "type") !== "dataBar") continue;
        const id = attr(rule, "id");
        const bar = firstChild(rule, "dataBar");
        if (id !== undefined && bar !== undefined) out.set(id.toUpperCase(), bar);
      }
    }
  }
  return out;
}

function linkedDataBarExtension(rule: XmlElement, extensions: ReadonlyMap<string, XmlElement>): XmlElement | undefined {
  const list = firstChild(rule, "extLst");
  if (list === undefined) return undefined;
  for (const extension of childElements(list, "ext")) {
    const id = firstChild(extension, "id");
    if (id === undefined) continue;
    const bar = extensions.get(textContent(id).trim().toUpperCase());
    if (bar !== undefined) return bar;
  }
  return undefined;
}

function dataBarLength(element: XmlElement, name: "minLength" | "maxLength", fallback: number): number {
  return numberAttr(element, name, "dataBar") ?? fallback;
}

function dataBarBool(element: XmlElement, name: string, fallback: boolean): boolean {
  return onOff(element, name) ?? fallback;
}

/**
 * Parse an `<iconSet>` (or `<x14:iconSet>`) element: family with Excel's
 * default (`3TrafficLights1` when omitted), value visibility, reversal and
 * custom per-tier icons. Thresholds are the `<cfvo>` children as written;
 * the first is ignored when tiers are assigned.
 */
function conditionalIconSet(element: XmlElement): ConditionalIconSet | undefined {
  const thresholds = childElements(element, "cfvo").map(conditionalBound);
  if (thresholds.length < 3 || thresholds.length > 5) return undefined;
  const family = attr(element, "iconSet") ?? "3TrafficLights1";
  if (family === "NoIcons") return undefined;
  if (ICON_TIERS[family] === undefined) return undefined;
  // A custom set names each tier's icon; without it the tier count must match the family.
  const custom = onOff(element, "custom") === true;
  const customIcons: ConditionalCustomIcon[] = [];
  if (custom) {
    for (const icon of childElements(element, "cfIcon")) {
      const iconFamily = attr(icon, "iconSet") ?? family;
      const iconId = numberAttr(icon, "iconId", "cfIcon") ?? 0;
      // A per-tier `NoIcons` intentionally paints nothing for that tier.
      if (iconFamily !== "NoIcons" && ICON_TIERS[iconFamily] === undefined) return undefined;
      if (iconFamily !== "NoIcons" && (!Number.isInteger(iconId) || iconId < 0 || iconId >= ICON_TIERS[iconFamily]!)) return undefined;
      customIcons.push({ family: iconFamily, icon: iconId });
    }
    if (customIcons.length !== thresholds.length) return undefined;
  } else if (ICON_TIERS[family] !== thresholds.length) {
    return undefined;
  }
  return {
    family,
    showValue: onOff(element, "showValue") !== false,
    reverse: onOff(element, "reverse") === true,
    custom,
    thresholds,
    customIcons,
  };
}

/** Standalone x14 rules: icon sets plus `expression` rules with `<xm:f>`/`<xm:sqref>` and inline `<dxf>` (Excel 16 stores cross-sheet formulas there, measured). */
function x14Rules(sheet: XmlElement, styles: XlsxStyles, skip: SkipDamaged): Array<{ ranges: CellRange[]; rule: ConditionalRule }> {
  const out: Array<{ ranges: CellRange[]; rule: ConditionalRule }> = [];
  const extensions = firstChild(sheet, "extLst");
  if (extensions === undefined) return out;
  for (const extension of childElements(extensions, "ext")) {
    const group = firstChild(extension, "conditionalFormattings");
    const formats = group === undefined
      ? childElements(extension, "conditionalFormatting")
      : childElements(group, "conditionalFormatting");
    for (const formatting of formats) {
      // Base x14 wraps sqref in `<xm:sqref>`; the older shape uses an attribute.
      const sqrefElement = firstChild(formatting, "sqref");
      const refs = (sqrefElement === undefined ? attr(formatting, "sqref") ?? "" : textContent(sqrefElement)).split(/\s+/).filter((part) => part !== "");
      const ranges: CellRange[] = [];
      for (const ref of refs) {
        const range = parseRef(ref);
        if (range === undefined) skip("conditional-format range");
        else ranges.push(range);
      }
      if (ranges.length === 0) continue;
      for (const element of childElements(formatting, "cfRule")) {
        const x14type = attr(element, "type");
        // Linked `dataBar` extensions carry the x14 appearance for a base rule
        // (matched by id); they are not standalone rules and must not report.
        if (x14type === "dataBar") continue;
        if (x14type === "expression") {
          // A cross-sheet (or otherwise x14-stored) expression: the formula
          // is `<xm:f>`, the ranges `<xm:sqref>`, the style inline `<dxf>`.
          // Element matching is local-name based, so the x14/xm prefixes ride along.
          const f = firstChild(element, "f");
          const formula = f === undefined ? "" : textContent(f);
          const dxfElement = firstChild(element, "dxf");
          const inline = dxfElement === undefined ? undefined : styles.inlineDxf(dxfElement);
          out.push({
            ranges,
            rule: {
              priority: numberAttr(element, "priority", "cfRule") ?? 0,
              stopIfTrue: onOff(element, "stopIfTrue") === true,
              type: "expression",
              percent: false,
              bottom: false,
              aboveAverage: true,
              equalAverage: false,
              formulas: [formula],
              ...(inline === undefined ? {} : { inline }),
            },
          });
          continue;
        }
        if (x14type !== "iconSet") {
          // Other x14 rules are not silently dropped: keep a placeholder so
          // the reader reports instead of guessing.
          if (x14type !== undefined) {
            out.push({
              ranges,
              rule: {
                priority: numberAttr(element, "priority", "cfRule") ?? 0,
                stopIfTrue: onOff(element, "stopIfTrue") === true,
                type: `x14-${x14type}`,
                percent: false,
                bottom: false,
                aboveAverage: true,
                equalAverage: false,
                formulas: [],
              },
            });
          }
          continue;
        }
        const setElement = firstChild(element, "iconSet");
        const iconSet = setElement === undefined ? undefined : conditionalIconSet(setElement);
        // An invalid set keeps its rule so `applyConditionalFormatting` reports
        // it as unsupported instead of silently dropping the rule. `NoIcons`
        // never reaches here: `conditionalIconSet` returns undefined for it,
        // and a rule that names no icon paints nothing by design.
        const rule: ConditionalRule = {
          priority: numberAttr(element, "priority", "cfRule") ?? 0,
          stopIfTrue: onOff(element, "stopIfTrue") === true,
          type: "iconSet",
          percent: false,
          bottom: false,
          aboveAverage: true,
          equalAverage: false,
          formulas: [],
          ...(iconSet === undefined ? {} : { iconSet }),
        };
        // A rule that explicitly names `NoIcons` paints nothing and is not
        // unsupported: it is the author's way of clearing an icon.
        if (iconSet === undefined && setElement !== undefined && attr(setElement, "iconSet") === "NoIcons") continue;
        out.push({ ranges, rule });
      }
    }
  }
  return out;
}

function conditionalRules(sheet: XmlElement, styles: XlsxStyles, skip: SkipDamaged): Array<{ ranges: CellRange[]; rule: ConditionalRule }> {
  const out: Array<{ ranges: CellRange[]; rule: ConditionalRule }> = [];
  const dataBarExtensions = conditionalDataBarExtensions(sheet);
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
      // CT_CfRule (ECMA-376 18.3.1.10) defaults the code relies on:
      // `aboveAverage` true, `percent`/`bottom`/`equalAverage`/`stopIfTrue`
      // false; `rank` is unsignedInt, `stdDev` int, at most three `<formula>`
      // children. `top10`'s N is `rank` (a percentage when `percent` is set,
      // MS-OI29500), falling back to `<formula>[0]` when a file omits it.
      const priority = numberAttr(element, "priority", "cfRule") ?? 0;
      const dxfId = numberAttr(element, "dxfId", "cfRule");
      const operator = attr(element, "operator");
      const type = attr(element, "type");
      const text = attr(element, "text");
      const formulas = childElements(element, "formula").map((formula) => textContent(formula));
      // `rank` is required by the schema for top10; a file that omits it carries the number in its formula.
      const rank = numberAttr(element, "rank", "cfRule") ?? (formulas[0] === undefined ? undefined : Number(formulas[0]));
      const rule: ConditionalRule = {
        priority,
        stopIfTrue: booleanAttribute(element, "stopIfTrue"),
        ...(type === undefined ? {} : { type }),
        ...(dxfId === undefined ? {} : { dxfId }),
        ...(operator === undefined ? {} : { operator }),
        ...(text === undefined ? {} : { text }),
        ...(rank === undefined || !Number.isFinite(rank) ? {} : { rank }),
        percent: booleanAttribute(element, "percent"),
        bottom: booleanAttribute(element, "bottom"),
        aboveAverage: onOff(element, "aboveAverage") !== false,
        equalAverage: booleanAttribute(element, "equalAverage"),
        ...(attr(element, "stdDev") === undefined ? {} : { stdDev: numberAttr(element, "stdDev", "cfRule") ?? Number.NaN }),
        ...(attr(element, "timePeriod") === undefined ? {} : { timePeriod: attr(element, "timePeriod")! }),
        formulas,
      };
      const scale = firstChild(element, "colorScale");
      if (scale !== undefined) {
        const values = childElements(scale, "cfvo").map(conditionalBound);
        const colors = childElements(scale, "color").map((entry) => rgbColor(entry)).filter((color): color is string => color !== undefined);
        rule.colorScale = { values, colors };
      }
      const setElement = firstChild(element, "iconSet");
      // `NoIcons` paints nothing by design and is never reported as unsupported.
      const noIcons = setElement !== undefined && attr(setElement, "iconSet") === "NoIcons";
      if (setElement !== undefined && !noIcons) {
        const iconSet = conditionalIconSet(setElement);
        if (iconSet !== undefined) rule.iconSet = iconSet;
      }
      const bar = firstChild(element, "dataBar");
      if (bar !== undefined) {
        const stops = childElements(bar, "cfvo").map(conditionalBound);
        const color = rgbColor(firstChild(bar, "color"));
        if (stops.length >= 2 && color !== undefined) {
          const extension = linkedDataBarExtension(element, dataBarExtensions);
          const extendedStops = extension === undefined ? [] : childElements(extension, "cfvo").map(conditionalBound);
          // Excel 16 resolves omitted x14 lengths to 10/90 and gradient,
          // showValue, and direction to true, true, and context (COM readback
          // on the Arabic coverage Hours rule).
          const lengthDefaults = extension === undefined ? { min: 0, max: 100 } : { min: 10, max: 90 };
          const borderColor = extension === undefined ? undefined : rgbColor(firstChild(extension, "borderColor"));
          const axisColor = extension === undefined ? undefined : rgbColor(firstChild(extension, "axisColor"));
          const negativeFillColor = extension === undefined ? undefined : rgbColor(firstChild(extension, "negativeFillColor"));
          const negativeBorderColor = extension === undefined ? undefined : rgbColor(firstChild(extension, "negativeBorderColor"));
          const direction = extension === undefined ? "context" : attr(extension, "direction") ?? "context";
          const axisPosition = extension === undefined ? "automatic" : attr(extension, "axisPosition") ?? "automatic";
          rule.dataBar = {
            min: extendedStops[0] ?? stops[0]!,
            max: extendedStops[1] ?? stops[1]!,
            color,
            minLength: dataBarLength(extension ?? bar, "minLength", dataBarLength(bar, "minLength", lengthDefaults.min)),
            maxLength: dataBarLength(extension ?? bar, "maxLength", dataBarLength(bar, "maxLength", lengthDefaults.max)),
            gradient: extension === undefined ? true : dataBarBool(extension, "gradient", true),
            showValue: extension === undefined
              ? dataBarBool(bar, "showValue", true)
              : dataBarBool(extension, "showValue", dataBarBool(bar, "showValue", true)),
            direction: direction === "leftToRight" || direction === "rightToLeft" ? direction : "context",
            border: extension === undefined ? false : dataBarBool(extension, "border", false),
            ...(borderColor === undefined ? {} : { borderColor }),
            axisPosition: axisPosition === "middle" || axisPosition === "none" ? axisPosition : "automatic",
            ...(axisColor === undefined ? {} : { axisColor }),
            ...(negativeFillColor === undefined ? {} : { negativeFillColor }),
            ...(negativeBorderColor === undefined ? {} : { negativeBorderColor }),
          };
        }
      }
      // A `NoIcons` set with no other content paints nothing by design.
      if (!(noIcons && rule.type === "iconSet" && rule.colorScale === undefined && rule.dataBar === undefined && rule.dxfId === undefined)) {
        out.push({ ranges, rule });
      }
    }
  }
  out.push(...x14Rules(sheet, styles, skip));
  // Excel evaluates lower priority numbers first; the first matching rule wins.
  out.sort((a, b) => a.rule.priority - b.rule.priority);
  return out;
}

function unquote(formula: string): string {
  return formula.length >= 2 && formula.startsWith('"') && formula.endsWith('"') ? formula.slice(1, -1) : formula;
}

/** Lookup an authored cell by position; absent columns are empty. */
function cellAt(grid: CellGrid, row: number, col: number): CellModel | undefined {
  const entries = grid.byRow.get(row);
  if (entries === undefined) return undefined;
  const at = lowerBound(entries, col, entryCol);
  return at < entries.length && entries[at]!.col === col ? entries[at]!.cell : undefined;
}

/** Materialize a blank cell for an expression/cellIs match (ISBLANK targets have no `<c>`). */
function materializeBlank(grid: CellGrid, row: number, col: number): CellModel {
  let entries = grid.byRow.get(row);
  if (entries === undefined) {
    const cell: CellModel = { value: null, formatted: "", type: "empty" };
    grid.byRow.set(row, [{ col, cell }]);
    const at = lowerBound(grid.rowOrder, row, identity);
    if (at >= grid.rowOrder.length || grid.rowOrder[at] !== row) grid.rowOrder.splice(at, 0, row);
    return cell;
  }
  const at = lowerBound(entries, col, entryCol);
  if (at < entries.length && entries[at]!.col === col) return entries[at]!.cell;
  const cell: CellModel = { value: null, formatted: "", type: "empty" };
  entries.splice(at, 0, { col, cell });
  return cell;
}

/** Total cells in ranges (for blank-materialization budget); capped to avoid overflow. */
function rangeArea(ranges: readonly CellRange[]): number {
  let total = 0;
  for (const range of ranges) {
    const rows = range.end.row - range.start.row + 1;
    const cols = range.end.col - range.start.col + 1;
    total += rows * cols;
    if (total > 50000) return total;
  }
  return total;
}

/** Numbers for aggregates: authored numbers only; errors/unknowns propagate, blanks/text/logicals ignored. */
function numbersInRanges(grid: CellGrid, ranges: readonly CellRange[]): { numbers: number[]; error?: { error: import("./formula.js").FormulaErrorCode }; unknown?: boolean } {
  const numbers: number[] = [];
  for (const range of ranges) {
    let sawError: { error: import("./formula.js").FormulaErrorCode } | undefined;
    let sawUnknown = false;
    forEachCellInRange(grid, range, (cell) => {
      if (sawError !== undefined || sawUnknown) return;
      if (cell.type === "error") {
        const code = String(cell.value);
        const known = (["#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!"] as string[]).includes(code)
          ? (code as import("./formula.js").FormulaErrorCode)
          : ("#VALUE!" as import("./formula.js").FormulaErrorCode);
        sawError = { error: known };
        return;
      }
      if (cell.value === null && cell.formula !== undefined && cell.formula !== "") {
        sawUnknown = true;
        return;
      }
      if (typeof cell.value === "number") numbers.push(cell.value);
    });
    if (sawError !== undefined) return { numbers, error: sawError };
    if (sawUnknown) return { numbers, unknown: true };
  }
  return { numbers };
}

/**
 * Everything conditional-formula evaluation needs beyond the sheet's own
 * grid: budgets, the compile cache, date snapshots, and preloaded
 * cross-sheet value grids (`evaluate` stays sync, so cross sheets load
 * before any rule runs).
 */
interface CfEngine {
  sheetLower: string;
  budget: FormulaBudget;
  cache: Map<string, CompiledFormula | undefined>;
  todaySerial: number;
  nowSerial: number;
  date1904: boolean;
  /** Other sheets' cached values by lower-cased name (values only, never their CF: no recursion). */
  crossGrids: Map<string, CellGrid>;
}

/** Authored values with positions plus the blank count for one range (blanks stay arithmetic, never materialized). */
function scanRangeIn(grid: CellGrid, range: CellRange, budget: FormulaBudget): ScanResult {
  const entries: RangeScan["entries"] = [];
  let authored = 0;
  let failed: "unsupported" | "unknown" | "budget" | undefined;
  forEachCellWithPos(grid, range, (row, col, cell) => {
    if (failed !== undefined) return;
    authored++;
    const r = cellResultFor(cell);
    if (!r.ok) {
      failed = r.reason;
      return;
    }
    entries.push({ row, col, value: r.value });
  });
  if (failed !== undefined) return { ok: false, reason: failed };
  budget.visits += authored;
  if (budget.visits > budget.maxVisits) return { ok: false, reason: "budget" };
  const rows = range.end.row - range.start.row + 1;
  const cols = range.end.col - range.start.col + 1;
  return { ok: true, scan: { entries, blanks: rows * cols - authored } };
}

/** Formula evaluation context over one sheet's cached grid, with the engine's snapshots and cross-sheet reads. */
function formulaSource(
  grid: CellGrid,
  engine: CfEngine,
  anchor: { row: number; col: number },
  target: { row: number; col: number },
): FormulaCellSource {
  const cross = (sheetLower: string): CellGrid | undefined => engine.crossGrids.get(sheetLower);
  return {
    sheetLower: engine.sheetLower,
    anchorRow: anchor.row,
    anchorCol: anchor.col,
    targetRow: target.row,
    targetCol: target.col,
    todaySerial: engine.todaySerial,
    nowSerial: engine.nowSerial,
    date1904: engine.date1904,
    budget: engine.budget,
    scalarAt: (row, col) => cellResultFor(cellAt(grid, row, col)),
    numbersIn: (range) => {
      const collected = numbersInRanges(grid, [range]);
      if (collected.error !== undefined) return { ok: true, value: collected.error };
      if (collected.unknown === true) return { ok: false, reason: "unknown" as const };
      return { ok: true, numbers: [...collected.numbers].sort((a, b) => a - b) };
    },
    scanIn: (range) => scanRangeIn(grid, range, engine.budget),
    crossScalarAt: (sheetLower, row, col) => {
      const other = cross(sheetLower);
      return other === undefined ? { ok: false, reason: "unsupported" as const } : cellResultFor(cellAt(other, row, col));
    },
    crossNumbersIn: (sheetLower, range) => {
      const other = cross(sheetLower);
      if (other === undefined) return { ok: false, reason: "unsupported" as const };
      const collected = numbersInRanges(other, [range]);
      if (collected.error !== undefined) return { ok: true, value: collected.error };
      if (collected.unknown === true) return { ok: false, reason: "unknown" as const };
      return { ok: true, numbers: [...collected.numbers].sort((a, b) => a - b) };
    },
    crossScanIn: (sheetLower, range) => {
      const other = cross(sheetLower);
      return other === undefined ? { ok: false, reason: "unsupported" as const } : scanRangeIn(other, range, engine.budget);
    },
  };
}

/** Compile once per distinct text; undefined means outside the subset (report, never guess). */
function compileCached(cache: Map<string, CompiledFormula | undefined>, text: string): CompiledFormula | undefined {
  const found = cache.get(text);
  if (found !== undefined) return found;
  // An empty slot is ambiguous (missing vs unsupported); store explicitly.
  if (cache.has(text)) return undefined;
  const compiled = compileFormula(text);
  cache.set(text, compiled);
  return compiled;
}

/** Original literal `cellIs` matching, kept for plain thresholds (no engine cost). */
function matchesCellIsLiteral(cell: CellModel, operator: string | undefined, formulas: string[]): boolean {
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
 * A `formula` bound evaluated over cached values: numbers, refs (same- and
 * cross-sheet), arithmetic, dates and the bounded aggregate subset (notably
 * `PERCENTILE`). Anything outside the subset, a missing cache, or an error
 * result is `undefined`, which the caller reports rather than guessing.
 */
function formulaBound(
  text: string,
  grid: CellGrid,
  engine: CfEngine,
  anchor: { row: number; col: number },
): number | undefined {
  if (text.trim() !== "" && Number.isFinite(Number(text))) return Number(text);
  const compiled = compileCached(engine.cache, text);
  if (compiled === undefined) return undefined;
  const ctx = formulaSource(grid, engine, anchor, anchor);
  const result = evaluate(compiled.node, ctx);
  if (!result.ok) return undefined;
  const value = result.value;
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (value === null) return 0;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed === "") return 0;
    const num = Number(trimmed);
    return Number.isFinite(num) ? num : undefined;
  }
  return undefined;
}

/**
 * The number a `cfvo` bound stands for, over the range's values in ascending
 * order. `percent` is a share of the min-to-max span; `percentile` is Excel's
 * PERCENTILE (linear between ranks). A `formula` bound resolves through the
 * bounded evaluator, so `PERCENTILE(...)` and cell expressions paint instead
 * of reporting when they are in-subset and cached.
 */
function scaleBound(
  bound: ConditionalBound,
  sorted: readonly number[],
  grid: CellGrid,
  engine: CfEngine,
  anchor: { row: number; col: number },
): number | undefined {
  if (bound.type === "num") return Number.isFinite(bound.value) ? bound.value : undefined;
  if (bound.type === "formula") return formulaBound(bound.text, grid, engine, anchor);
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

/**
 * Top/bottom N. The N-th largest (or smallest) value is the threshold, and
 * every cell tied with it is painted, so a rule can paint more than N cells.
 * A percentage counts `floor(numbers x percent / 100)` and is never below
 * one (measured in Excel: 0.5 of ten numbers paints the largest one). Only
 * numbers count, so blanks and text are neither ranked nor in the denominator.
 * A rank below one paints nothing: its meaning is not measured.
 */
function top10Test(values: readonly number[], rule: ConditionalRule): (value: number) => boolean {
  const rank = rule.rank ?? 0;
  const n = values.length;
  if (rank < 1 || n === 0) return () => false;
  const count = Math.min(rule.percent ? Math.max(1, Math.floor((n * rank) / 100)) : rank, n);
  const sorted = [...values].sort((a, b) => a - b);
  if (rule.bottom) {
    const threshold = sorted[count - 1]!;
    return (value) => value <= threshold;
  }
  const threshold = sorted[n - count]!;
  return (value) => value >= threshold;
}

/**
 * Above or below the mean of the numbers in the rule's range. Blanks and text
 * are left out of the mean (measured: text is not ranked and does not count
 * as zero), and `equalAverage` admits cells equal to it.
 *
 * With `stdDev`, the threshold moves to mean +/- k * population standard
 * deviation (measured in Excel 16: 10,9,8,8,8,7,6,5,4,3 with k=1 paints 10,9
 * above and 4,3 below, which is mean +/- pstdev, not sample stdev).
 */
function aboveAverageTest(values: readonly number[], rule: ConditionalRule): (value: number) => boolean {
  if (values.length === 0) return () => false;
  const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
  if (rule.stdDev !== undefined && Number.isFinite(rule.stdDev)) {
    const variance = values.reduce((sum, value) => sum + (value - mean) * (value - mean), 0) / values.length;
    const threshold = mean + (rule.aboveAverage ? 1 : -1) * rule.stdDev * Math.sqrt(variance);
    if (rule.aboveAverage) return rule.equalAverage ? (value) => value >= threshold : (value) => value > threshold;
    return rule.equalAverage ? (value) => value <= threshold : (value) => value < threshold;
  }
  if (rule.aboveAverage) return rule.equalAverage ? (value) => value >= mean : (value) => value > mean;
  return rule.equalAverage ? (value) => value <= mean : (value) => value < mean;
}

/** Ten `ST_TimePeriod` values Excel writes (there is no `next7Days`). */
const TIME_PERIODS = new Set([
  "today",
  "yesterday",
  "tomorrow",
  "last7Days",
  "thisWeek",
  "lastWeek",
  "nextWeek",
  "thisMonth",
  "lastMonth",
  "nextMonth",
]);

function dayNumber(part: { year: number; month: number; day: number }): number {
  return Math.floor(Date.UTC(part.year, part.month - 1, part.day) / 86_400_000);
}

/**
 * Whether a serial's calendar date falls in `period` around `today`.
 * Time of day is ignored (Excel's own formulas floor to the date), weeks run
 * Sunday to Saturday and months are calendar months (measured in Excel 16).
 */
function matchesTimePeriod(serial: number, period: string, today: XlsxToday, date1904: boolean): boolean {
  const part = serialToDate(Math.floor(serial), date1904);
  const day = dayNumber(part);
  const now = dayNumber(today);
  switch (period) {
    case "today": return day === now;
    case "yesterday": return day === now - 1;
    case "tomorrow": return day === now + 1;
    case "last7Days": return day <= now && day >= now - 6;
    case "thisMonth": return part.year === today.year && part.month === today.month;
    case "lastMonth": {
      const first = new Date(Date.UTC(today.year, today.month - 1, 1));
      first.setUTCMonth(first.getUTCMonth() - 1);
      return part.year === first.getUTCFullYear() && part.month === first.getUTCMonth() + 1;
    }
    case "nextMonth": {
      const first = new Date(Date.UTC(today.year, today.month - 1, 1));
      first.setUTCMonth(first.getUTCMonth() + 1);
      return part.year === first.getUTCFullYear() && part.month === first.getUTCMonth() + 1;
    }
    case "thisWeek":
    case "lastWeek":
    case "nextWeek": {
      const weekday = new Date(Date.UTC(today.year, today.month - 1, today.day)).getUTCDay();
      const start = now - weekday + (period === "lastWeek" ? -7 : period === "nextWeek" ? 7 : 0);
      return day >= start && day <= start + 6;
    }
    default: return false;
  }
}

/**
 * The icon tier a value takes: the count of real thresholds (the first cfvo
 * is ignored) it meets or beats (`Operator 7`, GreaterEqual, in every
 * Excel-authored rule). Returns `undefined` when a threshold needs a formula
 * engine.
 */
function iconTier(
  value: number,
  set: ConditionalIconSet,
  sorted: readonly number[],
  grid: CellGrid,
  engine: CfEngine,
  anchor: { row: number; col: number },
): number | undefined {
  const bounds: number[] = [];
  for (let i = 1; i < set.thresholds.length; i++) {
    const resolved = scaleBound(set.thresholds[i]!, sorted, grid, engine, anchor);
    if (resolved === undefined) return undefined;
    bounds.push(resolved);
  }
  let tier = 0;
  for (const bound of bounds) if (value >= bound) tier++;
  return tier;
}

/** Rule types the viewer evaluates from cached values; anything else is reported, not guessed. */
const EVALUATED_RULES = new Set([
  "cellIs",
  "containsText",
  "notContainsText",
  "beginsWith",
  "endsWith",
  "duplicateValues",
  "uniqueValues",
  "top10",
  "aboveAverage",
  "timePeriod",
  "expression",
]);

function numericValuesIn(grid: CellGrid, ranges: readonly CellRange[]): number[] {
  const values: number[] = [];
  for (const range of ranges) {
    forEachCellInRange(grid, range, (cell) => {
      if (typeof cell.value === "number") values.push(cell.value);
    });
  }
  return values;
}

/** Every formula text one rule can evaluate: conditions plus scale/bar/icon bounds (all compiled once and pre-scanned for cross-sheet reads). */
function ruleFormulaTexts(rule: ConditionalRule): string[] {
  const out = [...rule.formulas];
  if (rule.colorScale !== undefined) for (const bound of rule.colorScale.values) out.push(bound.text);
  if (rule.dataBar !== undefined) out.push(rule.dataBar.min.text, rule.dataBar.max.text);
  if (rule.iconSet !== undefined) for (const bound of rule.iconSet.thresholds) out.push(bound.text);
  return out;
}

/**
 * Values-only grid of one worksheet: cached `<v>` models with their formula
 * source (so a missing cache still reads unknown), sorted and de-duplicated
 * like the main parse. No conditional formats, drawings, or merges: cross
 * sheets answer reads without evaluating their own rules, so CF can never
 * recurse.
 */
async function readSheetValues(pkg: OpcPackage, workbook: XlsxWorkbook, part: string, signal?: AbortSignal): Promise<CellGrid> {
  const { styles, sharedStrings, date1904 } = workbook;
  const byRow = new Map<number, GridEntry[]>();
  const unsortedRows = new Set<number>();
  const root = await pkg.readXml(part, signal, {
    streamChildrenOf: "sheetData",
    onChild: (row) => {
      if (row.local !== "row") return;
      throwIfAborted(signal);
      const rowIndex = (numberAttr(row, "r", "row") ?? 0) - 1;
      if (rowIndex < 0) return;
      for (const cell of row.children) {
        if (cell.local !== "c") continue;
        const ref = attr(cell, "r");
        if (ref === undefined) continue;
        const position = parseCell(ref);
        if (position === undefined) continue;
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
        let entries = byRow.get(position.row);
        if (entries === undefined) {
          entries = [];
          byRow.set(position.row, entries);
        } else if (entries[entries.length - 1]!.col >= position.col) {
          unsortedRows.add(position.row);
        }
        entries.push({ col: position.col, cell: model });
      }
    },
  });
  if (root.local !== "worksheet") throw damagedDocument(`"${part}" root must be <worksheet>`);
  for (const row of unsortedRows) {
    const sorted = byRow.get(row)!.sort((a, b) => a.col - b.col);
    byRow.set(row, sorted.filter((entry, i) => sorted[i + 1]?.col !== entry.col));
  }
  return { byRow, rowOrder: [...byRow.keys()].sort((a, b) => a - b) };
}

/** Apply cached conditional-formatting results onto parsed cells (view-only snapshot). */
function applyConditionalFormatting(
  grid: CellGrid,
  formats: Array<{ ranges: CellRange[]; rule: ConditionalRule }>,
  styles: XlsxStyles,
  unsupported: SheetUnsupportedModel[],
  today: XlsxToday,
  date1904: boolean,
  sheetName?: string,
  cross?: { pkg: OpcPackage; workbook: XlsxWorkbook; selfIndex: number; signal?: AbortSignal },
): Promise<void> {
  const run = async (): Promise<void> => {
    const sheetLower = (sheetName ?? "").toLowerCase();
    const budget = makeFormulaBudget();
    const compiledCache = new Map<string, CompiledFormula | undefined>();
    // `evaluate` is sync, so every cross-sheet value grid loads up front.
    const crossGrids = new Map<string, CellGrid>();
    if (cross !== undefined) {
      const wanted = new Set<string>();
      for (const { rule } of formats) {
        for (const text of ruleFormulaTexts(rule)) {
          const compiled = compileCached(compiledCache, text);
          if (compiled === undefined) continue;
          for (const name of referencedSheets(compiled)) if (name !== sheetLower) wanted.add(name);
        }
      }
      for (const name of wanted) {
        throwIfAborted(cross.signal);
        const at = cross.workbook.sheets.findIndex((sheet) => sheet.name.toLowerCase() === name);
        if (at < 0 || at === cross.selfIndex) continue;
        try {
          crossGrids.set(name, await readSheetValues(cross.pkg, cross.workbook, cross.workbook.sheets[at]!.part, cross.signal));
        } catch (error) {
          // A damaged cross sheet costs only the rules that read it (they
          // report `unsupported`), never the sheet being parsed — but an
          // abort still aborts.
          if (isAbort(error)) throw error;
        }
      }
    }
    const rawToday = dateToSerial(today.year, today.month, today.day, date1904);
    const todaySerial = typeof rawToday === "string" ? Number.NaN : rawToday;
    const now = new Date();
    const nowSerial = todaySerial + (now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds() + now.getMilliseconds() / 1000) / 86400;
    const engine: CfEngine = { sheetLower, budget, cache: compiledCache, todaySerial, nowSerial, date1904, crossGrids };
    let reportedUnsupported = false;
    const reportOnce = (detail: string): void => {
      if (reportedUnsupported) return;
      reportedUnsupported = true;
      unsupported.push({ kind: "conditional-format", detail });
    };
  // Cells a `stopIfTrue` rule has matched: no later (lower-priority) rule touches them.
  const stopped = new WeakSet<CellModel>();
  for (const { ranges, rule } of formats) {
    const anchor = ranges[0]!.start;
    if (rule.colorScale !== undefined) {
      const { values, colors } = rule.colorScale;
      if (values.length < 2 || values.length !== colors.length) {
        reportOnce("a color-scale rule is malformed; it is shown without its gradient");
        continue;
      }
      const sorted = numericValuesIn(grid, ranges).sort((a, b) => a - b);
      const stops = values.map((bound) => scaleBound(bound, sorted, grid, engine, anchor));
      if (stops.some((stop) => stop === undefined)) {
        reportOnce("a color-scale rule uses an unsupported formula bound; it is shown without its gradient");
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
      if (sorted[0]! < 0 || bar.axisPosition === "middle") {
        reportOnce("a data-bar rule with negative values or a centered axis is not drawn");
        continue;
      }
      const low = scaleBound(bar.min, sorted, grid, engine, anchor);
      const high = scaleBound(bar.max, sorted, grid, engine, anchor);
      if (low === undefined || high === undefined) {
        reportOnce("a data-bar rule uses unsupported bounds; its bars are not drawn");
        continue;
      }
      const span = high - low === 0 ? 1 : high - low;
      for (const range of ranges) {
        forEachCellInRange(grid, range, (cell) => {
          if (typeof cell.value !== "number" || stopped.has(cell) || cell.conditional?.dataBar !== undefined) return;
          const fraction = Math.min(Math.max((cell.value - low) / span, 0), 1);
          const ratio = (bar.minLength + fraction * (bar.maxLength - bar.minLength)) / 100;
          cell.conditional = {
            ...(cell.conditional ?? {}),
            dataBar: {
              ratio,
              color: bar.color,
              gradient: bar.gradient,
              showValue: bar.showValue,
              direction: bar.direction,
              border: bar.border,
              ...(bar.borderColor === undefined ? {} : { borderColor: bar.borderColor }),
              axisPosition: bar.axisPosition,
              ...(bar.axisColor === undefined ? {} : { axisColor: bar.axisColor }),
              ...(bar.negativeFillColor === undefined ? {} : { negativeFillColor: bar.negativeFillColor }),
              ...(bar.negativeBorderColor === undefined ? {} : { negativeBorderColor: bar.negativeBorderColor }),
            },
          };
          if (rule.stopIfTrue) stopped.add(cell);
        });
      }
      continue;
    }
    if (rule.iconSet !== undefined) {
      const set = rule.iconSet;
      // A malformed custom set leaves the rule unpainted rather than guessed.
      if (set.custom && set.customIcons.length !== set.thresholds.length) {
        reportOnce("an icon-set rule with a malformed custom set is shown without its icons");
        continue;
      }
      const sorted = numericValuesIn(grid, ranges).sort((a, b) => a - b);
      const bounds: Array<number | undefined> = [];
      for (let i = 1; i < set.thresholds.length; i++) bounds.push(scaleBound(set.thresholds[i]!, sorted, grid, engine, anchor));
      if (bounds.some((bound) => bound === undefined)) {
        reportOnce("an icon-set rule uses an unsupported formula bound; its icons are not drawn");
        continue;
      }
      const count = set.thresholds.length;
      for (const range of ranges) {
        forEachCellInRange(grid, range, (cell) => {
          if (typeof cell.value !== "number" || stopped.has(cell) || cell.conditional?.icon !== undefined) return;
          const tier = iconTier(cell.value, set, sorted, grid, engine, anchor);
          if (tier === undefined) return;
          if (set.custom) {
            const picked = set.customIcons[tier] ?? set.customIcons[set.customIcons.length - 1]!;
            // A per-tier NoIcons intentionally paints nothing for that tier.
            if (picked.family === "NoIcons") {
              if (rule.stopIfTrue) stopped.add(cell);
              return;
            }
            cell.conditional = {
              ...(cell.conditional ?? {}),
              icon: { family: picked.family, icon: picked.icon, showValue: set.showValue },
            };
          } else {
            const icon = set.reverse ? count - 1 - tier : tier;
            cell.conditional = {
              ...(cell.conditional ?? {}),
              icon: { family: set.family, icon, showValue: set.showValue },
            };
          }
          if (rule.stopIfTrue) stopped.add(cell);
        });
      }
      continue;
    }
    if (rule.type !== undefined && !EVALUATED_RULES.has(rule.type)) {
      reportOnce(`a "${rule.type}" conditional rule is not evaluated; its highlighting is not shown`);
      continue;
    }
    if (rule.type === "timePeriod" && (rule.timePeriod === undefined || !TIME_PERIODS.has(rule.timePeriod))) {
      reportOnce("a date-based conditional rule with an unknown period is shown without highlighting");
      continue;
    }
    // Base rules name their style by `dxfId`; x14 rules carry it inline.
    const differential = rule.inline ?? (rule.dxfId !== undefined ? styles.differential(rule.dxfId) : undefined);
    if (differential === undefined) {
      reportOnce("a conditional rule without a differential style is shown without highlighting");
      continue;
    }
    if (differential.fill === undefined && differential.font === undefined) continue;
    // Top/bottom and above-average compare each cell with the numbers the rule covers.
    const ranked =
      rule.type === "top10" ? top10Test(numericValuesIn(grid, ranges), rule)
        : rule.type === "aboveAverage" ? aboveAverageTest(numericValuesIn(grid, ranges), rule)
          : undefined;
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
    const isExpression = rule.type === "expression";
    const isCellIs = rule.type === "cellIs";
    // Expression and cellIs operands shift per target cell from the first range start (measured).
    let expressionCompiled: CompiledFormula | undefined;
    let expressionFailed: "unsupported" | "unknown" | "budget" | undefined;
    const cellIsCompiled: Array<CompiledFormula | undefined> = [];
    let cellIsFailed: "unsupported" | "unknown" | "budget" | undefined;
    if (isExpression) {
      const text = rule.formulas[0] ?? "";
      if (text.trim() === "") {
        reportOnce("a conditional expression is empty; it is shown without highlighting");
        continue;
      }
      expressionCompiled = compileCached(compiledCache, text);
      if (expressionCompiled === undefined) {
        reportOnce("an expression conditional rule is outside the supported subset; its highlighting is not shown");
        continue;
      }
    }
    if (isCellIs) {
      // Fast literal path first (plain numbers/text); formulas fall through to the engine per cell.
      // Compilation is still once per distinct operand.
      for (const text of rule.formulas) {
        if (text.trim() !== "" && Number.isFinite(Number(text))) {
          cellIsCompiled.push(undefined);
          continue;
        }
        // Plain quoted text without operators stays literal; anything richer compiles.
        if (/^[A-Za-z0-9 ._-]+$/.test(unquote(text)) && !/[A-Za-z]+\(/.test(text) && !/[!:]/.test(text) && !/[+\-*/^&<>=]/.test(text)) {
          cellIsCompiled.push(undefined);
          continue;
        }
        const compiled = compileCached(compiledCache, text);
        if (compiled === undefined) {
          cellIsFailed = "unsupported";
          break;
        }
        cellIsCompiled.push(compiled);
      }
      if (cellIsFailed !== undefined) {
        reportOnce("a cell-value conditional rule is outside the supported subset; its highlighting is not shown");
        continue;
      }
    }
    const matchesAt = (row: number, col: number, cell: CellModel): boolean | "fail" => {
      if (ranked !== undefined) return typeof cell.value === "number" && ranked(cell.value);
      if (rule.type === "timePeriod") {
        if (typeof cell.value !== "number" || rule.timePeriod === undefined) return false;
        const code = cell.style?.numberFormat;
        if (code === undefined || !isDateFormatCode(code)) return false;
        return matchesTimePeriod(cell.value, rule.timePeriod, today, date1904);
      }
      if (textRule) return matchesText(cell, rule.type!, rule.text);
      if (isExpression) {
        const ctx = formulaSource(grid, engine, anchor, { row, col });
        const result = evaluate(expressionCompiled!.node, ctx);
        if (!result.ok) {
          expressionFailed = result.reason;
          return "fail";
        }
        const truth = isTruthyForCF(result.value);
        if (typeof truth === "string") return false;
        return truth;
      }
      if (isCellIs) {
        // Literal operands use the original path; compiled operands evaluate per cell.
        let useEngine = false;
        for (const c of cellIsCompiled) if (c !== undefined) useEngine = true;
        if (!useEngine) return matchesCellIsLiteral(cell, rule.operator, rule.formulas);
        const ctx = formulaSource(grid, engine, anchor, { row, col });
        const left = cellResultFor(cell);
        if (!left.ok) {
          cellIsFailed = left.reason;
          return "fail";
        }
        const leftScalar = left.value;
        const evalOperand = (index: number): { scalar: import("./formula.js").FormulaScalar } | { fail: "unsupported" | "unknown" | "budget" } | { error: true } => {
          const compiled = cellIsCompiled[index];
          if (compiled === undefined) {
            const text = rule.formulas[index] ?? "";
            const t = unquote(text).trim();
            if (t === "") return { scalar: "" as unknown as import("./formula.js").FormulaScalar };
            const n = Number(t);
            if (Number.isFinite(n) && !/[A-Za-z!]/.test(t)) return { scalar: n as unknown as import("./formula.js").FormulaScalar };
            return { scalar: unquote(rule.formulas[index] ?? "") as unknown as import("./formula.js").FormulaScalar };
          }
          const r = evaluate(compiled.node, ctx);
          if (!r.ok) return { fail: r.reason };
          return { scalar: r.value };
        };
        const first = evalOperand(0);
        if ("fail" in first) {
          cellIsFailed = first.fail;
          return "fail";
        }
        if ("error" in first) return false;
        const op = rule.operator ?? "";
        const opMap: Record<string, string> = {
          equal: "=",
          notEqual: "<>",
          greaterThan: ">",
          greaterThanOrEqual: ">=",
          lessThan: "<",
          lessThanOrEqual: "<=",
        };
        if (op === "between" || op === "notBetween") {
          const second = evalOperand(1);
          if ("fail" in second) {
            cellIsFailed = second.fail;
            return "fail";
          }
          if (isFormulaError(first.scalar) || isFormulaError((second as { scalar: import("./formula.js").FormulaScalar }).scalar)) return false;
          const a = first.scalar;
          const b = (second as { scalar: import("./formula.js").FormulaScalar }).scalar;
          if (isFormulaError(leftScalar)) return false;
          const lo = compareScalars(leftScalar, a, ">=");
          const hi = compareScalars(leftScalar, b, "<=");
          if (typeof lo === "string" || typeof hi === "string") return false;
          const inside = lo && hi;
          return op === "between" ? inside : !inside;
        }
        const mapped = opMap[op];
        if (mapped === undefined) return false;
        if (isFormulaError(first.scalar) || isFormulaError(leftScalar)) return false;
        const cmp = compareScalars(leftScalar, first.scalar, mapped);
        return typeof cmp === "string" ? false : cmp;
      }
      if (occurrences === undefined) return matchesCellIsLiteral(cell, rule.operator, rule.formulas);
      const key = duplicateKey(cell);
      if (key === undefined) return false;
      return rule.type === "duplicateValues" ? occurrences.get(key)! > 1 : occurrences.get(key) === 1;
    };
    // Helpers for error checks without importing internals twice.
    const isFormulaError = (v: unknown): boolean => typeof v === "object" && v !== null && "error" in (v as Record<string, unknown>);
    const blankArea = isExpression || isCellIs ? rangeArea(ranges) : 0;
    const paintBlankMatches = blankArea > 0 && blankArea <= 50000;
    for (const range of ranges) {
      let failed: "unsupported" | "unknown" | "budget" | undefined;
      forEachCellWithPos(grid, range, (row, col, cell) => {
        if (failed !== undefined || stopped.has(cell) || cell.conditional?.fill !== undefined) return;
        const hit = matchesAt(row, col, cell);
        if (hit === "fail") {
          failed = expressionFailed ?? cellIsFailed ?? "unsupported";
          return;
        }
        if (!hit) return;
        cell.conditional = {
          ...(cell.conditional ?? {}),
          ...(differential.fill === undefined ? {} : { fill: differential.fill }),
          ...(differential.font?.color === undefined ? {} : { color: differential.font.color }),
          ...(differential.font?.bold === undefined ? {} : { bold: differential.font.bold }),
        };
        if (rule.stopIfTrue) stopped.add(cell);
      });
      if (failed !== undefined) {
        if (failed === "budget") reportOnce("a conditional rule exceeds the evaluation budget; its highlighting is not shown");
        else if (failed === "unknown") reportOnce("a conditional rule needs a missing cached value; its highlighting is not shown");
        else reportOnce(isExpression
          ? "an expression conditional rule is outside the supported subset; its highlighting is not shown"
          : "a cell-value conditional rule is outside the supported subset; its highlighting is not shown");
        break;
      }
      // Blank recipients (no `<c>`): e.g. ISBLANK targets. Bounded to 50k cells per rule;
      // larger blank-sensitive ranges stay authored-only rather than materializing a sheet.
      if (paintBlankMatches && (isExpression || isCellIs)) {
        const blankCell: CellModel = { value: null, formatted: "", type: "empty" };
        for (let row = range.start.row; row <= range.end.row; row++) {
          if (failed !== undefined) break;
          for (let col = range.start.col; col <= range.end.col; col++) {
            if (cellAt(grid, row, col) !== undefined) continue;
            const hit = matchesAt(row, col, blankCell);
            if (hit === "fail") {
              failed = expressionFailed ?? cellIsFailed ?? "unsupported";
              break;
            }
            if (!hit) continue;
            const made = materializeBlank(grid, row, col);
            made.conditional = {
              ...(made.conditional ?? {}),
              ...(differential.fill === undefined ? {} : { fill: differential.fill }),
              ...(differential.font?.color === undefined ? {} : { color: differential.font.color }),
              ...(differential.font?.bold === undefined ? {} : { bold: differential.font.bold }),
            };
            if (rule.stopIfTrue) stopped.add(made);
          }
        }
        if (failed !== undefined) {
          if (failed === "budget") reportOnce("a conditional rule exceeds the evaluation budget; its highlighting is not shown");
          else if (failed === "unknown") reportOnce("a conditional rule needs a missing cached value; its highlighting is not shown");
          else reportOnce(isExpression
            ? "an expression conditional rule is outside the supported subset; its highlighting is not shown"
            : "a cell-value conditional rule is outside the supported subset; its highlighting is not shown");
          break;
        }
      }
    }
  }
  };
  return run();
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
      // An embedded SVG (`asvg:svgBlip`) paints first; the main `blip` stays
      // as its raster fallback.
      const source = resolveImageSource(blipFill === undefined ? undefined : firstChild(blipFill, "blip"), {
        relationships: drawingRels,
        mediaContentType: (target) => pkg.mediaContentType(target),
        hasPart: (target) => pkg.hasPart(target),
      });
      const nameElement = firstChild(firstChild(picture, "nvPicPr") ?? picture, "cNvPr");
      const name = nameElement === undefined ? "Picture" : (attr(nameElement, "name") ?? "Picture");
      if (source === undefined || !pkg.hasPart(source.mediaId)) {
        unsupported.push({ kind: "damaged", detail: `an image ("${name}") points at a missing part and is not previewed; the file may be damaged`, range });
        continue;
      }
      // Excel has no colour map: picture DrawingML uses the default one, as charts do.
      chartTheme ??= themeForColorMap(workbook.theme, undefined);
      const spPr = firstChild(picture, "spPr");
      const effects = readDrawingEffects(spPr, chartTheme);
      const xfrm = spPr === undefined ? undefined : firstChild(spPr, "xfrm");
      const transform = drawingShapeTransform(xfrm);
      const pictureModel: SheetPictureModel = {
        kind: "picture",
        name,
        mediaId: source.mediaId,
        contentType: source.contentType,
        ...(source.fallbackMediaId === undefined || source.fallbackContentType === undefined ? {} : {
          fallbackMediaId: source.fallbackMediaId,
          fallbackContentType: source.fallbackContentType,
        }),
        ...(effects === undefined ? {} : { effects }),
        ...(transform === undefined ? {} : { transform }),
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
  today?: XlsxToday,
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

  const formats = conditionalRules(root, styles, skip);
  if (formats.length > 0) {
    const selfIndex = workbook.sheets.findIndex((sheet) => sheet.part === reference.part);
    await applyConditionalFormatting(grid, formats, styles, unsupported, today ?? localToday(), date1904, reference.name, {
      pkg,
      workbook,
      selfIndex: selfIndex < 0 ? index : selfIndex,
      ...(signal === undefined ? {} : { signal }),
    });
  }

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
