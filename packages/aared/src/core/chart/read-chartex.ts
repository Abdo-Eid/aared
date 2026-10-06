/**
 * `cx:chartSpace` (the Office 2016 chart types) -> `ChartSpecModel`, in the
 * Worker, next to `read.ts`. Every type lowers onto the same model the page
 * lays out: a histogram is a bar group of bins, a Pareto a bar group plus a
 * cumulative-percentage line group on a second axis, a waterfall a bar group
 * of floating spans; funnel, box and whisker, treemap and sunburst have group
 * kinds of their own (spec.ts). Everything derived from the data (bins,
 * quartiles, running totals, the hierarchy) is computed here, so the page
 * receives summaries, not a hundred thousand raw values.
 *
 * Where the data lives. A chartex part embedded in PPTX or DOCX carries its
 * cache (`cx:lvl`/`cx:pt`). Excel writes an XLSX part with NO cache: each
 * `cx:f` names a hidden `_xlchart.v1.N` defined name and the values stay in
 * the cells. The XLSX route resolves those formulas before reading
 * (`xlsx/chart-data.ts`) and passes the grids in as `ChartexData`.
 *
 * Rules below marked "observed" were read off Excel's own render of
 * `tools/chart-oracle/chartex-oracle.xlsx`.
 */

import { attr, childElements, firstChild, textContent } from "../xml.js";
import type { XmlElement } from "../xml.js";
import { mixColors } from "../color.js";
import { themeFontFamily } from "../theme.js";
import type { DrawingTheme } from "../theme.js";
import type { FillOwner } from "../fill.js";
import { DISPLAY_UNITS, MAX_CATEGORY_LEVELS, MAX_SERIES_POINTS, POSITIONS, claimPoints, paletteColor, plainText, readLook, readText, releasePoints, themeColor, tooManyPoints, tx1 } from "./read.js";
import type { Ctx } from "./read.js";
import type { Axis, BoxSummary, ChartSpecModel, DataLabels, Group, LabelPosition, Legend, Line, Look, Series, TextLook, Title, TreeNode } from "./spec.js";

/** A resolved cell range, rows of cells: what an XLSX `cx:f` formula points at. */
export type ChartexGrid = Array<Array<string | number | null>>;
/** Resolved ranges by formula text. */
export type ChartexData = Record<string, ChartexGrid>;

/** True for a `cx:chartSpace`: it holds `cx:chartData`, which `c:chartSpace` never does. */
export function isChartEx(space: XmlElement): boolean {
  return firstChild(space, "chartData") !== undefined;
}

/** The formulas of every dimension that has no cached values: what an XLSX read must resolve first. */
export function chartexFormulas(space: XmlElement): string[] {
  const out: string[] = [];
  const data = firstChild(space, "chartData");
  if (data === undefined) return out;
  for (const set of childElements(data, "data")) {
    for (const dim of set.children) {
      if (dim.local !== "strDim" && dim.local !== "numDim") continue;
      if (firstChild(dim, "lvl") !== undefined) continue;
      const f = firstChild(dim, "f");
      if (f === undefined) continue;
      const text = textContent(f).trim();
      if (text !== "" && !out.includes(text)) out.push(text);
    }
  }
  return out;
}

// ---------------------------------------------------------------- data dimensions

/** One dimension: its levels, outermost first, each `count` long (`undefined` where a cell is empty). */
interface Dim {
  levels: Array<Array<string | undefined>>;
  count: number;
  formatCode?: string;
}

function distinct(level: ReadonlyArray<string | undefined>): number {
  const seen: string[] = [];
  for (const v of level) if (v !== undefined && v !== "" && !seen.includes(v)) seen.push(v);
  return seen.length;
}

/** An outer level names a group once and leaves the rows below it empty: carry the name down. */
function carryDown(level: Array<string | undefined>, count: number): Array<string | undefined> {
  const out: Array<string | undefined> = new Array(count);
  let last: string | undefined;
  for (let i = 0; i < count; i++) {
    const v = level[i];
    if (v !== undefined && v !== "") last = v;
    out[i] = last;
  }
  return out;
}

function readDim(dim: XmlElement, data: ChartexData | undefined): Dim | undefined {
  const lvls = childElements(dim, "lvl");
  if (lvls.length > 0) {
    if (lvls.length > MAX_CATEGORY_LEVELS) lvls.length = MAX_CATEGORY_LEVELS;
    let formatCode: string | undefined;
    const levels = lvls.map((lvl) => {
      const points: Array<string | undefined> = [];
      for (const pt of lvl.children) {
        if (pt.local !== "pt") continue;
        const idx = Number(attr(pt, "idx") ?? points.length);
        if (!Number.isInteger(idx) || idx < 0) continue;
        if (idx >= MAX_SERIES_POINTS) throw tooManyPoints();
        points[idx] = textContent(pt);
      }
      const claimed = Number(attr(lvl, "ptCount") ?? 0);
      if (claimed > MAX_SERIES_POINTS) throw tooManyPoints();
      formatCode ??= attr(lvl, "formatCode");
      return { points, claimed: Number.isFinite(claimed) ? claimed : 0 };
    });
    const count = Math.max(...levels.map((l) => Math.max(l.claimed, l.points.length)));
    const all = levels.map((l) => l.points);
    // The cache lists the innermost level first (an assumption: no PowerPoint output to check it against). The
    // innermost level has the most distinct names, so a file that lists the other way round is read correctly too.
    const leafFirst = all.length < 2 || distinct(all[0]!) >= distinct(all[all.length - 1]!);
    const ordered = leafFirst ? [...all].reverse() : all;
    const out: Dim = { levels: ordered.map((l, i) => (i < ordered.length - 1 ? carryDown(l, count) : l)), count };
    if (formatCode !== undefined) out.formatCode = formatCode;
    return out;
  }
  const f = firstChild(dim, "f");
  const grid = f === undefined || data === undefined ? undefined : data[textContent(f).trim()];
  if (grid === undefined) return undefined;
  const byRow = attr(f!, "dir") === "row";
  const rows = grid.length;
  const cols = grid.reduce((m, r) => Math.max(m, r.length), 0);
  const cell = (v: string | number | null | undefined): string | undefined => (v === null || v === undefined ? undefined : String(v));
  // A column range reads left to right, outermost level first.
  const levelCount = Math.min(byRow ? rows : cols, MAX_CATEGORY_LEVELS);
  const count = byRow ? cols : rows;
  if (count > MAX_SERIES_POINTS) throw tooManyPoints();
  const levels: Array<Array<string | undefined>> = [];
  for (let l = 0; l < levelCount; l++) {
    const level: Array<string | undefined> = new Array(count);
    for (let i = 0; i < count; i++) level[i] = cell(byRow ? grid[l]?.[i] : grid[i]?.[l]);
    levels.push(l < levelCount - 1 ? carryDown(level, count) : level);
  }
  return { levels, count };
}

interface DataSet {
  str?: Dim;
  num?: Dim;
  /** Points claimed for it (see `readDataSet`). */
  claimed: number;
}

function readDataSet(space: XmlElement, id: string | undefined, data: ChartexData | undefined, ctx: Ctx): DataSet {
  const chartData = firstChild(space, "chartData");
  const sets = chartData === undefined ? [] : childElements(chartData, "data");
  const set = sets.find((s) => id !== undefined && attr(s, "id") === id) ?? sets[0];
  const out: DataSet = { claimed: 0 };
  if (set === undefined) return out;
  for (const dim of set.children) {
    if (dim.local === "strDim" && out.str === undefined && (attr(dim, "type") ?? "cat") === "cat") {
      const read = readDim(dim, data);
      if (read !== undefined) out.str = read;
    } else if (dim.local === "numDim" && out.num === undefined) {
      const type = attr(dim, "type") ?? "val";
      if (type === "val" || type === "size") {
        const read = readDim(dim, data);
        if (read !== undefined) out.num = read;
      }
    }
  }
  // What a series made of this data set holds: its names and its values (the longer one). Each series
  // that reads the set builds its own arrays from it, so each claims it (a binned series gives the raw values back).
  out.claimed = Math.max(out.str?.count ?? 0, out.num?.count ?? 0);
  claimPoints(ctx, out.claimed);
  return out;
}

function numbersOf(dim: Dim | undefined): Array<number | null> {
  if (dim === undefined) return [];
  const level = dim.levels[dim.levels.length - 1] ?? [];
  const out: Array<number | null> = new Array(dim.count);
  for (let i = 0; i < dim.count; i++) {
    const text = level[i];
    const n = text === undefined || text.trim() === "" ? NaN : Number(text);
    out[i] = Number.isFinite(n) ? n : null;
  }
  return out;
}

/** The innermost level's names: one per row. */
function namesOf(dim: Dim | undefined, count: number): string[] {
  const level = dim === undefined ? [] : dim.levels[dim.levels.length - 1] ?? [];
  const out: string[] = new Array(count);
  for (let i = 0; i < count; i++) out[i] = level[i] ?? "";
  return out;
}

// ---------------------------------------------------------------- small readers

const clean = (v: number): number => (v === 0 ? 0 : Number(v.toPrecision(12)));

function numAttr(el: XmlElement | undefined, name: string): number | undefined {
  const raw = el === undefined ? undefined : attr(el, name);
  if (raw === undefined) return undefined;
  const n = Number(raw);
  return Number.isFinite(n) ? n : undefined;
}

function flag(el: XmlElement | undefined, name: string, fallback: boolean): boolean {
  const raw = el === undefined ? undefined : attr(el, name);
  return raw === undefined ? fallback : raw === "1" || raw === "true";
}

function readLabels(el: XmlElement | undefined, ctx: Ctx, text: TextLook): DataLabels | undefined {
  if (el === undefined) return undefined;
  const vis = firstChild(el, "visibility");
  const out: DataLabels = {
    showValue: flag(vis, "value", vis === undefined),
    showCategory: flag(vis, "categoryName", false),
    showSeries: flag(vis, "seriesName", false),
    showPercent: false,
    showBubbleSize: false,
    showLegendKey: false,
    text: readText(firstChild(el, "txPr"), text, ctx),
    look: readLook(firstChild(el, "spPr"), ctx, {}),
    deleted: false,
    fromCells: false,
  };
  const pos = attr(el, "pos");
  if (pos !== undefined && POSITIONS.has(pos)) out.position = pos as LabelPosition;
  const fmt = firstChild(el, "numFmt");
  if (fmt !== undefined && attr(fmt, "sourceLinked") !== "1" && attr(fmt, "sourceLinked") !== "true") {
    const code = attr(fmt, "formatCode");
    if (code !== undefined) out.formatCode = code;
  }
  const sep = firstChild(el, "separator");
  if (sep !== undefined) out.separator = textContent(sep);
  return out;
}

function readTitle(el: XmlElement | undefined, ctx: Ctx, base: TextLook, auto: string): Title | undefined {
  if (el === undefined) return undefined;
  const tx = firstChild(el, "tx");
  const rich = tx === undefined ? undefined : firstChild(tx, "rich");
  const data = tx === undefined ? undefined : firstChild(tx, "txData");
  let lines = [auto];
  let text = readText(firstChild(el, "txPr"), base, ctx);
  if (rich !== undefined) {
    lines = plainText(rich);
    text = readText(rich, base, ctx, lines.join(""));
  } else if (data !== undefined) {
    const v = firstChild(data, "v");
    if (v !== undefined) lines = [textContent(v)];
  }
  if (lines.every((line) => line.trim() === "")) return undefined;
  return { lines, text, look: readLook(firstChild(el, "spPr"), ctx, {}), overlay: flag(el, "overlay", false) };
}

function seriesName(ser: XmlElement, index: number): string {
  const tx = firstChild(ser, "tx");
  const data = tx === undefined ? undefined : firstChild(tx, "txData");
  const v = data === undefined ? undefined : firstChild(data, "v");
  return v === undefined ? `Series${index + 1}` : textContent(v);
}

// ---------------------------------------------------------------- axes

function readAxes(plotArea: XmlElement, ctx: Ctx): Axis[] {
  const axes: Axis[] = [];
  const titleBase: TextLook = ctx.text;
  const grid = (el: XmlElement | undefined, base: Line): Line | undefined => {
    if (el === undefined) return undefined;
    return readLook(firstChild(el, "spPr"), ctx, { line: base }).line;
  };
  const gridLine: Line = { color: tx1(ctx, 0.15, 0.85), width: 0.75 };
  for (const el of childElements(plotArea, "axis")) {
    const id = numAttr(el, "id");
    if (id === undefined) continue;
    const cat = firstChild(el, "catScaling");
    const valScale = firstChild(el, "valScaling");
    const numFmt = firstChild(el, "numFmt");
    const units = firstChild(el, "units");
    const unit = units === undefined ? undefined : attr(units, "unit");
    const axis: Axis = {
      id,
      kind: cat !== undefined ? "cat" : "val",
      position: cat !== undefined ? "b" : "l",
      deleted: flag(el, "hidden", false),
      crossAxisId: 0,
      crosses: "autoZero",
      reversed: false,
      majorTick: "none",
      minorTick: "none",
      labelPosition: firstChild(el, "tickLabels") === undefined ? "none" : "nextTo",
      text: readText(firstChild(el, "txPr"), ctx.text, ctx),
      formatCode: numFmt === undefined ? "General" : attr(numFmt, "formatCode") ?? "General",
      formatLinked: numFmt === undefined || attr(numFmt, "sourceLinked") === "1" || attr(numFmt, "sourceLinked") === "true",
      between: true,
      labelOffset: 100,
      displayUnit: 1,
      auto: false,
      multiLevel: false,
    };
    if (unit === "percentage") {
      axis.formatCode = "0%";
      axis.formatLinked = false;
    } else if (unit !== undefined && DISPLAY_UNITS[unit] !== undefined) {
      axis.displayUnit = DISPLAY_UNITS[unit]!;
    }
    // Observed: a category axis has a pale line, a value axis none.
    const line = readLook(firstChild(el, "spPr"), ctx, cat !== undefined ? { line: gridLine } : {}).line;
    if (line !== undefined) axis.line = line;
    for (const key of ["min", "max"] as const) {
      const n = numAttr(valScale, key);
      if (n !== undefined) axis[key] = n;
    }
    const major = numAttr(valScale, "majorUnit");
    if (major !== undefined && major > 0) axis.majorUnit = major;
    const minor = numAttr(valScale, "minorUnit");
    if (minor !== undefined && minor > 0) axis.minorUnit = minor;
    const majorGrid = grid(firstChild(el, "majorGridlines"), gridLine);
    if (majorGrid !== undefined) axis.majorGrid = majorGrid;
    const minorGrid = grid(firstChild(el, "minorGridlines"), { color: tx1(ctx, 1, 0.75), width: 0.75 });
    if (minorGrid !== undefined) axis.minorGrid = minorGrid;
    const title = readTitle(firstChild(el, "title"), ctx, titleBase, "Axis Title");
    if (title !== undefined) axis.title = title;
    axes.push(axis);
  }
  return axes;
}

// ---------------------------------------------------------------- group builders

function emptyGroup(kind: Group["kind"], series: Series[], axisIds: Group["axisIds"]): Group {
  return {
    kind, direction: "col", grouping: kind === "bar" ? "clustered" : "standard", varyColors: false, gapWidth: 150, overlap: 0,
    holeSize: 0, firstSliceAngle: 0, bubbleScale: 100, sizeRepresents: "area", scatterStyle: "marker", radarStyle: "standard",
    series, axisIds,
  };
}

function emptySeries(index: number, name: string, categories: string[], values: Array<number | null>, look: Look): Series {
  return {
    index, order: index, name, categories, values, formatCode: "General", look,
    marker: { symbol: "none", size: 5, look: {} }, smooth: false, explosion: 0, invertIfNegative: false, points: {},
  };
}

/** Per-point fills the file overrides (`cx:dataPt`). */
function applyDataPoints(series: Series, ser: XmlElement, ctx: Ctx, base: (i: number) => Look): void {
  for (const dPt of childElements(ser, "dataPt")) {
    const idx = numAttr(dPt, "idx");
    if (idx === undefined || idx < 0) continue;
    series.points[idx] = { look: readLook(firstChild(dPt, "spPr"), ctx, base(idx)) };
  }
}

/** Gap width as the file writes it (a ratio) -> the percent `Group.gapWidth` holds. */
function gapOf(axis: Axis | undefined, plotArea: XmlElement, fallback: number): number {
  const el = childElements(plotArea, "axis").find((a) => numAttr(a, "id") === axis?.id);
  const cat = el === undefined ? undefined : firstChild(el, "catScaling");
  const n = numAttr(cat, "gapWidth");
  return Math.round((n === undefined ? fallback : n) * 100);
}

function pickAxes(axes: Axis[], ser: XmlElement | undefined): { cat: Axis | undefined; val: Axis | undefined } {
  const named = ser === undefined ? undefined : numAttr(firstChild(ser, "axisId"), "val");
  const own = named === undefined ? undefined : axes.find((a) => a.id === named);
  const cat = axes.find((a) => a.kind === "cat");
  const val = own !== undefined && own.kind === "val" ? own : axes.find((a) => a.kind === "val");
  return { cat, val };
}

/** A category/value axis pair for a type whose file omits one (funnel has no value axis). */
function ensureAxes(axes: Axis[], ctx: Ctx, kinds: { cat: boolean; val: boolean }): void {
  const make = (id: number, kind: "cat" | "val"): Axis => ({
    id, kind, position: kind === "cat" ? "b" : "l", deleted: true, crossAxisId: 0, crosses: "autoZero", reversed: false,
    majorTick: "none", minorTick: "none", labelPosition: "none", text: ctx.text, formatCode: "General", formatLinked: true,
    between: true, labelOffset: 100, displayUnit: 1, auto: false, multiLevel: false,
  });
  const nextId = (): number => Math.min(-1, ...axes.map((a) => a.id)) - 1;
  if (kinds.cat && !axes.some((a) => a.kind === "cat")) axes.push(make(nextId(), "cat"));
  if (kinds.val && !axes.some((a) => a.kind === "val")) axes.push(make(nextId(), "val"));
}

/** Point the axes at one another: a category axis crosses the first value axis, each value axis the category axis. */
function crossAxes(axes: Axis[]): void {
  const cat = axes.find((a) => a.kind === "cat");
  const vals = axes.filter((a) => a.kind === "val");
  if (cat !== undefined && vals[0] !== undefined) cat.crossAxisId = vals[0].id;
  vals.forEach((v, i) => {
    if (cat === undefined) return;
    v.crossAxisId = cat.id;
    if (i > 0) {
      // A second value axis sits on the far side, like Excel's secondary axis.
      v.position = "r";
      v.crosses = "max";
    }
  });
}

// ---------------------------------------------------------------- waterfall

function readWaterfall(ser: XmlElement, set: DataSet, axes: Axis[], plotArea: XmlElement, ctx: Ctx): Group {
  const values = numbersOf(set.num);
  const categories = namesOf(set.str, values.length);
  const layout = firstChild(ser, "layoutPr");
  const totals = new Set<number>();
  const subtotals = layout === undefined ? undefined : firstChild(layout, "subtotals");
  for (const idx of subtotals === undefined ? [] : childElements(subtotals, "idx")) {
    const n = numAttr(idx, "val");
    if (n !== undefined && n >= 0) totals.add(n);
  }
  const increase: Look = { fill: paletteColor(ctx, 0, 3) };
  const decrease: Look = { fill: paletteColor(ctx, 1, 3) };
  const total: Look = { fill: paletteColor(ctx, 2, 3) };
  const series = emptySeries(0, seriesName(ser, 0), categories, values, increase);
  const spans: Array<[number, number] | null> = [];
  let level = 0;
  values.forEach((v, i) => {
    if (v === null) {
      spans.push(null);
      return;
    }
    // Observed: a point set as a total draws its own value from zero (not the running sum), and the running level continues from it.
    if (totals.has(i)) {
      spans.push([0, v]);
      level = v;
      series.points[i] = { look: total };
    } else {
      spans.push([level, level + v]);
      level += v;
      series.points[i] = { look: v < 0 ? decrease : increase };
    }
  });
  series.spans = spans;
  series.totals = [...totals].filter((i) => i < values.length && values[i] !== null).sort((a, b) => a - b);
  applyDataPoints(series, ser, ctx, (i) => series.points[i]?.look ?? increase);
  const labels = readLabels(firstChild(ser, "dataLabels"), ctx, ctx.text);
  if (labels !== undefined) series.labels = labels;
  const { cat, val } = pickAxes(axes, ser);
  const group = emptyGroup("bar", [series], cat === undefined || val === undefined ? undefined : [cat.id, val.id]);
  group.gapWidth = gapOf(cat, plotArea, 0.5);
  // Observed: connector lines are on unless the file turns them off.
  const visibility = layout === undefined ? undefined : firstChild(layout, "visibility");
  if (flag(visibility, "connectorLines", true)) group.connectors = { color: tx1(ctx, 0.25, 0.75), width: 0.75 };
  group.legendEntries = [
    { label: "Increase", look: increase }, { label: "Decrease", look: decrease }, { label: "Total", look: total },
  ];
  return group;
}

// ---------------------------------------------------------------- funnel

function readFunnel(ser: XmlElement, set: DataSet, axes: Axis[], plotArea: XmlElement, ctx: Ctx): Group {
  const values = numbersOf(set.num);
  const categories = namesOf(set.str, values.length);
  const look: Look = { fill: paletteColor(ctx, 0, 1) };
  const series = emptySeries(0, seriesName(ser, 0), categories, values, look);
  applyDataPoints(series, ser, ctx, () => look);
  // Observed: funnel labels are white, centred in their bars.
  const labels = readLabels(firstChild(ser, "dataLabels"), ctx, { ...ctx.text, color: themeColor("lt1", ctx, "#FFFFFF") });
  if (labels !== undefined) series.labels = labels;
  const { cat } = pickAxes(axes, ser);
  const group = emptyGroup("funnel", [series], cat === undefined ? undefined : [cat.id, axes.find((a) => a.kind === "val")!.id]);
  group.direction = "bar";
  group.gapWidth = gapOf(cat, plotArea, 0.06);
  if (cat !== undefined) {
    // The first stage sits on top.
    cat.reversed = true;
    cat.position = "l";
  }
  return group;
}

// ---------------------------------------------------------------- histogram and pareto

/** Excel's automatic bin width (Scott's normal reference rule, 3.5 s / n^(1/3)), rounded up to two significant digits (observed on two datasets). */
function scottWidth(values: readonly number[]): number {
  const n = values.length;
  if (n < 2) return 1;
  const mean = values.reduce((s, v) => s + v, 0) / n;
  const variance = values.reduce((s, v) => s + (v - mean) ** 2, 0) / (n - 1);
  const raw = (3.5 * Math.sqrt(variance)) / Math.cbrt(n);
  if (!(raw > 0)) return 1;
  const magnitude = 10 ** (Math.floor(Math.log10(raw)) - 1);
  return clean(Math.ceil(raw / magnitude - 1e-9) * magnitude);
}

/** Most bins one histogram makes: a tiny bin size over a wide range must not allocate millions. */
const MAX_BINS = 1000;

function histogramBins(values: number[], binning: XmlElement | undefined): { labels: string[]; counts: number[] } {
  const closedRight = binning === undefined || attr(binning, "intervalClosed") !== "l";
  const under = numAttr(binning, "underflow");
  const over = numAttr(binning, "overflow");
  let min = Infinity;
  let max = -Infinity;
  for (const v of values) {
    if (v < min) min = v;
    if (v > max) max = v;
  }
  if (min === Infinity) return { labels: [], counts: [] };
  const lo = under ?? min;
  const hi = over ?? max;
  const size = numAttr(binning === undefined ? undefined : firstChild(binning, "binSize"), "val");
  const wanted = numAttr(binning === undefined ? undefined : firstChild(binning, "binCount"), "val");
  let width: number;
  let count: number;
  if (size !== undefined && size > 0) {
    width = size;
    count = Math.max(1, Math.ceil((hi - lo) / width - 1e-9));
  } else if (wanted !== undefined && wanted >= 1) {
    count = Math.min(Math.round(wanted), MAX_BINS);
    width = hi > lo ? (hi - lo) / count : 1;
  } else {
    width = scottWidth(values);
    count = Math.max(1, Math.ceil((hi - lo) / width - 1e-9));
  }
  if (count > MAX_BINS) {
    count = MAX_BINS;
    width = (hi - lo) / count;
  }
  const edge = (k: number): string => String(clean(lo + k * width));
  const labels: string[] = [];
  const counts: number[] = new Array(count).fill(0);
  for (let k = 0; k < count; k++) {
    const first = k === 0;
    const last = k === count - 1;
    labels.push(closedRight ? `${first ? "[" : "("}${edge(k)}, ${edge(k + 1)}]` : `[${edge(k)}, ${edge(k + 1)}${last ? "]" : ")"}`);
  }
  let below = 0;
  let above = 0;
  for (const v of values) {
    if (under !== undefined && (closedRight ? v <= under : v < under)) {
      below++;
      continue;
    }
    if (over !== undefined && (closedRight ? v > over : v >= over)) {
      above++;
      continue;
    }
    const x = (v - lo) / width;
    const k = closedRight ? Math.ceil(x - 1e-9) - 1 : Math.floor(x + 1e-9);
    counts[Math.min(Math.max(k, 0), count - 1)]!++;
  }
  // Observed in Office's own labels: an underflow bin is "<= n", an overflow bin "> n".
  if (under !== undefined) {
    labels.unshift(`<= ${clean(under)}`);
    counts.unshift(below);
  }
  if (over !== undefined) {
    labels.push(`> ${clean(over)}`);
    counts.push(above);
  }
  return { labels, counts };
}

/** `cx:aggregation`: sum by default, a count when it says so (or when there is nothing to sum). */
function aggregate(set: DataSet, layout: XmlElement | undefined): { names: string[]; values: number[] } {
  const rows = set.str?.count ?? 0;
  const names = namesOf(set.str, rows);
  const nums = set.num === undefined ? undefined : numbersOf(set.num);
  const agg = layout === undefined ? undefined : firstChild(layout, "aggregation");
  const count = nums === undefined || (agg !== undefined && attr(agg, "val") === "count");
  const order: string[] = [];
  const sums = new Map<string, number>();
  for (let i = 0; i < rows; i++) {
    const add = count ? 1 : nums![i];
    if (add === null || add === undefined) continue;
    const name = names[i]!;
    if (!sums.has(name)) order.push(name);
    sums.set(name, (sums.get(name) ?? 0) + add);
  }
  return { names: order, values: order.map((n) => sums.get(n)!) };
}

function readColumns(sers: XmlElement[], space: XmlElement, data: ChartexData | undefined, axes: Axis[], plotArea: XmlElement, ctx: Ctx): Group[] | string {
  const bars = sers.filter((s) => attr(s, "layoutId") === "clusteredColumn");
  const pareto = sers.find((s) => attr(s, "layoutId") === "paretoLine");
  const groups: Group[] = [];
  const barSeries: Series[] = [];
  let cumulative: { categories: string[]; values: number[] } | undefined;
  let ids: Group["axisIds"];
  for (const [index, ser] of bars.entries()) {
    const set = readDataSet(space, attr(firstChild(ser, "dataId") ?? ser, "val"), data, ctx);
    const layout = firstChild(ser, "layoutPr");
    const binning = layout === undefined ? undefined : firstChild(layout, "binning");
    let categories: string[];
    let values: number[];
    if (set.str === undefined || binning !== undefined) {
      const raw = numbersOf(set.num).filter((v): v is number => v !== null);
      ({ labels: categories, counts: values } = histogramBins(raw, binning));
      // Only the bins stay once the raw values are counted into them.
      releasePoints(ctx, set.claimed);
      claimPoints(ctx, values.length);
    } else {
      ({ names: categories, values } = aggregate(set, layout));
    }
    // Observed: a Pareto draws its categories largest first.
    if (pareto !== undefined) {
      const order = categories.map((_, i) => i).sort((a, b) => values[b]! - values[a]! || a - b);
      categories = order.map((i) => categories[i]!);
      values = order.map((i) => values[i]!);
      cumulative = { categories, values };
    }
    const look: Look = { fill: paletteColor(ctx, index, bars.length), line: { color: themeColor("lt1", ctx, "#FFFFFF"), width: 0.75 } };
    const series = emptySeries(index, seriesName(ser, index), categories, values, look);
    applyDataPoints(series, ser, ctx, () => look);
    const labels = readLabels(firstChild(ser, "dataLabels"), ctx, ctx.text);
    if (labels !== undefined) series.labels = labels;
    barSeries.push(series);
    const { cat, val } = pickAxes(axes, ser);
    if (ids === undefined && cat !== undefined && val !== undefined) ids = [cat.id, val.id];
  }
  const first = bars[0];
  if (first === undefined) return "this chart has no data to preview";
  const bar = emptyGroup("bar", barSeries, ids);
  const catAxis = axes.find((a) => a.kind === "cat");
  bar.gapWidth = gapOf(catAxis, plotArea, 0);
  groups.push(bar);
  if (pareto !== undefined && cumulative !== undefined && catAxis !== undefined) {
    const total = cumulative.values.reduce((s, v) => s + v, 0);
    let running = 0;
    const shares = cumulative.values.map((v) => {
      running += v;
      return total > 0 ? running / total : 0;
    });
    const line: Line = { color: paletteColor(ctx, 1, 2), width: 2.25, cap: "round", join: "round" };
    const series = emptySeries(1, "Cumulative %", cumulative.categories, shares, { line });
    series.formatCode = "0%";
    const valAxis = axes.find((a) => a.kind === "val" && a.id === numAttr(firstChild(pareto, "axisId"), "val"));
    const group = emptyGroup("line", [series], valAxis === undefined ? undefined : [catAxis.id, valAxis.id]);
    groups.push(group);
  }
  return groups;
}

// ---------------------------------------------------------------- box and whisker

function quantile(sorted: readonly number[], p: number, inclusive: boolean): number {
  const n = sorted.length;
  // QUARTILE.INC / QUARTILE.EXC, as Excel's `quartileMethod` names them.
  const pos = inclusive ? p * (n - 1) + 1 : p * (n + 1);
  if (pos <= 1) return sorted[0]!;
  if (pos >= n) return sorted[n - 1]!;
  const below = Math.floor(pos);
  return sorted[below - 1]! + (pos - below) * (sorted[below]! - sorted[below - 1]!);
}

function summarize(sorted: number[], inclusive: boolean, keepInner: boolean): BoxSummary {
  const q1 = quantile(sorted, 0.25, inclusive);
  const median = quantile(sorted, 0.5, true);
  const q3 = quantile(sorted, 0.75, inclusive);
  const fence = 1.5 * (q3 - q1);
  let low = Infinity;
  let high = -Infinity;
  const outliers: number[] = [];
  for (const v of sorted) {
    if (v < q1 - fence || v > q3 + fence) outliers.push(v);
    else {
      if (v < low) low = v;
      if (v > high) high = v;
    }
  }
  const box: BoxSummary = {
    count: sorted.length, q1, median, q3, low: low === Infinity ? q1 : low, high: high === -Infinity ? q3 : high,
    mean: sorted.reduce((s, v) => s + v, 0) / sorted.length, outliers,
  };
  if (keepInner) box.inner = sorted;
  return box;
}

function readBoxes(sers: XmlElement[], space: XmlElement, data: ChartexData | undefined, axes: Axis[], plotArea: XmlElement, ctx: Ctx): Group | string {
  const series: Series[] = [];
  let shown = { meanMarker: true, meanLine: false, outliers: true, innerPoints: false };
  sers.forEach((ser, index) => {
    const set = readDataSet(space, attr(firstChild(ser, "dataId") ?? ser, "val"), data, ctx);
    const values = numbersOf(set.num);
    const layout = firstChild(ser, "layoutPr");
    const vis = layout === undefined ? undefined : firstChild(layout, "visibility");
    const stats = layout === undefined ? undefined : firstChild(layout, "statistics");
    shown = {
      meanMarker: flag(vis, "meanMarker", true), meanLine: flag(vis, "meanLine", false),
      outliers: flag(vis, "outliers", true), innerPoints: flag(vis, "nonoutliers", false),
    };
    const inclusive = stats !== undefined && attr(stats, "quartileMethod") === "inclusive";
    const names = set.str === undefined ? undefined : namesOf(set.str, values.length);
    const order: string[] = [];
    const groups = new Map<string, number[]>();
    values.forEach((v, i) => {
      if (v === null) return;
      const key = names === undefined ? "" : names[i]!;
      let bucket = groups.get(key);
      if (bucket === undefined) {
        bucket = [];
        groups.set(key, bucket);
        order.push(key);
      }
      bucket.push(v);
    });
    const fill = paletteColor(ctx, index, sers.length);
    // Observed: the outline, whiskers, median and outliers are a darker shade of the box.
    const look: Look = { fill, line: { color: mixColors(fill, "#000000", 0.35), width: 0.75 } };
    const boxes = order.map((key) => summarize(groups.get(key)!.sort((a, b) => a - b), inclusive, shown.innerPoints));
    const s = emptySeries(index, seriesName(ser, index), order, boxes.map((b) => b.median), look);
    s.boxes = boxes;
    applyDataPoints(s, ser, ctx, () => look);
    series.push(s);
  });
  const { cat, val } = pickAxes(axes, sers[0]);
  if (cat === undefined || val === undefined) return "this chart has no axes to preview";
  const group = emptyGroup("box", series, [cat.id, val.id]);
  group.gapWidth = gapOf(cat, plotArea, 1);
  group.box = shown;
  return group;
}

// ---------------------------------------------------------------- treemap and sunburst

interface Build {
  node: TreeNode;
  kids: Map<string, Build>;
}

function readTree(ser: XmlElement, set: DataSet, ctx: Ctx): { group: Group; leaves: number } | string {
  const dim = set.str;
  const sizes = numbersOf(set.num);
  if (dim === undefined || sizes.length === 0) return "this chart has no hierarchy to preview";
  const roots = new Map<string, Build>();
  let leaves = 0;
  for (let row = 0; row < dim.count; row++) {
    const size = sizes[row];
    if (size === null || size === undefined || size <= 0) continue;
    let level = roots;
    for (let l = 0; l < dim.levels.length; l++) {
      const name = dim.levels[l]![row] ?? "";
      let entry = level.get(name);
      if (entry === undefined) {
        entry = { node: { name, value: 0, color: "" }, kids: new Map() };
        level.set(name, entry);
      }
      entry.node.value += size;
      level = entry.kids;
    }
    leaves++;
  }
  // Colours follow the order the branches first appear in the data; the drawing order sorts by size (observed on both types).
  const tops = [...roots.values()];
  const legend = tops.map((t, i) => ({ label: t.node.name, look: { fill: paletteColor(ctx, i, tops.length) } as Look }));
  const finish = (builds: Build[], color: (i: number) => string | undefined, inherited: string): TreeNode[] => {
    const sorted = builds.map((b, i) => ({ b, i })).sort((a, b) => b.b.node.value - a.b.node.value || a.i - b.i);
    return sorted.map(({ b, i }) => {
      const node = b.node;
      node.color = color(i) ?? inherited;
      if (b.kids.size > 0) node.children = finish([...b.kids.values()], () => undefined, node.color);
      return node;
    });
  };
  const tree = finish(tops, (i) => legend[i]!.look.fill as string, "");
  const first = legend[0]?.look;
  const series = emptySeries(0, seriesName(ser, 0), tops.map((t) => t.node.name), tops.map((t) => t.node.value), first ?? {});
  const labels = readLabels(firstChild(ser, "dataLabels"), ctx, { ...ctx.text, color: themeColor("lt1", ctx, "#FFFFFF") });
  if (labels !== undefined) series.labels = labels;
  const kind = attr(ser, "layoutId") === "sunburst" ? "sunburst" : "treemap";
  const group = emptyGroup(kind, [series], undefined);
  group.tree = tree;
  group.legendEntries = legend;
  const layout = firstChild(ser, "layoutPr");
  const parents = layout === undefined ? undefined : firstChild(layout, "parentLabelLayout");
  const mode = parents === undefined ? undefined : attr(parents, "val");
  group.parentLabels = mode === "banner" ? "banner" : mode === "none" ? "none" : "overlapping";
  // Measured on Excel's own render (chartex-oracle.xlsx, treemap-banner): one grey band for every branch.
  if (group.parentLabels === "banner") group.bannerFill = tx1(ctx, 0.35, 0.65);
  return { group, leaves };
}

// ---------------------------------------------------------------- entry

/**
 * Read a `cx:chartSpace`. Returns the reason in words when the chart cannot be
 * drawn (a type the preview does not draw, data it cannot find), the model
 * otherwise. Throws `ChartTooLarge` past the point bounds.
 */
export function readChartEx(space: XmlElement, theme: DrawingTheme, data: ChartexData | undefined, owner?: FillOwner): ChartSpecModel | string {
  const chart = firstChild(space, "chart");
  const plotArea = chart === undefined ? undefined : firstChild(chart, "plotArea");
  const region = plotArea === undefined ? undefined : firstChild(plotArea, "plotAreaRegion");
  if (chart === undefined || plotArea === undefined || region === undefined) return "this chart has no plot area to preview";
  const sers = childElements(region, "series");
  const first = sers[0];
  if (first === undefined) return "this chart has no series to preview";
  const layoutId = attr(first, "layoutId") ?? "";
  if (layoutId === "regionMap") return "Map charts are not previewed";
  if (!["waterfall", "funnel", "clusteredColumn", "boxWhisker", "treemap", "sunburst"].includes(layoutId)) {
    return `chart type ${layoutId === "" ? "unknown" : layoutId} is not previewed`;
  }

  // The Office 2016 chart style: 9pt grey text, 14pt title, bg1 chart area with a pale border.
  const base: TextLook = { size: 9, bold: false, italic: false, color: "", family: theme.fonts.minor.latin || "Calibri" };
  // The theme's complex-script slot, as for the classic charts: what Office sets Arabic chart text in.
  const themeCs = themeFontFamily(theme.fonts.minor, "cs", true);
  if (themeCs !== undefined && themeCs !== "") base.csFamily = themeCs;
  const pre: Ctx = { theme, style: 2, text: base, owner, budget: { points: 0, read: 0 } };
  const grey = tx1(pre, 0.65, 0.35);
  const ctx: Ctx = { ...pre, text: readText(firstChild(space, "txPr"), { ...base, color: grey }, pre) };

  const axes = readAxes(plotArea, ctx);
  const sets = (ser: XmlElement): DataSet => readDataSet(space, attr(firstChild(ser, "dataId") ?? ser, "val"), data, ctx);
  let groups: Group[];
  if (layoutId === "waterfall" || layoutId === "funnel") {
    const set = sets(first);
    if (set.num === undefined) return "this chart has no cached values to preview";
    if (layoutId === "funnel") ensureAxes(axes, ctx, { cat: true, val: true });
    crossAxes(axes);
    groups = [layoutId === "waterfall" ? readWaterfall(first, set, axes, plotArea, ctx) : readFunnel(first, set, axes, plotArea, ctx)];
  } else if (layoutId === "clusteredColumn") {
    crossAxes(axes);
    const read = readColumns(sers, space, data, axes, plotArea, ctx);
    if (typeof read === "string") return read;
    groups = read;
  } else if (layoutId === "boxWhisker") {
    crossAxes(axes);
    const read = readBoxes(sers.filter((s) => attr(s, "layoutId") === "boxWhisker"), space, data, axes, plotArea, ctx);
    if (typeof read === "string") return read;
    groups = [read];
  } else {
    const read = readTree(first, sets(first), ctx);
    if (typeof read === "string") return read;
    groups = [read.group];
  }

  const spec: ChartSpecModel = {
    groups,
    axes: groups.some((g) => g.axisIds !== undefined) ? axes : [],
    plot: { look: readLook(firstChild(plotArea, "spPr"), ctx, {}) },
    look: readLook(firstChild(space, "spPr"), ctx, { fill: themeColor("lt1", ctx, "#FFFFFF"), line: { color: tx1(ctx, 0.15, 0.85), width: 0.75 } }),
    roundedCorners: false,
    blanks: "gap",
    date1904: false,
    text: ctx.text,
  };
  // Observed: a title element with no text of its own shows "Chart Title".
  const title = readTitle(firstChild(chart, "title"), ctx, { ...ctx.text, size: 14 }, "Chart Title");
  if (title !== undefined) spec.title = title;
  const legendEl = firstChild(chart, "legend");
  if (legendEl !== undefined) {
    const pos = attr(legendEl, "pos");
    const legend: Legend = {
      position: pos === "b" || pos === "l" || pos === "r" || pos === "tr" ? pos : "t",
      overlay: flag(legendEl, "overlay", false),
      text: readText(firstChild(legendEl, "txPr"), ctx.text, ctx),
      look: readLook(firstChild(legendEl, "spPr"), ctx, {}),
      deleted: [],
    };
    spec.legend = legend;
  }
  return spec;
}
