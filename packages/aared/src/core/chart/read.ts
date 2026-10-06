/**
 * `c:chartSpace` -> `ChartSpecModel`, in the Worker. One pass over the part;
 * every property comes out resolved, so the page never looks at XML or at the
 * theme again.
 *
 * Absent formatting falls back the way Office does for the chart's
 * `c:style` (default 2): theme accents for series, `tx1` greys for axes and
 * gridlines, 10pt minor-font text, an 18pt bold title. `tx1`/`bg1` are
 * colour-map names: the caller passes a theme mapped through the slide's
 * `p:clrMap` (PPTX) or the default map (XLSX), `themeForColorMap`.
 */

import { attr, childElements, firstChild, textContent } from "../xml.js";
import type { XmlElement } from "../xml.js";
import { ARABIC_PROBE } from "../chars.js";
import { resolveDrawingColor, resolveThemeTypeface, shadeColor, themeFontFamily } from "../theme.js";
import type { DrawingTheme } from "../theme.js";
import { isDateFormatCode } from "../xlsx/numberfmt.js";
import { resolveGradientFill, resolvePatternFill, resolvePictureFill } from "../fill.js";
import type { FillOwner } from "../fill.js";
import type { ShapeFillModel } from "../model.js";
import type {
  Axis, ChartSpecModel, DataLabels, ErrorBars, Fill, Group, GroupKind, LabelField, LabelPart, LabelPosition, Legend, Line, Look,
  ManualLayout, Marker, MarkerSymbol, Point, Series, TextLook, TickMark, TimeUnit, Title, Trendline,
} from "./spec.js";
import { isRadialKind } from "./spec.js";

const EMU_PER_PT = 12700;
/**
 * Plot types the engine draws. The 3D kinds draw flat (ARCHITECTURE, Known
 * limits). Stock and of-pie are kinds of their own: a stock group is its
 * series plus the hi-low lines and up-down bars between them, an of-pie group
 * is a pie plus the second plot the split-off points go to. Drawn as a plain
 * line or pie they would tell the reader something else.
 */
const GROUP_TAGS: Record<string, GroupKind> = {
  barChart: "bar", bar3DChart: "bar", lineChart: "line", line3DChart: "line", areaChart: "area", area3DChart: "area",
  pieChart: "pie", pie3DChart: "pie", doughnutChart: "doughnut", scatterChart: "scatter",
  bubbleChart: "bubble", radarChart: "radar", stockChart: "stock", ofPieChart: "ofPie",
};
export const SUPPORTED_GROUP_TAGS = Object.keys(GROUP_TAGS);

// ---------------------------------------------------------------- primitives

function val(el: XmlElement | undefined, name: string): string | undefined {
  const child = el === undefined ? undefined : firstChild(el, name);
  return child === undefined ? undefined : attr(child, "val");
}

function num(el: XmlElement | undefined, name: string): number | undefined {
  const raw = val(el, name);
  if (raw === undefined) return undefined;
  const n = Number(raw);
  return Number.isFinite(n) ? n : undefined;
}

/** An enumerated `val` of `el`'s `name` child, or `fallback` when absent or not one of `allowed` (the schema default). */
function choice<T extends string>(el: XmlElement | undefined, name: string, allowed: readonly T[], fallback: T): T {
  const raw = val(el, name);
  return raw !== undefined && (allowed as readonly string[]).includes(raw) ? raw as T : fallback;
}

/** CT_Boolean: an element with no `val` means true; an absent element gives `fallback`. */
function bool(el: XmlElement | undefined, name: string, fallback: boolean): boolean {
  const child = el === undefined ? undefined : firstChild(el, name);
  if (child === undefined) return fallback;
  const raw = attr(child, "val");
  return raw === undefined || raw === "1" || raw === "true";
}

// ---------------------------------------------------------------- color & look

export interface Ctx {
  theme: DrawingTheme;
  style: number;
  text: TextLook;
  /** The chart part's relationships, for picture fills; absent when the caller has none (the fill is then `unsupported`). */
  owner: FillOwner | undefined;
  /** Points the chart holds now, and points it has read in all; shared by every copy of this context (see `claimPoints`). */
  budget: { points: number; read: number };
}

function color(container: XmlElement | undefined, ctx: Ctx): string | undefined {
  return resolveDrawingColor(container, ctx.theme);
}

/** A shared-reader fill as a chart `Fill`: a solid is its color, a one-stop gradient is that stop's color, the rest stays a model. */
function chartFill(fill: ShapeFillModel): Fill {
  if (fill.kind === "solid") return fill.color;
  if ((fill.kind === "linear-gradient" || fill.kind === "path-gradient") && fill.stops.length === 1) return fill.stops[0]!.color;
  return fill;
}

/**
 * The single color a stroke or text run takes from a fill: a stroke or a text
 * run has no paint server here, so a gradient gives its first stop, a pattern
 * its foreground, and a picture or unreadable fill gives none (the caller's
 * default color stands). The chart text and outline limit in ARCHITECTURE.
 */
function strokeColor(fill: Fill): string | undefined {
  if (typeof fill === "string") return fill;
  switch (fill.kind) {
    case "linear-gradient":
    case "path-gradient":
      return fill.stops[0]!.color;
    case "pattern":
      return fill.foreground;
    default:
      return undefined;
  }
}

/**
 * The fill a `c:spPr`-like element states: a Fill, `null` for an explicit
 * noFill, `undefined` when silent. Read by the same functions as a shape's
 * fill (`core/fill.ts`), so a pattern, a path gradient or a picture keeps its
 * real definition; one that cannot be read comes back `unsupported`, for the
 * page to mark rather than to replace with a color.
 */
function readFill(spPr: XmlElement | undefined, ctx: Ctx): Fill | null | undefined {
  if (spPr === undefined) return undefined;
  for (const child of spPr.children) {
    switch (child.local) {
      case "noFill":
        return null;
      case "solidFill":
        return color(child, ctx) ?? null;
      case "gradFill":
        return chartFill(resolveGradientFill(child, ctx.theme));
      case "pattFill":
        return chartFill(resolvePatternFill(child, ctx.theme));
      case "blipFill":
        return chartFill(resolvePictureFill(child, ctx.owner));
      case "grpFill":
        return { kind: "unsupported", source: "group" };
    }
  }
  return undefined;
}

const DASHES: Record<string, number[]> = {
  dot: [1, 1], sysDot: [1, 1], dash: [4, 3], sysDash: [3, 1], lgDash: [8, 3], dashDot: [4, 3, 1, 3],
  sysDashDot: [3, 1, 1, 1], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

/** `a:ln` merged over `base`: `null` for an explicit noFill line, `undefined` when silent. */
function readLine(spPr: XmlElement | undefined, ctx: Ctx, base: Line | undefined): Line | null | undefined {
  const ln = spPr === undefined ? undefined : firstChild(spPr, "ln");
  if (ln === undefined) return undefined;
  const widthEmu = Number(attr(ln, "w"));
  const fill = readFill(ln, ctx);
  if (fill === null) return null;
  const stroke = (fill === undefined ? undefined : strokeColor(fill)) ?? base?.color;
  if (stroke === undefined) return null;
  const width = Number.isFinite(widthEmu) && attr(ln, "w") !== undefined ? widthEmu / EMU_PER_PT : base?.width ?? 0.75;
  const line: Line = { color: stroke, width };
  const preset = val(ln, "prstDash");
  if (preset !== undefined && preset !== "solid" && DASHES[preset] !== undefined) {
    line.dash = DASHES[preset]!.map((d) => d * Math.max(width, 0.75));
  } else if (preset === undefined && base?.dash !== undefined) {
    line.dash = base.dash;
  }
  const cap = attr(ln, "cap");
  if (cap === "rnd") line.cap = "round";
  else if (cap === "sq") line.cap = "square";
  if (firstChild(ln, "round") !== undefined) line.join = "round";
  else if (firstChild(ln, "bevel") !== undefined) line.join = "bevel";
  return line;
}

/** Resolve a `c:spPr` over the element's default look. */
export function readLook(spPr: XmlElement | undefined, ctx: Ctx, base: Look): Look {
  const fill = readFill(spPr, ctx);
  const line = readLine(spPr, ctx, base.line);
  const look: Look = {};
  const f = fill === undefined ? base.fill : fill ?? undefined;
  const l = line === undefined ? base.line : line ?? undefined;
  if (f !== undefined) look.fill = f;
  if (l !== undefined) look.line = l;
  return look;
}

// ---------------------------------------------------------------- text

export function plainText(rich: XmlElement): string[] {
  return childElements(rich, "p").map((p) =>
    p.children.map((run) => (run.local === "r" || run.local === "fld") ? textContent(firstChild(run, "t") ?? run) : run.local === "br" ? "\n" : "").join(""),
  ).flatMap((line) => line.split("\n"));
}

/** Merge `a:defRPr`/`a:rPr` attributes over `base`. */
function applyRunProps(rPr: XmlElement | undefined, base: TextLook, ctx: Ctx, sample: string): TextLook {
  if (rPr === undefined) return base;
  const out = { ...base };
  const sz = Number(attr(rPr, "sz"));
  if (attr(rPr, "sz") !== undefined && Number.isFinite(sz) && sz > 0) out.size = sz / 100;
  const b = attr(rPr, "b");
  if (b !== undefined) out.bold = b === "1" || b === "true";
  const i = attr(rPr, "i");
  if (i !== undefined) out.italic = i === "1" || i === "true";
  const fill = readFill(rPr, ctx);
  const textColor = fill === undefined || fill === null ? undefined : strokeColor(fill);
  if (textColor !== undefined) out.color = textColor;
  const latin = firstChild(rPr, "latin");
  const family = resolveThemeTypeface(latin === undefined ? undefined : attr(latin, "typeface"), sample, ctx.theme);
  if (family !== undefined && family !== "") out.family = family;
  const cs = csTypeface(rPr, ctx.theme);
  if (cs !== undefined) out.csFamily = cs;
  return out;
}

/** The complex-script face a run's `a:cs` names, theme tokens resolved to the Arabic slot; undefined when it names none. */
function csTypeface(rPr: XmlElement, theme: DrawingTheme): string | undefined {
  const node = firstChild(rPr, "cs");
  const family = resolveThemeTypeface(node === undefined ? undefined : attr(node, "typeface"), ARABIC_PROBE, theme);
  return family === undefined || family === "" ? undefined : family;
}

/** A `c:txPr` or `c:rich` body over `base`; `sample` picks the theme's script slot. */
export function readText(body: XmlElement | undefined, base: TextLook, ctx: Ctx, sample = ""): TextLook {
  if (body === undefined) return base;
  let out = base;
  const bodyPr = firstChild(body, "bodyPr");
  const rot = bodyPr === undefined ? undefined : attr(bodyPr, "rot");
  const vert = bodyPr === undefined ? undefined : attr(bodyPr, "vert");
  const p = firstChild(body, "p");
  const pPr = p === undefined ? undefined : firstChild(p, "pPr");
  out = applyRunProps(pPr === undefined ? undefined : firstChild(pPr, "defRPr"), out, ctx, sample);
  const run = p === undefined ? undefined : firstChild(p, "r");
  out = applyRunProps(run === undefined ? undefined : firstChild(run, "rPr"), out, ctx, sample);
  // Text turns at most a quarter turn either way; anything else (Office is
  // believed to write -60000000 for "automatic") leaves the layout to choose.
  const degrees = rot === undefined ? NaN : Number(rot) / 60000;
  if (Math.abs(degrees) <= 90) out = { ...out, rotation: degrees };
  else if (vert === "vert" || vert === "eaVert") out = { ...out, rotation: 90 };
  else if (vert === "vert270") out = { ...out, rotation: -90 };
  return out;
}

// ---------------------------------------------------------------- Office's c:style defaults

const ACCENTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
/** Office's series color variations past the sixth series: [lumMod, lumOff] per round. */
const VARIATIONS: Array<[number, number]> = [[1, 0], [0.6, 0], [0.8, 0.2], [0.8, 0], [0.6, 0.4], [0.5, 0], [0.7, 0.3], [0.7, 0], [0.5, 0.5]];

function shadeHex(hex: string, lumMod: number, lumOff: number): string {
  return shadeColor(hex, lumMod, lumOff) ?? hex;
}

export function themeColor(name: string, from: { theme: DrawingTheme }, fallback: string): string {
  const raw = from.theme.colors[name];
  return raw === undefined ? fallback : raw.startsWith("#") ? raw : `#${raw}`;
}

/** Text and line grey: `tx1` lightened, the way the built-in styles derive axis and gridline colors. */
export function tx1(ctx: Ctx, lumMod: number, lumOff: number): string {
  return shadeHex(themeColor("dk1", ctx, "#000000"), lumMod, lumOff);
}

/** Automatic color of series (or point) `index` of `count` under the chart's `c:style`. */
export function paletteColor(ctx: Ctx, index: number, count: number): string {
  const row = (ctx.style - 1) % 8;
  if (row === 0) {
    // Greyscale styles: tx1 from dark to light.
    const t = count <= 1 ? 0.5 : index / (count - 1);
    return tx1(ctx, 1, 0.25 + 0.5 * t);
  }
  if (row === 1) {
    const base = themeColor(ACCENTS[index % 6]!, ctx, "#4472C4");
    const [mod, off] = VARIATIONS[Math.floor(index / 6) % VARIATIONS.length]!;
    return mod === 1 && off === 0 ? base : shadeHex(base, mod, off);
  }
  // Monochrome styles: one accent from darker to lighter.
  const base = themeColor(ACCENTS[row - 2]!, ctx, "#4472C4");
  if (count <= 1) return base;
  const t = index / (count - 1);
  return t < 0.5 ? shadeHex(base, 0.5 + t, 0) : shadeHex(base, 1 - (t - 0.5), (t - 0.5) * 0.8);
}

/** Built-in style 41-48: dark chart area. */
function darkStyle(ctx: Ctx): boolean {
  return ctx.style >= 41;
}

export function defaultLine(ctx: Ctx, width = 0.75): Line {
  return { color: tx1(ctx, 1, 0.5), width };
}

/**
 * The line Office draws for an unstyled axis, trendline, error bar, leader
 * line or hi-low line: thin and `tx1` itself (black). Measured on the oracle's
 * bare cases in PowerPoint's PDF and Excel's export alike (DECISION_LOG §74).
 */
function inkLine(ctx: Ctx): Line {
  return { color: tx1(ctx, 1, 0), width: 0.75 };
}

// ---------------------------------------------------------------- layout

function readLayout(el: XmlElement | undefined): ManualLayout | undefined {
  const layout = el === undefined ? undefined : firstChild(el, "layout");
  const manual = layout === undefined ? undefined : firstChild(layout, "manualLayout");
  if (manual === undefined) return undefined;
  const out: ManualLayout = {
    xMode: val(manual, "xMode") === "edge" ? "edge" : "factor",
    yMode: val(manual, "yMode") === "edge" ? "edge" : "factor",
    target: val(manual, "layoutTarget") === "inner" ? "inner" : "outer",
  };
  for (const key of ["x", "y", "w", "h"] as const) {
    const n = num(manual, key);
    if (n !== undefined) out[key] = n;
  }
  return out;
}

// ---------------------------------------------------------------- data

/**
 * Most points one series may hold, and one chart in all. A cache's
 * `ptCount` and `pt/@idx` are only claims: a 600-byte part can claim nine
 * million points, and laying that out runs on the page. Past either bound
 * the chart is not drawn (its placeholder says why) rather than drawn in part,
 * which would tell the reader something else.
 */
export const MAX_SERIES_POINTS = 100_000;
export const MAX_CHART_POINTS = 250_000;
/** Points read in all, released or not: bounds the work, where `MAX_CHART_POINTS` bounds what is held. */
export const MAX_CHART_READ = 1_000_000;

/** A chart bigger than the preview draws; its message is the placeholder's reason. */
export class ChartTooLarge extends Error {}

/**
 * Count `count` more points against the chart's budget, before they are
 * allocated: many series that each claim a large `ptCount` would otherwise
 * allocate far more than the bound before a sum taken afterwards could stop it.
 */
export function claimPoints(ctx: Ctx, count: number): void {
  ctx.budget.points += count;
  ctx.budget.read += count;
  if (ctx.budget.points > MAX_CHART_POINTS) {
    throw new ChartTooLarge(`this chart has more than ${MAX_CHART_POINTS.toLocaleString("en-US")} points in all, more than the preview draws`);
  }
  if (ctx.budget.read > MAX_CHART_READ) {
    throw new ChartTooLarge(`this chart reads more than ${MAX_CHART_READ.toLocaleString("en-US")} points in all, more than the preview draws`);
  }
}

/**
 * Give back points that are no longer held: a histogram's raw values once they
 * are binned. What was read stays counted, so many large series still stop.
 */
export function releasePoints(ctx: Ctx, count: number): void {
  ctx.budget.points -= count;
}

/** Most outer category levels (hierarchy columns) kept; Excel itself stops well short of this. */
export const MAX_CATEGORY_LEVELS = 8;

interface Cache {
  count: number;
  points: Array<string | undefined>;
  formatCode?: string;
  numeric: boolean;
  /** Outer levels of grouped categories, nearest the inner one first. */
  levels?: Array<Array<{ start: number; text: string }>>;
}

export function tooManyPoints(): ChartTooLarge {
  return new ChartTooLarge(`this chart has a series of more than ${MAX_SERIES_POINTS.toLocaleString("en-US")} points, more than the preview draws`);
}

function collectPoints(container: XmlElement): Array<string | undefined> {
  const points: Array<string | undefined> = [];
  for (const pt of container.children) {
    if (pt.local !== "pt") continue;
    const idx = Number(attr(pt, "idx") ?? points.length);
    const v = firstChild(pt, "v");
    if (!Number.isInteger(idx) || idx < 0 || v === undefined) continue;
    if (idx >= MAX_SERIES_POINTS) throw tooManyPoints();
    points[idx] = textContent(v);
  }
  return points;
}

/** Points a cache spans: its claimed `ptCount`, or further when a point sits past it. */
function cacheCount(container: XmlElement, points: readonly unknown[]): number {
  const count = Math.max(num(container, "ptCount") ?? 0, points.length);
  if (count > MAX_SERIES_POINTS) throw tooManyPoints();
  return count;
}

function readCache(ref: XmlElement | undefined): Cache | undefined {
  if (ref === undefined) return undefined;
  const source = ref.children[0];
  if (source === undefined) return undefined;
  let container: XmlElement | undefined;
  let numeric = false;
  switch (source.local) {
    case "numRef": container = firstChild(source, "numCache"); numeric = true; break;
    case "numLit": container = source; numeric = true; break;
    case "strRef": container = firstChild(source, "strCache"); break;
    case "strLit": container = source; break;
    case "multiLvlStrRef": {
      // The first level is the innermost and labels each point; the rest group them, each `pt/@idx` opening a group.
      const cache = firstChild(source, "multiLvlStrCache");
      const levels = cache === undefined ? [] : childElements(cache, "lvl");
      if (cache === undefined || levels.length === 0) return undefined;
      const points = collectPoints(levels[0]!);
      const out: Cache = { count: cacheCount(cache, points), points, numeric: false };
      const outer: NonNullable<Cache["levels"]> = [];
      for (const level of levels.slice(1, MAX_CATEGORY_LEVELS)) {
        const starts: Array<{ start: number; text: string }> = [];
        collectPoints(level).forEach((text, start) => { if (text !== undefined && start < out.count) starts.push({ start, text }); });
        if (starts.length > 0) outer.push(starts);
      }
      if (outer.length > 0) out.levels = outer;
      return out;
    }
    default: return undefined;
  }
  if (container === undefined) return undefined;
  const points = collectPoints(container);
  const count = cacheCount(container, points);
  const formatEl = firstChild(container, "formatCode");
  const formatCode = formatEl === undefined ? undefined : textContent(formatEl);
  return formatCode === undefined ? { count, points, numeric } : { count, points, numeric, formatCode };
}

function numbers(cache: Cache | undefined): Array<number | null> {
  if (cache === undefined) return [];
  const out: Array<number | null> = new Array(cache.count);
  for (let i = 0; i < cache.count; i++) {
    const text = cache.points[i];
    const n = text === undefined || text.trim() === "" ? NaN : Number(text);
    out[i] = Number.isFinite(n) ? n : null;
  }
  return out;
}

function strings(cache: Cache | undefined): string[] {
  if (cache === undefined) return [];
  const out: string[] = new Array(cache.count);
  for (let i = 0; i < cache.count; i++) out[i] = cache.points[i] ?? "";
  return out;
}

// ---------------------------------------------------------------- series parts

const SYMBOLS = new Set(["none", "circle", "square", "diamond", "triangle", "x", "star", "plus", "dot", "dash", "picture"]);
/** Office's automatic marker cycle. */
const AUTO_SYMBOLS: MarkerSymbol[] = ["diamond", "square", "triangle", "x", "star", "circle", "plus", "dot", "dash"];

function readMarker(el: XmlElement | undefined, ctx: Ctx, base: Marker): Marker {
  if (el === undefined) return base;
  const raw = val(el, "symbol");
  const symbol = raw !== undefined && SYMBOLS.has(raw) ? raw as MarkerSymbol : base.symbol;
  const size = num(el, "size") ?? base.size;
  return { symbol, size: Math.min(Math.max(size, 2), 72), look: readLook(firstChild(el, "spPr"), ctx, base.look) };
}

export const POSITIONS = new Set(["bestFit", "b", "ctr", "inBase", "inEnd", "l", "outEnd", "r", "t"]);
const LABEL_FIELDS = new Set(["VALUE", "SERIESNAME", "CATEGORYNAME", "PERCENTAGE", "CELLRANGE", "XVALUE", "YVALUE", "BUBBLESIZE"]);
/** A rich label is a sentence at most; the bound keeps a hostile part from sizing the model. */
const MAX_LABEL_PARTS = 64;
const MAX_LABEL_TEXT = 500;

/** Author-typed label text: runs as literal text, fields as the placeholder Office cached, paragraphs split by "\n". */
function richParts(rich: XmlElement): LabelPart[] {
  const parts: LabelPart[] = [];
  const push = (part: LabelPart): void => {
    if (parts.length < MAX_LABEL_PARTS) parts.push(typeof part === "string" ? part.slice(0, MAX_LABEL_TEXT) : part);
  };
  childElements(rich, "p").forEach((p, index) => {
    if (index > 0) push("\n");
    for (const run of p.children) {
      if (run.local === "r") {
        const t = firstChild(run, "t");
        push(t === undefined ? "" : textContent(t));
      } else if (run.local === "fld") {
        const t = firstChild(run, "t");
        const cached = (t === undefined ? "" : textContent(t)).slice(0, MAX_LABEL_TEXT);
        const type = attr(run, "type") ?? "";
        push(LABEL_FIELDS.has(type) ? { field: type as LabelField, cached } : cached);
      } else if (run.local === "br") {
        push("\n");
      }
    }
  });
  return parts;
}

/** A child of `el`'s `c:extLst/c:ext` entries, by local name (Office 2010+ extensions such as c15:showDataLabelsRange). */
function extChild(el: XmlElement, local: string): XmlElement | undefined {
  const list = firstChild(el, "extLst");
  if (list === undefined) return undefined;
  for (const ext of list.children) {
    const found = firstChild(ext, local);
    if (found !== undefined) return found;
  }
  return undefined;
}

function readLabels(el: XmlElement | undefined, ctx: Ctx, base: DataLabels | undefined): DataLabels | undefined {
  if (el === undefined) return base;
  const start: DataLabels = base ?? {
    showValue: false, showCategory: false, showSeries: false, showPercent: false, showBubbleSize: false, showLegendKey: false,
    text: ctx.text, look: {}, deleted: false, fromCells: false,
  };
  const out: DataLabels = {
    ...start,
    showValue: bool(el, "showVal", start.showValue),
    showCategory: bool(el, "showCatName", start.showCategory),
    showSeries: bool(el, "showSerName", start.showSeries),
    showPercent: bool(el, "showPercent", start.showPercent),
    showBubbleSize: bool(el, "showBubbleSize", start.showBubbleSize),
    showLegendKey: bool(el, "showLegendKey", start.showLegendKey),
    deleted: bool(el, "delete", false),
    text: readText(firstChild(el, "txPr"), start.text, ctx),
    look: readLook(firstChild(el, "spPr"), ctx, start.look),
    fromCells: start.fromCells,
  };
  const range = extChild(el, "showDataLabelsRange");
  if (range !== undefined) out.fromCells = attr(range, "val") === undefined || attr(range, "val") === "1" || attr(range, "val") === "true";
  // Per-label author text and position only come from a `c:dLbl`; they are never inherited by another point.
  delete out.custom;
  delete out.offset;
  const tx = firstChild(el, "tx");
  const rich = tx === undefined ? undefined : firstChild(tx, "rich");
  if (rich !== undefined) {
    const parts = richParts(rich);
    out.custom = parts;
    out.text = readText(rich, out.text, ctx, parts.map((p) => (typeof p === "string" ? p : p.cached)).join(""));
  }
  const layout = firstChild(el, "layout");
  const manual = layout === undefined ? undefined : firstChild(layout, "manualLayout");
  if (manual !== undefined) {
    const x = num(manual, "x");
    const y = num(manual, "y");
    // Office writes label positions as offsets from the default spot ("factor" mode); an "edge" mode is not a label offset.
    if ((x !== undefined || y !== undefined) && val(manual, "xMode") !== "edge" && val(manual, "yMode") !== "edge") out.offset = { x: x ?? 0, y: y ?? 0 };
  }
  const position = val(el, "dLblPos");
  if (position !== undefined && POSITIONS.has(position)) out.position = position as LabelPosition;
  const sep = firstChild(el, "separator");
  if (sep !== undefined) out.separator = textContent(sep);
  const fmt = firstChild(el, "numFmt");
  if (fmt !== undefined && attr(fmt, "sourceLinked") !== "1") {
    const code = attr(fmt, "formatCode");
    if (code !== undefined) out.formatCode = code;
  }
  if (firstChild(el, "showLeaderLines") !== undefined) {
    if (bool(el, "showLeaderLines", false)) {
      // Unstyled, Office draws leader lines thin and black (the oracle's g-pie-leaders).
      const line = readLook(firstChild(firstChild(el, "leaderLines") ?? el, "spPr"), ctx, { line: inkLine(ctx) }).line;
      if (line !== undefined) out.leaderLines = line;
      else delete out.leaderLines;
    } else {
      delete out.leaderLines;
    }
  }
  return out;
}

const TREND_TYPES = new Set(["linear", "exp", "log", "poly", "power", "movingAvg"]);

function readTrendlines(ser: XmlElement, ctx: Ctx): Trendline[] | undefined {
  const out: Trendline[] = [];
  for (const el of childElements(ser, "trendline")) {
    const raw = val(el, "trendlineType");
    const type = raw !== undefined && TREND_TYPES.has(raw) ? raw as Trendline["type"] : "linear";
    // Unstyled, Office draws a trendline thin, solid and black (the oracle's g-trend-* cases); the dotted
    // series-coloured line is what its Insert command writes out explicitly, so it arrives as spPr.
    const base: Line = inkLine(ctx);
    const look = readLook(firstChild(el, "spPr"), ctx, { line: base });
    const trend: Trendline = {
      type,
      order: Math.min(Math.max(Math.round(num(el, "order") ?? 2), 2), 6),
      period: Math.max(Math.round(num(el, "period") ?? 2), 2),
      forward: Math.max(num(el, "forward") ?? 0, 0),
      backward: Math.max(num(el, "backward") ?? 0, 0),
      dispEq: bool(el, "dispEq", false),
      dispRSqr: bool(el, "dispRSqr", false),
      line: look.line ?? base,
      labelText: ctx.text,
    };
    const intercept = num(el, "intercept");
    if (intercept !== undefined) trend.intercept = intercept;
    const name = firstChild(el, "name");
    if (name !== undefined) trend.name = textContent(name);
    const lbl = firstChild(el, "trendlineLbl");
    if (lbl !== undefined) {
      trend.labelText = readText(firstChild(lbl, "txPr"), ctx.text, ctx);
      const fmt = firstChild(lbl, "numFmt");
      if (fmt !== undefined && attr(fmt, "sourceLinked") !== "1") {
        const code = attr(fmt, "formatCode");
        if (code !== undefined) trend.labelFormatCode = code;
      }
    }
    out.push(trend);
  }
  return out.length === 0 ? undefined : out;
}

function readErrBars(ser: XmlElement, ctx: Ctx, kind: GroupKind): ErrorBars[] | undefined {
  const out: ErrorBars[] = [];
  for (const el of childElements(ser, "errBars")) {
    const rawType = val(el, "errBarType");
    const rawVal = val(el, "errValType");
    const valType = rawVal === "fixedVal" || rawVal === "percentage" || rawVal === "stdDev" || rawVal === "stdErr" || rawVal === "cust" ? rawVal : "fixedVal";
    const xy = kind === "scatter" || kind === "bubble";
    const plus = firstChild(el, "plus");
    const minus = firstChild(el, "minus");
    // Unstyled, Office draws error bars thin and black (the oracle's g-err-* cases).
    const line = readLook(firstChild(el, "spPr"), ctx, { line: inkLine(ctx) }).line;
    if (line === undefined) continue;
    const eb: ErrorBars = {
      // Only XY plots have an errDir; elsewhere the bars run along the value axis.
      dir: xy && val(el, "errDir") === "x" ? "x" : "y",
      type: rawType === "plus" || rawType === "minus" ? rawType : "both",
      valType,
      value: num(el, "val") ?? (valType === "percentage" ? 5 : 1),
      noEndCap: bool(el, "noEndCap", false),
      line,
    };
    if (valType === "cust") {
      if (plus !== undefined) eb.plus = numbers(readCache(plus));
      if (minus !== undefined) eb.minus = numbers(readCache(minus));
    }
    out.push(eb);
  }
  return out.length === 0 ? undefined : out;
}

function seriesName(ser: XmlElement, index: number): string {
  const tx = firstChild(ser, "tx");
  if (tx !== undefined) {
    const v = firstChild(tx, "v");
    if (v !== undefined) return textContent(v);
    const cache = readCache(tx);
    const parts = cache === undefined ? [] : cache.points.filter((p): p is string => p !== undefined);
    if (parts.length > 0) return parts.join(" ");
  }
  return `Series${index + 1}`;
}

/** A series drawn as a line (its colour is the line's), not a filled shape: lines, stock, scatter, and radar unless filled. */
function isLineLike(kind: GroupKind, radarStyle: string | undefined): boolean {
  return kind === "line" || kind === "stock" || kind === "scatter" || (kind === "radar" && radarStyle !== "filled");
}

function readSeries(group: XmlElement, kind: GroupKind, ctx: Ctx, groupLabels: DataLabels | undefined, vary: boolean, offset: number, total: number): Series[] {
  const sers = childElements(group, "ser");
  // A stock series with no c:marker gets automatic markers too, as a line series does (the oracle's
  // g-stock-bare); Excel's own stock charts write theirs out.
  const lineLike = isLineLike(kind, val(group, "radarStyle"));
  const groupMarkers = kind === "line" ? bool(group, "marker", true) : true;
  return sers.map((ser, position) => {
    const index = num(ser, "idx") ?? position + offset;
    const order = num(ser, "order") ?? index;
    const isXY = kind === "scatter" || kind === "bubble";
    const catCache = readCache(firstChild(ser, isXY ? "xVal" : "cat"));
    const valCache = readCache(firstChild(ser, isXY ? "yVal" : "val"));
    const sizeCache = kind === "bubble" ? readCache(firstChild(ser, "bubbleSize")) : undefined;
    // Each of the series' arrays (values, category names or numbers, bubble sizes) is as long as its cache;
    // grouped categories add a point per group. Claimed before any is allocated.
    claimPoints(ctx, Math.max(valCache?.count ?? 0, catCache?.count ?? 0, sizeCache?.count ?? 0)
      + (catCache?.levels?.reduce((sum, level) => sum + level.length, 0) ?? 0));
    const values = numbers(valCache);
    const auto = paletteColor(ctx, index, total);
    const defaultLook: Look = lineLike
      ? { line: { color: auto, width: 2.25, cap: "round", join: "round" } }
      : isRadialKind(kind)
        ? { fill: auto, line: { color: themeColor("lt1", ctx, "#FFFFFF"), width: 0.75 } }
        : { fill: auto };
    if (ctx.style >= 9 && ctx.style <= 16 && !lineLike) defaultLook.line = { color: tx1(ctx, 1, 0), width: 0.75 };
    const look = readLook(firstChild(ser, "spPr"), ctx, defaultLook);
    const seriesColor = look.line?.color ?? (typeof look.fill === "string" ? look.fill : auto);
    const autoSymbol = AUTO_SYMBOLS[index % AUTO_SYMBOLS.length]!;
    const scatterStyle = val(group, "scatterStyle");
    const showsMarkers = lineLike && groupMarkers && !(kind === "scatter" && (scatterStyle === "line" || scatterStyle === "smooth"))
      && !(kind === "radar" && val(group, "radarStyle") !== "marker");
    const baseMarker: Marker = {
      symbol: showsMarkers ? autoSymbol : "none",
      size: 5,
      look: { fill: seriesColor, line: { color: seriesColor, width: 0.75 } },
    };
    const marker = readMarker(firstChild(ser, "marker"), ctx, baseMarker);
    const smoothDefault = kind === "scatter" && (scatterStyle === "smooth" || scatterStyle === "smoothMarker");
    const series: Series = {
      index,
      order,
      name: seriesName(ser, index),
      categories: catCache === undefined ? [] : catCache.numeric ? [] : strings(catCache),
      values,
      formatCode: valCache?.formatCode ?? "General",
      look,
      marker,
      smooth: bool(ser, "smooth", smoothDefault),
      explosion: num(ser, "explosion") ?? 0,
      invertIfNegative: bool(ser, "invertIfNegative", false),
      points: {},
    };
    if (catCache?.numeric === true) {
      series.categoryValues = numbers(catCache);
      if (catCache.formatCode !== undefined) series.categoryFormatCode = catCache.formatCode;
    }
    if (catCache?.levels !== undefined) series.categoryGroups = catCache.levels;
    if (kind === "bubble") series.bubbleSizes = numbers(sizeCache);
    const labelRange = extChild(ser, "datalabelsRange");
    const rangeCache = labelRange === undefined ? undefined : firstChild(labelRange, "dlblRangeCache");
    if (rangeCache !== undefined) {
      const points = collectPoints(rangeCache);
      const count = cacheCount(rangeCache, points);
      claimPoints(ctx, count);
      series.labelRange = strings({ count, points, numeric: false });
    }
    const labels = readLabels(firstChild(ser, "dLbls"), ctx, groupLabels);
    if (labels !== undefined) series.labels = labels;
    if (!isRadialKind(kind)) {
      const trends = readTrendlines(ser, ctx);
      if (trends !== undefined) series.trendlines = trends;
      const bars = readErrBars(ser, ctx, kind);
      if (bars !== undefined) series.errBars = bars;
    }
    // An of-pie series styles one point more than it holds: the "Other" slice (Office writes its `c:dPt` at idx = point count).
    const count = Math.max(values.length, series.categories.length) + (kind === "ofPie" ? 1 : 0);
    for (const dPt of childElements(ser, "dPt")) {
      const idx = num(dPt, "idx");
      if (idx === undefined) continue;
      const base = vary ? pointLook(ctx, look, idx, count, lineLike) : look;
      const point: Point = { look: readLook(firstChild(dPt, "spPr"), ctx, base) };
      const marker = firstChild(dPt, "marker");
      if (marker !== undefined) point.marker = readMarker(marker, ctx, series.marker);
      const explosion = num(dPt, "explosion");
      if (explosion !== undefined) point.explosion = explosion;
      series.points[idx] = point;
    }
    if (series.labels !== undefined) {
      const dLbls = firstChild(ser, "dLbls");
      for (const dLbl of dLbls === undefined ? [] : childElements(dLbls, "dLbl")) {
        const idx = num(dLbl, "idx");
        if (idx === undefined) continue;
        const point = series.points[idx] ?? {};
        const labels = readLabels(dLbl, ctx, series.labels);
        if (labels !== undefined) point.labels = labels;
        series.points[idx] = point;
      }
    }
    if (vary) {
      // A point that only carries a label (a moved or retyped one) still takes its own colour.
      for (let idx = 0; idx < count; idx++) {
        const point = (series.points[idx] ??= {});
        point.look ??= pointLook(ctx, look, idx, count, lineLike);
      }
    }
    return series;
  }).sort((a, b) => a.order - b.order);
}

/** varyColors: each point takes the palette color its index would have as a series. */
function pointLook(ctx: Ctx, seriesLook: Look, idx: number, count: number, lineLike: boolean): Look {
  const auto = paletteColor(ctx, idx, count);
  if (lineLike) {
    return seriesLook.line === undefined ? seriesLook : { ...seriesLook, line: { ...seriesLook.line, color: auto } };
  }
  return { ...seriesLook, fill: auto };
}

// ---------------------------------------------------------------- groups & axes

function readGroup(el: XmlElement, kind: GroupKind, ctx: Ctx, offset: number, total: number): Group {
  const direction = val(el, "barDir") === "bar" ? "bar" : "col";
  const rawGrouping = val(el, "grouping");
  const grouping = rawGrouping === "stacked" || rawGrouping === "percentStacked" || rawGrouping === "clustered" || rawGrouping === "standard"
    ? rawGrouping
    : kind === "bar" ? "clustered" : "standard";
  // CT_Boolean defaults to true, but an absent varyColors leaves a non-pie chart one color per series.
  const radial = isRadialKind(kind);
  const varyColors = bool(el, "varyColors", radial) && (radial || childElements(el, "ser").length === 1);
  const labels = readLabels(firstChild(el, "dLbls"), ctx, undefined);
  const ids = childElements(el, "axId").map((ax) => Number(attr(ax, "val")));
  // Hi-low, drop and series lines with no style: thin and black, as Office draws them.
  const lineOf = (name: string): Line | undefined => {
    const child = firstChild(el, name);
    if (child === undefined) return undefined;
    return readLook(firstChild(child, "spPr"), ctx, { line: inkLine(ctx) }).line;
  };
  const stacked = grouping === "stacked" || grouping === "percentStacked";
  const group: Group = {
    kind,
    direction,
    grouping,
    varyColors,
    gapWidth: num(el, "gapWidth") ?? 150,
    // Stacked bars always sit on one another: Office ignores the authored overlap.
    overlap: stacked && kind === "bar" ? 100 : num(el, "overlap") ?? 0,
    holeSize: num(el, "holeSize") ?? (kind === "doughnut" ? 50 : 0),
    firstSliceAngle: num(el, "firstSliceAng") ?? 0,
    bubbleScale: num(el, "bubbleScale") ?? 100,
    sizeRepresents: val(el, "sizeRepresents") === "w" ? "w" : "area",
    scatterStyle: choice(el, "scatterStyle", ["line", "lineMarker", "marker", "none", "smooth", "smoothMarker"], "marker"),
    radarStyle: choice(el, "radarStyle", ["standard", "marker", "filled"], "standard"),
    series: readSeries(el, kind, ctx, labels, varyColors, offset, total),
    axisIds: ids.length >= 2 && ids.every(Number.isFinite) ? [ids[0]!, ids[1]!] : undefined,
  };
  if (labels !== undefined) group.labels = labels;
  const drop = lineOf("dropLines");
  if (drop !== undefined) group.dropLines = drop;
  const hiLow = lineOf("hiLowLines");
  if (hiLow !== undefined) group.hiLowLines = hiLow;
  const serLines = lineOf("serLines");
  if (serLines !== undefined) group.serLines = serLines;
  const bars = firstChild(el, "upDownBars");
  if (bars !== undefined) {
    // Office's automatic up-down bars: white with a dark outline when the last series closes above the first, dark when below.
    const outline: Line = { color: tx1(ctx, 1, 0), width: 0.75 };
    const look = (name: string, fill: string): Look => {
      const side = firstChild(bars, name);
      return readLook(side === undefined ? undefined : firstChild(side, "spPr"), ctx, { fill, line: outline });
    };
    group.upDownBars = {
      gapWidth: Math.min(Math.max(num(bars, "gapWidth") ?? 150, 0), 500),
      up: look("upBars", themeColor("lt1", ctx, "#FFFFFF")),
      // Unstyled down bars are dark grey, not black (PowerPoint and Excel alike, the oracle's g-stock-bare).
      down: look("downBars", tx1(ctx, 1, 0.25)),
    };
  }
  if (kind === "ofPie") group.ofPie = readOfPie(el, group, ctx);
  return group;
}

// ---------------------------------------------------------------- of-pie

/**
 * The points that move to the second plot, by index. `c:splitType` says how
 * they are chosen; absent (`auto`) Office takes the last third of the points,
 * rounded up (measured: 2-3 points give 1, 4-6 give 2, 7-9 give 3, 10-12
 * give 4). `pos` takes the last `splitPos` points, `val` the points under
 * `splitPos`, `percent` those whose share of the total is under `splitPos`
 * percent, `cust` the listed points.
 */
function splitPoints(el: XmlElement, values: ReadonlyArray<number | null>): number[] {
  const n = values.length;
  const type = val(el, "splitType") ?? "auto";
  const pos = num(el, "splitPos");
  const flags = new Array<boolean>(n).fill(false);
  if (type === "cust") {
    const list = firstChild(el, "custSplit");
    for (const pt of list === undefined ? [] : childElements(list, "secondPiePt")) {
      const i = Number(attr(pt, "val"));
      if (Number.isInteger(i) && i >= 0 && i < n) flags[i] = true;
    }
  } else if (type === "val") {
    for (let i = 0; i < n; i++) if (values[i] !== null && Math.abs(values[i]!) < (pos ?? 0)) flags[i] = true;
  } else if (type === "percent") {
    const total = values.reduce<number>((sum, v) => sum + Math.abs(v ?? 0), 0);
    for (let i = 0; i < n; i++) if (values[i] !== null && total > 0 && (Math.abs(values[i]!) / total) * 100 < (pos ?? 0)) flags[i] = true;
  } else {
    const last = type === "pos" && pos !== undefined ? Math.min(Math.max(Math.round(pos), 0), n) : Math.ceil(n / 3);
    for (let i = Math.max(n - last, 0); i < n; i++) flags[i] = true;
  }
  const out: number[] = [];
  flags.forEach((f, i) => { if (f && values[i] !== null) out.push(i); });
  return out;
}

function readOfPie(el: XmlElement, group: Group, ctx: Ctx): NonNullable<Group["ofPie"]> {
  const series = group.series[0];
  const n = series === undefined ? 0 : Math.max(series.values.length, series.categories.length);
  const other = series?.points[n]?.look ?? { fill: paletteColor(ctx, n, n + 1), line: { color: themeColor("lt1", ctx, "#FFFFFF"), width: 0.75 } };
  return {
    type: val(el, "ofPieType") === "bar" ? "bar" : "pie",
    secondary: series === undefined ? [] : splitPoints(el, series.values),
    secondSize: Math.min(Math.max(num(el, "secondPieSize") ?? 75, 5), 200),
    other,
  };
}

export const DISPLAY_UNITS: Record<string, number> = {
  hundreds: 1e2, thousands: 1e3, tenThousands: 1e4, hundredThousands: 1e5, millions: 1e6, tenMillions: 1e7,
  hundredMillions: 1e8, billions: 1e9, trillions: 1e12,
};
const DISPLAY_UNIT_NAMES: Record<string, string> = {
  hundreds: "Hundreds", thousands: "Thousands", tenThousands: "x 10000", hundredThousands: "x 100000", millions: "Millions",
  tenMillions: "x 10000000", hundredMillions: "x 100000000", billions: "Billions", trillions: "Trillions",
};

function readTitle(el: XmlElement | undefined, ctx: Ctx, base: TextLook, autoText: string[]): Title | undefined {
  if (el === undefined) return undefined;
  const tx = firstChild(el, "tx");
  const rich = tx === undefined ? undefined : firstChild(tx, "rich");
  let lines = autoText;
  let text = readText(firstChild(el, "txPr"), base, ctx);
  if (rich !== undefined) {
    lines = plainText(rich);
    text = readText(rich, base, ctx, lines.join(""));
  } else if (tx !== undefined) {
    const cache = readCache(tx);
    lines = cache === undefined ? autoText : [cache.points.filter((p): p is string => p !== undefined).join(" ")];
  }
  if (lines.every((line) => line.trim() === "")) return undefined;
  const title: Title = { lines, text, look: readLook(firstChild(el, "spPr"), ctx, {}), overlay: bool(el, "overlay", false) };
  const layout = readLayout(el);
  if (layout !== undefined) title.layout = layout;
  return title;
}

function readAxis(el: XmlElement, ctx: Ctx, titleBase: TextLook): Axis {
  const kind = el.local === "catAx" ? "cat" : el.local === "dateAx" ? "date" : el.local === "serAx" ? "ser" : "val";
  const scaling = firstChild(el, "scaling");
  const crossesRaw = val(el, "crosses");
  const tick = (name: string, fallback: TickMark): TickMark => {
    const raw = val(el, name);
    return raw === "none" || raw === "in" || raw === "out" || raw === "cross" ? raw : fallback;
  };
  const numFmt = firstChild(el, "numFmt");
  // An unstyled axis line is black, and absent tick marks are the schema's default, "cross", major and
  // minor alike: what PowerPoint and Excel both draw for a bare axis (the oracle's bare cases).
  const axisLine = readLook(firstChild(el, "spPr"), ctx, { line: inkLine(ctx) }).line;
  const gridLook = (name: string, base: Line): Line | undefined => {
    const grid = firstChild(el, name);
    return grid === undefined ? undefined : readLook(firstChild(grid, "spPr"), ctx, { line: base }).line;
  };
  const posRaw = val(el, "axPos");
  const axis: Axis = {
    id: num(el, "axId") ?? 0,
    kind,
    position: posRaw === "l" || posRaw === "r" || posRaw === "t" ? posRaw : "b",
    deleted: bool(el, "delete", false),
    crossAxisId: num(el, "crossAx") ?? 0,
    crosses: crossesRaw === "min" || crossesRaw === "max" ? crossesRaw : "autoZero",
    reversed: val(scaling, "orientation") === "maxMin",
    majorTick: tick("majorTickMark", "cross"),
    minorTick: tick("minorTickMark", "cross"),
    labelPosition: choice(el, "tickLblPos", ["nextTo", "low", "high", "none"], "nextTo"),
    text: readText(firstChild(el, "txPr"), ctx.text, ctx),
    formatCode: numFmt === undefined ? "General" : attr(numFmt, "formatCode") ?? "General",
    formatLinked: numFmt === undefined || attr(numFmt, "sourceLinked") === "1",
    between: val(el, "crossBetween") !== "midCat",
    labelOffset: num(el, "lblOffset") ?? 100,
    displayUnit: 1,
    auto: bool(el, "auto", false),
    multiLevel: !bool(el, "noMultiLvlLbl", false),
  };
  if (axisLine !== undefined) axis.line = axisLine;
  const crossesAt = num(el, "crossesAt");
  if (crossesAt !== undefined) axis.crossesAt = crossesAt;
  for (const key of ["min", "max", "logBase"] as const) {
    const n = num(scaling, key);
    if (n !== undefined) axis[key] = n;
  }
  const major = num(el, "majorUnit");
  if (major !== undefined && major > 0) axis.majorUnit = major;
  const minor = num(el, "minorUnit");
  if (minor !== undefined && minor > 0) axis.minorUnit = minor;
  const majorGrid = gridLook("majorGridlines", defaultLine(ctx));
  if (majorGrid !== undefined) axis.majorGrid = majorGrid;
  const minorGrid = gridLook("minorGridlines", { color: tx1(ctx, 1, 0.75), width: 0.75 });
  if (minorGrid !== undefined) axis.minorGrid = minorGrid;
  const skip = num(el, "tickLblSkip");
  if (skip !== undefined && skip >= 1) axis.labelSkip = Math.round(skip);
  const markSkip = num(el, "tickMarkSkip");
  if (markSkip !== undefined && markSkip >= 1) axis.markSkip = Math.round(markSkip);
  const disp = firstChild(el, "dispUnits");
  if (disp !== undefined) {
    const builtIn = val(disp, "builtInUnit");
    const custom = num(disp, "custUnit");
    axis.displayUnit = custom ?? DISPLAY_UNITS[builtIn ?? ""] ?? 1;
    const lbl = firstChild(disp, "dispUnitsLbl");
    if (lbl !== undefined && builtIn !== undefined) axis.displayUnitLabel = DISPLAY_UNIT_NAMES[builtIn] ?? "";
  }
  const timeUnit = (name: string): TimeUnit | undefined => {
    const raw = val(el, name);
    return raw === "days" || raw === "months" || raw === "years" ? raw : undefined;
  };
  const base = timeUnit("baseTimeUnit");
  if (base !== undefined) axis.baseUnit = base;
  const majorTime = timeUnit("majorTimeUnit");
  if (majorTime !== undefined) axis.majorTimeUnit = majorTime;
  const minorTime = timeUnit("minorTimeUnit");
  if (minorTime !== undefined) axis.minorTimeUnit = minorTime;
  const vertical = axis.position === "l" || axis.position === "r";
  const title = readTitle(firstChild(el, "title"), ctx, { ...titleBase, ...(vertical ? { rotation: -90 } : {}) }, ["Axis Title"]);
  if (title !== undefined) axis.title = title;
  return axis;
}

// ---------------------------------------------------------------- entry

/**
 * Read a chart part, or `undefined` when it has no plot area or plots a type
 * the engine does not draw: a chart is drawn whole or not at all, never with
 * one of its plots silently missing. Throws `ChartTooLarge` past the point
 * bounds.
 */
export function readChart(chartSpace: XmlElement, theme: DrawingTheme, owner?: FillOwner): ChartSpecModel | undefined {
  const chart = chartSpace.local === "chart" ? chartSpace : firstChild(chartSpace, "chart");
  if (chart === undefined) return undefined;
  const plotArea = firstChild(chart, "plotArea");
  if (plotArea === undefined) return undefined;
  if (plotArea.children.some((child) => child.local.endsWith("Chart") && GROUP_TAGS[child.local] === undefined)) return undefined;
  const style = Math.min(Math.max(Math.round(num(chartSpace, "style") ?? styleFromAlternate(chartSpace) ?? 2), 1), 48);
  const baseText: TextLook = {
    size: 10, bold: false, italic: false,
    color: style >= 41 ? "#FFFFFF" : themeColor("dk1", { theme }, "#000000"),
    family: theme.fonts.minor.latin || "Calibri",
  };
  // The theme's complex-script slot, or its Arabic script font: what Office sets Arabic chart text in.
  const themeCs = themeFontFamily(theme.fonts.minor, "cs", true);
  if (themeCs !== undefined && themeCs !== "") baseText.csFamily = themeCs;
  const pre: Ctx = { theme, style, text: baseText, owner, budget: { points: 0, read: 0 } };
  const ctx: Ctx = { ...pre, text: readText(firstChild(chartSpace, "txPr"), baseText, pre) };

  const kinds = plotArea.children.filter((child) => GROUP_TAGS[child.local] !== undefined);
  const total = kinds.reduce((sum, el) => sum + childElements(el, "ser").length, 0);
  let offset = 0;
  const groups: Group[] = [];
  for (const el of kinds) {
    groups.push(readGroup(el, GROUP_TAGS[el.local]!, ctx, offset, total));
    offset += childElements(el, "ser").length;
  }
  const axisTitleText: TextLook = { ...ctx.text, bold: true };
  const axes = new Map<number, Axis>();
  for (const el of plotArea.children) {
    if (el.local === "catAx" || el.local === "valAx" || el.local === "dateAx" || el.local === "serAx") {
      const axis = readAxis(el, ctx, axisTitleText);
      axes.set(axis.id, axis);
    }
  }
  // `c:auto` on a category axis: Excel makes it a date axis when the data are dates (numeric categories in a date format).
  for (const axis of axes.values()) {
    if (axis.kind !== "cat" || !axis.auto) continue;
    const first = groups.filter((g) => g.axisIds?.[0] === axis.id && (g.kind === "bar" || g.kind === "line" || g.kind === "area"))
      .flatMap((g) => g.series).find((s) => s.categoryValues !== undefined);
    if (first === undefined) continue;
    if (isDateFormatCode(axis.formatLinked ? first.categoryFormatCode ?? "General" : axis.formatCode)) axis.kind = "date";
  }
  // Generators that omit the axes still get a drawable chart: Office's default pair, shared by every cartesian group.
  const cartesian = groups.filter((g) => !isRadialKind(g.kind));
  if (cartesian.some((g) => g.axisIds === undefined || !axes.has(g.axisIds[0]) || !axes.has(g.axisIds[1]))) {
    const xy = cartesian.some((g) => g.kind === "scatter" || g.kind === "bubble");
    const make = (id: number, cross: number, tag: string, pos: string, grid: boolean): Axis => {
      const parts: Array<{ local: string; attrs: Record<string, string> }> = [
        { local: "axId", attrs: { val: String(id) } }, { local: "axPos", attrs: { val: pos } }, { local: "crossAx", attrs: { val: String(cross) } },
        ...(grid ? [{ local: "majorGridlines", attrs: {} }] : []),
      ];
      const children: XmlElement[] = parts.map((c) => ({ name: c.local, local: c.local, attrs: c.attrs, children: [], text: "" }));
      return readAxis({ name: tag, local: tag, attrs: {}, children, text: "" }, ctx, axisTitleText);
    };
    const horizontalBars = cartesian.every((g) => g.kind === "bar" && g.direction === "bar");
    axes.set(-1, make(-1, -2, xy ? "valAx" : "catAx", horizontalBars ? "l" : "b", false));
    axes.set(-2, make(-2, -1, "valAx", horizontalBars ? "b" : "l", true));
    for (const g of cartesian) if (g.axisIds === undefined || !axes.has(g.axisIds[0]) || !axes.has(g.axisIds[1])) g.axisIds = [-1, -2];
  }
  const dark = darkStyle(ctx);
  const spec: ChartSpecModel = {
    groups,
    axes: [...axes.values()],
    plot: { look: readLook(firstChild(plotArea, "spPr"), ctx, style >= 33 && style <= 40 ? { fill: tx1(ctx, 1, 0.85) } : {}) },
    look: readLook(firstChild(chartSpace, "spPr"), ctx, dark
      ? { fill: tx1(ctx, 1, 0.15) }
      : { fill: themeColor("lt1", ctx, "#FFFFFF"), line: { color: tx1(ctx, 1, 0.5), width: 0.75 } }),
    roundedCorners: bool(chartSpace, "roundedCorners", false),
    blanks: choice(chart, "dispBlanksAs", ["gap", "zero", "span"], "gap"),
    date1904: bool(chartSpace, "date1904", false),
    text: ctx.text,
  };
  const plotLayout = readLayout(plotArea);
  if (plotLayout !== undefined) spec.plot.layout = plotLayout;

  // Title: an explicit c:title, or Office's automatic one (the lone series' name).
  const titleEl = firstChild(chart, "title");
  const autoDeleted = bool(chart, "autoTitleDeleted", false);
  const allSeries = groups.flatMap((g) => g.series);
  const titleBase: TextLook = { ...ctx.text, size: ctx.text.size * 1.8, bold: true };
  if (titleEl !== undefined) {
    const title = readTitle(titleEl, ctx, titleBase, [allSeries.length === 1 ? allSeries[0]!.name : "Chart Title"]);
    if (title !== undefined) spec.title = title;
  } else if (!autoDeleted && allSeries.length === 1 && groups.length === 1 && isRadialKind(groups[0]!.kind)) {
    spec.title = { lines: [allSeries[0]!.name], text: titleBase, look: {}, overlay: false };
  }

  const legendEl = firstChild(chart, "legend");
  if (legendEl !== undefined) {
    const pos = val(legendEl, "legendPos");
    const legend: Legend = {
      position: pos === "b" || pos === "t" || pos === "l" || pos === "tr" ? pos : "r",
      overlay: bool(legendEl, "overlay", false),
      text: readText(firstChild(legendEl, "txPr"), ctx.text, ctx),
      look: readLook(firstChild(legendEl, "spPr"), ctx, {}),
      deleted: [],
    };
    for (const entry of childElements(legendEl, "legendEntry")) {
      const idx = num(entry, "idx");
      if (idx !== undefined && bool(entry, "delete", false) && !legend.deleted.includes(idx)) legend.deleted.push(idx);
    }
    const layout = readLayout(legendEl);
    if (layout !== undefined) legend.layout = layout;
    spec.legend = legend;
  }
  return spec;
}

/** `mc:AlternateContent/mc:Choice/c14:style` carries the style when `c:style` is absent (Office 2010+ writes it there). */
function styleFromAlternate(chartSpace: XmlElement): number | undefined {
  for (const child of chartSpace.children) {
    if (child.local !== "AlternateContent") continue;
    for (const branch of child.children) {
      const style = firstChild(branch, "style");
      const n = style === undefined ? NaN : Number(attr(style, "val"));
      if (Number.isFinite(n)) return n > 100 ? n - 100 : n;
    }
  }
  return undefined;
}
