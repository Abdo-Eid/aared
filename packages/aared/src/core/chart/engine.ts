/**
 * Layout and painting, on the page: `ChartSpecModel` + frame size (points) ->
 * primitives. Pure (no DOM), so it stays testable headless; the page reaches it
 * through `draw.ts`, and only the page loads that.
 *
 * The chart is laid out the way Office does it, outside in:
 *   chart area -> title band -> legend band -> outer plot area -> axis
 *   decorations (titles, tick labels, tick marks) -> inner plot area,
 * iterating the last step because axis labels both depend on and consume the
 * inner size (more room -> more value ticks; narrower slots -> wrapped or
 * rotated category labels). Every Office-specific distance lives in `K`, the
 * one calibration table, fitted to PowerPoint's own PDF output
 * (`tools/chart-oracle/`).
 */

import { formatNumberValue } from "../xlsx/numberfmt.js";
import { textBox } from "./bounds.js";
import { ascent, blockSize, fitText, isRtlText, lineHeight, textWidth, wrapFit, wrapText } from "./metrics.js";
import { equation, errorAmounts, fitTrend, movingAverage, rSquaredLabel } from "./trend.js";
import { categoryFraction, categorySlot, clean, majorTicks, minorTicks, timeScale, timeSlot, timeTicks, valueFraction, valueScale } from "./scale.js";
import type { CategoryScale, Scale, TimeTick, ValueScale } from "./scale.js";
import { Scene, fmt } from "./scene.js";
import type { ChartPrimitive } from "./scene.js";
import type { Axis, ChartSpecModel, DataLabels, Group, LabelPosition, Line, Look, ManualLayout, Marker, Series, TextLook, Title, Trendline, TreeNode } from "./spec.js";
import { isRadialKind } from "./spec.js";

function axisById(spec: ChartSpecModel, id: number): Axis | undefined {
  return spec.axes.find((axis) => axis.id === id);
}

/** Office's layout constants, in points unless noted. Fitted against the PDFs of `tools/chart-oracle/`. */
export const K = {
  /** Chart area edge → any element. */
  pad: 7,
  /** Title box bottom → next element. */
  titleGap: 3.3,
  /** Title box inset above and below its text, in title font sizes. */
  titleInset: 0.15,
  /** Office insets the plot area's right side this much more than the others. */
  plotRightInset: 4,
  /** Vertical legend: extra space between entries, on top of the line height. */
  legendRowGap: 5.9,
  /** Legend band ↔ plot area. */
  legendGap: 7,
  /** Legend box inner padding (horizontal, vertical). */
  legendPadX: 3,
  legendPadY: 2,
  /** Legend: key width for fill keys, as a fraction of the font size. */
  legendKey: 0.6,
  /** Legend: key length for line keys, as a fraction of the font size. */
  legendLineKey: 1.63,
  /** Legend: key → text gap, fraction of font size. */
  legendKeyGap: 0.35,
  /** Legend: gap between horizontal entries, fraction of font size. */
  legendEntryGap: 0.78,
  /** Major tick mark length. */
  tick: 3.75,
  minorTick: 2,
  /** Axis line (or tick end) → tick label, in label font sizes: beside a vertical axis, below/above a horizontal one. */
  labelGapBeside: 0.92,
  labelGapBelow: 0.73,
  /** Radar: spoke end → category label. */
  labelGap: 3.5,
  /** Axis tick labels → axis title. */
  axisTitleGap: 4,
  /** Value axis: minimum spacing between major ticks, in label font sizes. */
  valueTickSpacing: 1.5,
  /** Value axis: most major intervals Office ever picks automatically. */
  maxValueIntervals: 10,
  /** Horizontal value axis: minimum gap between neighbouring labels. */
  valueLabelGap: 6,
  /** Horizontal value axis: minimum spacing between major ticks, in label line heights. */
  valueTickSpacingH: 3,
  /** Data label gap from the bar end / point. */
  dataLabelGap: 3,
  /** Pie: data label gap outside the circle. */
  pieLabelGap: 6,
  /** Pie of pie / bar of pie: room left on each side of the plot box (Excel: a 466 pt plot box holds 458 pt of plots). */
  ofPieInset: 4,
  /** Pie of pie / bar of pie: the gap between the main pie and the second plot is `gapWidth` percent of this many main radii (measured: 100% leaves one radius before a second pie, half a radius before a bar). */
  ofPieGapPie: 1,
  ofPieGapBar: 0.5,
  /** Bar of pie: the stacked bar is this fraction as wide as it is tall. */
  ofPieBarAspect: 0.5,
  /**
   * Pie: a label wraps at this fraction of the plot's width, inside the pie or out. PowerPoint kept an
   * 83pt name whole and wrapped 120pt and 165pt ones near 105pt in a 562pt plot (g-pie-bestfit-long).
   */
  pieLabelMax: 0.2,
  /** Pie: a label inside its slice wraps to the slice's width at this fraction of the radius (g-pie-bestfit-long, ar-pie-labels). */
  pieInsideWrapAt: 0.85,
  /** Pie: smallest radius labels may shrink the pie to, as a fraction of the unshrunk radius (assumption). */
  pieMinRadius: 0.4,
  /** Pie: minimum vertical gap between spread labels. */
  pieLabelSpacing: 1,
  /** Pie: a label moved less than this from its natural spot gets no leader line. */
  pieLeaderMin: 4,
  /** Pie: leader line end to the label's edge. */
  pieLeaderGap: 2,
  /** Pie: the level stroke a moved label's leader line ends in. */
  pieLeaderElbow: 6,
  /** Legend: widest a side legend may grow, as a fraction of the chart width (assumption). */
  legendMaxWidth: 1 / 3,
  /** Legend: tallest a top or bottom legend may grow, as a fraction of the chart height (assumption). */
  legendMaxHeight: 1 / 3,
  /**
   * Legend that wraps (top or bottom): Office lays it on a grid. A cell is the widest entry plus
   * this many font sizes; rows sit this many font sizes further apart than lines of text. Fitted
   * on one PowerPoint case (20 series at 12pt: 57pt columns, 20.4pt rows); thin evidence.
   */
  legendGridGap: 0.14,
  legendGridRowGap: 0.48,
  /** Legend: most lines an entry wraps to in a side legend. */
  legendEntryLines: 3,
  /** Error bars: half the width of an end cap. */
  errCap: 3.5,
  /** Trendline: samples along a curved fit, at most (a straight fit is two points). */
  trendSamples: 120,
  /** Long lines: the column width a run is thinned to (half a point, under one device pixel at 96 dpi). */
  thinColumn: 0.5,
  /** Bubble chart: largest bubble diameter as a fraction of the plot's smaller side. */
  bubbleMax: 0.25,
  /** Category labels: widest a wrapped label may grow, as a fraction of the chart width (vertical axes). */
  maxLabelWidth: 0.4,
  /** Grouped categories: space a group label row adds to its text, above and below together. Not yet fitted to Office. */
  groupLabelPad: 4,
  /** Treemap: a parent's banner, in label font sizes (Excel: 18.7pt for 9pt labels). */
  treemapBanner: 2,
  /** Date axis: labels lie flat only with this many line heights between them, else turn 45 degrees (g-date-years). */
  dateLabelGap: 1,
};

interface Box { x: number; y: number; w: number; h: number }

// ------------------------------------------------------------------ helpers

function formatValue(v: number, code: string, spec: ChartSpecModel): string {
  try {
    return formatNumberValue(clean(v), code === "" ? "General" : code, spec.date1904);
  } catch {
    return String(clean(v));
  }
}

function pointLook(series: Series, i: number, value: number | null): Look {
  const point = series.points[i];
  let look = point?.look ?? series.look;
  if ((point?.invertIfNegative ?? series.invertIfNegative) && value !== null && value < 0 && typeof look.fill === "string") {
    look = { ...look, fill: "#FFFFFF" };
  }
  return look;
}

function deg(a: number): number {
  return (a * Math.PI) / 180;
}

/** Text block anchored at a point, horizontal. `vAlign` places the block relative to y. */
function drawLines(scene: Scene, lines: readonly string[], x: number, y: number, look: TextLook, anchor: "start" | "middle" | "end", vAlign: "top" | "middle" | "bottom"): void {
  const lh = lineHeight(look);
  const total = lh * lines.length;
  let top = vAlign === "top" ? y : vAlign === "middle" ? y - total / 2 : y - total;
  const asc = ascent(look);
  // The lines are one paragraph: its first strong character sets every line's direction.
  const rtl = isRtlText(lines.join("\n"));
  for (const line of lines) {
    scene.text(x, top + asc, line, look, anchor, 0, rtl);
    top += lh;
  }
}

/** Rotated text block: rotation in degrees (clockwise), anchored at the block center. */
function drawRotated(scene: Scene, lines: readonly string[], cx: number, cy: number, look: TextLook, rotation: number): void {
  const lh = lineHeight(look);
  const asc = ascent(look);
  const total = lh * lines.length;
  const rad = deg(rotation);
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const rtl = isRtlText(lines.join("\n"));
  lines.forEach((line, i) => {
    // Baseline point in the unrotated block, relative to its center, then rotated.
    const dy = -total / 2 + i * lh + asc;
    scene.text(cx - dy * sin, cy + dy * cos, line, look, "middle", rotation, rtl);
  });
}

// ------------------------------------------------------------------ axes

interface AxisPlan {
  axis: Axis;
  cross?: AxisPlan;
  role: "cat" | "val";
  horizontal: boolean;
  groups: Group[];
  scale: Scale;
  labels: Array<{ f: number; lines: string[] }>;
  labelRotation: number;
  labelLook: TextLook;
  /** Perpendicular room the labels need (including gap and outward ticks), points. */
  thickness: number;
  /** Room needed past each end of the axis, points. */
  overhangStart: number;
  overhangEnd: number;
  /** Where the axis line sits along the cross axis, 0..1 (before reversal). */
  crossF: number;
  /** +1: labels toward the cross axis' max side; -1: toward its min. */
  side: 1 | -1;
  /** Where labels hang: at the axis line, or pinned to an inner-plot edge. */
  labelAt: "line" | "low" | "high";
  visible: boolean;
  categories: string[];
  /** Category label widths and longest-word widths, measured once per chart. */
  widths?: number[];
  longestWord?: number[];
  /** The label interval actually used; tick marks follow it unless the file sets its own. */
  labelSkipAuto: number;
  /** Date axis: where the major ticks fall, and the dates the axis' points carry when a series has none of its own. */
  timeTicks?: TimeTick[];
  dates?: Array<number | null>;
  /** Grouped categories: the outer label rows (innermost first) and where the first one starts, from the axis line. */
  levels: Array<{ height: number; spans: Array<{ from: number; to: number; lines: string[] }> }>;
  levelStart: number;
}

function axisVisible(plan: AxisPlan): boolean {
  return !plan.axis.deleted;
}

function isPercent(groups: Group[]): boolean {
  return groups.some((g) => g.grouping === "percentStacked");
}

/** Data extent of a value axis across its groups (stacking aware). */
function valueExtent(groups: Group[], xAxis: boolean): [number, number] {
  let lo = Infinity;
  let hi = -Infinity;
  const see = (v: number): void => {
    if (v < lo) lo = v;
    if (v > hi) hi = v;
  };
  // The axis reaches the error bars' ends too: Office widens it for them (the oracle's g-err-* cases).
  const seeErrorBars = (s: Series, dir: "x" | "y", values: ReadonlyArray<number | null>): void => {
    for (const eb of s.errBars ?? []) {
      if (eb.dir !== dir) continue;
      errorAmounts(eb, values).forEach((amt, i) => {
        const v = values[i];
        if (amt === undefined || v === null || v === undefined) return;
        const centre = amt.centre ?? v;
        see(centre + amt.plus);
        see(centre - amt.minus);
      });
    }
  };
  for (const g of groups) {
    if (xAxis) {
      for (const s of g.series) {
        const xs = s.categoryValues ?? s.values.map((_, i) => i + 1);
        for (const v of xs) if (v !== null) see(v);
        seeErrorBars(s, "x", xs);
      }
      continue;
    }
    if (g.grouping !== "percentStacked" && g.grouping !== "stacked") for (const s of g.series) seeErrorBars(s, "y", s.values);
    const stacked = g.grouping === "stacked" || g.grouping === "percentStacked";
    if (stacked && g.kind !== "scatter" && g.kind !== "bubble") {
      const n = Math.max(0, ...g.series.map((s) => s.values.length));
      for (let i = 0; i < n; i++) {
        let pos = 0;
        let neg = 0;
        for (const s of g.series) {
          const v = s.values[i] ?? 0;
          if (v >= 0) pos += v;
          else neg += v;
        }
        if (g.grouping === "percentStacked") {
          const total = pos - neg;
          if (total > 0) { see(pos / total); see(neg / total); }
        } else {
          see(pos); see(neg);
          if (g.kind === "area" || g.kind === "line") see(0);
        }
      }
    } else {
      for (const s of g.series) {
        if (s.spans !== undefined) {
          // A waterfall's bars float: the data's extent is where they reach.
          for (const span of s.spans) if (span !== null) { see(span[0]); see(span[1]); }
        } else if (s.boxes !== undefined) {
          for (const b of s.boxes) {
            see(b.low); see(b.high); see(b.q1); see(b.q3);
            for (const o of b.outliers) see(o);
          }
        } else if (g.kind === "funnel") {
          see(0);
          for (const v of s.values) if (v !== null) see(v);
        } else {
          for (const v of s.values) if (v !== null) see(v);
        }
      }
    }
  }
  if (lo === Infinity) return [0, 1];
  return [lo, hi];
}

/** The series whose categories name an axis' points (the first that has any), and how many points the axis spans. */
function categorySource(groups: Group[]): { best: Series | undefined; count: number } {
  let best: Series | undefined;
  let count = 0;
  for (const g of groups) {
    for (const s of g.series) {
      const n = Math.max(s.values.length, s.categories.length, s.categoryValues?.length ?? 0);
      if (n > count) count = n;
      if (best === undefined && (s.categories.length > 0 || s.categoryValues !== undefined)) best = s;
    }
  }
  return { best, count };
}

/**
 * The number format a date axis shows its dates in: the axis' own, else its
 * source's. Not a date format, Office shows the serial numbers themselves
 * ("45292"; the oracle's g-date-general), so no date format is put in.
 */
function dateFormat(axis: Pick<Axis, "formatLinked" | "formatCode">, best: Series | undefined): string {
  return axis.formatLinked ? best?.categoryFormatCode ?? "General" : axis.formatCode;
}

function categoryLabels(groups: Group[], axis: Pick<Axis, "kind" | "formatLinked" | "formatCode">, spec: ChartSpecModel): string[] {
  const { best, count } = categorySource(groups);
  const out: string[] = [];
  for (let i = 0; i < count; i++) {
    if (best?.categoryValues !== undefined) {
      const v = best.categoryValues[i];
      const code = axis.kind === "date" ? dateFormat(axis, best) : axis.formatLinked ? best.categoryFormatCode ?? "General" : axis.formatCode;
      out.push(v === null || v === undefined ? "" : formatValue(v, code, spec));
    } else {
      out.push(best?.categories[i] ?? String(i + 1));
    }
  }
  return out;
}

function planAxes(spec: ChartSpecModel): AxisPlan[] {
  const plans = new Map<number, AxisPlan>();
  for (const g of spec.groups) {
    if (g.axisIds === undefined) continue;
    const [catId, valId] = g.axisIds;
    const catAxis = axisById(spec, catId);
    const valAxis = axisById(spec, valId);
    if (catAxis === undefined || valAxis === undefined) continue;
    const xy = g.kind === "scatter" || g.kind === "bubble";
    const horizontalCat = !((g.kind === "bar" || g.kind === "funnel") && g.direction === "bar");
    for (const [axis, role, horizontal] of [[catAxis, xy ? "val" : "cat", horizontalCat], [valAxis, "val", !horizontalCat]] as const) {
      let plan = plans.get(axis.id);
      if (plan === undefined) {
        plan = {
          axis, role, horizontal, groups: [], scale: { kind: "category", count: 0, between: true },
          labels: [], labelRotation: 0, labelLook: axis.text, thickness: 0, overhangStart: 0, overhangEnd: 0,
          crossF: 0, side: -1, labelAt: "line", visible: !axis.deleted, categories: [], labelSkipAuto: 1,
          levels: [], levelStart: 0,
        };
        plans.set(axis.id, plan);
      }
      if (!plan.groups.includes(g)) plan.groups.push(g);
    }
  }
  for (const plan of plans.values()) {
    const cross = plans.get(plan.axis.crossAxisId);
    if (cross !== undefined) plan.cross = cross;
  }
  return [...plans.values()];
}

/** Value labels for a value-axis plan. */
function valueLabels(plan: AxisPlan, scale: ValueScale, spec: ChartSpecModel): Array<{ f: number; lines: string[] }> {
  const axis = plan.axis;
  const percent = isPercent(plan.groups);
  const linkedCode = plan.role === "val" && plan.groups[0]?.kind !== "scatter" && plan.groups[0]?.kind !== "bubble"
    ? plan.groups[0]?.series[0]?.formatCode
    : plan.horizontal ? plan.groups[0]?.series[0]?.categoryFormatCode : plan.groups[0]?.series[0]?.formatCode;
  const code = !axis.formatLinked ? axis.formatCode : percent ? "0%" : linkedCode ?? "General";
  return majorTicks(scale).map((v) => ({ f: valueFraction(scale, v), lines: [formatValue(v / axis.displayUnit, code, spec)] }));
}

/** Pixel position along an axis for fraction f. */
function along(plan: AxisPlan, inner: Box, f: number): number {
  const g = plan.axis.reversed ? 1 - f : f;
  return plan.horizontal ? inner.x + g * inner.w : inner.y + inner.h - g * inner.h;
}

/** Fraction along `plan.cross` where `plan` crosses it. */
function crossFraction(plan: AxisPlan): number {
  const cross = plan.cross;
  if (cross === undefined) return 0;
  const a = plan.axis;
  if (cross.scale.kind === "category") {
    if (a.crosses === "max") return 1;
    if (a.crossesAt !== undefined && cross.scale.count > 0) return Math.min(Math.max((a.crossesAt - 1) / cross.scale.count, 0), 1);
    return 0;
  }
  const s = cross.scale;
  if (a.crosses === "max") return 1;
  if (a.crosses === "min") return 0;
  const at = a.crossesAt ?? 0;
  if (s.log !== undefined) return a.crossesAt !== undefined ? Math.min(Math.max(valueFraction(s, at), 0), 1) : 0;
  return Math.min(Math.max(valueFraction(s, at), 0), 1);
}

/** Pixel coordinate (perpendicular) of an axis line. */
function axisLinePos(plan: AxisPlan, inner: Box): number {
  const cross = plan.cross;
  if (cross === undefined) return plan.horizontal ? inner.y + inner.h : inner.x;
  return along(cross, inner, plan.crossF);
}

/** Perpendicular pixel direction of the label side: +1 = increasing x/y. */
function labelDirection(plan: AxisPlan): 1 | -1 {
  const cross = plan.cross;
  // The cross axis' max end in pixels: for a vertical cross axis increasing value goes up (−y).
  const crossMaxDir = cross === undefined ? 1 : cross.horizontal ? (cross.axis.reversed ? -1 : 1) : (cross.axis.reversed ? 1 : -1);
  return (plan.side === 1 ? crossMaxDir : -crossMaxDir) as 1 | -1;
}

/** Where the labels of `plan` start (perpendicular coordinate) and grow from. */
function labelOrigin(plan: AxisPlan, inner: Box): number {
  if (plan.labelAt === "line") return axisLinePos(plan, inner);
  const cross = plan.cross;
  if (cross === undefined) return axisLinePos(plan, inner);
  return along(cross, inner, plan.labelAt === "low" ? 0 : 1);
}

/** Fit labels along an axis of `length` points; sets labels, rotation, thickness, overhang. */
function fitAxis(plan: AxisPlan, length: number, spec: ChartSpecModel, chartWidth: number): void {
  const axis = plan.axis;
  const look = axis.text;
  plan.labelLook = look;
  plan.levels = [];
  plan.levelStart = 0;
  delete plan.timeTicks;
  const lh = lineHeight(look);
  if (plan.role === "val") {
    const [lo, hi] = valueExtent(plan.groups, plan.horizontal && (plan.groups[0]?.kind === "scatter" || plan.groups[0]?.kind === "bubble"));
    const percent = isPercent(plan.groups);
    // How many intervals fit: vertical axes by line height, horizontal by label width (estimated from the range ends).
    let maxIntervals: number;
    if (!plan.horizontal) {
      maxIntervals = Math.max(1, Math.floor(length / (look.size * K.valueTickSpacing)));
    } else {
      const probe = valueScale(axis, lo, hi, 10, percent);
      const widest = Math.max(...valueLabels(plan, probe, spec).map((l) => textWidth(l.lines[0]!, look)));
      maxIntervals = Math.max(1, Math.floor(length / Math.max(widest + K.valueLabelGap, lh * K.valueTickSpacingH)));
    }
    const scale = valueScale(axis, lo, hi, Math.min(maxIntervals, K.maxValueIntervals), percent);
    plan.scale = scale;
    plan.labels = axis.labelPosition === "none" ? [] : valueLabels(plan, scale, spec);
    plan.labelRotation = look.rotation ?? 0;
  } else {
    const cats = categoryLabels(plan.groups, axis, spec);
    plan.categories = cats;
    const crossVal = plan.cross?.axis;
    const between = plan.groups.some((g) => g.kind === "bar") || (crossVal?.between ?? true);
    const { best } = categorySource(plan.groups);
    // A date axis places points by date; text categories (or dates the axis cannot span) go by index.
    const timed = axis.kind === "date" && best?.categoryValues !== undefined
      ? timeScale(plan.groups.flatMap((g) => g.series.flatMap((s) => s.categoryValues ?? [])), axis, between, spec.date1904)
      : undefined;
    if (timed !== undefined) {
      fitTimeLabels(plan, timed, best!, length, spec);
      measureAxis(plan, length);
      return;
    }
    const scale: CategoryScale = { kind: "category", count: cats.length, between };
    plan.scale = scale;
    const slot = length * categorySlot(scale);
    let skip = axis.labelSkip ?? 1;
    let rotation = look.rotation ?? 0;
    let lines: string[][] | undefined;
    if (axis.labelPosition !== "none" && cats.length > 0) {
      const widths = plan.widths ??= cats.map((c) => textWidth(c, look));
      let widest = 0;
      for (const w of widths) if (w > widest) widest = w;
      if (plan.horizontal) {
        if (look.rotation === undefined && axis.labelSkip === undefined) {
          if (widest > slot) {
            // Wrap at spaces first, when every word fits the slot and no label needs more than three lines;
            // otherwise rotate 45° and thin the labels out until they clear each other.
            const words = plan.longestWord ??= cats.map((c) => Math.max(...c.split(" ").map((w) => textWidth(w, look))));
            let wordsFit = cats.some((c) => c.includes(" "));
            for (let i = 0; wordsFit && i < words.length; i++) if (words[i]! > slot + 0.01) wordsFit = false;
            const wrapped = wordsFit ? cats.map((c) => wrapText(c, look, slot)) : undefined;
            if (wrapped !== undefined && wrapped.every((ls) => ls.length <= 3)) {
              lines = wrapped;
            } else if (lh / Math.SQRT1_2 <= slot) {
              rotation = -45;
            } else {
              // Too tight for 45 degrees: Office stands the labels upright and thins them only when even that overlaps.
              rotation = -90;
              skip = Math.max(1, Math.ceil(lh / Math.max(slot, 0.01)));
            }
          }
        } else if (axis.labelSkip === undefined) {
          const needed = rotation === 0 ? widest : lh / Math.max(Math.abs(Math.sin(deg(rotation))), 0.1);
          skip = Math.max(1, Math.ceil(needed / Math.max(slot, 0.01)));
        }
      } else {
        // Vertical category axis: horizontal labels, wrapped to a share of the chart width; thin slots skip.
        const limit = chartWidth * K.maxLabelWidth;
        if (widest > limit) lines = cats.map((c, i) => (widths[i]! > limit ? wrapText(c, look, limit) : [c]));
        if (axis.labelSkip === undefined) {
          let most = 1;
          if (lines !== undefined) for (const l of lines) if (l.length > most) most = l.length;
          skip = Math.max(1, Math.ceil((most * lh) / Math.max(slot, 0.01)));
        }
      }
    }
    plan.labelSkipAuto = skip;
    plan.labelRotation = rotation;
    plan.labels = [];
    if (axis.labelPosition !== "none") {
      for (let i = 0; i < cats.length; i += skip) plan.labels.push({ f: categoryFraction(scale, i), lines: lines?.[i] ?? [cats[i]!] });
    }
    if (axis.multiLevel && axis.labelPosition !== "none" && best?.categoryGroups !== undefined) {
      plan.levels = groupRows(best.categoryGroups, cats.length, scale, length, look);
    }
  }
  measureAxis(plan, length);
  if (plan.levels.length > 0) {
    // The outer rows hang below the inner labels, so the axis needs their room as well.
    plan.levelStart = plan.thickness;
    for (const row of plan.levels) plan.thickness += row.height;
  }
}

/** The outer rows of grouped categories: each group's label, wrapped to its span, and the row's height. */
function groupRows(groups: NonNullable<Series["categoryGroups"]>, count: number, scale: CategoryScale, length: number, look: TextLook): AxisPlan["levels"] {
  const lh = lineHeight(look);
  const slot = length * categorySlot(scale);
  const rows: AxisPlan["levels"] = [];
  for (const level of groups) {
    const spans: AxisPlan["levels"][number]["spans"] = [];
    let most = 1;
    level.forEach((group, i) => {
      const from = group.start;
      const to = Math.min(level[i + 1]?.start ?? count, count);
      if (!(to > from)) return;
      const lines = wrapText(group.text, look, Math.max(slot * (to - from) - 2, 1));
      most = Math.max(most, lines.length);
      spans.push({ from, to, lines });
    });
    if (spans.length > 0) rows.push({ height: most * lh + K.groupLabelPad, spans });
  }
  return rows;
}

/**
 * A date axis' labels: major ticks by the file's unit, or by the finest step
 * whose labels clear each other. Office labels every base unit while it can,
 * lying flat when a line height is left between labels, turned 45 degrees
 * when a line height (along the slant) fits, and steps up a unit only after
 * that; it never stands date labels upright (the oracle's g-date-* cases).
 */
function fitTimeLabels(plan: AxisPlan, scale: CategoryScale, best: Series, length: number, spec: ChartSpecModel): void {
  const axis = plan.axis;
  const look = axis.text;
  const time = scale.time!;
  plan.scale = scale;
  const code = dateFormat(axis, best);
  // Month names differ in width: probe the widest across the span.
  const perSlot = time.base === "days" ? 1 : time.base === "months" ? 30.44 : 365.25;
  let widest = 0;
  for (let k = 0; k <= 12; k++) widest = Math.max(widest, textWidth(formatValue(time.start + (perSlot * scale.count * k) / 12, code, spec), look));
  const lh = lineHeight(look);
  const slotLength = length * categorySlot(scale);
  const pitch = (ticks: TimeTick[]): number => {
    let gap = Infinity;
    for (let i = 1; i < ticks.length; i++) gap = Math.min(gap, ticks[i]!.slot - ticks[i - 1]!.slot);
    return gap * slotLength;
  };
  const flat = (ticks: TimeTick[]): boolean => ticks.length <= 1 || widest + lh * K.dateLabelGap <= pitch(ticks);
  const slanted = (ticks: TimeTick[]): boolean => ticks.length <= 1 || lh / Math.SQRT1_2 <= pitch(ticks);
  const turnable = plan.horizontal && look.rotation === undefined;
  const ticks = timeTicks(scale, axis, plan.horizontal
    ? (t) => flat(t) || (turnable && slanted(t))
    : (t) => t.length <= 1 || look.size * K.valueTickSpacing <= pitch(t));
  plan.timeTicks = ticks;
  plan.dates = best.categoryValues ?? [];
  plan.labelSkipAuto = 1;
  let rotation = look.rotation ?? 0;
  if (turnable && !flat(ticks)) rotation = -45;
  plan.labelRotation = rotation;
  plan.labels = axis.labelPosition === "none" ? [] : ticks.map((t) => ({ f: categoryFraction(scale, t.slot), lines: [formatValue(t.serial, code, spec)] }));
}

/** Axis line (or tick end) to the labels, by the side the labels sit on. */
function labelGap(plan: AxisPlan): number {
  return plan.labelLook.size * (plan.horizontal ? K.labelGapBelow : K.labelGapBeside);
}

/**
 * Perpendicular thickness of an axis' labels, and how far they stick out past
 * each end of an axis `length` points long. A label centred inside the axis
 * needs no room past its ends; only what crosses an end counts. Along the
 * axis a line of text occupies its font size, not its line height: Office
 * lets the leading of the end labels overlap the frame padding.
 */
function measureAxis(plan: AxisPlan, length: number): void {
  const look = plan.labelLook;
  const leading = lineHeight(look) - look.size;
  let thick = 0;
  let overStart = 0;
  let overEnd = 0;
  const rot = deg(plan.labelRotation);
  const cos = Math.abs(Math.cos(rot));
  const sin = Math.abs(Math.sin(rot));
  for (const label of plan.labels) {
    const { width, height } = blockSize(label.lines, look);
    const w = width * cos + height * sin;
    const h = width * sin + height * cos;
    // Where the label's centre sits along the axis, from its start.
    const at = (plan.axis.reversed ? 1 - label.f : label.f) * length;
    if (plan.horizontal) {
      thick = Math.max(thick, h);
      // Rotated labels end at their tick and hang toward the start by their full projected width.
      const back = plan.labelRotation < 0 ? width * cos : plan.labelRotation === 0 ? w / 2 : 0;
      const fore = plan.labelRotation === 0 ? w / 2 : 0;
      overStart = Math.max(overStart, back - at);
      overEnd = Math.max(overEnd, fore - (length - at));
    } else {
      thick = Math.max(thick, w);
      const half = (h - leading) / 2;
      overStart = Math.max(overStart, half - at);
      overEnd = Math.max(overEnd, half - (length - at));
    }
  }
  // An outward tick sits inside the label gap: Office does not push the labels past it.
  const ticks = plan.axis.majorTick === "out" || plan.axis.majorTick === "cross" ? K.tick : 0;
  plan.thickness = plan.labels.length === 0 ? ticks : thick + Math.max(labelGap(plan), ticks);
  plan.overhangStart = plan.labels.length === 0 ? 0 : overStart;
  plan.overhangEnd = plan.labels.length === 0 ? 0 : overEnd;
}

function titleBlock(title: Title, maxWidth: number): { lines: string[]; width: number; height: number } {
  const look = title.text;
  const vertical = title.text.rotation !== undefined && Math.abs(title.text.rotation) === 90;
  // A word wider than the room is cut, not left to spill.
  const lines = title.lines.flatMap((l) => wrapFit(l, look, Math.max(maxWidth, 1), 99));
  const size = blockSize(lines, look);
  return vertical ? { lines, width: size.height, height: size.width } : { lines, ...size };
}

// ------------------------------------------------------------------ legend

interface LegendEntry {
  label: string;
  kind: "fill" | "line";
  look: Look;
  marker?: Marker;
}

function legendEntries(spec: ChartSpecModel): LegendEntry[] {
  const out: LegendEntry[] = [];
  // Office 2016 charts name their own entries (a waterfall's Increase, Decrease and Total; a treemap's branches).
  const named = spec.groups.find((g) => g.legendEntries !== undefined);
  if (named !== undefined) return named.legendEntries!.map((e) => ({ label: e.label, kind: "fill", look: e.look }));
  const single = spec.groups.length === 1 ? spec.groups[0]! : undefined;
  if (single !== undefined && single.varyColors && single.series.length >= 1) {
    const s = single.series[0]!;
    const n = Math.max(s.values.length, s.categories.length);
    for (let i = 0; i < n; i++) {
      const look = s.points[i]?.look ?? s.look;
      out.push({ label: s.categories[i] ?? String(i + 1), kind: single.kind === "line" || single.kind === "scatter" ? "line" : "fill", look });
    }
    return out;
  }
  for (const g of spec.groups) {
    for (const s of g.series) {
      const lineKind = g.kind === "line" || g.kind === "stock" || g.kind === "scatter" || (g.kind === "radar" && g.radarStyle !== "filled");
      const entry: LegendEntry = { label: s.name, kind: lineKind ? "line" : "fill", look: s.look };
      if (lineKind && s.marker.symbol !== "none") entry.marker = s.marker;
      out.push(entry);
    }
  }
  // Trendlines follow every series, each with the name Office gives it.
  for (const g of spec.groups) {
    for (const s of g.series) {
      for (const t of s.trendlines ?? []) {
        out.push({ label: t.name ?? `${TREND_NAMES[t.type]!.replace("#", String(t.period))} (${s.name})`, kind: "line", look: { line: t.line } });
      }
    }
  }
  return out;
}

const TREND_NAMES: Record<Trendline["type"], string> = {
  linear: "Linear", exp: "Expon.", log: "Log.", poly: "Poly.", power: "Power", movingAvg: "# per. Mov. Avg.",
};

/** `span`: the width the row is centred as, when it belongs to a grid (its columns line up with the other rows'). */
interface LegendRow { items: Array<{ entry: LegendEntry; x: number; width: number; lines: string[] }>; h: number; span?: number }
interface LegendLayout { box: Box; rows: LegendRow[]; vertical: boolean }

function layoutLegend(spec: ChartSpecModel, area: Box, chartW: number, chartH: number): LegendLayout | undefined {
  const legend = spec.legend;
  if (legend === undefined) return undefined;
  let entries = legendEntries(spec).filter((_, i) => !legend.deleted.includes(i));
  if (entries.length === 0) return undefined;
  const look = legend.text;
  const fs = look.size;
  const vertical = legend.position === "r" || legend.position === "l" || legend.position === "tr";
  const reverse = vertical && spec.groups.some((g) => (g.kind === "bar" && (g.direction === "bar" || g.grouping !== "clustered")) || (g.kind === "area" && g.grouping !== "standard"));
  if (reverse && !(spec.groups.length === 1 && spec.groups[0]!.varyColors) && !spec.groups.some((g) => g.legendEntries !== undefined)) entries = [...entries].reverse();
  const keyW = (e: LegendEntry): number => (e.kind === "line" ? K.legendLineKey : K.legendKey) * fs;
  const lh = lineHeight(look);
  // How big the legend may grow (assumptions): a side legend a third of the chart wide, a top or bottom
  // one a third of its height; a manual layout's own size wins. What does not fit is dropped.
  const manual = legend.layout;
  const capW = manual?.w !== undefined ? manual.w * chartW : vertical ? chartW * K.legendMaxWidth : area.w;
  const capH = manual?.h !== undefined ? manual.h * chartH : vertical ? area.h : chartH * K.legendMaxHeight;
  const innerW = Math.max(capW - 2 * K.legendPadX, 0);
  const innerH = Math.max(capH - 2 * K.legendPadY, 0);
  const rows: LegendRow[] = [];
  let width = 0;
  let used = 0;
  if (vertical) {
    for (const e of entries) {
      const textMax = innerW - keyW(e) - K.legendKeyGap * fs;
      if (textMax <= 0) break;
      // Stacked entries sit further apart than lines of text; a long name wraps to a few lines.
      const lines = wrapFit(e.label, look, textMax, K.legendEntryLines);
      const h = lines.length * lh + K.legendRowGap;
      if (used + h > innerH + 0.01) break;
      used += h;
      const w = keyW(e) + K.legendKeyGap * fs + Math.max(...lines.map((l) => textWidth(l, look)));
      width = Math.max(width, w);
      rows.push({ items: [{ entry: e, x: 0, width: w, lines }], h });
    }
  } else {
    // One entry wider than the whole legend is cut.
    const cut = entries.map((e) => {
      const label = fitText(e.label, look, Math.max(innerW - keyW(e) - K.legendKeyGap * fs, 0));
      return { entry: e, label, w: keyW(e) + K.legendKeyGap * fs + textWidth(label, look) };
    });
    const packed = cut.reduce((sum, c) => sum + c.w, 0) + K.legendEntryGap * fs * (cut.length - 1);
    if (packed <= innerW) {
      // One row: each entry as wide as its own text.
      let x = 0;
      const row: LegendRow = { items: [], h: lh };
      for (const c of cut) {
        row.items.push({ entry: c.entry, x, width: c.w, lines: [c.label] });
        x += c.w + K.legendEntryGap * fs;
      }
      rows.push(row);
      used = lh;
      width = packed;
    } else {
      // More than a row: Office lays the entries on a grid of equal cells, as few columns as
      // keep the fewest rows (20 entries that fit 9 a row go 7, 7, 6), rows further apart.
      const cell = Math.max(...cut.map((c) => c.w)) + K.legendGridGap * fs;
      const maxCols = Math.max(Math.floor((innerW + K.legendGridGap * fs) / cell), 1);
      const cols = Math.ceil(cut.length / Math.ceil(cut.length / maxCols));
      const pitch = lh + K.legendGridRowGap * fs;
      const span = cols * cell - K.legendGridGap * fs;
      // Rows are a pitch apart; the gap goes between them only, so the block is one gap shorter.
      for (let start = 0; start < cut.length; start += cols) {
        if (used + pitch - (pitch - lh) > innerH + 0.01) break;
        used += pitch;
        rows.push({ items: cut.slice(start, start + cols).map((c, k) => ({ entry: c.entry, x: k * cell, width: c.w, lines: [c.label] })), h: pitch, span });
      }
      used -= pitch - lh;
      width = span;
    }
  }
  if (rows.length === 0) return undefined;
  const boxW = Math.min(width + 2 * K.legendPadX, chartW);
  const boxH = Math.min(used + 2 * K.legendPadY, chartH);
  let box: Box;
  if (legend.layout !== undefined) {
    const l = legend.layout;
    box = {
      x: (l.x ?? 0) * chartW,
      y: (l.y ?? 0) * chartH,
      w: l.w !== undefined ? l.w * chartW : boxW,
      h: l.h !== undefined ? l.h * chartH : boxH,
    };
  } else {
    switch (legend.position) {
      case "r": box = { x: area.x + area.w - boxW, y: area.y + (area.h - boxH) / 2, w: boxW, h: boxH }; break;
      case "l": box = { x: area.x, y: area.y + (area.h - boxH) / 2, w: boxW, h: boxH }; break;
      case "tr": box = { x: area.x + area.w - boxW, y: area.y, w: boxW, h: boxH }; break;
      case "t": box = { x: area.x + (area.w - boxW) / 2, y: area.y, w: boxW, h: boxH }; break;
      default: box = { x: area.x + (area.w - boxW) / 2, y: area.y + area.h - boxH, w: boxW, h: boxH };
    }
  }
  // Whatever the layout asked for, the legend stays in the frame.
  box.w = Math.min(box.w, chartW);
  box.h = Math.min(box.h, chartH);
  box.x = Math.min(Math.max(box.x, 0), chartW - box.w);
  box.y = Math.min(Math.max(box.y, 0), chartH - box.h);
  return { box, rows, vertical };
}

function paintLegend(scene: Scene, spec: ChartSpecModel, layout: LegendLayout): void {
  const legend = spec.legend!;
  const look = legend.text;
  const fs = look.size;
  const { box, rows, vertical } = layout;
  scene.rect(box.x, box.y, box.w, box.h, legend.look, false);
  const lh = lineHeight(look);
  // Grid rows are a pitch tall with the gap split above and below: the first row's half gap sits outside the box.
  const first = rows[0];
  let y = box.y + K.legendPadY - (first?.span !== undefined ? (first.h - lh) / 2 : 0);
  for (const row of rows) {
    const last = row.items[row.items.length - 1];
    const rowWidth = row.span ?? (last === undefined ? 0 : last.x + last.width);
    const startX = vertical ? box.x + K.legendPadX : box.x + (box.w - rowWidth) / 2;
    const cy = y + row.h / 2;
    for (const item of row.items) {
      const x = startX + item.x;
      const e = item.entry;
      const k = (e.kind === "fill" ? K.legendKey : K.legendLineKey) * fs;
      if (e.kind === "fill") {
        scene.rect(x, cy - k / 2, k, k, e.look, false);
      } else {
        scene.segment(x, cy, x + k, cy, e.look.line);
        if (e.marker !== undefined) paintMarker(scene, x + k / 2, cy, e.marker, e.look);
      }
      const top = cy - (item.lines.length * lh) / 2;
      const rtl = isRtlText(item.lines.join("\n"));
      item.lines.forEach((line, n) => scene.text(x + k + K.legendKeyGap * fs, top + n * lh + ascent(look), line, look, "start", 0, rtl));
      scene.breakRun();
    }
    y += row.h;
  }
}

// ------------------------------------------------------------------ markers

function markerPath(symbol: Marker["symbol"], cx: number, cy: number, size: number): { d: string; stroked: boolean } {
  const h = size / 2;
  const p = (x: number, y: number): string => `${fmt(cx + x)} ${fmt(cy + y)}`;
  switch (symbol) {
    case "circle":
    case "dot": {
      const rr = symbol === "dot" ? h / 2 : h;
      return { d: `M${p(-rr, 0)}A${fmt(rr)} ${fmt(rr)} 0 1 0 ${p(rr, 0)}A${fmt(rr)} ${fmt(rr)} 0 1 0 ${p(-rr, 0)}Z`, stroked: false };
    }
    case "square": return { d: `M${p(-h, -h)}L${p(h, -h)}L${p(h, h)}L${p(-h, h)}Z`, stroked: false };
    case "diamond": return { d: `M${p(0, -h)}L${p(h, 0)}L${p(0, h)}L${p(-h, 0)}Z`, stroked: false };
    case "triangle": return { d: `M${p(0, -h)}L${p(h, h)}L${p(-h, h)}Z`, stroked: false };
    case "dash": return { d: `M${p(-h, -h / 4)}L${p(h, -h / 4)}L${p(h, h / 4)}L${p(-h, h / 4)}Z`, stroked: false };
    case "x": return { d: `M${p(-h, -h)}L${p(h, h)}M${p(-h, h)}L${p(h, -h)}`, stroked: true };
    case "plus": return { d: `M${p(-h, 0)}L${p(h, 0)}M${p(0, -h)}L${p(0, h)}`, stroked: true };
    case "star": return { d: `M${p(-h, -h)}L${p(h, h)}M${p(-h, h)}L${p(h, -h)}M${p(0, -h)}L${p(0, h)}`, stroked: true };
    default: return { d: "", stroked: false };
  }
}

function paintMarker(scene: Scene, cx: number, cy: number, marker: Marker, seriesLook: Look): void {
  if (marker.symbol === "none" || marker.symbol === "picture") return;
  const { d, stroked } = markerPath(marker.symbol, cx, cy, marker.size);
  if (d === "") return;
  if (stroked) {
    const line = marker.look.line ?? (typeof marker.look.fill === "string" ? { color: marker.look.fill, width: 0.75 } : seriesLook.line);
    if (line !== undefined) scene.path(d, { line });
    return;
  }
  scene.path(d, marker.look);
}

// ------------------------------------------------------------------ data labels

function labelText(labels: DataLabels, series: Series, i: number, value: number, percent: number | undefined, category: string, spec: ChartSpecModel, group: Group): string {
  const range = series.labelRange?.[i];
  const code = labels.formatCode ?? series.formatCode;
  if (labels.custom !== undefined) {
    // Author-typed text: literal runs as typed, fields filled from this point (their cached text when the point has no such thing).
    return labels.custom.map((part) => {
      if (typeof part === "string") return part;
      switch (part.field) {
        case "VALUE": case "YVALUE": return formatValue(value, code, spec);
        case "SERIESNAME": return series.name;
        case "CATEGORYNAME": case "XVALUE": return category;
        case "PERCENTAGE": return percent === undefined ? part.cached : formatValue(percent, "0%", spec);
        case "CELLRANGE": return range ?? part.cached;
        default: {
          const b = series.bubbleSizes?.[i];
          return b === null || b === undefined ? part.cached : formatValue(b, series.formatCode, spec);
        }
      }
    }).join("");
  }
  const parts: string[] = [];
  // "Value from cells" leads the label, ahead of the series name.
  if (labels.fromCells && range !== undefined) parts.push(range);
  if (labels.showSeries) parts.push(series.name);
  if (labels.showCategory) parts.push(category);
  if (labels.showValue) parts.push(formatValue(value, code, spec));
  if (labels.showPercent && percent !== undefined) parts.push(formatValue(percent, labels.formatCode !== undefined && !labels.showValue ? labels.formatCode : "0%", spec));
  if (labels.showBubbleSize && series.bubbleSizes !== undefined) {
    const b = series.bubbleSizes[i];
    if (b !== null && b !== undefined) parts.push(formatValue(b, series.formatCode, spec));
  }
  // Office stacks a pie label's parts on separate lines unless the file names a separator.
  return parts.join(labels.separator ?? (isRadialKind(group.kind) ? "\n" : ", "));
}

/**
 * Each slice's share of `total`. With the default whole-percent format Office
 * makes the labels add up to 100%: the rounding error goes to the largest
 * slice (58.6/22.9/10/8.6 reads 58/23/10/9, not 59/23/10/9).
 */
function sliceShares(values: ReadonlyArray<number | null>, total: number, adjust: boolean): number[] {
  const shares = values.map((v) => (total > 0 ? Math.abs(v ?? 0) / total : 0));
  if (!adjust || total <= 0) return shares;
  let rounded = 0;
  let largest = 0;
  shares.forEach((s, i) => {
    rounded += Math.round(s * 100);
    if (s > shares[largest]!) largest = i;
  });
  if (rounded !== 100 && shares[largest]! > 0) shares[largest] = (Math.round(shares[largest]! * 100) + 100 - rounded) / 100;
  return shares;
}

function labelsFor(series: Series, i: number): DataLabels | undefined {
  const l = series.points[i]?.labels ?? series.labels;
  if (l === undefined || l.deleted) return undefined;
  if (l.custom !== undefined || (l.fromCells && series.labelRange !== undefined)) return l;
  if (!l.showValue && !l.showCategory && !l.showSeries && !l.showPercent && !l.showBubbleSize) return undefined;
  return l;
}

/**
 * `look.offset` (a manual label position) is applied when the label is painted, against the chart's size.
 * `leader` is where a leader line starts (a pie slice's rim), drawn only when the label was moved off its spot.
 */
interface PendingLabel {
  text: string; x: number; y: number; look: DataLabels; anchor: "start" | "middle" | "end"; v: "top" | "middle" | "bottom";
  leader?: { x: number; y: number; line: Line; centre: { x: number; y: number; r: number } };
}

function placeLabel(pending: PendingLabel[], labels: DataLabels, text: string, x: number, y: number, pos: LabelPosition, gap: number, size: number): void {
  // `size` is the half-extent of what the label sits beside (a marker radius), along the placement direction.
  switch (pos) {
    case "t": pending.push({ text, x, y: y - size - gap, look: labels, anchor: "middle", v: "bottom" }); break;
    case "b": pending.push({ text, x, y: y + size + gap, look: labels, anchor: "middle", v: "top" }); break;
    case "l": pending.push({ text, x: x - size - gap, y, look: labels, anchor: "end", v: "middle" }); break;
    case "r": pending.push({ text, x: x + size + gap, y, look: labels, anchor: "start", v: "middle" }); break;
    default: pending.push({ text, x, y, look: labels, anchor: "middle", v: "middle" });
  }
}

/** `frame` is the whole chart: a manual label offset is a fraction of it, and no label leaves it. */
function paintPendingLabels(scene: Scene, pending: PendingLabel[], frame: Box): void {
  for (const p of pending) {
    const lines = p.text.split("\n");
    const { width, height } = blockSize(lines, p.look.text);
    const px = p.x + (p.look.offset?.x ?? 0) * frame.w;
    const py = p.y + (p.look.offset?.y ?? 0) * frame.h;
    const left = p.anchor === "start" ? px : p.anchor === "end" ? px - width : px - width / 2;
    const top = p.v === "top" ? py : p.v === "bottom" ? py - height : py - height / 2;
    // A data label never leaves the chart area: it slides back in (the backing box counts too).
    const pad = p.look.look.fill !== undefined || p.look.look.line !== undefined ? 2 : 0;
    const dx = Math.max(frame.x - (left - pad), 0) - Math.max(left + width + pad - (frame.x + frame.w), 0);
    const dy = Math.max(frame.y - (top - pad / 2), 0) - Math.max(top + height + pad / 2 - (frame.y + frame.h), 0);
    if (pad > 0) scene.rect(left + dx - 2, top + dy - 1, width + 4, height + 2, p.look.look, false);
    drawLines(scene, lines, px + dx, py + dy, p.look.text, p.anchor, p.v);
    const offset = p.look.offset;
    if (p.leader !== undefined && offset !== undefined && (offset.x !== 0 || offset.y !== 0)) {
      paintLeader(scene, p.leader, { x: left + dx, y: top + dy, w: width, h: height });
    }
  }
}

/**
 * A moved label's leader line, as Office draws it: from the slice's rim toward
 * the label, ending in a short level stroke into the side that faces the rim
 * (the top or bottom middle when the rim sits within the label's width). A
 * label that still touches the pie gets none (the oracle's g-pie-leaders).
 */
function paintLeader(scene: Scene, from: NonNullable<PendingLabel["leader"]>, box: Box): void {
  const gap = K.pieLeaderGap;
  const nearX = Math.min(Math.max(from.centre.x, box.x), box.x + box.w);
  const nearY = Math.min(Math.max(from.centre.y, box.y), box.y + box.h);
  if (Math.hypot(nearX - from.centre.x, nearY - from.centre.y) < from.centre.r + gap) return;
  let d: string;
  if (from.x < box.x || from.x > box.x + box.w) {
    const side = from.x < box.x ? -1 : 1;
    const edgeX = side < 0 ? box.x - gap : box.x + box.w + gap;
    const cy = box.y + box.h / 2;
    d = `M${fmt(from.x)} ${fmt(from.y)}L${fmt(edgeX + side * K.pieLeaderElbow)} ${fmt(cy)}L${fmt(edgeX)} ${fmt(cy)}`;
  } else {
    const edgeY = from.y < box.y ? box.y - gap : box.y + box.h + gap;
    d = `M${fmt(from.x)} ${fmt(from.y)}L${fmt(box.x + box.w / 2)} ${fmt(edgeY)}`;
  }
  scene.breakRun();
  scene.path(d, { line: from.line });
}

/**
 * The last guard: text the layout still put partly outside the frame (a long
 * label, an overlay, a manual layout) is cut to the frame's width if it is
 * wider, then slid back in. Office lets labels overhang the chart area
 * slightly (the frame clips them); here nothing does, so a reader never sees
 * half a label.
 */
function confineText(primitives: ChartPrimitive[], width: number, height: number): void {
  for (const p of primitives) {
    if (p.kind !== "text") continue;
    let b = textBox(p);
    if (b.x0 >= 0 && b.y0 >= 0 && b.x1 <= width && b.y1 <= height) continue;
    if (p.rotate === undefined && b.x1 - b.x0 > width) {
      p.text = fitText(p.text, { size: p.fontSize, bold: p.fontWeight === "bold", italic: false, color: "", family: p.fontFamily }, width);
      b = textBox(p);
    }
    p.x = Math.round((p.x + Math.max(-b.x0, 0) - Math.max(b.x1 - width, 0)) * 100) / 100;
    p.y = Math.round((p.y + Math.max(-b.y0, 0) - Math.max(b.y1 - height, 0)) * 100) / 100;
  }
}

// ------------------------------------------------------------------ cartesian painters

interface Cartesian {
  inner: Box;
  cat: AxisPlan;
  val: AxisPlan;
}

/**
 * Where point `i` of series `s` sits along the category axis. On a date axis
 * the point goes by its date, not its index; one whose date is missing or
 * falls outside the axis has no position.
 */
function slotPos(c: Cartesian, s: Series, i: number): number | undefined {
  const scale = c.cat.scale;
  if (scale.kind !== "category") return along(c.cat, c.inner, valueFraction(scale, i + 1));
  if (scale.time === undefined) return along(c.cat, c.inner, categoryFraction(scale, i));
  const date = s.categoryValues?.[i] ?? c.cat.dates?.[i];
  if (date === null || date === undefined || !Number.isFinite(date)) return undefined;
  const slot = timeSlot(scale.time, date);
  return slot < 0 || slot >= scale.count ? undefined : along(c.cat, c.inner, categoryFraction(scale, slot));
}

/** What "category name" means for point `i`: its category, or on an XY chart its x value. */
function categoryText(c: Cartesian, s: Series, i: number, xy: boolean, spec: ChartSpecModel): string {
  if (!xy) return c.cat.categories[i] ?? "";
  const x = s.categoryValues?.[i] ?? i + 1;
  return x === null ? "" : formatValue(x, s.categoryFormatCode ?? "General", spec);
}

/** Points a group draws along its category axis: its longest series. */
function pointCount(g: Group): number {
  let n = 0;
  for (const s of g.series) if (s.values.length > n) n = s.values.length;
  return n;
}

function valPos(c: Cartesian, v: number): number {
  const s = c.val.scale as ValueScale;
  return along(c.val, c.inner, valueFraction(s, v));
}

/** The value bars and areas grow from: where the category axis crosses. */
function baseValue(c: Cartesian): number {
  const s = c.val.scale as ValueScale;
  const a = c.cat.axis;
  if (a.crosses === "max") return s.max;
  if (a.crosses === "min") return s.min;
  const at = a.crossesAt ?? 0;
  if (s.log !== undefined) return a.crossesAt ?? s.min;
  return Math.min(Math.max(at, s.min), s.max);
}

function clampValue(c: Cartesian, v: number): number {
  const s = c.val.scale as ValueScale;
  return Math.min(Math.max(v, s.min), s.max);
}

// ------------------------------------------------------------------ trendlines and error bars

/**
 * One plotted point as the overlays need it: `x` is the position along the
 * category (1-based) or x axis, `y` the series value, `at` where it was drawn
 * on the value axis (the stack top for stacked groups), `px` its pixel along
 * the category axis when the painter already knows it (a bar's own centre).
 */
interface ExtraPoint { x: number; y: number; at: number; px?: number }

function catPixel(c: Cartesian, x: number): number {
  const s = c.cat.scale;
  if (s.kind === "category") return along(c.cat, c.inner, categoryFraction(s, x - 1));
  return along(c.cat, c.inner, valueFraction(s, x));
}

/** Pixel of a point (category-axis coordinate, value); bars of a horizontal chart swap the two. */
function plotPoint(c: Cartesian, horizontalBars: boolean, x: number, y: number, px?: number): [number, number] {
  const a = px ?? catPixel(c, x);
  const b = valPos(c, y);
  return horizontalBars ? [b, a] : [a, b];
}

/** Liang-Barsky: the part of a segment inside `box`, or undefined. */
function clipSegment(x0: number, y0: number, x1: number, y1: number, box: Box): [number, number, number, number] | undefined {
  let t0 = 0;
  let t1 = 1;
  const dx = x1 - x0;
  const dy = y1 - y0;
  const edges: Array<[number, number]> = [[-dx, x0 - box.x], [dx, box.x + box.w - x0], [-dy, y0 - box.y], [dy, box.y + box.h - y0]];
  for (const [p, q] of edges) {
    if (p === 0) {
      if (q < 0) return undefined;
    } else {
      const t = q / p;
      if (p < 0) t0 = Math.max(t0, t);
      else t1 = Math.min(t1, t);
    }
  }
  return t0 > t1 ? undefined : [x0 + t0 * dx, y0 + t0 * dy, x0 + t1 * dx, y0 + t1 * dy];
}

/** A polyline clipped to `box`: runs break where the curve leaves it. `last` is the last visible point. */
function clippedPath(points: ReadonlyArray<[number, number] | undefined>, box: Box): { d: string; last?: [number, number] } {
  let d = "";
  let last: [number, number] | undefined;
  let prev: [number, number] | undefined;
  for (const p of points) {
    if (p === undefined) { prev = undefined; continue; }
    if (prev !== undefined) {
      const seg = clipSegment(prev[0], prev[1], p[0], p[1], box);
      if (seg !== undefined) {
        const connected = last !== undefined && Math.abs(last[0] - seg[0]) < 1e-6 && Math.abs(last[1] - seg[1]) < 1e-6;
        d += `${connected ? "" : `M${fmt(seg[0])} ${fmt(seg[1])}`}L${fmt(seg[2])} ${fmt(seg[3])}`;
        last = [seg[2], seg[3]];
      }
    }
    prev = p;
  }
  return last === undefined ? { d } : { d, last };
}

function paintTrendlines(scene: Scene, g: Group, c: Cartesian, s: Series, pts: ExtraPoint[], pending: PendingLabel[], spec: ChartSpecModel, hb: boolean): void {
  // Office offers trendlines on unstacked series only.
  if (g.grouping === "stacked" || g.grouping === "percentStacked" || pts.length < 2) return;
  const xs = pts.map((p) => p.x);
  const ys = pts.map((p) => p.y);
  for (const t of s.trendlines ?? []) {
    const curve: Array<[number, number] | undefined> = [];
    const text: string[] = [];
    if (t.type === "movingAvg") {
      for (const [x, y] of movingAverage(xs, ys, t.period)) curve.push([x, y]);
    } else {
      const fit = fitTrend(t, xs, ys);
      if (fit === undefined) continue;
      // A loop, not a spread: a series can hold up to 100,000 points, which is near an engine's argument limit.
      let low = Infinity;
      let high = -Infinity;
      for (const x of xs) {
        if (x < low) low = x;
        if (x > high) high = x;
      }
      const x0 = low - t.backward;
      const x1 = high + t.forward;
      const n = t.type === "linear" ? 1 : K.trendSamples;
      for (let k = 0; k <= n; k++) {
        const x = x0 + ((x1 - x0) * k) / n;
        const y = fit.at(x);
        curve.push(Number.isFinite(y) ? [x, y] : undefined);
      }
      // "General" on a trendline label is Office's own rounding (5 significant digits, R-squared to 4
      // places: "1.3214x2 - 9.25x + 17"), not the cells' General, which would print every digit.
      const code = t.labelFormatCode;
      const format = code === undefined || code === "General" ? undefined : (v: number): string => formatValue(v, code, spec);
      if (t.dispEq) text.push(equation(t, fit, format));
      if (t.dispRSqr) text.push(rSquaredLabel(fit, format));
    }
    const px = curve.map((p) => (p === undefined ? undefined : plotPoint(c, hb, p[0], p[1])));
    const { d, last } = clippedPath(px, c.inner);
    scene.breakRun();
    scene.path(d, { line: t.line });
    if (text.length > 0 && last !== undefined) {
      const look: DataLabels = { showValue: false, showCategory: false, showSeries: false, showPercent: false, showBubbleSize: false, showLegendKey: false, text: t.labelText, look: {}, deleted: false, fromCells: false };
      const { width, height } = blockSize(text, t.labelText);
      // As PowerPoint places it (g-trend-poly-eq): the block's right edge at the line's end, centred on
      // it vertically, each line centred in the block; kept inside the plot.
      const right = Math.min(Math.max(last[0], c.inner.x + width), c.inner.x + c.inner.w);
      const y = Math.min(Math.max(last[1], c.inner.y + height / 2), c.inner.y + c.inner.h - height / 2);
      pending.push({ text: text.join("\n"), x: right - width / 2, y, look, anchor: "middle", v: "middle" });
    }
  }
}

function paintErrorBars(scene: Scene, g: Group, c: Cartesian, s: Series, pts: ExtraPoint[], hb: boolean): void {
  if (g.grouping === "percentStacked") return;
  for (const eb of s.errBars ?? []) {
    const alongX = eb.dir === "x";
    const amounts = errorAmounts(eb, pts.map((p) => (alongX ? p.x : p.y)));
    // The bar runs along the value axis, which is horizontal for horizontal bars and for x bars.
    const vertical = !alongX && !hb;
    let d = "";
    pts.forEach((p, k) => {
      const amt = amounts[k];
      if (amt === undefined) return;
      const centre = amt.centre ?? (alongX ? p.x : p.at);
      const at = (v: number): [number, number] => (alongX ? plotPoint(c, hb, v, p.y, undefined) : plotPoint(c, hb, p.x, v, p.px));
      const ends: Array<[number, number]> = [];
      const lo = at(centre - amt.minus);
      const hi = at(centre + amt.plus);
      const seg = clipSegment(lo[0], lo[1], hi[0], hi[1], c.inner);
      if (seg === undefined) return;
      d += `M${fmt(seg[0])} ${fmt(seg[1])}L${fmt(seg[2])} ${fmt(seg[3])}`;
      if (eb.noEndCap) return;
      // A cap sits on an end that was not cut off.
      if (amt.minus > 0 && Math.abs(seg[0] - lo[0]) + Math.abs(seg[1] - lo[1]) < 0.01) ends.push([lo[0], lo[1]]);
      if (amt.plus > 0 && Math.abs(seg[2] - hi[0]) + Math.abs(seg[3] - hi[1]) < 0.01) ends.push([hi[0], hi[1]]);
      for (const [x, y] of ends) d += vertical ? `M${fmt(x - K.errCap)} ${fmt(y)}L${fmt(x + K.errCap)} ${fmt(y)}` : `M${fmt(x)} ${fmt(y - K.errCap)}L${fmt(x)} ${fmt(y + K.errCap)}`;
    });
    scene.breakRun();
    scene.path(d, { line: eb.line });
  }
}

/** Error bars, then trendlines, over a series that was just drawn. */
function paintExtras(scene: Scene, g: Group, c: Cartesian, s: Series, pts: ExtraPoint[], pending: PendingLabel[], spec: ChartSpecModel): void {
  if (s.errBars === undefined && s.trendlines === undefined) return;
  const hb = g.kind === "bar" && g.direction === "bar";
  paintErrorBars(scene, g, c, s, pts, hb);
  paintTrendlines(scene, g, c, s, pts, pending, spec, hb);
}

function paintBars(scene: Scene, g: Group, c: Cartesian, pending: PendingLabel[], spec: ChartSpecModel): void {
  const horizontalBars = g.direction === "bar";
  const scale = c.cat.scale as CategoryScale;
  const slot = (horizontalBars ? c.inner.h : c.inner.w) * categorySlot(scale);
  const stacked = g.grouping !== "clustered";
  const n = stacked ? 1 : g.series.length;
  const overlap = Math.min(Math.max(g.overlap / 100, -1), 1);
  const gap = Math.max(g.gapWidth, 0) / 100;
  const bar = slot / (n - overlap * (n - 1) + gap);
  const groupWidth = bar * (n - overlap * (n - 1));
  const base = baseValue(c);
  const count = pointCount(g);
  const posSum = new Array<number>(count).fill(0);
  const negSum = new Array<number>(count).fill(0);
  const totals = new Array<number>(count).fill(0);
  if (g.grouping === "percentStacked") {
    for (const s of g.series) s.values.forEach((v, i) => { if (v !== null && i < count) totals[i]! += Math.abs(v); });
  }
  const extras: ExtraPoint[][] = g.series.map(() => []);
  // Category order along the pixel axis: the first category's slot starts where the reversed-or-not axis begins.
  g.series.forEach((s, si) => {
    scene.breakRun();
    for (let i = 0; i < count; i++) {
      const raw = s.values[i];
      if (raw === null || raw === undefined) continue;
      const center = slotPos(c, s, i);
      if (center === undefined) continue;
      const span = s.spans?.[i];
      if (s.spans !== undefined && (span === null || span === undefined)) continue;
      let v = raw;
      if (g.grouping === "percentStacked") v = totals[i]! === 0 ? 0 : raw / totals[i]!;
      let from = base;
      let to = v;
      if (stacked) {
        const acc = v >= 0 ? posSum : negSum;
        from = acc[i]!;
        to = from + v;
        acc[i] = to;
      }
      if (span !== undefined && span !== null) [from, to] = span;
      const offset = stacked ? 0 : si * bar * (1 - overlap);
      // Office lays series out from the axis start: reversed axes mirror the order with the slot.
      const dir = horizontalBars ? (c.cat.axis.reversed ? 1 : -1) : (c.cat.axis.reversed ? -1 : 1);
      const lead = center - (dir * groupWidth) / 2;
      const a = lead + dir * offset;
      const b = a + dir * bar;
      const p0 = valPos(c, clampValue(c, from));
      const p1 = valPos(c, clampValue(c, to));
      const look = pointLook(s, i, raw);
      extras[si]!.push({ x: i + 1, y: raw, at: stacked ? to : v, px: (a + b) / 2 });
      if (horizontalBars) scene.rect(Math.min(p0, p1), Math.min(a, b), Math.abs(p1 - p0), Math.abs(b - a), look);
      else scene.rect(Math.min(a, b), Math.min(p0, p1), Math.abs(b - a), Math.abs(p1 - p0), look);
      const labels = labelsFor(s, i);
      if (labels !== undefined) {
        const text = labelText(labels, s, i, raw, g.grouping === "percentStacked" ? v : undefined, c.cat.categories[i] ?? "", spec, g);
        const pos = labels.position ?? (stacked ? "ctr" : "outEnd");
        const mid = (a + b) / 2;
        // Direction along the value axis from base toward the end, in pixels.
        const sgn = Math.sign(p1 - p0) || (horizontalBars ? 1 : -1);
        const gapPx = K.dataLabelGap;
        let at: number;
        let anchorPos: number;
        switch (pos) {
          case "inEnd": at = p1 - sgn * gapPx; anchorPos = -1; break;
          case "inBase": at = p0 + sgn * gapPx; anchorPos = 1; break;
          case "ctr": at = (p0 + p1) / 2; anchorPos = 0; break;
          default: at = p1 + sgn * gapPx; anchorPos = 1;
        }
        if (horizontalBars) {
          const anchor = anchorPos === 0 ? "middle" : (anchorPos * sgn > 0 ? "start" : "end");
          pending.push({ text, x: at, y: mid, look: labels, anchor, v: "middle" });
        } else {
          const v2 = anchorPos === 0 ? "middle" : (anchorPos * sgn > 0 ? "top" : "bottom");
          pending.push({ text, x: mid, y: at, look: labels, anchor: "middle", v: v2 });
        }
      }
    }
  });
  g.series.forEach((s, si) => paintExtras(scene, g, c, s, extras[si]!, pending, spec));
  if (!horizontalBars) paintConnectors(scene, g, c, bar);
}

/**
 * Waterfall connector lines: from where one bar ends (the running level) to
 * where the next begins. Observed: a total bar is entered at its own value, so
 * the line slopes when the total is not the running sum.
 */
function paintConnectors(scene: Scene, g: Group, c: Cartesian, barWidth: number): void {
  const s = g.series[0];
  if (g.connectors === undefined || s?.spans === undefined) return;
  scene.breakRun();
  const totals = s.totals ?? [];
  let d = "";
  let previous = -1;
  for (let i = 0; i < s.spans.length; i++) {
    const span = s.spans[i];
    if (span === null || span === undefined) continue;
    const at = slotPos(c, s, i);
    if (at === undefined) continue;
    if (previous >= 0) {
      const before = s.spans[previous]!;
      const entry = totals.includes(i) ? span[1] : span[0];
      const x0 = slotPos(c, s, previous)! + barWidth / 2;
      const x1 = at - barWidth / 2;
      d += `M${fmt(x0)} ${fmt(valPos(c, clampValue(c, before[1])))}L${fmt(x1)} ${fmt(valPos(c, clampValue(c, entry)))}`;
    }
    previous = i;
  }
  scene.path(d, { line: g.connectors });
}

function smoothPath(points: number[]): string {
  if (points.length < 6) return polyPath(points);
  let d = `M${fmt(points[0]!)} ${fmt(points[1]!)}`;
  const n = points.length / 2;
  for (let i = 0; i < n - 1; i++) {
    const x0 = points[Math.max(i - 1, 0) * 2]!, y0 = points[Math.max(i - 1, 0) * 2 + 1]!;
    const x1 = points[i * 2]!, y1 = points[i * 2 + 1]!;
    const x2 = points[(i + 1) * 2]!, y2 = points[(i + 1) * 2 + 1]!;
    const x3 = points[Math.min(i + 2, n - 1) * 2]!, y3 = points[Math.min(i + 2, n - 1) * 2 + 1]!;
    const t = 1 / 6;
    d += `C${fmt(x1 + (x2 - x0) * t)} ${fmt(y1 + (y2 - y0) * t)} ${fmt(x2 - (x3 - x1) * t)} ${fmt(y2 - (y3 - y1) * t)} ${fmt(x2)} ${fmt(y2)}`;
  }
  return d;
}

function polyPath(points: number[]): string {
  let d = "";
  for (let i = 0; i < points.length; i += 2) d += `${i === 0 ? "M" : "L"}${fmt(points[i]!)} ${fmt(points[i + 1]!)}`;
  return d;
}

/**
 * Hi-low lines and up-down bars of a line or stock group. They sit under the
 * series' own lines and markers: first one line per category from its highest
 * to its lowest value, then the bars from the first series' value to the last
 * one's (the open and the close of a stock chart), over the line.
 */
function paintBetweenSeries(scene: Scene, g: Group, c: Cartesian): void {
  if (g.hiLowLines === undefined && g.upDownBars === undefined) return;
  const scale = c.cat.scale;
  const count = pointCount(g);
  const slot = scale.kind === "category" ? c.inner.w * categorySlot(scale) : c.inner.w / Math.max(count, 1);
  // Every series of the group shares its categories (and on a date axis its dates).
  const at = (i: number): number | undefined => slotPos(c, g.series[0]!, i);
  scene.breakRun();
  if (g.hiLowLines !== undefined) {
    let d = "";
    for (let i = 0; i < count; i++) {
      let lo = Infinity;
      let hi = -Infinity;
      let seen = 0;
      for (const s of g.series) {
        const v = s.values[i];
        if (v === null || v === undefined) continue;
        seen++;
        if (v < lo) lo = v;
        if (v > hi) hi = v;
      }
      const x = at(i);
      if (seen < 2 || x === undefined) continue;
      d += `M${fmt(x)} ${fmt(valPos(c, clampValue(c, hi)))}V${fmt(valPos(c, clampValue(c, lo)))}`;
    }
    scene.path(d, { line: g.hiLowLines });
  }
  const bars = g.upDownBars;
  const first = g.series[0];
  const last = g.series[g.series.length - 1];
  if (bars === undefined || first === undefined || last === undefined || first === last) return;
  const width = slot / (1 + bars.gapWidth / 100);
  let up = "";
  let down = "";
  for (let i = 0; i < count; i++) {
    const open = first.values[i];
    const close = last.values[i];
    const x = at(i);
    if (open === null || open === undefined || close === null || close === undefined || x === undefined) continue;
    const y0 = valPos(c, clampValue(c, open));
    const y1 = valPos(c, clampValue(c, close));
    const rect = `M${fmt(x - width / 2)} ${fmt(Math.min(y0, y1))}h${fmt(width)}v${fmt(Math.abs(y1 - y0))}h${fmt(-width)}Z`;
    if (close >= open) up += rect;
    else down += rect;
  }
  scene.path(up, bars.up, false);
  scene.path(down, bars.down, false);
}

/** Line, stacked line, scatter and the line part of radar share this: runs of points, broken at gaps. */
function paintLines(scene: Scene, g: Group, c: Cartesian, pending: PendingLabel[], spec: ChartSpecModel): void {
  const xy = g.kind === "scatter" || g.kind === "bubble";
  paintBetweenSeries(scene, g, c);
  const stackedSums: number[] = [];
  const totals: number[] = [];
  if (g.grouping === "percentStacked") {
    for (const s of g.series) s.values.forEach((v, i) => { totals[i] = (totals[i] ?? 0) + Math.abs(v ?? 0); });
  }
  const placed: Array<{ s: Series; pts: Array<[number, number, number, number]>; extra: ExtraPoint[] }> = [];
  for (const s of g.series) {
    scene.breakRun();
    const runs: number[][] = [[]];
    const pts: Array<[number, number, number, number]> = [];
    const extra: ExtraPoint[] = [];
    const n = s.values.length;
    for (let i = 0; i < n; i++) {
      let v = s.values[i] ?? null;
      if (v === null) {
        if (spec.blanks === "zero" && !xy) v = 0;
        else {
          if (spec.blanks !== "span") runs.push([]);
          continue;
        }
      }
      let y = v;
      if (g.grouping === "stacked" || g.grouping === "percentStacked") {
        y = (stackedSums[i] ?? 0) + v;
        stackedSums[i] = y;
        if (g.grouping === "percentStacked") y = totals[i] ? y / totals[i]! : 0;
      }
      let x: number;
      if (xy) {
        const xv = s.categoryValues?.[i] ?? i + 1;
        if (xv === null) continue;
        x = along(c.cat, c.inner, valueFraction(c.cat.scale as ValueScale, xv));
      } else {
        const at = slotPos(c, s, i);
        if (at === undefined) continue;
        x = at;
      }
      const py = valPos(c, y);
      runs[runs.length - 1]!.push(x, py);
      pts.push([x, py, i, v]);
      extra.push({ x: xy ? s.categoryValues?.[i] ?? i + 1 : i + 1, y: v, at: y, px: x });
    }
    // A series with no line (a marker-only scatter) draws none.
    if (s.look.line !== undefined) {
      for (const run of runs) {
        if (run.length < 4) continue;
        const shown = thinRun(run, c.inner.w);
        scene.path(s.smooth ? smoothPath(shown) : polyPath(shown), { line: s.look.line });
      }
    }
    placed.push({ s, pts, extra });
  }
  // Markers sit above every line of the group.
  for (const { s, pts, extra } of placed) {
    scene.breakRun();
    for (const [x, y, i] of pts) {
      const marker = s.points[i]?.marker ?? s.marker;
      paintMarker(scene, x, y, marker, s.look);
    }
    paintExtras(scene, g, c, s, extra, pending, spec);
    for (const [x, y, i, v] of pts) {
      const labels = labelsFor(s, i);
      if (labels === undefined) continue;
      const text = labelText(labels, s, i, v, undefined, categoryText(c, s, i, xy, spec), spec, g);
      const marker = s.points[i]?.marker ?? s.marker;
      placeLabel(pending, labels, text, x, y, labels.position ?? "r", K.dataLabelGap, marker.symbol === "none" ? 0 : marker.size / 2);
    }
  }
}

/**
 * A line run (flat x, y pairs) with more points than the plot has columns,
 * cut to what can show: per column of `K.thinColumn` points, the first, the
 * lowest, the highest and the last point, in their order (the M4 reduction,
 * Jugel et al., VLDB 2014). Every peak, dip and step stays where it was; only
 * points no reader could tell apart go. Runs whose x turns back (an unsorted
 * scatter) are left whole.
 */
function thinRun(run: number[], width: number): number[] {
  const columns = Math.max(Math.ceil(width / K.thinColumn), 1);
  if (run.length / 2 <= columns * 4) return run;
  // A reversed axis runs right to left: either way, x must not turn back.
  const sign = Math.sign(run[run.length - 2]! - run[0]!);
  for (let i = 2; i < run.length; i += 2) if (Math.sign(run[i]! - run[i - 2]!) === -sign) return run;
  const out: number[] = [];
  let column = Number.NaN;
  let first = 0, last = 0, low = 0, high = 0;
  const flush = (): void => {
    const keep = [...new Set([first, low, high, last])].sort((a, b) => a - b);
    for (const k of keep) out.push(run[k]!, run[k + 1]!);
  };
  for (let i = 0; i < run.length; i += 2) {
    const at = Math.floor(run[i]! / K.thinColumn);
    if (at !== column) {
      if (!Number.isNaN(column)) flush();
      column = at;
      first = low = high = last = i;
      continue;
    }
    last = i;
    // SVG y grows downward: the highest value has the smallest y.
    if (run[i + 1]! > run[low + 1]!) low = i;
    if (run[i + 1]! < run[high + 1]!) high = i;
  }
  flush();
  return out;
}

function paintAreas(scene: Scene, g: Group, c: Cartesian, pending: PendingLabel[], spec: ChartSpecModel): void {
  // An area runs along every category, empty ones at zero; on a date axis it runs through the points that have a date.
  const scale = c.cat.scale;
  const count = scale.kind === "category" && scale.time === undefined ? scale.count : pointCount(g);
  const lower = new Array<number>(count).fill(0);
  const totals = new Array<number>(count).fill(0);
  const stacked = g.grouping !== "standard";
  if (g.grouping === "percentStacked") for (const s of g.series) s.values.forEach((v, i) => { if (i < count) totals[i]! += Math.abs(v ?? 0); });
  const base = baseValue(c);
  // Standard areas: later series paint over earlier ones, so Office draws them in order.
  for (const s of g.series) {
    scene.breakRun();
    const top: number[] = [];
    const bottom: number[] = [];
    // Index of each point in `top`/`bottom`: a point whose date is off a date axis has none.
    const at: Array<number | undefined> = [];
    const extra: ExtraPoint[] = [];
    for (let i = 0; i < count; i++) {
      let v = s.values[i] ?? 0;
      if (g.grouping === "percentStacked") v = totals[i] ? v / totals[i]! : 0;
      const from = stacked ? lower[i]! : base;
      const to = stacked ? from + v : v;
      if (stacked) lower[i] = to;
      const x = slotPos(c, s, i);
      if (x === undefined) continue;
      at[i] = top.length / 2;
      top.push(x, valPos(c, clampValue(c, to)));
      bottom.push(x, valPos(c, clampValue(c, from)));
      if (s.values[i] !== null && s.values[i] !== undefined) extra.push({ x: i + 1, y: s.values[i]!, at: to, px: x });
    }
    if (top.length < 2) continue;
    let d = polyPath(top);
    for (let i = bottom.length - 2; i >= 0; i -= 2) d += `L${fmt(bottom[i]!)} ${fmt(bottom[i + 1]!)}`;
    scene.path(`${d}Z`, s.look);
    for (let i = 0; i < count; i++) {
      const labels = labelsFor(s, i);
      const v = s.values[i];
      const k = at[i];
      if (labels === undefined || v === null || v === undefined || k === undefined) continue;
      const text = labelText(labels, s, i, v, undefined, c.cat.categories[i] ?? "", spec, g);
      placeLabel(pending, labels, text, top[k * 2]!, (top[k * 2 + 1]! + bottom[k * 2 + 1]!) / 2, "ctr", 0, 0);
    }
    paintExtras(scene, g, c, s, extra, pending, spec);
  }
}

function paintBubbles(scene: Scene, g: Group, c: Cartesian, pending: PendingLabel[], spec: ChartSpecModel): void {
  let maxSize = 0;
  for (const s of g.series) for (const b of s.bubbleSizes ?? []) if (b !== null && Math.abs(b) > maxSize) maxSize = Math.abs(b);
  if (maxSize === 0) maxSize = 1;
  const dMax = Math.min(c.inner.w, c.inner.h) * K.bubbleMax * (g.bubbleScale / 100);
  for (const s of g.series) {
    scene.breakRun();
    const extra: ExtraPoint[] = [];
    s.values.forEach((v, i) => {
      const size = s.bubbleSizes?.[i];
      const xv = s.categoryValues?.[i] ?? i + 1;
      if (v === null || xv === null || size === null || size === undefined || size <= 0) return;
      const ratio = size / maxSize;
      const d = dMax * (g.sizeRepresents === "w" ? ratio : Math.sqrt(ratio));
      const x = along(c.cat, c.inner, valueFraction(c.cat.scale as ValueScale, xv));
      const y = valPos(c, v);
      const rr = d / 2;
      extra.push({ x: xv, y: v, at: v, px: x });
      scene.path(`M${fmt(x - rr)} ${fmt(y)}A${fmt(rr)} ${fmt(rr)} 0 1 0 ${fmt(x + rr)} ${fmt(y)}A${fmt(rr)} ${fmt(rr)} 0 1 0 ${fmt(x - rr)} ${fmt(y)}Z`, pointLook(s, i, v));
      const labels = labelsFor(s, i);
      if (labels !== undefined) placeLabel(pending, labels, labelText(labels, s, i, v, undefined, categoryText(c, s, i, true, spec), spec, g), x, y, labels.position ?? "ctr", 0, rr);
    });
    paintExtras(scene, g, c, s, extra, pending, spec);
  }
}

// ------------------------------------------------------------------ Office 2016 chart types

/** Funnel: one bar per stage, centred, as wide as its value is large (the largest spans the plot). */
function paintFunnel(scene: Scene, g: Group, c: Cartesian, pending: PendingLabel[], spec: ChartSpecModel): void {
  const s = g.series[0];
  if (s === undefined) return;
  const scale = c.cat.scale as CategoryScale;
  const bar = (c.inner.h * categorySlot(scale)) / (1 + Math.max(g.gapWidth, 0) / 100);
  let max = 0;
  for (const v of s.values) if (v !== null && v > max) max = v;
  if (max <= 0) return;
  scene.breakRun();
  for (let i = 0; i < scale.count; i++) {
    const v = s.values[i];
    if (v === null || v === undefined || v <= 0) continue;
    const cy = slotPos(c, s, i);
    if (cy === undefined) continue;
    const width = (c.inner.w * v) / max;
    scene.rect(c.inner.x + (c.inner.w - width) / 2, cy - bar / 2, width, bar, pointLook(s, i, v));
    const labels = labelsFor(s, i);
    if (labels !== undefined) {
      pending.push({ text: labelText(labels, s, i, v, undefined, c.cat.categories[i] ?? "", spec, g), x: c.inner.x + c.inner.w / 2, y: cy, look: labels, anchor: "middle", v: "middle" });
    }
  }
}

/** Box and whisker: box from the first to the third quartile, median line, whiskers with caps, mean marker, outliers. */
function paintBoxes(scene: Scene, g: Group, c: Cartesian): void {
  const scale = c.cat.scale as CategoryScale;
  const options = g.box ?? { meanMarker: true, meanLine: false, outliers: true, innerPoints: false };
  const ns = Math.max(g.series.length, 1);
  const slot = c.inner.w * categorySlot(scale);
  const bar = slot / (ns + Math.max(g.gapWidth, 0) / 100);
  const meanPoints: Array<Array<[number, number]>> = g.series.map(() => []);
  g.series.forEach((s, si) => {
    scene.breakRun();
    const fill = s.look.fill ?? "#156082";
    // The reader resolves the outline, whiskers and median to a darker shade of the box.
    const line: Line = s.look.line ?? { color: "#000000", width: 0.75 };
    const ink = line.color;
    (s.boxes ?? []).forEach((b, i) => {
      const at = slotPos(c, s, i);
      if (at === undefined) return;
      const cx = at - (ns * bar) / 2 + (si + 0.5) * bar;
      const half = bar / 2;
      const y = (v: number): number => valPos(c, v);
      scene.rect(cx - half, y(b.q3), bar, y(b.q1) - y(b.q3), { fill, line });
      scene.breakRun();
      const cap = bar * 0.15;
      scene.path(
        `M${fmt(cx - half)} ${fmt(y(b.median))}H${fmt(cx + half)}` +
        `M${fmt(cx)} ${fmt(y(b.q3))}V${fmt(y(b.high))}M${fmt(cx - cap)} ${fmt(y(b.high))}H${fmt(cx + cap)}` +
        `M${fmt(cx)} ${fmt(y(b.q1))}V${fmt(y(b.low))}M${fmt(cx - cap)} ${fmt(y(b.low))}H${fmt(cx + cap)}`,
        { line },
      );
      const dots = options.innerPoints ? (b.inner ?? []).filter((v) => !b.outliers.includes(v)) : [];
      const marks = [...(options.outliers ? b.outliers : []), ...dots];
      if (marks.length > 0) {
        let d = "";
        for (const v of marks) {
          const r = 2;
          d += `M${fmt(cx - r)} ${fmt(y(v))}A${fmt(r)} ${fmt(r)} 0 1 0 ${fmt(cx + r)} ${fmt(y(v))}A${fmt(r)} ${fmt(r)} 0 1 0 ${fmt(cx - r)} ${fmt(y(v))}Z`;
        }
        scene.path(d, { fill: ink });
      }
      meanPoints[si]!.push([cx, y(b.mean)]);
      if (options.meanMarker) scene.path(markerPath("x", cx, y(b.mean), 7).d, { line });
    });
  });
  if (options.meanLine) {
    scene.breakRun();
    for (const [si, pts] of meanPoints.entries()) scene.polyline(pts.flat(), g.series[si]!.look.line ?? { color: "#000000", width: 0.75 });
  }
}

/** Squarified treemap layout (Bruls, Huizing and van Wijk): rows along the shorter side, kept as square as the areas allow. */
function squarify(values: readonly number[], box: Box): Box[] {
  const total = values.reduce((sum, v) => sum + v, 0);
  if (!(total > 0) || box.w <= 0 || box.h <= 0) return values.map(() => ({ x: box.x, y: box.y, w: 0, h: 0 }));
  const areas = values.map((v) => (v / total) * box.w * box.h);
  const out: Box[] = new Array(values.length);
  let { x, y, w, h } = box;
  let i = 0;
  while (i < areas.length) {
    const side = Math.min(w, h);
    // Rows hold areas sorted largest first: the extremes are the row's ends.
    const worst = (sum: number, max: number, min: number): number => Math.max((side * side * max) / (sum * sum), (sum * sum) / (side * side * min));
    let sum = areas[i]!;
    const max = areas[i]!;
    let min = areas[i]!;
    let j = i + 1;
    while (j < areas.length) {
      const next = areas[j]!;
      if (worst(sum + next, max, next) > worst(sum, max, min)) break;
      sum += next;
      min = next;
      j++;
    }
    if (w >= h) {
      const width = sum / h;
      let top = y;
      for (let k = i; k < j; k++) {
        const height = areas[k]! / width;
        out[k] = { x, y: top, w: width, h: height };
        top += height;
      }
      x += width;
      w -= width;
    } else {
      const height = sum / w;
      let left = x;
      for (let k = i; k < j; k++) {
        const width = areas[k]! / height;
        out[k] = { x: left, y, w: width, h: height };
        left += width;
      }
      y += height;
      h -= height;
    }
    i = j;
  }
  return out;
}

function nodeText(labels: DataLabels, series: Series, node: TreeNode, spec: ChartSpecModel): string {
  const parts: string[] = [];
  if (labels.showSeries) parts.push(series.name);
  if (labels.showCategory) parts.push(node.name);
  if (labels.showValue) parts.push(formatValue(node.value, labels.formatCode ?? series.formatCode, spec));
  return parts.join(labels.separator ?? ", ");
}

function treeDepth(nodes: readonly TreeNode[]): number {
  let deepest = 0;
  for (const n of nodes) deepest = Math.max(deepest, n.children === undefined ? 1 : 1 + treeDepth(n.children));
  return deepest;
}

/** Treemap: leaves tile their parents, parents tile the plot; every level sorted largest first. */
function treemapPlot(scene: Scene, spec: ChartSpecModel, g: Group, plot: Box): void {
  const series = g.series[0]!;
  const labels = series.labels;
  const edge = edgeColor(spec);
  const looks = new Map<string, Look>();
  const lookOf = (color: string): Look => {
    let look = looks.get(color);
    if (look === undefined) {
      look = { fill: color, line: { color: edge, width: 1.5 } };
      looks.set(color, look);
    }
    return look;
  };
  const paintLabels: Array<() => void> = [];
  const place = (nodes: readonly TreeNode[], box: Box): void => {
    const rects = squarify(nodes.map((n) => n.value), box);
    nodes.forEach((node, i) => {
      const r = rects[i]!;
      const parent = node.children !== undefined;
      let inner = r;
      const text = labels === undefined || labels.deleted ? "" : nodeText(labels, series, node, spec);
      if (!parent) scene.rect(r.x, r.y, r.w, r.h, lookOf(node.color));
      else if (g.parentLabels === "banner" && text !== "" && labels !== undefined) {
        // Excel's band is two label font sizes tall, the reader's grey (treemap-banner in chartex-oracle.xlsx).
        const band = labels.text.size * K.treemapBanner;
        scene.rect(r.x, r.y, r.w, Math.min(band, r.h), lookOf(g.bannerFill ?? node.color));
        inner = { x: r.x, y: r.y + band, w: r.w, h: Math.max(r.h - band, 0) };
      }
      if (parent) place(node.children!, inner);
      if (text === "" || labels === undefined || (parent && g.parentLabels === "none")) return;
      const look = labels.text;
      const lh = lineHeight(look);
      const shown = fitText(text, look, r.w - 8);
      if (shown === "" || r.h < lh + 4) return;
      // Observed: a leaf's name sits at its bottom-left, a parent's at its top-left, over its first child,
      // or centred in its banner.
      const top = !parent ? r.y + r.h - 2 - lh : g.parentLabels === "banner" ? r.y + (Math.min(look.size * K.treemapBanner, r.h) - lh) / 2 : r.y + 2;
      paintLabels.push(() => scene.text(r.x + 4, top + ascent(look), shown, look, "start"));
    });
  };
  place(g.tree ?? [], plot);
  for (const paint of paintLabels) paint();
}

/** The chart area's own fill is what shows between treemap and sunburst pieces. */
function edgeColor(spec: ChartSpecModel): string {
  return typeof spec.look.fill === "string" ? spec.look.fill : "#FFFFFF";
}

/** Sunburst: one ring per hierarchy level from the hole outward, the first branch starting at twelve o'clock and running clockwise. */
function sunburstPlot(scene: Scene, spec: ChartSpecModel, g: Group, plot: Box): void {
  const tree = g.tree ?? [];
  const depth = treeDepth(tree);
  if (depth === 0) return;
  const series = g.series[0]!;
  const labels = series.labels;
  const radius = Math.min(plot.w, plot.h) / 2;
  // Observed: the hole and every ring are one share of the radius each.
  const step = radius / (depth + 1);
  const cx = plot.x + plot.w / 2;
  const cy = plot.y + plot.h / 2;
  const edge = edgeColor(spec);
  const looks = new Map<string, Look>();
  const lookOf = (color: string): Look => {
    let look = looks.get(color);
    if (look === undefined) {
      look = { fill: color, line: { color: edge, width: 1 } };
      looks.set(color, look);
    }
    return look;
  };
  const paintLabels: Array<() => void> = [];
  const ring = (nodes: readonly TreeNode[], level: number, start: number, sweep: number, whole: number): void => {
    let angle = start;
    for (const node of nodes) {
      const part = whole > 0 ? (node.value / whole) * sweep : 0;
      const ri = step * (level + 1);
      const ro = ri + step;
      if (part > 0) {
        scene.path(sectorPath(cx, cy, ro, ri, angle, part), lookOf(node.color), false);
        const text = labels === undefined || labels.deleted ? "" : nodeText(labels, series, node, spec);
        if (text !== "" && labels !== undefined) {
          const look = labels.text;
          const mid = angle + part / 2;
          const rm = (ri + ro) / 2;
          // Observed: text runs along the radius, turned upright.
          let rotation = mid;
          if (rotation > 90) rotation -= 180;
          if (rotation < -90) rotation += 180;
          const along = (mid * Math.PI) / 180;
          const shown = fitText(text, look, step - 6);
          if (shown !== "" && (rm * part * Math.PI) / 180 >= lineHeight(look)) {
            const x = cx + rm * Math.cos(along);
            const y = cy + rm * Math.sin(along);
            paintLabels.push(() => drawRotated(scene, [shown], x, y, look, rotation));
          }
        }
        if (node.children !== undefined) ring(node.children, level + 1, angle, part, node.value);
      }
      angle += part;
    }
  };
  const total = tree.reduce((sum, n) => sum + n.value, 0);
  ring(tree, 0, -90, 360, total);
  for (const paint of paintLabels) paint();
}

// ------------------------------------------------------------------ axes painting

/** Where a date axis' major ticks fall, as 0..1 fractions: at the start of their slot, or on the point when points sit on ticks. */
function timeBoundaries(plan: AxisPlan): number[] {
  const s = plan.scale as CategoryScale;
  return (plan.timeTicks ?? []).map((t) => (s.between ? t.slot / s.count : categoryFraction(s, t.slot)));
}

function paintGrid(scene: Scene, plan: AxisPlan, inner: Box): void {
  const axis = plan.axis;
  const lines: Array<[Line | undefined, number[]]> = [];
  if (plan.scale.kind === "value") {
    if (axis.minorGrid !== undefined) lines.push([axis.minorGrid, minorTicks(plan.scale).map((v) => valueFraction(plan.scale as ValueScale, v))]);
    if (axis.majorGrid !== undefined) lines.push([axis.majorGrid, majorTicks(plan.scale).map((v) => valueFraction(plan.scale as ValueScale, v))]);
  } else if (plan.timeTicks !== undefined) {
    if (axis.majorGrid !== undefined) lines.push([axis.majorGrid, timeBoundaries(plan)]);
  } else {
    const s = plan.scale;
    const boundaries: number[] = [];
    const n = s.between ? s.count : s.count - 1;
    for (let i = 0; i <= n; i++) boundaries.push(n === 0 ? 0 : i / n);
    if (axis.majorGrid !== undefined) lines.push([axis.majorGrid, boundaries]);
  }
  for (const [line, fs] of lines) {
    if (line === undefined) continue;
    scene.breakRun();
    let d = "";
    for (const f of fs) {
      if (f < -1e-9 || f > 1 + 1e-9) continue;
      const p = along(plan, inner, f);
      d += plan.horizontal ? `M${fmt(p)} ${fmt(inner.y)}V${fmt(inner.y + inner.h)}` : `M${fmt(inner.x)} ${fmt(p)}H${fmt(inner.x + inner.w)}`;
    }
    scene.path(d, { line });
  }
}

function paintAxis(scene: Scene, plan: AxisPlan, inner: Box): void {
  if (!plan.visible) return;
  const axis = plan.axis;
  const pos = axisLinePos(plan, inner);
  const dir = labelDirection(plan);
  // Tick marks.
  const tickFs: number[] = [];
  if (plan.scale.kind === "value") {
    for (const v of majorTicks(plan.scale)) tickFs.push(valueFraction(plan.scale, v));
  } else if (plan.timeTicks !== undefined) {
    tickFs.push(...timeBoundaries(plan));
  } else {
    const s = plan.scale;
    const n = s.between ? s.count : s.count - 1;
    // Office thins tick marks with the labels unless the file sets its own interval.
    const step = axis.markSkip ?? plan.labelSkipAuto;
    for (let i = 0; i <= n; i += step) tickFs.push(n === 0 ? 0 : i / n);
  }
  const line = axis.line;
  if (line !== undefined) {
    scene.breakRun();
    const tickSpan = (kind: Axis["majorTick"], len: number): [number, number] => {
      switch (kind) {
        case "out": return [0, dir * len];
        case "in": return [-dir * len, 0];
        case "cross": return [-len, len];
        default: return [0, 0];
      }
    };
    const [t0, t1] = tickSpan(axis.majorTick, K.tick);
    let d = plan.horizontal
      ? `M${fmt(inner.x)} ${fmt(pos)}H${fmt(inner.x + inner.w)}`
      : `M${fmt(pos)} ${fmt(inner.y)}V${fmt(inner.y + inner.h)}`;
    if (t0 !== t1) {
      for (const f of tickFs) {
        const p = along(plan, inner, f);
        d += plan.horizontal ? `M${fmt(p)} ${fmt(pos + t0)}V${fmt(pos + t1)}` : `M${fmt(pos + t0)} ${fmt(p)}H${fmt(pos + t1)}`;
      }
    }
    if (axis.minorTick !== "none" && plan.scale.kind === "value") {
      const [m0, m1] = tickSpan(axis.minorTick, K.minorTick);
      for (const v of minorTicks(plan.scale)) {
        const p = along(plan, inner, valueFraction(plan.scale, v));
        d += plan.horizontal ? `M${fmt(p)} ${fmt(pos + m0)}V${fmt(pos + m1)}` : `M${fmt(pos + m0)} ${fmt(p)}H${fmt(pos + m1)}`;
      }
    }
    scene.path(d, { line });
  }
}

function paintAxisLabels(scene: Scene, plan: AxisPlan, inner: Box): void {
  if (!plan.visible || (plan.labels.length === 0 && plan.levels.length === 0)) return;
  const axis = plan.axis;
  const look = plan.labelLook;
  const dir = labelDirection(plan);
  const outward = plan.labelAt === "line" && (axis.majorTick === "out" || axis.majorTick === "cross") ? K.tick : 0;
  const origin = labelOrigin(plan, inner) + dir * Math.max(outward, labelGap(plan));
  const rot = plan.labelRotation;
  for (const label of plan.labels) {
    const p = along(plan, inner, label.f);
    if (plan.horizontal) {
      if (rot === 0) {
        drawLines(scene, label.lines, p, origin, look, "middle", dir > 0 ? "top" : "bottom");
      } else {
        // Rotated labels end at the tick: their far end touches the axis side, text runs away from it.
        const { width, height } = blockSize(label.lines, look);
        const rad = deg(rot);
        const cx = p - (dir * (width / 2) * Math.cos(rad)) * Math.sign(rot) * -1;
        const cy = origin + dir * ((width / 2) * Math.abs(Math.sin(rad)) + (height / 2) * Math.abs(Math.cos(rad)));
        drawRotated(scene, label.lines, cx, cy, look, rot);
      }
    } else {
      drawLines(scene, label.lines, origin, p, look, dir > 0 ? "start" : "end", "middle");
    }
  }
  paintGroupRows(scene, plan, inner);
}

/**
 * Outer rows of grouped categories: each group's label centred under its span,
 * and a tick between groups reaching from the axis line through every row up to
 * its own (so a group boundary of an outer row crosses the inner rows too).
 * Rows beside a vertical axis read bottom to top.
 */
function paintGroupRows(scene: Scene, plan: AxisPlan, inner: Box): void {
  if (plan.levels.length === 0) return;
  const look = plan.labelLook;
  const scale = plan.scale as CategoryScale;
  const dir = labelDirection(plan);
  const base = labelOrigin(plan, inner);
  const half = categorySlot(scale) / 2;
  const edge = (index: number): number => Math.min(Math.max(categoryFraction(scale, index) - half, 0), 1);
  const lastEdge = (index: number): number => Math.min(Math.max(categoryFraction(scale, index - 1) + half, 0), 1);
  let offset = plan.levelStart;
  for (const row of plan.levels) {
    const near = base + dir * offset;
    const far = base + dir * (offset + row.height);
    let d = "";
    for (const span of row.spans) {
      const f0 = edge(span.from);
      const f1 = lastEdge(span.to);
      const p0 = along(plan, inner, f0);
      const p1 = along(plan, inner, f1);
      const mid = (p0 + p1) / 2;
      if (plan.horizontal) {
        drawLines(scene, span.lines, mid, near + dir * (K.groupLabelPad / 2), look, "middle", dir > 0 ? "top" : "bottom");
        d += `M${fmt(p0)} ${fmt(base)}V${fmt(far)}M${fmt(p1)} ${fmt(base)}V${fmt(far)}`;
      } else {
        drawRotated(scene, span.lines, (near + far) / 2, mid, look, -90);
        d += `M${fmt(base)} ${fmt(p0)}H${fmt(far)}M${fmt(base)} ${fmt(p1)}H${fmt(far)}`;
      }
    }
    scene.breakRun();
    if (plan.axis.line !== undefined) scene.path(d, { line: plan.axis.line });
    offset += row.height;
  }
}

function paintAxisTitle(scene: Scene, plan: AxisPlan, inner: Box, outer: Box): void {
  const title = plan.axis.title;
  if (title === undefined || !plan.visible) return;
  const vertical = !plan.horizontal;
  const maxLen = vertical ? inner.h : inner.w;
  const block = titleBlock(title, maxLen);
  const dir = labelDirection(plan);
  const labelsEdge = labelOrigin(plan, inner) + dir * plan.thickness;
  if (vertical) {
    const cx = (dir < 0 ? Math.max(labelsEdge - K.axisTitleGap - block.width / 2, outer.x + block.width / 2) : Math.min(labelsEdge + K.axisTitleGap + block.width / 2, outer.x + outer.w - block.width / 2));
    drawRotated(scene, block.lines, cx, inner.y + inner.h / 2, title.text, title.text.rotation ?? -90);
  } else {
    const cy = dir > 0 ? labelsEdge + K.axisTitleGap : labelsEdge - K.axisTitleGap;
    drawLines(scene, block.lines, inner.x + inner.w / 2, cy, title.text, "middle", dir > 0 ? "top" : "bottom");
  }
}

// ------------------------------------------------------------------ pie & doughnut

function piePlot(scene: Scene, spec: ChartSpecModel, groups: Group[], plot: Box, pending: PendingLabel[]): void {
  const g = groups[0]!;
  if (g.kind === "ofPie") {
    ofPiePlot(scene, spec, g, plot, pending);
    return;
  }
  const doughnut = g.kind === "doughnut";
  const series = doughnut ? g.series : g.series.slice(0, 1);
  if (series.length === 0) return;
  const first = series[0]!;
  const anyOutside = series.some((s) => s.values.some((_, i) => {
    const l = labelsFor(s, i);
    return l !== undefined && (l.position === "outEnd" || l.position === "bestFit" || l.position === undefined);
  }));
  // Room for outside labels: one label line plus the gap, all round.
  const labelRoom = anyOutside && !doughnut ? lineHeight(first.labels?.text ?? spec.text) * 1.5 + K.pieLabelGap : 0;
  let maxExplode = 0;
  for (const s of series) {
    maxExplode = Math.max(maxExplode, s.explosion);
    for (const p of Object.values(s.points)) maxExplode = Math.max(maxExplode, p.explosion ?? 0);
  }
  const side = Math.max(Math.min(plot.w, plot.h) - 2 * labelRoom, 1);
  const startRadius = side / 2 / (1 + maxExplode / 100);
  const cx = plot.x + plot.w / 2;
  const cy = plot.y + plot.h / 2;
  const hole = doughnut ? Math.min(Math.max(g.holeSize, 10), 90) / 100 : 0;

  // Slices and their label text first: the angles do not depend on the radius, the label positions do.
  /**
   * A slice and its label: `lines` wrapped at the outside width (`maxW`), with their size; `raw`, the
   * label's own lines unwrapped, and `rawW`, the widest of them, which decide inside or outside.
   */
  interface Slice {
    ring: number; i: number; start: number; sweep: number; look: Look; explosion: number;
    label?: { labels: DataLabels; lines: string[]; w: number; h: number; raw: string[]; rawW: number; maxW: number };
  }
  const rings: Slice[][] = series.map((s, ring) => {
    const total = s.values.reduce<number>((sum, v) => sum + Math.abs(v ?? 0), 0);
    const out: Slice[] = [];
    if (total <= 0) return out;
    const shares = sliceShares(s.values, total, (s.labels ?? g.labels)?.formatCode === undefined);
    const maxW = Math.max(plot.w * K.pieLabelMax, 1);
    let angle = g.firstSliceAngle - 90;
    s.values.forEach((raw, i) => {
      const sweep = (Math.abs(raw ?? 0) / total) * 360;
      if (sweep <= 0) return;
      const slice: Slice = { ring, i, start: angle, sweep, look: s.points[i]?.look ?? s.look, explosion: (s.points[i]?.explosion ?? s.explosion) / 100 };
      const labels = labelsFor(s, i);
      if (labels !== undefined && raw !== null) {
        // A label wider than its share of the plot wraps, then is cut: it must fit beside the pie at some radius.
        const text = labelText(labels, s, i, raw, shares[i], s.categories[i] ?? "", spec, g).split("\n");
        const lines = text.flatMap((l) => wrapFit(l, labels.text, maxW, 3));
        const { width, height } = blockSize(lines, labels.text);
        slice.label = { labels, lines, w: width, h: height, raw: text, rawW: Math.max(...text.map((l) => textWidth(l, labels.text))), maxW };
      }
      out.push(slice);
      angle += sweep;
    });
    return out;
  });

  interface Placed { slice: Slice; mid: number; ox: number; oy: number; outerR: number; x0: number; cy: number; natX0: number; natCy: number; anchor: "start" | "middle" | "end" }
  /** Lay every label out for `radius`: inside ones where they sit, outside ones spread so none overlap. */
  const arrange = (radius: number): { inside: PendingLabel[]; outside: Placed[]; overflow: number } => {
    const ringWidth = series.length > 0 ? (radius * (1 - hole)) / series.length : radius;
    const inside: PendingLabel[] = [];
    const outside: Placed[] = [];
    for (const slices of rings) {
      for (const slice of slices) {
        const label = slice.label;
        if (label === undefined) continue;
        const outerR = radius - slice.ring * ringWidth;
        const innerR = doughnut ? outerR - ringWidth : 0;
        const mid = deg(slice.start + slice.sweep / 2);
        const ox = cx + slice.explosion * outerR * Math.cos(mid);
        const oy = cy + slice.explosion * outerR * Math.sin(mid);
        const { labels, raw } = label;
        let { lines, w, h } = label;
        const pos = labels.position ?? "bestFit";
        // Inside when the label, unwrapped, is no wider than its slice at the rim, as Office judges it
        // (PowerPoint keeps a 76pt Arabic name in a 36 degree slice and moves a 120pt one out).
        const half = Math.min(deg(slice.sweep) / 2, Math.PI / 2);
        const chord = 2 * outerR * Math.sin(half);
        const fits = label.rawW <= chord && h <= outerR * 0.6;
        if (fits && pos !== "outEnd") {
          // Inside, it wraps to the slice's width near where it sits (`pieInsideWrapAt`): a 76pt Arabic
          // name breaks there in PowerPoint's 36 degree slice, a 56pt one in a 32 degree slice does not.
          const room = Math.min(label.maxW, 2 * outerR * K.pieInsideWrapAt * Math.sin(half));
          lines = raw.flatMap((l) => wrapFit(l, labels.text, Math.max(room, 1), 3));
          ({ width: w, height: h } = blockSize(lines, labels.text));
        }
        // Office has no outside labels on a doughnut; a pie label that does not fit its slice moves out.
        const out = !doughnut && (pos === "outEnd" || !fits);
        if (out) {
          const rr = outerR + K.pieLabelGap;
          const c = Math.cos(mid);
          const sn = Math.sin(mid);
          const lx = ox + rr * c;
          const anchor = Math.abs(c) < 0.2 ? "middle" : c > 0 ? "start" : "end";
          const x0 = anchor === "start" ? lx : anchor === "end" ? lx - w : lx - w / 2;
          const ly = oy + rr * sn + (sn * h) / 2;
          outside.push({ slice, mid, ox, oy, outerR, x0, cy: ly, natX0: x0, natCy: ly, anchor });
        } else {
          // A best-fit label that fits inside sits where an inside-end one does, against the rim.
          const rr = pos === "inEnd" || (pos === "bestFit" && !doughnut) ? outerR - Math.max(w, h) / 2 - K.dataLabelGap : pos === "ctr" && !doughnut ? outerR / 2 : doughnut ? (outerR + innerR) / 2 : outerR * 0.6;
          const label: PendingLabel = { text: lines.join("\n"), x: ox + rr * Math.cos(mid), y: oy + rr * Math.sin(mid), look: labels, anchor: "middle", v: "middle" };
          // A label dragged off its slice keeps a leader line to the rim.
          if (labels.leaderLines !== undefined && labels.offset !== undefined) {
            label.leader = { x: ox + outerR * Math.cos(mid), y: oy + outerR * Math.sin(mid), line: labels.leaderLines, centre: { x: ox, y: oy, r: outerR } };
          }
          inside.push(label);
        }
      }
    }
    // Spread the outside labels so none overlaps another: push down from the top, pull up from the
    // bottom, twice. A label moved along its column keeps clear of the circle (its near edge follows
    // the rim), so a move can change which labels share a horizontal extent: settle() updates x too.
    const top = plot.y;
    const bottom = plot.y + plot.h;
    const settle = (p: Placed, y: number): void => {
      p.cy = y;
      if (p.anchor === "middle" || Math.abs(y - p.natCy) <= 0.01) return;
      const { w, h } = p.slice.label!;
      const dy = Math.max(Math.abs(y - p.oy) - h / 2, 0);
      const reach = Math.sqrt(Math.max((p.outerR + K.pieLabelGap) ** 2 - dy * dy, 0));
      p.x0 = p.anchor === "start" ? Math.max(p.natX0, p.ox + reach) : Math.min(p.natX0, p.ox - reach - w);
    };
    const across = (a: Placed, b: Placed): boolean => a.x0 < b.x0 + b.slice.label!.w && b.x0 < a.x0 + a.slice.label!.w;
    const order = [...outside].sort((a, b) => a.natCy - b.natCy);
    for (let round = 0; round < 4; round++) {
      order.forEach((p, k) => {
        for (let again = 0; again < 3; again++) {
          let min = top + p.slice.label!.h / 2;
          for (let j = 0; j < k; j++) if (across(order[j]!, p)) min = Math.max(min, order[j]!.cy + (order[j]!.slice.label!.h + p.slice.label!.h) / 2 + K.pieLabelSpacing);
          if (p.cy >= min - 0.01) break;
          settle(p, min);
        }
      });
      for (let k = order.length - 1; k >= 0; k--) {
        const p = order[k]!;
        for (let again = 0; again < 3; again++) {
          let max = bottom - p.slice.label!.h / 2;
          for (let j = k + 1; j < order.length; j++) if (across(p, order[j]!)) max = Math.min(max, order[j]!.cy - (order[j]!.slice.label!.h + p.slice.label!.h) / 2 - K.pieLabelSpacing);
          if (p.cy <= max + 0.01) break;
          settle(p, max);
        }
      }
    }
    // Too many labels for the plot's height cannot all clear each other: stay inside the plot at least.
    for (const p of order) settle(p, Math.min(Math.max(p.cy, top + p.slice.label!.h / 2), bottom - p.slice.label!.h / 2));
    let overflow = 0;
    for (const p of outside) overflow = Math.max(overflow, plot.x - p.x0, p.x0 + p.slice.label!.w - (plot.x + plot.w));
    return { inside, outside, overflow };
  };

  // Office shrinks the pie to make room for labels beside it: step the radius in until they fit the plot.
  let radius = startRadius;
  let arranged = arrange(radius);
  const floor = startRadius * K.pieMinRadius;
  for (let pass = 0; pass < 8 && arranged.overflow > 0.5 && radius > floor; pass++) {
    radius = Math.max(radius - arranged.overflow, floor);
    arranged = arrange(radius);
  }

  const ringWidth = series.length > 0 ? (radius * (1 - hole)) / series.length : radius;
  rings.forEach((slices, ring) => {
    scene.breakRun();
    const outerR = radius - ring * ringWidth;
    const innerR = doughnut ? outerR - ringWidth : 0;
    for (const slice of slices) {
      const mid = deg(slice.start + slice.sweep / 2);
      scene.path(sectorPath(cx + slice.explosion * outerR * Math.cos(mid), cy + slice.explosion * outerR * Math.sin(mid), outerR, innerR, slice.start, slice.sweep), slice.look, false);
    }
  });
  pending.push(...arranged.inside);
  for (const p of arranged.outside) {
    const { labels, lines, w, h } = p.slice.label!;
    const x = p.anchor === "start" ? p.x0 : p.anchor === "end" ? p.x0 + w : p.x0 + w / 2;
    const rim = { x: p.ox + p.outerR * Math.cos(p.mid), y: p.oy + p.outerR * Math.sin(p.mid) };
    const label: PendingLabel = { text: lines.join("\n"), x, y: p.cy, look: labels, anchor: p.anchor, v: "middle" };
    // A label dragged by hand gets its leader where it is painted (`paintPendingLabels`).
    if (labels.leaderLines !== undefined && labels.offset !== undefined) label.leader = { ...rim, line: labels.leaderLines, centre: { x: p.ox, y: p.oy, r: p.outerR } };
    pending.push(label);
    // A leader line joins a label the spreading moved off its slice to the slice's rim.
    if (labels.leaderLines !== undefined && labels.offset === undefined && Math.hypot(p.cy - p.natCy, p.x0 - p.natX0) > K.pieLeaderMin) {
      const rimX = rim.x;
      const rimY = rim.y;
      const toX = p.anchor === "start" ? p.x0 - K.pieLeaderGap : p.anchor === "end" ? p.x0 + w + K.pieLeaderGap : p.x0 + w / 2;
      const toY = p.anchor === "middle" ? p.cy + (Math.sin(p.mid) > 0 ? -h / 2 - K.pieLeaderGap : h / 2 + K.pieLeaderGap) : p.cy;
      scene.breakRun();
      scene.segment(rimX, rimY, toX, toY, labels.leaderLines);
    }
  }
}

/** One slice of an of-pie ring: the series point it stands for, or the "Other" slice (`index` past the points, with its own look). */
interface RingSlice { index: number; value: number | null; category: string; look?: Look }

/**
 * The slices of one of-pie ring, clockwise from `ring.start` (SVG degrees: 0
 * is 3 o'clock). A slice spans its share of `total`; `shares` holds the number
 * each label shows as its percentage, parallel to `slices`.
 */
function paintSlices(
  scene: Scene, spec: ChartSpecModel, g: Group, s: Series, slices: readonly RingSlice[], total: number, shares: readonly number[],
  ring: { cx: number; cy: number; outerR: number; innerR: number; start: number }, pending: PendingLabel[],
): void {
  const { cx, cy, outerR, innerR } = ring;
  let angle = ring.start;
  slices.forEach((slice, k) => {
    const raw = slice.value;
    const i = slice.index;
    const v = Math.abs(raw ?? 0);
    const sweep = (v / total) * 360;
    if (sweep <= 0) return;
    const explosion = ((s.points[i]?.explosion ?? s.explosion) / 100) * outerR;
    const mid = deg(angle + sweep / 2);
    const ox = cx + explosion * Math.cos(mid);
    const oy = cy + explosion * Math.sin(mid);
    const look = slice.look ?? s.points[i]?.look ?? s.look;
    scene.path(sectorPath(ox, oy, outerR, innerR, angle, sweep), look, false);
    const labels = labelsFor(s, i);
    if (labels !== undefined && raw !== null) {
      const text = labelText(labels, s, i, raw, shares[k], slice.category, spec, g);
      const lines = text.split("\n");
      const { width, height } = blockSize(lines, labels.text);
      const pos = labels.position ?? "bestFit";
      let fitsInside = true;
      if (pos === "bestFit") {
        // Inside when the label's box fits within the slice near its centroid.
        const rr = outerR * 0.6;
        const chord = 2 * rr * Math.sin(Math.min(deg(sweep) / 2, Math.PI / 2));
        fitsInside = width <= chord * 1.1 && height <= outerR * 0.6;
      }
      if (pos === "outEnd" || (pos === "bestFit" && !fitsInside)) {
        const rr = outerR + K.pieLabelGap;
        const c = Math.cos(mid);
        const sn = Math.sin(mid);
        pending.push({ text, x: ox + rr * c, y: oy + rr * sn + (sn * height) / 2, look: labels, anchor: Math.abs(c) < 0.2 ? "middle" : c > 0 ? "start" : "end", v: "middle" });
      } else {
        // A best-fit label that fits inside sits where an inside-end one does, against the rim.
        const rr = pos === "inEnd" || pos === "bestFit" ? outerR - Math.max(width, height) / 2 - K.dataLabelGap : pos === "ctr" ? outerR / 2 : outerR * 0.6;
        pending.push({ text, x: ox + rr * Math.cos(mid), y: oy + rr * Math.sin(mid), look: labels, anchor: "middle", v: "middle" });
      }
    }
    angle += sweep;
  });
}

// ------------------------------------------------------------------ pie of pie, bar of pie

/**
 * A pie with the points that moved to the second plot replaced by one "Other"
 * slice, and that second plot (a pie, or one stacked bar) to its right. The
 * rules are Excel's, measured:
 *  - the Other slice is centred on 3 o'clock, so the main pie starts at 90
 *    degrees plus half of Other's sweep; the second pie starts at the same angle;
 *  - the plots sit side by side, centred as a block, both on one centre line;
 *    the second pie is `secondSize` percent of the main one, the bar that tall
 *    and half as wide;
 *  - the gap is `gapWidth` percent of the main radius (of half of it, for a bar);
 *  - the series lines run from the Other slice's rim corners to the bar's
 *    top and bottom left corners, or tangent to the second pie;
 *  - labels show the share of the whole series, Other's included.
 */
function ofPiePlot(scene: Scene, spec: ChartSpecModel, g: Group, plot: Box, pending: PendingLabel[]): void {
  const s = g.series[0];
  const of = g.ofPie;
  if (s === undefined || of === undefined) return;
  const n = s.values.length;
  const total = s.values.reduce<number>((sum, v) => sum + Math.abs(v ?? 0), 0);
  if (total <= 0) return;
  const moved = new Array<boolean>(n).fill(false);
  for (const i of of.secondary) if (i >= 0 && i < n) moved[i] = true;
  const second = s.values.flatMap((_, i) => (moved[i] ? [i] : []));
  const secondTotal = second.reduce((sum, i) => sum + Math.abs(s.values[i] ?? 0), 0);
  const hasSecond = secondTotal > 0;
  const adjust = (s.labels ?? g.labels)?.formatCode === undefined;

  const anyOutside = s.values.some((_, i) => {
    const l = labelsFor(s, i);
    return l !== undefined && (l.position === "outEnd" || l.position === "bestFit" || l.position === undefined);
  });
  const labelRoom = anyOutside ? lineHeight(s.labels?.text ?? spec.text) * 1.5 + K.pieLabelGap : 0;
  const size = of.secondSize / 100;
  const gap = Math.min(Math.max(g.gapWidth, 0), 500) / 100;
  const bar = of.type === "bar";
  // Width of everything in main diameters: the pie, the gap, the second plot.
  const span = hasSecond ? 1 + gap * (bar ? K.ofPieGapBar : K.ofPieGapPie) / 2 + (bar ? size * K.ofPieBarAspect : size) : 1;
  const availW = Math.max(plot.w - 2 * (K.ofPieInset + labelRoom), 1);
  const availH = Math.max(plot.h - 2 * labelRoom, 1);
  const diameter = Math.max(Math.min(availH / (hasSecond ? Math.max(1, size) : 1), availW / span), 1);
  const r1 = diameter / 2;
  const cy = plot.y + plot.h / 2;
  const left = plot.x + plot.w / 2 - (diameter * span) / 2;
  const cx1 = left + r1;

  const mainIndices = s.values.flatMap((_, i) => (moved[i] ? [] : [i]));
  const main: RingSlice[] = mainIndices.map((index) => ({ index, value: s.values[index] ?? null, category: s.categories[index] ?? "" }));
  const otherSweep = hasSecond ? (secondTotal / total) * 360 : 0;
  const start = otherSweep / 2;
  if (hasSecond) main.push({ index: n, value: secondTotal, category: "Other", look: of.other });
  scene.breakRun();
  const mainShares = sliceShares(main.map((m) => m.value), total, adjust);
  paintSlices(scene, spec, g, s, main, total, mainShares, { cx: cx1, cy, outerR: r1, innerR: 0, start }, pending);
  if (!hasSecond) return;

  // The second plot, and where the series lines meet it.
  const wholeShare = (i: number): number => Math.abs(s.values[i] ?? 0) / total;
  const slices: RingSlice[] = second.map((index) => ({ index, value: s.values[index] ?? null, category: s.categories[index] ?? "" }));
  const shares = slices.map((sl) => wholeShare(sl.index));
  const gapPx = gap * (bar ? K.ofPieGapBar : K.ofPieGapPie) * r1;
  const rim = (a: number): [number, number] => [cx1 + r1 * Math.cos(deg(a)), cy + r1 * Math.sin(deg(a))];
  const corners: Array<[number, number]> = [];
  scene.breakRun();
  if (bar) {
    const h = diameter * size;
    const w = h * K.ofPieBarAspect;
    const x = cx1 + r1 + gapPx;
    let y = cy - h / 2;
    for (const sl of slices) {
      const part = (Math.abs(sl.value ?? 0) / secondTotal) * h;
      if (part <= 0) continue;
      const look = s.points[sl.index]?.look ?? s.look;
      scene.rect(x, y, w, part, look, false);
      const labels = labelsFor(s, sl.index);
      if (labels !== undefined && sl.value !== null) {
        const text = labelText(labels, s, sl.index, sl.value, wholeShare(sl.index), sl.category, spec, g);
        if (labels.position === "outEnd") pending.push({ text, x: x + w + K.pieLabelGap, y: y + part / 2, look: labels, anchor: "start", v: "middle" });
        else pending.push({ text, x: x + w / 2, y: y + part / 2, look: labels, anchor: "middle", v: "middle" });
      }
      y += part;
    }
    corners.push([x, cy - h / 2], [x, cy + h / 2]);
  } else {
    const r2 = (diameter * size) / 2;
    const cx2 = cx1 + r1 + gapPx + r2;
    paintSlices(scene, spec, g, s, slices, secondTotal, shares, { cx: cx2, cy, outerR: r2, innerR: 0, start }, pending);
    // Outer tangents from the Other slice's rim corners to the second pie.
    for (const sign of [-1, 1]) {
      const [px, py] = rim(sign * (otherSweep / 2));
      const d = Math.hypot(cx2 - px, cy - py);
      if (d <= r2) continue;
      const along = Math.atan2(cy - py, cx2 - px) + sign * Math.asin(r2 / d);
      const len = Math.sqrt(d * d - r2 * r2);
      corners.push([px + len * Math.cos(along), py + len * Math.sin(along)]);
    }
  }
  if (g.serLines !== undefined && corners.length === 2) {
    scene.breakRun();
    // The upper corner joins the Other slice's upper rim corner (SVG angle -sweep/2).
    corners.forEach(([x, y], k) => {
      const [px, py] = rim((k === 0 ? -1 : 1) * (otherSweep / 2));
      scene.segment(px, py, x, y, g.serLines);
    });
  }
}

function sectorPath(cx: number, cy: number, ro: number, ri: number, start: number, sweep: number): string {
  if (sweep >= 359.999) {
    const full = `M${fmt(cx - ro)} ${fmt(cy)}A${fmt(ro)} ${fmt(ro)} 0 1 1 ${fmt(cx + ro)} ${fmt(cy)}A${fmt(ro)} ${fmt(ro)} 0 1 1 ${fmt(cx - ro)} ${fmt(cy)}Z`;
    return ri > 0 ? `${full}M${fmt(cx - ri)} ${fmt(cy)}A${fmt(ri)} ${fmt(ri)} 0 1 0 ${fmt(cx + ri)} ${fmt(cy)}A${fmt(ri)} ${fmt(ri)} 0 1 0 ${fmt(cx - ri)} ${fmt(cy)}Z` : full;
  }
  const a0 = deg(start);
  const a1 = deg(start + sweep);
  const large = sweep > 180 ? 1 : 0;
  const pt = (rr: number, a: number): string => `${fmt(cx + rr * Math.cos(a))} ${fmt(cy + rr * Math.sin(a))}`;
  if (ri <= 0) return `M${fmt(cx)} ${fmt(cy)}L${pt(ro, a0)}A${fmt(ro)} ${fmt(ro)} 0 ${large} 1 ${pt(ro, a1)}Z`;
  return `M${pt(ro, a0)}A${fmt(ro)} ${fmt(ro)} 0 ${large} 1 ${pt(ro, a1)}L${pt(ri, a1)}A${fmt(ri)} ${fmt(ri)} 0 ${large} 0 ${pt(ri, a0)}Z`;
}

// ------------------------------------------------------------------ radar

function radarPlot(scene: Scene, spec: ChartSpecModel, g: Group, plot: Box, pending: PendingLabel[]): void {
  const catAxis = g.axisIds === undefined ? undefined : axisById(spec, g.axisIds[0]);
  const valAxis = g.axisIds === undefined ? undefined : axisById(spec, g.axisIds[1]);
  const cats = categoryLabels([g], catAxis ?? { kind: "cat", formatLinked: true, formatCode: "General" }, spec);
  const n = Math.max(cats.length, 1);
  const catLook = catAxis?.text ?? spec.text;
  const lh = lineHeight(catLook);
  const widest = Math.max(0, ...cats.map((c) => textWidth(c, catLook)));
  const radius = Math.max(Math.min(plot.w / 2 - widest - K.labelGap, plot.h / 2 - lh - K.labelGap), 4);
  const cx = plot.x + plot.w / 2;
  const cy = plot.y + plot.h / 2;
  const [lo, hi] = valueExtent([g], false);
  const fit = Math.floor(radius / ((valAxis?.text ?? spec.text).size * K.valueTickSpacing));
  const scale = valueScale(valAxis, Math.min(lo, 0), hi, Math.max(1, Math.min(fit, K.maxValueIntervals)), false, false);
  const angleOf = (i: number): number => deg(-90 + (360 * i) / n);
  const point = (i: number, v: number): [number, number] => {
    const rr = valueFraction(scale, v) * radius;
    return [cx + rr * Math.cos(angleOf(i)), cy + rr * Math.sin(angleOf(i))];
  };
  // Web: value gridlines as polygons, category spokes.
  const grid = valAxis?.majorGrid;
  if (grid !== undefined) {
    let d = "";
    for (const v of majorTicks(scale)) {
      if (v === scale.min) continue;
      for (let i = 0; i <= n; i++) {
        const [x, y] = point(i % n, v);
        d += `${i === 0 ? "M" : "L"}${fmt(x)} ${fmt(y)}`;
      }
    }
    scene.path(d, { line: grid });
  }
  const spoke = catAxis?.majorGrid ?? catAxis?.line;
  if (spoke !== undefined && catAxis?.deleted !== true) {
    let d = "";
    for (let i = 0; i < n; i++) {
      const [x, y] = point(i, scale.max);
      d += `M${fmt(cx)} ${fmt(cy)}L${fmt(x)} ${fmt(y)}`;
    }
    scene.path(d, { line: spoke });
  }
  scene.breakRun();
  for (const s of g.series) {
    scene.breakRun();
    const pts: number[] = [];
    s.values.forEach((v, i) => {
      if (v === null) return;
      pts.push(...point(i, v));
    });
    if (pts.length < 4) continue;
    if (g.radarStyle === "filled") scene.path(`${polyPath(pts)}Z`, s.look);
    else scene.polyline(pts, s.look.line, true);
    s.values.forEach((v, i) => {
      if (v === null) return;
      const [x, y] = point(i, v);
      paintMarker(scene, x, y, s.points[i]?.marker ?? s.marker, s.look);
      const labels = labelsFor(s, i);
      if (labels !== undefined) placeLabel(pending, labels, labelText(labels, s, i, v, undefined, cats[i] ?? "", spec, g), x, y, labels.position ?? "t", K.dataLabelGap, s.marker.size / 2);
    });
  }
  if (catAxis !== undefined && !catAxis.deleted && catAxis.labelPosition !== "none") {
    cats.forEach((c, i) => {
      const a = angleOf(i);
      const x = cx + (radius + K.labelGap) * Math.cos(a);
      const y = cy + (radius + K.labelGap) * Math.sin(a);
      const cos = Math.cos(a);
      const sin = Math.sin(a);
      drawLines(scene, [c], x, y, catLook, Math.abs(cos) < 0.1 ? "middle" : cos > 0 ? "start" : "end", Math.abs(sin) < 0.1 ? "middle" : sin > 0 ? "top" : "bottom");
    });
  }
  if (valAxis !== undefined && !valAxis.deleted && valAxis.labelPosition !== "none") {
    const code = valAxis.formatLinked ? g.series[0]?.formatCode ?? "General" : valAxis.formatCode;
    for (const v of majorTicks(scale)) {
      const [, y] = point(0, v);
      drawLines(scene, [formatValue(v, code, spec)], cx - K.labelGap, y, valAxis.text, "end", "middle");
    }
  }
}

// ------------------------------------------------------------------ entry

function manualBox(l: ManualLayout, w: number, h: number, auto: Box): Box {
  const x = l.x === undefined ? auto.x : l.xMode === "edge" ? l.x * w : auto.x + l.x * w;
  const y = l.y === undefined ? auto.y : l.yMode === "edge" ? l.y * h : auto.y + l.y * h;
  const bw = l.w === undefined ? auto.w : l.w * w;
  const bh = l.h === undefined ? auto.h : l.h * h;
  // A manual box past the frame is cut to it: nothing is drawn outside the chart area.
  const cx = Math.min(Math.max(x, 0), w);
  const cy = Math.min(Math.max(y, 0), h);
  return { x: cx, y: cy, w: Math.max(Math.min(bw, w - cx), 0), h: Math.max(Math.min(bh, h - cy), 0) };
}

/** `box` cut to what is left of `frame` after `margin` all round: room for the axis line's ticks. */
function insetBox(box: Box, frame: Box, margin: number): Box {
  const x = Math.max(box.x, frame.x + margin);
  const y = Math.max(box.y, frame.y + margin);
  return { x, y, w: Math.max(Math.min(box.x + box.w, frame.x + frame.w - margin) - x, 4), h: Math.max(Math.min(box.y + box.h, frame.y + frame.h - margin) - y, 4) };
}

export interface DrawnChart {
  primitives: ChartPrimitive[];
  plot: Box;
}

export function drawChart(spec: ChartSpecModel, width: number, height: number): DrawnChart {
  const scene = new Scene();
  const frame: Box = { x: 0, y: 0, w: width, h: height };
  // Chart area: Office strokes the outline inside the frame.
  if (spec.look.fill !== undefined || spec.look.line !== undefined) {
    const inset = spec.look.line === undefined ? 0 : spec.look.line.width / 2;
    if (spec.roundedCorners) {
      const rr = Math.min(width, height) * 0.06;
      const x0 = inset, y0 = inset, x1 = width - inset, y1 = height - inset;
      scene.path(`M${fmt(x0 + rr)} ${fmt(y0)}H${fmt(x1 - rr)}A${fmt(rr)} ${fmt(rr)} 0 0 1 ${fmt(x1)} ${fmt(y0 + rr)}V${fmt(y1 - rr)}A${fmt(rr)} ${fmt(rr)} 0 0 1 ${fmt(x1 - rr)} ${fmt(y1)}H${fmt(x0 + rr)}A${fmt(rr)} ${fmt(rr)} 0 0 1 ${fmt(x0)} ${fmt(y1 - rr)}V${fmt(y0 + rr)}A${fmt(rr)} ${fmt(rr)} 0 0 1 ${fmt(x0 + rr)} ${fmt(y0)}Z`, spec.look, false);
    } else {
      scene.rect(inset, inset, width - 2 * inset, height - 2 * inset, spec.look, false);
    }
  }
  let area: Box = { x: K.pad, y: K.pad, w: width - 2 * K.pad, h: height - 2 * K.pad };

  // Title band.
  let titleDraw: (() => void) | undefined;
  if (spec.title !== undefined) {
    const title = spec.title;
    const block = titleBlock(title, area.w * 0.8);
    // The title is a text box: its text sits inside an inset above and below.
    const inset = title.text.size * K.titleInset;
    const boxH = block.height + 2 * inset;
    const auto: Box = { x: area.x + (area.w - block.width) / 2, y: area.y, w: block.width, h: boxH };
    const placed = title.layout === undefined ? auto : { ...manualBox(title.layout, width, height, auto), w: block.width, h: boxH };
    // The title box stays in the frame (its 3pt side margins included).
    const box = { ...placed, x: Math.min(Math.max(placed.x, 3), Math.max(width - placed.w - 3, 3)), y: Math.min(Math.max(placed.y, 0), Math.max(height - placed.h, 0)) };
    titleDraw = () => {
      scene.rect(box.x - 3, box.y, box.w + 6, box.h, title.look, false);
      drawLines(scene, block.lines, box.x + box.w / 2, box.y + inset, title.text, "middle", "top");
    };
    if (!title.overlay && title.layout === undefined) {
      area = { ...area, y: area.y + boxH + K.titleGap, h: area.h - boxH - K.titleGap };
    }
  }

  // Legend band.
  const legendLayout = layoutLegend(spec, area, width, height);
  if (legendLayout !== undefined && spec.legend !== undefined && !spec.legend.overlay && spec.legend.layout === undefined) {
    const b = legendLayout.box;
    switch (spec.legend.position) {
      case "r": area = { ...area, w: area.w - b.w - K.legendGap }; break;
      // Top right: Office gives the plot neither the legend's column nor the band beside it.
      case "tr": area = { x: area.x, y: b.y + b.h, w: area.w - b.w - K.legendGap, h: area.y + area.h - b.y - b.h }; break;
      case "l": area = { ...area, x: area.x + b.w + K.legendGap, w: area.w - b.w - K.legendGap }; break;
      case "t": area = { ...area, y: area.y + b.h + K.legendGap, h: area.h - b.h - K.legendGap }; break;
      default: area = { ...area, h: area.h - b.h - K.legendGap };
    }
  }

  // Plot area.
  const pending: PendingLabel[] = [];
  const plans = planAxes(spec);
  const radial = spec.groups.filter((g) => isRadialKind(g.kind));
  const radar = spec.groups.find((g) => g.kind === "radar");
  const tree = spec.groups.find((g) => g.tree !== undefined);
  let plotBox: Box = area;
  if (spec.plot.layout !== undefined && spec.plot.layout.target === "outer") plotBox = manualBox(spec.plot.layout, width, height, area);

  if (radial.length > 0 || radar !== undefined || plans.length === 0) {
    const plot = spec.plot.layout !== undefined ? manualBox(spec.plot.layout, width, height, area) : plotBox;
    scene.rect(plot.x, plot.y, plot.w, plot.h, spec.plot.look, false);
    if (radial.length > 0) piePlot(scene, spec, radial, plot, pending);
    else if (radar !== undefined) radarPlot(scene, spec, radar, plot, pending);
    else if (tree !== undefined) {
      if (tree.kind === "sunburst") sunburstPlot(scene, spec, tree, plot);
      else treemapPlot(scene, spec, tree, plot);
    }
    paintPendingLabels(scene, pending, frame);
    titleDraw?.();
    if (legendLayout !== undefined) paintLegend(scene, spec, legendLayout);
    confineText(scene.primitives, width, height);
    return { primitives: scene.primitives, plot };
  }

  // Cartesian: settle the inner plot box by iterating axis fitting.
  const manualInner = spec.plot.layout !== undefined && spec.plot.layout.target === "inner" ? manualBox(spec.plot.layout, width, height, area) : undefined;
  let inner: Box = manualInner === undefined ? { ...plotBox } : insetBox(manualInner, frame, K.tick);
  for (let pass = 0; pass < 3; pass++) {
    for (const plan of plans) fitAxis(plan, plan.horizontal ? inner.w : inner.h, spec, width);
    for (const plan of plans) {
      plan.crossF = crossFraction(plan);
      plan.side = plan.crossF >= 1 - 1e-9 ? 1 : -1;
      plan.labelAt = plan.axis.labelPosition === "low" ? "low" : plan.axis.labelPosition === "high" ? "high" : "line";
      if (plan.labelAt === "high") plan.side = 1;
      if (plan.labelAt === "low") plan.side = -1;
    }
    if (manualInner !== undefined) break;
    // Reserve, per side of the inner box, what the axes on that side need.
    const need = { l: 0, r: 0, t: 0, b: 0 };
    const over = { l: 0, r: 0, t: 0, b: 0 };
    for (const plan of plans) {
      if (!plan.visible) continue;
      const dir = labelDirection(plan);
      const atEdge = plan.labelAt !== "line" || plan.crossF <= 1e-9 || plan.crossF >= 1 - 1e-9 || plan.cross === undefined;
      let thick = atEdge ? plan.thickness : 0;
      if (plan.axis.title !== undefined) {
        const tb = titleBlock(plan.axis.title, plan.horizontal ? inner.w : inner.h);
        thick += (plan.horizontal ? tb.height : tb.width) + K.axisTitleGap;
      }
      if (plan.horizontal) {
        if (dir > 0) need.b = Math.max(need.b, thick); else need.t = Math.max(need.t, thick);
        // measureAxis already counts from the pixel left, reversed or not.
        over.l = Math.max(over.l, plan.overhangStart);
        over.r = Math.max(over.r, plan.overhangEnd);
      } else {
        if (dir > 0) need.r = Math.max(need.r, thick); else need.l = Math.max(need.l, thick);
        over.t = Math.max(over.t, plan.overhangEnd);
        over.b = Math.max(over.b, plan.overhangStart);
      }
    }
    const l = Math.max(need.l, over.l);
    // Office keeps a little more room on the right, unless axis labels already sit there.
    const rgt = need.r > 0 ? Math.max(need.r, over.r) : over.r + K.plotRightInset;
    const t = Math.max(need.t, over.t);
    const b = Math.max(need.b, over.b);
    inner = { x: plotBox.x + l, y: plotBox.y + t, w: Math.max(plotBox.w - l - rgt, 4), h: Math.max(plotBox.h - t - b, 4) };
  }

  scene.rect(inner.x, inner.y, inner.w, inner.h, spec.plot.look, false);
  for (const plan of plans) if (plan.role === "cat" || plan.scale.kind === "category") paintGrid(scene, plan, inner);
  for (const plan of plans) if (plan.role === "val" && plan.scale.kind === "value") paintGrid(scene, plan, inner);

  // Series, in the order Office stacks chart groups: areas, bars, then lines and XY on top.
  const rank = (g: Group): number => (g.kind === "area" ? 0 : g.kind === "bar" ? 1 : g.kind === "bubble" ? 2 : 3);
  for (const g of [...spec.groups].sort((a, b) => rank(a) - rank(b))) {
    if (g.axisIds === undefined) continue;
    const xPlan = plans.find((p) => p.axis.id === g.axisIds![0]);
    const yPlan = plans.find((p) => p.axis.id === g.axisIds![1]);
    if (xPlan === undefined || yPlan === undefined) continue;
    const c: Cartesian = { inner, cat: xPlan, val: yPlan };
    switch (g.kind) {
      case "bar": paintBars(scene, g, c, pending, spec); break;
      case "area": paintAreas(scene, g, c, pending, spec); break;
      case "bubble": paintBubbles(scene, g, c, pending, spec); break;
      case "funnel": paintFunnel(scene, g, c, pending, spec); break;
      case "box": paintBoxes(scene, g, c); break;
      default: paintLines(scene, g, c, pending, spec);
    }
  }
  for (const plan of plans) paintAxis(scene, plan, inner);
  for (const plan of plans) paintAxisLabels(scene, plan, inner);
  for (const plan of plans) paintAxisTitle(scene, plan, inner, plotBox);
  paintPendingLabels(scene, pending, frame);
  titleDraw?.();
  if (legendLayout !== undefined) paintLegend(scene, spec, legendLayout);
  confineText(scene.primitives, width, height);
  return { primitives: scene.primitives, plot: inner };
}
