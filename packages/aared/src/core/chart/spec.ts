/**
 * The resolved chart model: what a DrawingML `c:chartSpace` *means*, with
 * every inherited or defaulted property already resolved, and nothing left
 * that depends on the file format it came from. The Worker reads it
 * (`read.ts`) for PPTX and XLSX alike and sends it in the unit model; the page
 * lays it out and paints it (`draw.ts`, loaded the first time a chart shows).
 *
 * Plain data only (no `Map`, `Set` or class instances), so it crosses the
 * Worker boundary and `JSON.stringify` alike.
 *
 * Units: lengths in points, angles in degrees clockwise, colors as CSS strings
 * (alpha folded in), fractions as 0..1.
 */

import type { ShapeFillModel, SolidFillModel } from "../model.js";

/**
 * A resolved fill. A solid one is its CSS color string; anything else is the
 * same `ShapeFillModel` a shape carries (linear or path gradient, pattern,
 * picture, or `unsupported` for one that cannot be read), so the page paints
 * both through one painter (`src/fill-paint.ts`) and marks what it cannot.
 */
export type Fill = string | Exclude<ShapeFillModel, SolidFillModel>;

export interface Line {
  color: string;
  width: number;
  /** Dash pattern in points (already multiplied by the width). */
  dash?: number[];
  cap?: "butt" | "round" | "square";
  join?: "round" | "bevel" | "miter";
}

/** A resolved `c:spPr`. `undefined` members mean "draw none". */
export interface Look {
  fill?: Fill;
  line?: Line;
}

export interface TextLook {
  size: number;
  bold: boolean;
  italic: boolean;
  color: string;
  family: string;
  /**
   * The complex-script face (`a:cs`, or the theme's `+mn-cs`, else its Arabic script font): Office paints
   * and measures Arabic runs with it, not with `family`. Absent when nothing names one; Arabic then takes
   * Arial, Office's own fallback, or Times New Roman beside a serif `family`.
   */
  csFamily?: string;
  /** `a:bodyPr/@rot`, degrees; undefined lets the layout pick. */
  rotation?: number;
  /**
   * `a:normAutofit/@lnSpcReduction` as a share of the line height Office removes
   * (20000 = 0.2), from a stored scale rather than a fit. Glyph sizes are
   * unaffected. Measured in Excel 16.0 build 20430: a 24pt title's PDF baseline
   * gap fell from 25.39pt to 20.30pt (0.799 of it) with 20000.
   */
  lineSpacingReduction?: number;
}

export interface ManualLayout {
  /** Fractions of the chart area. x/y are edge positions unless the mode is factor. */
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  xMode: "edge" | "factor";
  yMode: "edge" | "factor";
  /** Plot area only: whether the rectangle bounds the axes' inner box or includes tick labels. */
  target: "inner" | "outer";
}

export interface Title {
  /** Paragraphs; one string per line. */
  lines: string[];
  text: TextLook;
  look: Look;
  overlay: boolean;
  layout?: ManualLayout;
}

export type MarkerSymbol =
  | "none" | "circle" | "square" | "diamond" | "triangle" | "x" | "star" | "plus" | "dot" | "dash" | "picture";

export interface Marker {
  symbol: MarkerSymbol;
  size: number;
  look: Look;
}

export type LabelPosition = "bestFit" | "b" | "ctr" | "inBase" | "inEnd" | "l" | "outEnd" | "r" | "t";

/** The fields a rich data label can embed (`a:fld/@type`); each is replaced by what it names. */
export type LabelField = "VALUE" | "SERIESNAME" | "CATEGORYNAME" | "PERCENTAGE" | "CELLRANGE" | "XVALUE" | "YVALUE" | "BUBBLESIZE";

/** A piece of author-typed label text: literal text ("\n" between lines), or a field with the text Office cached for it. */
export type LabelPart = string | { field: LabelField; cached: string };

export interface DataLabels {
  showValue: boolean;
  showCategory: boolean;
  showSeries: boolean;
  showPercent: boolean;
  showBubbleSize: boolean;
  showLegendKey: boolean;
  separator?: string;
  position?: LabelPosition;
  formatCode?: string;
  text: TextLook;
  look: Look;
  /** Whether the label is deleted (`c:delete`). */
  deleted: boolean;
  /** Author-typed text (`c:tx/c:rich`) replacing the generated one; fields inside it are filled from the point. */
  custom?: LabelPart[];
  /** "Value from cells" (`c15:showDataLabelsRange`): the series' `labelRange` text leads the label. */
  fromCells: boolean;
  /** Manual position (`c:layout/c:manualLayout` x/y): offset from the default spot, as fractions of the chart. */
  offset?: { x: number; y: number };
  /**
   * Pie leader lines (`c:showLeaderLines` on, with `c:leaderLines/c:spPr`):
   * present when a label moved off its slice is joined back to it.
   */
  leaderLines?: Line;
}

export type TrendlineType = "linear" | "exp" | "log" | "poly" | "power" | "movingAvg";

/** `c:trendline`: parameters only; the page computes the curve. */
export interface Trendline {
  type: TrendlineType;
  /** Polynomial order (2..6). */
  order: number;
  /** Moving-average period (2 or more). */
  period: number;
  /** Extension past the last / before the first x, in x units (categories on a category axis). */
  forward: number;
  backward: number;
  /** `c:intercept`: the curve is forced through (0, intercept). */
  intercept?: number;
  dispEq: boolean;
  dispRSqr: boolean;
  /** `c:name`; undefined lets the legend name it the way Office does. */
  name?: string;
  line: Line;
  /** The equation / R-squared label's text and number format (`c:trendlineLbl`). */
  labelText: TextLook;
  labelFormatCode?: string;
}

/** `c:errBars`. */
export interface ErrorBars {
  /** Direction: x or y. Bar, line and area charts have only the value direction (always "y"). */
  dir: "x" | "y";
  type: "both" | "plus" | "minus";
  valType: "fixedVal" | "percentage" | "stdDev" | "stdErr" | "cust";
  /** The fixed amount, the percentage, or the standard-deviation multiple. */
  value: number;
  /** `c:plus` / `c:minus` caches for custom amounts. A one-point cache applies to every point. */
  plus?: Array<number | null>;
  minus?: Array<number | null>;
  noEndCap: boolean;
  line: Line;
}

export interface Point {
  look?: Look;
  marker?: Marker;
  explosion?: number;
  invertIfNegative?: boolean;
  labels?: DataLabels;
}

/**
 * `stock` is a line group drawn the way a stock chart is: hi-low lines and
 * up-down bars between its series. `ofPie` is a pie with some points split off
 * into a second pie or a stacked bar.
 *
 * The Office 2016 chart types (`cx:chartSpace`) lower onto the same model:
 * a histogram is a `bar` group of bins, a Pareto a `bar` group plus a `line`
 * group on a second value axis, a waterfall a `bar` group whose series carry
 * `spans`. A funnel, box and whisker, treemap and sunburst draw differently
 * and have kinds of their own.
 */
export type GroupKind =
  | "bar" | "line" | "area" | "pie" | "doughnut" | "scatter" | "bubble" | "radar" | "stock" | "ofPie"
  | "funnel" | "box" | "treemap" | "sunburst";

/** Kinds drawn as slices of a circle, with no axes. */
export function isRadialKind(kind: GroupKind): boolean {
  return kind === "pie" || kind === "doughnut" || kind === "ofPie";
}

/** One box and whisker: the five-number summary of a category's values, with the points outside the whiskers. */
export interface BoxSummary {
  /** Number of values summarised. */
  count: number;
  q1: number;
  median: number;
  q3: number;
  /** Ends of the whiskers: the lowest and highest value within 1.5 interquartile ranges of the box. */
  low: number;
  high: number;
  mean: number;
  /** Values past the whiskers. */
  outliers: number[];
  /** Every value, kept only when the chart shows its inner points. */
  inner?: number[];
}

/** A treemap or sunburst node; a parent's value is the sum of its children. */
export interface TreeNode {
  name: string;
  value: number;
  /** The branch colour (the top-level ancestor's), as drawn. */
  color: string;
  children?: TreeNode[];
}

export interface Series {
  /** `c:idx`: drives the automatic palette. */
  index: number;
  /** `c:order`: drawing / legend order. */
  order: number;
  name: string;
  categories: string[];
  /**
   * Outer levels of grouped categories (`c:multiLvlStrRef`), nearest the inner
   * level first. Each level lists where its groups start (an index into
   * `categories`); a group runs to the next start, the last to the end.
   * `categories` holds the innermost level.
   */
  categoryGroups?: Array<Array<{ start: number; text: string }>>;
  /** Text of "value from cells" labels, by point (`c15:datalabelsRange` cache). */
  labelRange?: string[];
  /** Numeric category values (dates, scatter x); undefined when the categories are text. */
  categoryValues?: Array<number | null>;
  values: Array<number | null>;
  bubbleSizes?: Array<number | null>;
  formatCode: string;
  categoryFormatCode?: string;
  look: Look;
  marker: Marker;
  smooth: boolean;
  explosion: number;
  invertIfNegative: boolean;
  /** Per-point overrides, by point index. */
  points: Record<number, Point>;
  labels?: DataLabels;
  trendlines?: Trendline[];
  errBars?: ErrorBars[];
  /**
   * Floating bars (waterfall): the `[from, to]` extent of each point's bar, in
   * value units; `values` keeps the signed amounts the labels show.
   */
  spans?: Array<[number, number] | null>;
  /** Waterfall: the points drawn as totals (their bar runs from zero to their own value). */
  totals?: number[];
  /** Box and whisker: one summary per category, in category order (`categories` names them). */
  boxes?: BoxSummary[];
}

export interface Group {
  kind: GroupKind;
  direction: "col" | "bar";
  grouping: "clustered" | "stacked" | "percentStacked" | "standard";
  varyColors: boolean;
  gapWidth: number;
  overlap: number;
  holeSize: number;
  firstSliceAngle: number;
  bubbleScale: number;
  sizeRepresents: "area" | "w";
  scatterStyle: "line" | "lineMarker" | "marker" | "none" | "smooth" | "smoothMarker";
  radarStyle: "standard" | "marker" | "filled";
  series: Series[];
  /** Ids of the category (or x) axis and the value (or y) axis. */
  axisIds: [number, number] | undefined;
  labels?: DataLabels;
  dropLines?: Line;
  hiLowLines?: Line;
  serLines?: Line;
  /** Bars between the first and the last series of a line or stock group: up when the last is higher. */
  upDownBars?: { gapWidth: number; up: Look; down: Look };
  /** `ofPie` only: what the second plot holds and how it looks. */
  ofPie?: {
    /** The second plot: a pie, or one stacked bar. */
    type: "pie" | "bar";
    /** Point indices of the first series that moved to the second plot, in source order. */
    secondary: number[];
    /** Second plot size as a percentage of the main pie (`c:secondPieSize`). */
    secondSize: number;
    /** The "Other" slice that stands for the secondary points in the main pie. */
    other: Look;
  };
  /** Legend entries that are not the series: a waterfall's Increase, Decrease and Total. */
  legendEntries?: Array<{ label: string; look: Look }>;
  /** Waterfall: the line joining each bar's end to the next bar's start. */
  connectors?: Line;
  /** Box and whisker: what the chart shows besides the box. */
  box?: { meanMarker: boolean; meanLine: boolean; outliers: boolean; innerPoints: boolean };
  /** Treemap and sunburst: the hierarchy, in drawing order (Office sorts every level by size, largest first; colours follow the data order). */
  tree?: TreeNode[];
  /** Treemap: where a parent's name goes. */
  parentLabels?: "overlapping" | "banner" | "none";
  /** Treemap banners: the band's fill (Excel's is `tx1` at 35%, the same grey for every branch). */
  bannerFill?: string;
}

export type TickMark = "none" | "in" | "out" | "cross";

export type TimeUnit = "days" | "months" | "years";

export interface Axis {
  id: number;
  kind: "cat" | "val" | "date" | "ser";
  position: "b" | "l" | "t" | "r";
  deleted: boolean;
  crossAxisId: number;
  crosses: "autoZero" | "min" | "max";
  crossesAt?: number;
  reversed: boolean;
  min?: number;
  max?: number;
  logBase?: number;
  majorUnit?: number;
  minorUnit?: number;
  majorGrid?: Line;
  minorGrid?: Line;
  majorTick: TickMark;
  minorTick: TickMark;
  labelPosition: "nextTo" | "low" | "high" | "none";
  line?: Line;
  text: TextLook;
  formatCode: string;
  /** True when the format follows the source data (`sourceLinked="1"` or no c:numFmt). */
  formatLinked: boolean;
  title?: Title;
  /** `c:crossBetween`: whether the crossing axis sits between categories or on them. */
  between: boolean;
  labelSkip?: number;
  markSkip?: number;
  /** Category label offset from the axis, percent (`c:lblOffset`). */
  labelOffset: number;
  /** Display units divisor (`c:dispUnits`), 1 when none. */
  displayUnit: number;
  displayUnitLabel?: string;
  /** Date axis base unit: one slot per unit (`c:baseTimeUnit`); absent lets the data choose. */
  baseUnit?: TimeUnit;
  /** Date axis: unit `majorUnit` / `minorUnit` count in (`c:majorTimeUnit`); absent means the base unit. */
  majorTimeUnit?: TimeUnit;
  minorTimeUnit?: TimeUnit;
  /** `c:auto`: a category axis whose data are dates becomes a date axis (resolved into `kind` when read). */
  auto: boolean;
  /** Grouped categories draw their outer levels (false when `c:noMultiLvlLbl` is set). */
  multiLevel: boolean;
}

export interface Legend {
  position: "b" | "t" | "l" | "r" | "tr";
  overlay: boolean;
  layout?: ManualLayout;
  text: TextLook;
  look: Look;
  /** Legend entries the author deleted, by entry index. */
  deleted: number[];
}

export interface ChartSpecModel {
  title?: Title;
  legend?: Legend;
  groups: Group[];
  /** Every axis once; groups name theirs by `Axis.id`. */
  axes: Axis[];
  plot: { look: Look; layout?: ManualLayout };
  /** Chart area fill and outline. */
  look: Look;
  roundedCorners: boolean;
  blanks: "gap" | "zero" | "span";
  date1904: boolean;
  /** The chart-wide text defaults every element inherits from. */
  text: TextLook;
}
