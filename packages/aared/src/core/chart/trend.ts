/**
 * Trendline fits and error amounts: pure numbers, computed on the page from
 * the parameters the spec carries (`Trendline`, `ErrorBars`).
 *
 * Fits follow Excel's: least squares on the data (linear, polynomial, log),
 * or on ln y (exp) and ln x, ln y (power). R-squared is the one of the fit
 * actually solved, so exp and power report it in log space, as Excel does.
 */

import type { ErrorBars, Trendline } from "./spec.js";

export interface Fit {
  /** The curve's y at x; NaN where it is undefined (log of a non-positive x). */
  at: (x: number) => number;
  /** y = ... coefficients, in the order `equation` prints them. */
  coefficients: number[];
  r2: number;
}

type Pair = [number, number];

function usable(xs: ReadonlyArray<number>, ys: ReadonlyArray<number | null>, positiveX: boolean, positiveY: boolean): Pair[] {
  const out: Pair[] = [];
  for (let i = 0; i < xs.length; i++) {
    const y = ys[i];
    const x = xs[i]!;
    if (y === null || y === undefined || !Number.isFinite(x) || !Number.isFinite(y)) continue;
    if ((positiveX && x <= 0) || (positiveY && y <= 0)) continue;
    out.push([x, y]);
  }
  return out;
}

/** Solve A c = b by Gaussian elimination with partial pivoting; undefined when singular. */
function solve(a: number[][], b: number[]): number[] | undefined {
  const n = b.length;
  for (let col = 0; col < n; col++) {
    let pivot = col;
    for (let r = col + 1; r < n; r++) if (Math.abs(a[r]![col]!) > Math.abs(a[pivot]![col]!)) pivot = r;
    if (Math.abs(a[pivot]![col]!) < 1e-12) return undefined;
    [a[col], a[pivot]] = [a[pivot]!, a[col]!];
    [b[col], b[pivot]] = [b[pivot]!, b[col]!];
    for (let r = col + 1; r < n; r++) {
      const f = a[r]![col]! / a[col]![col]!;
      for (let k = col; k < n; k++) a[r]![k]! -= f * a[col]![k]!;
      b[r]! -= f * b[col]!;
    }
  }
  const out = new Array<number>(n).fill(0);
  for (let r = n - 1; r >= 0; r--) {
    let s = b[r]!;
    for (let k = r + 1; k < n; k++) s -= a[r]![k]! * out[k]!;
    out[r] = s / a[r]![r]!;
  }
  return out;
}

/** Least-squares polynomial of `order` (coefficients lowest power first), x scaled for conditioning. */
function polyFit(points: Pair[], order: number, intercept: number | undefined): number[] | undefined {
  let scale = 0;
  for (const [x] of points) scale = Math.max(scale, Math.abs(x));
  if (scale === 0) scale = 1;
  const first = intercept === undefined ? 0 : 1;
  const m = order + 1 - first;
  if (points.length < m) return undefined;
  const a: number[][] = Array.from({ length: m }, () => new Array<number>(m).fill(0));
  const b = new Array<number>(m).fill(0);
  for (const [x, y0] of points) {
    const xs = x / scale;
    const y = y0 - (intercept ?? 0);
    const pow: number[] = [1];
    for (let k = 1; k <= 2 * order; k++) pow[k] = pow[k - 1]! * xs;
    for (let r = 0; r < m; r++) {
      b[r]! += y * pow[r + first]!;
      for (let c = 0; c < m; c++) a[r]![c]! += pow[r + c + 2 * first]!;
    }
  }
  const solved = solve(a, b);
  if (solved === undefined) return undefined;
  const coefficients: number[] = intercept === undefined ? [] : [intercept];
  for (const [k, c] of solved.entries()) coefficients.push(c / scale ** (k + first));
  return coefficients;
}

function evalPoly(c: readonly number[], x: number): number {
  let y = 0;
  for (let k = c.length - 1; k >= 0; k--) y = y * x + c[k]!;
  return y;
}

function rSquared(points: Pair[], predict: (x: number) => number): number {
  let mean = 0;
  for (const [, y] of points) mean += y;
  mean /= points.length;
  let ssTot = 0;
  let ssRes = 0;
  for (const [x, y] of points) {
    ssTot += (y - mean) ** 2;
    ssRes += (y - predict(x)) ** 2;
  }
  return ssTot === 0 ? 1 : Math.max(0, 1 - ssRes / ssTot);
}

/**
 * Fit `trend` to the points. Moving averages have no formula: use
 * `movingAverage`. Undefined when the data cannot be fitted (too few points,
 * or no positive values for a log-type fit).
 */
export function fitTrend(trend: Trendline, xs: ReadonlyArray<number>, ys: ReadonlyArray<number | null>): Fit | undefined {
  const type = trend.type;
  if (type === "movingAvg") return undefined;
  if (type === "linear" || type === "poly") {
    const points = usable(xs, ys, false, false);
    const order = type === "linear" ? 1 : Math.min(Math.max(Math.round(trend.order), 2), 6);
    if (points.length < 2) return undefined;
    const c = polyFit(points, order, trend.intercept);
    if (c === undefined) return undefined;
    const at = (x: number): number => evalPoly(c, x);
    return { at, coefficients: c, r2: rSquared(points, at) };
  }
  if (type === "log") {
    const points = usable(xs, ys, true, false).map(([x, y]): Pair => [Math.log(x), y]);
    if (points.length < 2) return undefined;
    const c = polyFit(points, 1, undefined);
    if (c === undefined) return undefined;
    const at = (x: number): number => (x > 0 ? c[0]! + c[1]! * Math.log(x) : NaN);
    return { at, coefficients: c, r2: rSquared(points, (lx) => c[0]! + c[1]! * lx) };
  }
  // exp: ln y = ln a + b x; power: ln y = ln a + b ln x.
  const power = type === "power";
  const points = usable(xs, ys, power, true).map(([x, y]): Pair => [power ? Math.log(x) : x, Math.log(y)]);
  if (points.length < 2) return undefined;
  const c = polyFit(points, 1, trend.intercept !== undefined && trend.intercept > 0 ? Math.log(trend.intercept) : undefined);
  if (c === undefined) return undefined;
  const a = Math.exp(c[0]!);
  const b = c[1]!;
  const at = power ? (x: number): number => (x > 0 ? a * x ** b : NaN) : (x: number): number => a * Math.exp(b * x);
  return { at, coefficients: [a, b], r2: rSquared(points, (v) => c[0]! + b * v) };
}

/** Moving average of `period` points, drawn at the last point of each window (Excel's). */
export function movingAverage(xs: ReadonlyArray<number>, ys: ReadonlyArray<number | null>, period: number): Pair[] {
  const out: Pair[] = [];
  const n = Math.max(2, Math.round(period));
  for (let i = n - 1; i < ys.length; i++) {
    let sum = 0;
    let ok = true;
    for (let k = i - n + 1; k <= i; k++) {
      const y = ys[k];
      if (y === null || y === undefined) { ok = false; break; }
      sum += y;
    }
    const x = xs[i];
    if (ok && x !== undefined && Number.isFinite(x)) out.push([x, sum / n]);
  }
  return out;
}

/** The superscript two and three (Latin-1, so a literal is shipped as is). */
const SUPERSCRIPT = ["²", "³"];

function num(v: number, format: ((v: number) => string) | undefined): string {
  return format !== undefined ? format(v) : String(Number(v.toPrecision(5)));
}

function signed(v: number, rest: string, format: ((v: number) => string) | undefined): string {
  return `${v < 0 ? " - " : " + "}${num(Math.abs(v), format)}${rest}`;
}

/** The label Office prints for `dispEq`: "y = 2x + 1", "y = 3e^0.5x", ... */
export function equation(trend: Trendline, fit: Fit, format?: (v: number) => string): string {
  const c = fit.coefficients;
  switch (trend.type) {
    case "exp": return `y = ${num(c[0]!, format)}e^${num(c[1]!, format)}x`;
    case "power": return `y = ${num(c[0]!, format)}x^${num(c[1]!, format)}`;
    case "log": return `y = ${num(c[1]!, format)}ln(x)${signed(c[0]!, "", format)}`;
    default: {
      // Highest power first; Office leaves a zero constant out.
      let out = "y = ";
      for (let k = c.length - 1; k >= 0; k--) {
        const v = c[k]!;
        if (k === 0 && v === 0 && c.length > 1) continue;
        const x = k === 0 ? "" : k === 1 ? "x" : k <= 3 ? `x${SUPERSCRIPT[k - 2]!}` : `x^${k}`;
        out += k === c.length - 1 ? `${num(v, format)}${x}` : signed(v, x, format);
      }
      return out;
    }
  }
}

export function rSquaredLabel(fit: Fit, format?: (v: number) => string): string {
  return `R${SUPERSCRIPT[0]!} = ${format !== undefined ? format(fit.r2) : fit.r2.toFixed(4)}`;
}

/**
 * Error amounts per point, `{ plus, minus }` (both >= 0). `values` are what
 * the bars are measured on (the y values, or the x values for x bars);
 * `centre` is where each bar is drawn from: the value itself, except standard
 * deviation bars, which Excel centres on the series mean (assumption: the
 * equation images on Microsoft's page could not be read).
 */
export function errorAmounts(eb: ErrorBars, values: ReadonlyArray<number | null>): Array<{ plus: number; minus: number; centre?: number } | undefined> {
  const present = values.filter((v): v is number => v !== null && Number.isFinite(v));
  const n = present.length;
  const mean = n === 0 ? 0 : present.reduce((a, b) => a + b, 0) / n;
  const sd = n < 2 ? 0 : Math.sqrt(present.reduce((a, b) => a + (b - mean) ** 2, 0) / (n - 1));
  return values.map((v, i) => {
    if (v === null || !Number.isFinite(v)) return undefined;
    let plus: number;
    let minus: number;
    let centre: number | undefined;
    switch (eb.valType) {
      case "percentage": plus = minus = Math.abs(v) * eb.value / 100; break;
      case "stdDev": plus = minus = sd * eb.value; centre = mean; break;
      case "stdErr": plus = minus = n === 0 ? 0 : sd / Math.sqrt(n); break;
      case "cust": {
        const pick = (arr: Array<number | null> | undefined): number => {
          const x = arr === undefined ? 0 : arr.length === 1 ? arr[0] : arr[i];
          return x === null || x === undefined || !Number.isFinite(x) ? 0 : Math.abs(x);
        };
        plus = pick(eb.plus);
        minus = pick(eb.minus);
        break;
      }
      default: plus = minus = Math.abs(eb.value);
    }
    if (eb.type === "plus") minus = 0;
    else if (eb.type === "minus") plus = 0;
    return centre === undefined ? { plus, minus } : { plus, minus, centre };
  });
}
