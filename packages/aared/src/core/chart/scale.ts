/**
 * Axis scales: the value domain each axis spans, its major/minor units, and
 * the tick values. Pure numbers; layout decides how many ticks fit.
 *
 * Value axes follow Office's automatic scaling:
 *  - zero is kept in range unless the data's own span is under 1/6 of its
 *    magnitude (all values within the top sixth), then the range floats;
 *  - a floating end moves half the data's span past the data; an end that
 *    holds or crosses zero gains 5% headroom;
 *  - the major unit is the smallest of 1, 2, 5 x 10^n giving no more
 *    intervals than the axis has room for (never more than 10);
 *  - the ends round outward to whole major units.
 */

import { serialToDate } from "../xlsx/numberfmt.js";
import type { Axis, TimeUnit } from "./spec.js";

export interface ValueScale {
  kind: "value";
  min: number;
  max: number;
  major: number;
  minor: number;
  log?: number;
}

export interface CategoryScale {
  kind: "category";
  /** Slots along the axis: one per category, or one per base time unit on a date axis. */
  count: number;
  /** Categories sit between tick marks (columns) or on them (area, scatter-like lines). */
  between: boolean;
  /** Present on a date axis: points land by date, not by index. */
  time?: TimeInfo;
}

/** How a date axis maps a date to its slot. */
export interface TimeInfo {
  base: TimeUnit;
  /** Absolute index of slot 0 (the base unit holding the axis' first date). */
  origin: number;
  /** Serial of the axis' first day. */
  start: number;
  date1904: boolean;
}

export interface TimeTick {
  /** Slot of the tick, counted from the axis start. */
  slot: number;
  /** Date the label shows: the first day of the tick's base unit. */
  serial: number;
}

export type Scale = ValueScale | CategoryScale;

const NICE = [1, 2, 5, 10];

/** The smallest 1-2-5 step at or above `raw`. */
export function niceStep(raw: number): number {
  if (!(raw > 0) || !Number.isFinite(raw)) return 1;
  const exp = Math.floor(Math.log10(raw));
  const base = 10 ** exp;
  for (const n of NICE) if (n * base >= raw * (1 - 1e-9)) return clean(n * base);
  return clean(10 * base);
}

/** Strip binary noise: 0.30000000000000004 → 0.3. */
export function clean(v: number): number {
  return v === 0 ? 0 : Number(v.toPrecision(12));
}

/** Fit a value scale to [lo, hi] given how many major intervals fit along the axis. */
export function valueScale(axis: Axis | undefined, lo: number, hi: number, maxIntervals: number, percent = false, headroom = true): ValueScale {
  const fixedMin = axis?.min;
  const fixedMax = axis?.max;
  const base = axis?.logBase;
  if (base !== undefined && base >= 2) return logScale(base, lo, hi, fixedMin, fixedMax);
  if (!Number.isFinite(lo) || !Number.isFinite(hi)) { lo = 0; hi = 1; }
  if (percent) {
    lo = lo < 0 ? -1 : 0;
    hi = hi > 0 ? 1 : 0;
  }
  if (fixedMin !== undefined) lo = Math.min(lo, fixedMin);
  if (fixedMax !== undefined) hi = Math.max(hi, fixedMax);
  if (lo === hi) {
    if (lo === 0) hi = 1;
    else if (lo > 0) lo = 0;
    else hi = 0;
  }
  // Keep zero unless every value sits in the outer sixth of its own magnitude.
  let floatLow = false;
  let floatHigh = false;
  if (lo >= 0 && fixedMin === undefined) {
    if ((hi - lo) / hi < 1 / 6) floatLow = true;
    else lo = 0;
  }
  if (hi <= 0 && fixedMax === undefined) {
    if ((hi - lo) / -lo < 1 / 6) floatHigh = true;
    else hi = 0;
  }
  const span = hi - lo;
  // A floating end drops half the data's span below the data (97-100 starts
  // at 95.5); an end that holds zero or crosses it gains 5% headroom, which
  // radar axes do without (they end on the data).
  const pad = percent || !headroom ? 0 : 0.05 * span;
  const top = fixedMax ?? (floatHigh ? hi + span / 2 : hi > 0 ? hi + pad : hi);
  const bottom = fixedMin ?? (floatLow ? lo - span / 2 : lo < 0 ? lo - pad : lo);
  let major = axis?.majorUnit ?? niceStep((top - bottom) / Math.max(1, maxIntervals));
  if (percent && axis?.majorUnit === undefined) major = Math.max(major, 0.05);
  let min = fixedMin ?? Math.floor(bottom / major + 1e-9) * major;
  let max = fixedMax ?? Math.ceil(top / major - 1e-9) * major;
  if (max <= min) max = min + major;
  // A floating end chosen from the unit may change the span; refit once.
  if (axis?.majorUnit === undefined && (max - min) / major > maxIntervals * 1.0001 && maxIntervals >= 1) {
    major = niceStep((max - min) / maxIntervals);
    min = fixedMin ?? Math.floor(bottom / major + 1e-9) * major;
    max = fixedMax ?? Math.ceil(top / major - 1e-9) * major;
  }
  const minor = axis?.minorUnit ?? major / 5;
  return { kind: "value", min: clean(min), max: clean(max), major: clean(major), minor: clean(minor) };
}

function logScale(base: number, lo: number, hi: number, fixedMin?: number, fixedMax?: number): ValueScale {
  const ln = Math.log(base);
  const positiveLo = lo > 0 ? lo : 1;
  const positiveHi = hi > 0 ? hi : base;
  const min = fixedMin !== undefined && fixedMin > 0 ? fixedMin : base ** Math.floor(Math.log(positiveLo) / ln + 1e-9);
  let max = fixedMax !== undefined && fixedMax > 0 ? fixedMax : base ** Math.ceil(Math.log(positiveHi) / ln - 1e-9);
  if (max <= min) max = min * base;
  return { kind: "value", min: clean(min), max: clean(max), major: base, minor: base, log: base };
}

/** Major tick values from min to max inclusive. */
export function majorTicks(scale: ValueScale): number[] {
  const out: number[] = [];
  if (scale.log !== undefined) {
    for (let v = scale.min; v <= scale.max * (1 + 1e-9) && out.length < 1000; v *= scale.log) out.push(clean(v));
    return out;
  }
  const n = Math.round((scale.max - scale.min) / scale.major);
  for (let i = 0; i <= n && i < 1000; i++) out.push(clean(scale.min + i * scale.major));
  return out;
}

/** Minor tick values strictly between major ticks. */
export function minorTicks(scale: ValueScale): number[] {
  const out: number[] = [];
  if (scale.log !== undefined) {
    for (let decade = scale.min; decade < scale.max && out.length < 2000; decade *= scale.log) {
      for (let k = 2; k < scale.log; k++) out.push(decade * k);
    }
    return out;
  }
  const n = Math.round((scale.max - scale.min) / scale.minor);
  for (let i = 1; i < n && i < 5000; i++) {
    const v = scale.min + i * scale.minor;
    if (Math.abs(v / scale.major - Math.round(v / scale.major)) > 1e-6) out.push(clean(v));
  }
  return out;
}

/** Position of `v` along a value scale as a 0..1 fraction (unclamped). */
export function valueFraction(scale: ValueScale, v: number): number {
  if (scale.log !== undefined) {
    const l = Math.log(scale.log);
    const safe = v > 0 ? v : scale.min;
    return (Math.log(safe) / l - Math.log(scale.min) / l) / (Math.log(scale.max) / l - Math.log(scale.min) / l);
  }
  return (v - scale.min) / (scale.max - scale.min);
}

// ------------------------------------------------------------------ date axes

/**
 * Most slots a date axis lays out. A day-based axis over decades would be a
 * hundred thousand slots; Excel moves to a larger base unit before that, and so
 * does this.
 */
export const MAX_TIME_SLOTS = 20_000;
const RANK: Record<TimeUnit, number> = { days: 0, months: 1, years: 2 };

/** Serial of a calendar date; `month` (0-based) and `day` may run past their range. */
function civilSerial(year: number, month: number, day: number, date1904: boolean): number {
  return Date.UTC(year, month, day) / 86_400_000 + (date1904 ? 24107 : 25569);
}

/** Index of the base unit holding `serial`, counted from year 0 / the serial origin. */
function absoluteSlot(serial: number, base: TimeUnit, date1904: boolean): number {
  if (base === "days") return Math.floor(serial);
  const d = serialToDate(serial, date1904);
  return base === "months" ? d.year * 12 + d.month - 1 : d.year;
}

/** Slot of `serial` on a date axis, counted from the axis start (negative before it). */
export function timeSlot(time: TimeInfo, serial: number): number {
  return absoluteSlot(serial, time.base, time.date1904) - time.origin;
}

/** Excel's automatic base unit: the coarsest unit the closest pair of dates still resolves. */
function autoBase(dates: readonly number[]): TimeUnit {
  const days = [...new Set(dates.map(Math.floor))].sort((a, b) => a - b);
  let gap = Infinity;
  for (let i = 1; i < days.length; i++) gap = Math.min(gap, days[i]! - days[i - 1]!);
  if (!Number.isFinite(gap)) return "days";
  return gap < 28 ? "days" : gap < 365 ? "months" : "years";
}

/**
 * A date axis over `dates` (serials): one slot per base unit from the first
 * date (or the axis minimum) to the last (or the maximum), so dates land by
 * how far apart they are. `undefined` when no date is usable or the span has
 * more slots than a chart lays out.
 */
export function timeScale(dates: ReadonlyArray<number | null | undefined>, axis: Axis, between: boolean, date1904: boolean): CategoryScale | undefined {
  let lo = Infinity;
  let hi = -Infinity;
  const finite: number[] = [];
  for (const d of dates) {
    if (d === null || d === undefined || !Number.isFinite(d)) continue;
    finite.push(d);
    if (d < lo) lo = d;
    if (d > hi) hi = d;
  }
  if (finite.length === 0) return undefined;
  if (axis.min !== undefined && Number.isFinite(axis.min)) lo = axis.min;
  if (axis.max !== undefined && Number.isFinite(axis.max)) hi = axis.max;
  if (hi < lo) hi = lo;
  let base = axis.baseUnit ?? autoBase(finite);
  const slots = (b: TimeUnit): number => absoluteSlot(hi, b, date1904) - absoluteSlot(lo, b, date1904) + 1;
  while (!(slots(base) <= MAX_TIME_SLOTS) && base !== "years") base = base === "days" ? "months" : "years";
  const count = slots(base);
  if (!(count >= 1 && count <= MAX_TIME_SLOTS)) return undefined;
  return { kind: "category", count, between, time: { base, origin: absoluteSlot(lo, base, date1904), start: Math.floor(lo), date1904 } };
}

/** What Excel steps a date axis by when the file names no major unit, finest first. */
const AUTO_STEPS: Array<[number, TimeUnit]> = [
  [1, "days"], [2, "days"], [7, "days"], [14, "days"],
  [1, "months"], [2, "months"], [3, "months"], [6, "months"],
  [1, "years"], [2, "years"], [5, "years"], [10, "years"], [20, "years"], [50, "years"], [100, "years"],
];

/** Every `n`th `unit` from the axis start, at most `limit` of them. */
function stepTicks(scale: CategoryScale, n: number, unit: TimeUnit, limit: number): TimeTick[] {
  const t = scale.time!;
  const start = serialToDate(t.start, t.date1904);
  const out: TimeTick[] = [];
  for (let k = 0; out.length < limit; k++) {
    let slot: number;
    let serial: number;
    if (t.base === "days") {
      if (unit === "days") {
        serial = t.start + k * n;
      } else {
        // Months and years step by the calendar from the start date, keeping its day (clamped to short months).
        const total = start.year * 12 + start.month - 1 + k * (unit === "years" ? 12 * n : n);
        const year = Math.floor(total / 12);
        const month = total - year * 12;
        const last = new Date(Date.UTC(year, month + 1, 0)).getUTCDate();
        serial = civilSerial(year, month, Math.min(start.day, last), t.date1904);
      }
      slot = Math.floor(serial) - t.origin;
    } else if (t.base === "months") {
      const total = start.year * 12 + start.month - 1 + k * (unit === "years" ? 12 * n : n);
      slot = total - t.origin;
      serial = civilSerial(Math.floor(total / 12), total - Math.floor(total / 12) * 12, 1, t.date1904);
    } else {
      const year = start.year + k * n;
      slot = year - t.origin;
      serial = civilSerial(year, 0, 1, t.date1904);
    }
    if (!(slot < scale.count)) break;
    out.push({ slot, serial });
  }
  return out;
}

/**
 * Major ticks of a date axis: the file's `majorUnit` of `majorTimeUnit` (the
 * base unit when absent), else the finest step `accept` takes. A unit finer
 * than the base unit is the base unit: a slot is the smallest thing the axis
 * can step by. `accept` sees each candidate, finest first (every base unit,
 * then Excel's coarser steps), and the coarsest is taken when it takes none.
 */
export function timeTicks(scale: CategoryScale, axis: Axis, accept: (ticks: TimeTick[]) => boolean): TimeTick[] {
  const time = scale.time!;
  const floor = (unit: TimeUnit): TimeUnit => (RANK[unit] < RANK[time.base] ? time.base : unit);
  if (axis.majorUnit !== undefined) {
    return stepTicks(scale, Math.max(1, Math.round(axis.majorUnit)), floor(axis.majorTimeUnit ?? time.base), 5000);
  }
  let last: TimeTick[] = [];
  const steps: Array<[number, TimeUnit]> = AUTO_STEPS.filter(([, unit]) => RANK[unit] >= RANK[time.base]);
  for (let n = 200; n <= 102400; n *= 2) steps.push([n, "years"]);
  for (const [n, unit] of steps) {
    // A step whose ticks number more than the slots would only repeat them.
    const ticks = stepTicks(scale, n, unit, Math.min(scale.count, 5000) + 1);
    last = ticks;
    if (accept(ticks)) return ticks;
    if (ticks.length <= 1) return ticks;
  }
  return last;
}

/** Position of category `i`'s center as a 0..1 fraction. */
export function categoryFraction(scale: CategoryScale, i: number): number {
  if (scale.between) return (i + 0.5) / Math.max(scale.count, 1);
  return scale.count <= 1 ? 0.5 : i / (scale.count - 1);
}

/** Width of one category slot as a fraction of the axis. */
export function categorySlot(scale: CategoryScale): number {
  return scale.between ? 1 / Math.max(scale.count, 1) : 1 / Math.max(scale.count - 1, 1);
}
