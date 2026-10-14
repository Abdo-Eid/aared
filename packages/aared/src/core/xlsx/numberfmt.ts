/**
 * Narrow Excel number-format application over cached cell values.
 *
 * View-only: this module formats the file's stored result; it never evaluates
 * formulas. Dates use UTC getters so output is deterministic across host
 * timezones. Unknown or out-of-scope codes fall back to a plain value string
 * rather than a wrong guess.
 */

import { ownString } from "../xml.js";

const BUILTIN_FORMATS: Record<number, string> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "mm-dd-yy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yy h:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss",
  46: "[h]:mm:ss",
  // ECMA-376's table prints "mmss.0"; Excel resolves id 47 to "mm:ss.0" and
  // shows 45123.76 as 08:38.4 (probe xlsx-builtin-numfmt), so the spec's
  // missing colon is a misprint.
  47: "mm:ss.0",
  48: "##0.0E+0",
  49: "@",
};

// Ids the file may name without defining a code. Excel resolves every one of
// them from the reader's regional settings, so there is no single right
// answer (probe `xlsx-builtin-numfmt`, Excel 16: on an AED, dd/mm/yyyy machine
// 5-8 and 42/44 read "AED", 14, 22 and the dates below read dd/mm/yyyy, and
// 32-35 read h:mm:ss). aared shows the ECMA-376 (en-US) code for the currency
// and accounting ids and the short date or time for the East Asian era-date
// and time ids, which is what Excel itself falls back to outside those locales.
const REGIONAL_FORMATS: Record<number, string> = {
  5: '"$"#,##0_);("$"#,##0)',
  6: '"$"#,##0_);[Red]("$"#,##0)',
  7: '"$"#,##0.00_);("$"#,##0.00)',
  8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
  41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  ...Object.fromEntries([27, 28, 29, 30, 31, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58].map((id) => [id, "mm-dd-yy"])),
  ...Object.fromEntries([32, 33, 34, 35].map((id) => [id, "h:mm:ss"])),
};

/** Resolve a numFmtId to its format code, preferring custom records. */
export function formatCodeForId(id: number, custom: ReadonlyMap<number, string>): string {
  const override = custom.get(id);
  if (override !== undefined) return override;
  return BUILTIN_FORMATS[id] ?? REGIONAL_FORMATS[id] ?? "General";
}

interface Section {
  code: string;
  color?: string;
}

/**
 * Format codes are few per workbook but applied per cell, so their parsed
 * forms are memoized. The caches are bounded: a pathological workbook with
 * endless distinct codes just starts over instead of growing without limit.
 * They live as long as the Worker, which outlives a workbook (it is reused
 * between opens), so they key and build from an owned copy of the code: a
 * code sliced from `styles.xml` would otherwise keep that part alive.
 */
const CACHE_LIMIT = 512;

function memo<T>(cache: Map<string, T>, key: string, build: (key: string) => T): T {
  let value = cache.get(key);
  if (value === undefined) {
    if (cache.size >= CACHE_LIMIT) cache.clear();
    const owned = ownString(key);
    value = build(owned);
    cache.set(owned, value);
  }
  return value;
}

interface ParsedCode {
  sections: Section[];
  /** `isDateFormatCode` of each section's code. */
  dates: boolean[];
}

const parsedCodes = new Map<string, ParsedCode>();

function parseCode(code: string): ParsedCode {
  return memo(parsedCodes, code, () => {
    const parts = sections(code);
    return { sections: parts, dates: parts.map((part) => isDateFormatCode(part.code)) };
  });
}

/** Split a format code on unquoted `;` into positive/negative/zero/text sections. */
function sections(code: string): Section[] {
  const parts: string[] = [];
  let current = "";
  let quoted: string | undefined;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]!;
    if (quoted !== undefined) {
      current += ch;
      if (ch === quoted) quoted = undefined;
      continue;
    }
    if (ch === '"' || ch === "'") {
      quoted = ch;
      current += ch;
      continue;
    }
    if (ch === "\\" && i + 1 < code.length) {
      current += ch + code[i + 1]!;
      i++;
      continue;
    }
    if (ch === ";") {
      parts.push(current);
      current = "";
      continue;
    }
    current += ch;
  }
  parts.push(current);
  return parts.map((part) => {
    const color = /^\[([A-Za-z]+)\]/.exec(part)?.[1];
    return { code: part, ...(color === undefined ? {} : { color }) };
  });
}

interface Masked {
  masked: string;
  literals: string[];
}

const maskedCodes = new Map<string, Masked>();

/** Remove quoted literals into placeholders so token scans ignore them (memoized; read-only result). */
function maskLiterals(code: string): Masked {
  return memo(maskedCodes, code, maskLiteralsUncached);
}

function maskLiteralsUncached(code: string): Masked {
  // Placeholders use control characters on purpose: a numeric index would
  // itself scan as a `0` digit slot in number sections.
  const literals: string[] = [];
  const token = (index: number): string => `\u0001${index}\u0002`;
  let masked = "";
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]!;
    if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      const literal = end === -1 ? code.slice(i + 1) : code.slice(i + 1, end);
      literals.push(literal);
      masked += token(literals.length - 1);
      i = end === -1 ? code.length : end;
      continue;
    }
    if (ch === "\\" && i + 1 < code.length) {
      literals.push(code[i + 1]!);
      masked += token(literals.length - 1);
      i++;
      continue;
    }
    if (ch === "_") {
      i++;
      continue;
    }
    if (ch === "*") {
      i++;
      continue;
    }
    masked += ch;
  }
  return { masked, literals };
}

function unmask(text: string, literals: string[]): string {
  return text.replace(/\u0001(\d+)\u0002/g, (_, index: string) => literals[Number(index)] ?? "");
}

/** Whether a format code renders a serial as a date/time rather than a number. */
export function isDateFormatCode(code: string): boolean {
  // Elapsed counters (`[h]:mm` durations) are date-system values even though
  // the brackets are stripped before token analysis below.
  if (/\[[hms]\]/i.test(code)) return true;
  const { masked } = maskLiterals(code);
  const stripped = masked.replace(/\[[^\]]*\]/g, "").replace(/AM\/PM|am\/pm|A\/P/i, "");
  if (/[yY]/.test(stripped)) return true;
  if (/[dD]/.test(stripped)) return true;
  if (/[hH]/.test(stripped)) return true;
  if (/[sS]/.test(stripped)) return true;
  // Bare `m` is ambiguous (month vs minute); treat as a date only when it sits
  // next to an unambiguous date/time token.
  return /m(?=[/:.\s,\-][hmsd])|(?<=[hmsd][/:.\s,\-])m/i.test(stripped);
}

const MONTHS_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const MONTHS_LONG = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

const WEEKDAYS_LONG = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const WEEKDAYS_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

function pad(value: number, width: number): string {
  return String(value).padStart(width, "0");
}

/** Excel serial (1900 or 1904 system) to UTC calendar parts; `weekday` is 0 for Sunday. */
export function serialToDate(serial: number, date1904: boolean): { year: number; month: number; day: number; weekday: number; hours: number; minutes: number; seconds: number; milliseconds: number } {
  // Serial 25569 (1900 system) and 24107 (1904 system) are 1970-01-01.
  const millis = (serial - (date1904 ? 24107 : 25569)) * 86_400_000;
  const date = new Date(Math.round(millis));
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    weekday: date.getUTCDay(),
    hours: date.getUTCHours(),
    minutes: date.getUTCMinutes(),
    seconds: date.getUTCSeconds(),
    milliseconds: date.getUTCMilliseconds(),
  };
}

/** Format an Excel serial with a date/time code. Month and day names stay English so output is host-locale independent. */
export function formatSerialDate(serial: number, code: string, date1904: boolean): string {
  // Elapsed counters (`[h]:mm` durations) are totals, not clock fields, so
  // they bypass the calendar token path entirely.
  const elapsed = /\[([hms])\]/i.exec(code);
  if (elapsed !== null) {
    const totalSeconds = Math.floor(serial * 86_400);
    switch (elapsed[1]!.toLowerCase()) {
      case "h": {
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        return /:ss/i.test(code)
          ? `${hours}:${pad(minutes, 2)}:${pad(totalSeconds % 60, 2)}`
          : `${hours}:${pad(minutes, 2)}`;
      }
      case "m": {
        const minutes = Math.floor(totalSeconds / 60);
        return /:ss/i.test(code) ? `${minutes}:${pad(totalSeconds % 60, 2)}` : `${minutes}`;
      }
      default:
        return `${totalSeconds}`;
    }
  }
  const { masked, literals } = maskLiterals(code);
  // `ss.0`, `ss.00`, `ss.000`: tenths to thousandths of a second (Excel shows
  // 18:08:38.4 under `mm:ss.0` as 08:38.4). The time is rounded at that
  // precision first, so 59.96s under `ss.0` carries into the next minute.
  const digits = Math.min(/s\.(0+)/i.exec(masked)?.[1]?.length ?? 0, 3);
  const steps = 86_400 * 10 ** digits;
  const part = serialToDate(digits === 0 ? serial : Math.round(serial * steps) / steps, date1904);
  const fraction = digits === 0 ? "" : pad(Math.round(part.milliseconds / 10 ** (3 - digits)), digits);
  const twelveHour = /AM\/PM/i.test(code);
  const hour24 = part.hours + part.minutes / 60 + part.seconds / 3600;
  const hour12 = part.hours % 12 === 0 ? 12 : part.hours % 12;
  // Replace longest tokens first; `m` after `h`/`s` context means minutes.
  let out = masked;
  out = out.replace(/yyyy+/gi, String(part.year));
  out = out.replace(/yy/gi, pad(part.year % 100, 2));
  out = out.replace(/mmmm+/gi, MONTHS_LONG[part.month - 1]!);
  out = out.replace(/mmm/gi, MONTHS_SHORT[part.month - 1]!);
  // Minutes vs month: an `m` token following an hour token or a colon, or
  // preceding a seconds token, is minutes; otherwise it is a month.
  out = out.replace(/mm/gi, (match, offset: number, text: string) => {
    const before = text.slice(Math.max(0, offset - 4), offset);
    const after = text.slice(offset + 2, offset + 6);
    const minutes = /h/i.test(before) || /:$/.test(before) || /:s/i.test(after) || /^s/i.test(after);
    return minutes ? pad(part.minutes, 2) : pad(part.month, 2);
  });
  out = out.replace(/(^|[^a-zA-Z])m(?![a-zA-Z])/gi, (match, prefix: string) => {
    const minutes = /h:?$/i.test(prefix) || /:$/.test(prefix);
    return `${prefix}${minutes ? part.minutes : part.month}`;
  });
  out = out.replace(/dddd+/gi, WEEKDAYS_LONG[part.weekday]!);
  out = out.replace(/ddd/gi, WEEKDAYS_SHORT[part.weekday]!);
  out = out.replace(/dd/gi, pad(part.day, 2));
  out = out.replace(/(^|[^a-zA-Z])d(?![a-zA-Z])/gi, (_, prefix: string) => `${prefix}${part.day}`);
  out = out.replace(/hh/gi, pad(twelveHour ? hour12 : part.hours, 2));
  out = out.replace(/(^|[^a-zA-Z])h(?![a-zA-Z])/gi, (_, prefix: string) => `${prefix}${twelveHour ? hour12 : part.hours}`);
  if (digits > 0) out = out.replace(/(ss|s)\.0+/i, (token) => `${token.startsWith("ss") ? pad(part.seconds, 2) : part.seconds}.${fraction}`);
  out = out.replace(/ss/gi, pad(part.seconds, 2));
  out = out.replace(/(^|[^a-zA-Z.])s(?![a-zA-Z])/gi, (_, prefix: string) => `${prefix}${part.seconds}`);
  if (twelveHour) out = out.replace(/AM\/PM/gi, hour24 < 12 ? "AM" : "PM").replace(/A\/P/gi, hour24 < 12 ? "A" : "P");
  return unmask(out.replace(/\[[^\]]*\]/g, ""), literals).trim();
}

function bestFraction(value: number, maxDenominator: number): { whole: number; numerator: number; denominator: number } {
  const whole = Math.floor(value);
  const rest = value - whole;
  let numerator = 0;
  let denominator = 1;
  let best = Number.POSITIVE_INFINITY;
  for (let candidate = 1; candidate <= maxDenominator; candidate++) {
    const rounded = Math.round(rest * candidate);
    const error = Math.abs(rest - rounded / candidate);
    if (error < best) {
      best = error;
      numerator = rounded;
      denominator = candidate;
    }
  }
  if (numerator === denominator) return { whole: whole + 1, numerator: 0, denominator: 1 };
  return { whole, numerator, denominator };
}

/** Matches a masked-literal token; its index digits are literals, not number slots. */
const TOKEN_PATTERN = /\u0001\d+\u0002/g;

/**
 * A number section analysed once: every regex scan of the code happens
 * here, so formatting a cell is only arithmetic and string assembly.
 */
type SectionPlan =
  | { kind: "literal"; text: string }
  | { kind: "percent"; inner: SectionPlan; appended: boolean }
  | { kind: "scientific"; decimals: number }
  | { kind: "fraction"; maxDenominator: number; integerPart: boolean; spacer: string }
  | { kind: "fixed"; decimals: number; decimalSlots: string; grouping: boolean; integerSlots: number; integerZero: boolean; prefix: string; suffix: string };

const sectionPlans = new Map<string, SectionPlan>();

function planSection(masked: string, literals: string[]): SectionPlan {
  const clean = masked.replace(/\[[^\]]*\]/g, "");
  // Skeleton analysis runs on the token-stripped section so literal digits
  // (token indexes, quoted text) never scan as number slots.
  const skeleton = clean.replace(TOKEN_PATTERN, "");
  if (!/[0#?]/.test(skeleton)) return { kind: "literal", text: unmask(clean, literals) };
  if (/%/.test(skeleton)) {
    // `%` scales by 100 and prints where the code writes it: `\(0.0%\)` shows
    // (81.5%), not (81.5)%. It becomes a literal, so a fixed section keeps it
    // in its prefix or suffix; the other kinds have no affixes and append it.
    const inner = planSection(clean.replace(/%/g, `\u0001${literals.length}\u0002`), [...literals, "%"]);
    return { kind: "percent", inner, appended: inner.kind !== "fixed" };
  }
  if (/E\+0+/i.test(skeleton)) return { kind: "scientific", decimals: /E\+(0+)/i.exec(skeleton)?.[1]?.length ?? 0 };
  if (skeleton.includes("/")) {
    const head = skeleton.split("/")[0] ?? "";
    return {
      kind: "fraction",
      maxDenominator: /\?\?/.test(skeleton) ? 99 : 9,
      integerPart: /#/.test(head) || /0/.test(head),
      spacer: skeleton.includes("  ") ? "  " : " ",
    };
  }
  const dot = skeleton.indexOf(".");
  const { prefix, suffix } = affixes(clean, literals);
  return {
    kind: "fixed",
    decimals: dot === -1 ? 0 : (skeleton.slice(dot + 1).match(/[0#?]/g) ?? []).length,
    // The slots after the point, in order: a trailing `#` or `?` that lands on a zero digit is dropped or blanked.
    decimalSlots: dot === -1 ? "" : (skeleton.slice(dot + 1).match(/[0#?]/g) ?? []).join(""),
    grouping: skeleton.includes(","),
    // `?` reserves digit slots; browsers render plain digits, so pad to the
    // requested integer width only when the code asks for leading slots.
    integerSlots: (dot === -1 ? skeleton : skeleton.slice(0, dot)).match(/[0?]/g)?.length ?? 0,
    // Only a `0` slot prints an integer part of zero; `#` and `?` leave it
    // empty (`#.00` shows .50, `"-"??` shows no digit for zero).
    integerZero: (dot === -1 ? skeleton : skeleton.slice(0, dot)).includes("0"),
    prefix,
    suffix,
  };
}

function applySection(value: number, plan: SectionPlan): string {
  switch (plan.kind) {
    case "literal":
      return plan.text;
    case "percent":
      return `${applySection(value * 100, plan.inner)}${plan.appended ? "%" : ""}`;
    case "scientific": {
      const [mantissa, exponent] = value.toExponential(plan.decimals).split("e");
      const sign = exponent!.startsWith("-") ? "-" : "+";
      return `${mantissa}E${sign}${exponent!.replace(/^[+-]/, "").padStart(2, "0")}`;
    }
    case "fraction": {
      const fraction = bestFraction(Math.abs(value), plan.maxDenominator);
      const sign = value < 0 ? "-" : "";
      if (fraction.numerator === 0) return `${sign}${fraction.whole}`;
      return plan.integerPart && fraction.whole !== 0
        ? `${sign}${fraction.whole}${plan.spacer}${fraction.numerator}/${fraction.denominator}`
        : `${sign}${fraction.numerator}/${fraction.denominator}`;
    }
    case "fixed": {
      const rounded = Math.abs(value).toFixed(plan.decimals);
      let [integer, fraction = ""] = rounded.split(".");
      if (integer === "0" && !plan.integerZero) integer = "";
      if (plan.grouping) integer = integer!.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
      if (integer!.length < plan.integerSlots) integer = integer!.padStart(plan.integerSlots, " ");
      // `0.0#` shows 0.5 for 0.5 and `??.??` pads with spaces, not zeros; the point stays (Excel shows "5." for `0.##`).
      let end = fraction.length;
      let pad = "";
      while (end > 0 && fraction[end - 1] === "0" && plan.decimalSlots[end - 1] !== "0") {
        if (plan.decimalSlots[end - 1] === "?") pad += " ";
        end--;
      }
      fraction = fraction.slice(0, end) + pad;
      const digits = plan.decimals === 0 ? integer! : `${integer}.${fraction}`;
      return `${value < 0 ? "-" : ""}${plan.prefix}${digits}${plan.suffix}`;
    }
  }
}

/** Format a finite number with one number-section code (no `;` branches). */
export function formatNumberSection(value: number, code: string): string {
  const plan = memo(sectionPlans, code, () => {
    const { masked, literals } = maskLiterals(code);
    return planSection(masked, literals);
  });
  return applySection(value, plan);
}

/**
 * Split a masked section into literal prefix/suffix around its digit
 * skeleton. Token spans (`\u0001n\u0002`) are skipped: their indexes contain
 * digits that are literals, not number slots.
 */
function affixes(clean: string, literals: string[]): { prefix: string; suffix: string } {
  const tokenAt = (index: number): number => {
    const end = clean.indexOf("\u0002", index + 1);
    return end === -1 ? index + 1 : end + 1;
  };
  let first = -1;
  let last = -1;
  for (let i = 0; i < clean.length;) {
    // Token markers are emitted as escape text, so compare against the
    // runtime control characters they decode to.
    if (clean[i] === "\u0001") {
      i = tokenAt(i);
      continue;
    }
    if (clean[i] === "0" || clean[i] === "#" || clean[i] === "?") {
      if (first === -1) first = i;
      last = i;
    }
    i++;
  }
  if (first === -1) return { prefix: unmask(clean, literals), suffix: "" };
  return { prefix: unmask(clean.slice(0, first), literals), suffix: unmask(clean.slice(last + 1), literals) };
}

/**
 * Format a cached numeric value with a full format code, honoring
 * positive/negative/zero `;` sections. Non-finite input falls back to a plain
 * string so a corrupt cache never renders as a wrong number.
 */
export function formatNumberValue(value: number, code: string, date1904 = false): string {
  if (!Number.isFinite(value)) return String(value);
  if (code === "General" || code === "@") return generalNumber(value);
  const { sections: parts, dates } = parseCode(code);
  const at = value > 0 ? 0 : value < 0 ? (parts[1] === undefined ? 0 : 1) : (parts[2] === undefined ? 0 : 2);
  const branch = parts[at]!;
  if (dates[at]!) {
    const magnitude = Math.abs(value);
    if (magnitude > 2958465) return generalNumber(value);
    return formatSerialDate(value, branch.code, date1904);
  }
  // Sections own their sign: the negative branch of `#,##0;[Red]-#,##0`
  // already carries its minus, while a single-section code renders `-5`.
  const rendered = formatNumberSection(Math.abs(value), branch.code);
  if (value < 0 && parts.length < 2) return rendered.startsWith("-") ? rendered : `-${rendered}`;
  return rendered;
}

function generalNumber(value: number): string {
  if (Object.is(value, -0)) return "0";
  if (Number.isInteger(value) && Math.abs(value) < 1e11) return String(value);
  // Excel General shows up to 10-11 significant digits; beyond that it uses
  // scientific notation. Match that boundary instead of JS's longer default.
  // Fast path: a plain decimal with at most 10 digits already round-trips
  // through `toPrecision(10)` unchanged, so skip that (slow) call.
  const plain = String(value);
  if (plain.length <= 12 && plain.indexOf("e") === -1) {
    let digits = 0;
    for (let i = 0; i < plain.length; i++) {
      const code = plain.charCodeAt(i);
      if (code >= 48 && code <= 57) digits++;
    }
    if (digits <= 10) return plain;
  }
  return String(Number(value.toPrecision(10)));
}
