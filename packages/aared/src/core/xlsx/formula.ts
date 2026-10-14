/**
 * Bounded conditional-format formula subset: parse and evaluate worksheet
 * formulas over cached values, never recalculating the workbook.
 *
 * View-only: every reference reads its cached `<v>` (or blank when no cell
 * was authored). A formula cell with no cached value is unknown, not blank;
 * anything outside the subset is unsupported, never guessed. Callers leave
 * the whole affected rule unpainted and report it when compilation or a
 * dependency is unsupported, unknown, or over budget.
 *
 * Ground truth: relative references shift per target cell from the first
 * range's start (tools/probe/cf-formulas/probe.py: `$B2` on `A2:A6` paints
 * per-row; `A2>5` on `B2:B4` paints only B2); `cellIs` operands shift the
 * same way (`6,1,6` vs `10,0,5` paints C3,C4; probe3.py). Blank equals `""`,
 * `0` and `FALSE` but not `TRUE`; `""` equals `""` but not `0`
 * (probe2.py). `-2^2` is `4`, `10%` is `0.1`, `MOD(-3,2)` is `1`
 * (probe.py). `AND("abc")`, `IF("",1,2)` and `SUM("a")` are `#VALUE!`;
 * `"TRUE"`/`"FALSE"` coerce in `AND`/`OR`/`SUM` (probe2/3.py).
 * Cross-sheet `expression` rules live in x14 with `<xm:f>`/`<xm:sqref>`
 * (probe3.py); same-sheet base rules use `<formula>`/sqref.
 *
 * Date, text-search and conditional-aggregate ground truth
 * (tools/probe/cf-engine-close/, Excel 16.0, deleted after codification):
 * `TODAY()` is the date serial and `TODAY()=INT(NOW())`; `DATE` rolls month
 * and day overflow over, adds 1900 to years 0-1899, keeps the fictitious
 * 1900-02-29 (serial 60), and is `#NUM!` for negative results; `YEAR(-5)`
 * and `WEEKDAY(x,99)` are `#NUM!` while `WEEKDAY(0)` is 7 and `DAY(0)` is 0.
 * `SEARCH` is case-insensitive with `*`/`?`/`~` wildcards, `FIND` is
 * case-sensitive and literal, and a miss is `#VALUE!`. `MID` is 1-based:
 * start 0 is `#VALUE!` but length 0 and a start past the end are `""`.
 * `CONCAT` joins ranges, `CONCATENATE` with a range spills (out of subset).
 * `COUNTIF` is case-insensitive with wildcards: `"5"` matches 5 and text
 * `"5"`, `"TRUE"` matches `TRUE`, `"#DIV/0!"` matches that error, `""`
 * matches blanks and zero-length strings while `"="` matches only blanks
 * and `"<>"` everything but blanks; a blank criteria matches nothing and
 * an error criteria matches its error; `SUMIF` aligns a short sum range at
 * the top-left and `AVERAGEIF` with no match is `#DIV/0!`.
 */

import type { CellModel, CellRange } from "../model.js";
import { isInsideGrid } from "./limits.js";

export type FormulaErrorCode =
  | "#DIV/0!"
  | "#N/A"
  | "#NAME?"
  | "#NULL!"
  | "#NUM!"
  | "#REF!"
  | "#VALUE!";

/** An Excel error value: valid data (ISERROR sees it), not an engine failure. */
export interface FormulaErrorValue {
  error: FormulaErrorCode;
}

/** Scalar worksheet value: blank is null, distinct from "". */
export type FormulaScalar = number | string | boolean | null | FormulaErrorValue;

export type FormulaFailReason = "unsupported" | "unknown" | "budget";

export type FormulaResult =
  | { ok: true; value: FormulaScalar }
  | { ok: false; reason: FormulaFailReason };

export interface FormulaBudget {
  visits: number;
  maxVisits: number;
  nodes: number;
  maxNodes: number;
}

export function makeFormulaBudget(maxVisits = 5000000, maxNodes = 2048): FormulaBudget {
  return { visits: 0, maxVisits, nodes: 0, maxNodes };
}

function fail(reason: FormulaFailReason): { ok: false; reason: FormulaFailReason } {
  return { ok: false, reason };
}

export function isFormulaError(value: FormulaScalar): value is FormulaErrorValue {
  return typeof value === "object" && value !== null && "error" in value;
}

function errorValue(code: FormulaErrorCode): FormulaResult {
  return { ok: true, value: { error: code } };
}

/** Raw cell read for the evaluator: authored cells, blanks, and missing caches. */
export interface FormulaCellSource {
  /** Lower-cased current sheet name for same-sheet checks. */
  sheetLower: string;
  /** First range start: relative references shift from here to each target. */
  anchorRow: number;
  anchorCol: number;
  /** Cell being evaluated (target for relative translation). */
  targetRow: number;
  targetCol: number;
  /** `TODAY()` serial: the sheet's parse-date snapshot, never recalculated. */
  todaySerial: number;
  /** `NOW()` serial: the same snapshot plus the parse-time day fraction. */
  nowSerial: number;
  /** True for the legacy Mac date system (serial epoch 1904, no 1900 quirk). */
  date1904: boolean;
  /** Scalar cached value, blank for absent cells, unknown for missing caches. */
  scalarAt: (row: number, col: number) => FormulaResult;
  /** Numbers in a same-sheet rectangular range: errors/unknown propagate, text/blanks/logicals ignored. */
  numbersIn: (range: CellRange) => { ok: true; numbers: number[] } | { ok: true; value: FormulaErrorValue } | { ok: false; reason: FormulaFailReason };
  /** Authored values in a same-sheet range plus the blank count (blanks stay arithmetic, never materialized). */
  scanIn: (range: CellRange) => ScanResult;
  /** Cross-sheet numbers, when the caller supports them; absent means report. */
  crossNumbersIn?: (sheetLower: string, range: CellRange) => { ok: true; numbers: number[] } | { ok: true; value: FormulaErrorValue } | { ok: false; reason: FormulaFailReason };
  crossScalarAt?: (sheetLower: string, row: number, col: number) => FormulaResult;
  /** Cross-sheet scan, when the caller preloads other sheets; absent means report. */
  crossScanIn?: (sheetLower: string, range: CellRange) => ScanResult;
  budget: FormulaBudget;
}

/** One range's authored cells with positions plus its blank count, for `COUNTIF`-family alignment. */
export interface RangeScan {
  entries: Array<{ row: number; col: number; value: FormulaScalar }>;
  /** Range cells with no authored `<c>`: blank, distinct from `""`. */
  blanks: number;
}

export type ScanResult = { ok: true; scan: RangeScan } | { ok: false; reason: FormulaFailReason };

type Node =
  | { t: "n"; v: number }
  | { t: "s"; v: string }
  | { t: "b"; v: boolean }
  | { t: "e"; code: FormulaErrorCode }
  | { t: "ref"; sheet?: string; row: number; col: number; absRow: boolean; absCol: boolean }
  | { t: "range"; sheet?: string; startRow: number; startCol: number; startAbsRow: boolean; startAbsCol: boolean; endRow: number; endCol: number; endAbsRow: boolean; endAbsCol: boolean }
  | { t: "f"; name: string; args: Node[] }
  | { t: "u"; op: "+" | "-"; arg: Node }
  | { t: "op"; op: string; l: Node; r: Node }
  | { t: "pct"; arg: Node };

const MAX_TEXT = 8192;
const MAX_DEPTH = 64;
const MAX_ARGS = 256;

const SUPPORTED_FUNCS = new Set([
  "AND", "OR", "NOT", "IF", "IFERROR",
  "ISBLANK", "ISNUMBER", "ISTEXT", "ISERROR",
  "LEN", "LEFT", "RIGHT", "TRIM", "MID", "SEARCH", "FIND", "CONCAT", "CONCATENATE",
  "ABS", "MOD",
  "TODAY", "NOW", "DATE", "YEAR", "MONTH", "DAY", "WEEKDAY",
  "SUM", "COUNT", "MIN", "MAX", "AVERAGE",
  "COUNTIF", "COUNTIFS", "SUMIF", "SUMIFS", "AVERAGEIF", "AVERAGEIFS",
  "PERCENTILE", "PERCENTILE.INC",
]);

const ERROR_CODES: FormulaErrorCode[] = ["#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!"];

function isSpace(code: number): boolean {
  return code === 32 || code === 9 || code === 10 || code === 13;
}

function isDigit(code: number): boolean {
  return code >= 48 && code <= 57;
}

function isLetter(code: number): boolean {
  return (code >= 65 && code <= 90) || (code >= 97 && code <= 122);
}

function upper(code: number): number {
  return code >= 97 && code <= 122 ? code - 32 : code;
}

class Parser {
  text: string;
  pos = 0;
  depth = 0;
  nodes = 0;
  constructor(text: string) {
    this.text = text;
  }
  fail(): undefined {
    return undefined;
  }
  skip(): void {
    while (this.pos < this.text.length && isSpace(this.text.charCodeAt(this.pos))) this.pos++;
  }
  peekChar(): string {
    return this.text[this.pos] ?? "";
  }
  makeNode(node: Node): Node | undefined {
    this.nodes++;
    if (this.nodes > 2048) return undefined;
    return node;
  }
  parse(): Node | undefined {
    this.skip();
    // Leading "=" is stored without it, but accept it.
    if (this.peekChar() === "=") this.pos++;
    this.skip();
    if (this.pos >= this.text.length) return undefined;
    const node = this.parseComparison();
    if (node === undefined) return undefined;
    this.skip();
    if (this.pos !== this.text.length) return undefined;
    return node;
  }
  parseComparison(): Node | undefined {
    let left = this.parseConcat();
    if (left === undefined) return undefined;
    for (;;) {
      this.skip();
      const start = this.pos;
      let op = "";
      const c0 = this.text[start] ?? "";
      const c1 = this.text[start + 1] ?? "";
      if (c0 === "<" && c1 === ">") op = "<>";
      else if (c0 === "<" && c1 === "=") op = "<=";
      else if (c0 === ">" && c1 === "=") op = ">=";
      else if (c0 === "=" || c0 === "<" || c0 === ">") op = c0;
      else return left;
      this.pos = start + op.length;
      const right = this.parseConcat();
      if (right === undefined) return undefined;
      const made = this.makeNode({ t: "op", op, l: left, r: right });
      if (made === undefined) return undefined;
      left = made;
    }
  }
  parseConcat(): Node | undefined {
    let left = this.parseAddSub();
    if (left === undefined) return undefined;
    for (;;) {
      this.skip();
      if (this.peekChar() !== "&") return left;
      this.pos++;
      const right = this.parseAddSub();
      if (right === undefined) return undefined;
      const made = this.makeNode({ t: "op", op: "&", l: left, r: right });
      if (made === undefined) return undefined;
      left = made;
    }
  }
  parseAddSub(): Node | undefined {
    let left = this.parseMulDiv();
    if (left === undefined) return undefined;
    for (;;) {
      this.skip();
      const c = this.peekChar();
      if (c !== "+" && c !== "-") return left;
      // Binary +/-, not part of exponent (exponent handled in number lexer).
      this.pos++;
      const right = this.parseMulDiv();
      if (right === undefined) return undefined;
      const made = this.makeNode({ t: "op", op: c, l: left, r: right });
      if (made === undefined) return undefined;
      left = made;
    }
  }
  parseMulDiv(): Node | undefined {
    let left = this.parsePower();
    if (left === undefined) return undefined;
    for (;;) {
      this.skip();
      const c = this.peekChar();
      if (c !== "*" && c !== "/") return left;
      this.pos++;
      const right = this.parsePower();
      if (right === undefined) return undefined;
      const made = this.makeNode({ t: "op", op: c, l: left, r: right });
      if (made === undefined) return undefined;
      left = made;
    }
  }
  parsePower(): Node | undefined {
    const base = this.parseUnary();
    if (base === undefined) return undefined;
    this.skip();
    if (this.peekChar() !== "^") return base;
    this.pos++;
    const exp = this.parsePower();
    if (exp === undefined) return undefined;
    return this.makeNode({ t: "op", op: "^", l: base, r: exp });
  }
  parseUnary(): Node | undefined {
    this.skip();
    const c = this.peekChar();
    if (c === "+" || c === "-") {
      this.pos++;
      const arg = this.parseUnary();
      if (arg === undefined) return undefined;
      return this.makeNode({ t: "u", op: c as "+" | "-", arg });
    }
    return this.parsePercent();
  }
  parsePercent(): Node | undefined {
    let node = this.parsePrimary();
    if (node === undefined) return undefined;
    for (;;) {
      this.skip();
      if (this.peekChar() !== "%") return node;
      this.pos++;
      const made = this.makeNode({ t: "pct", arg: node });
      if (made === undefined) return undefined;
      node = made;
    }
  }
  parsePrimary(): Node | undefined {
    this.skip();
    if (this.depth >= MAX_DEPTH) return undefined;
    const c = this.peekChar();
    if (c === "") return undefined;
    if (c === "(") {
      this.pos++;
      this.depth++;
      const inner = this.parseComparison();
      this.depth--;
      if (inner === undefined) return undefined;
      this.skip();
      if (this.peekChar() !== ")") return undefined;
      this.pos++;
      return inner;
    }
    if (c === '"') return this.parseString();
    if (c === "#") return this.parseErrorLiteral();
    if (c === "'") {
      // Quoted sheet prefix: 'O''Brien'!A1
      return this.parseRefOrFunc();
    }
    if (isDigit(this.text.charCodeAt(this.pos)) || (c === "." && isDigit(this.text.charCodeAt(this.pos + 1)))) {
      return this.parseNumber();
    }
    if (isLetter(this.text.charCodeAt(this.pos)) || c === "_" || c === "$") {
      return this.parseRefOrFunc();
    }
    return undefined;
  }
  parseString(): Node | undefined {
    // Opening quote already peeked.
    this.pos++;
    let out = "";
    for (;;) {
      if (this.pos >= this.text.length) return undefined;
      const ch = this.text[this.pos]!;
      if (ch === '"') {
        if (this.text[this.pos + 1] === '"') {
          out += '"';
          this.pos += 2;
          continue;
        }
        this.pos++;
        const made = this.makeNode({ t: "s", v: out });
        return made ?? undefined;
      }
      out += ch;
      this.pos++;
      if (out.length > MAX_TEXT) return undefined;
    }
  }
  parseErrorLiteral(): Node | undefined {
    const rest = this.text.slice(this.pos).toUpperCase();
    for (const code of ERROR_CODES) {
      if (rest.startsWith(code)) {
        this.pos += code.length;
        return this.makeNode({ t: "e", code });
      }
    }
    return undefined;
  }
  parseNumber(): Node | undefined {
    const start = this.pos;
    let i = this.pos;
    const n = this.text.length;
    while (i < n && isDigit(this.text.charCodeAt(i))) i++;
    if (i < n && this.text[i] === ".") {
      i++;
      while (i < n && isDigit(this.text.charCodeAt(i))) i++;
    }
    if (i < n && (this.text[i] === "E" || this.text[i] === "e")) {
      let j = i + 1;
      if (j < n && (this.text[j] === "+" || this.text[j] === "-")) j++;
      const expStart = j;
      while (j < n && isDigit(this.text.charCodeAt(j))) j++;
      if (j > expStart) i = j;
    }
    if (i === start) return undefined;
    // A number directly followed by a letter is a name, not a number.
    if (i < n && isLetter(this.text.charCodeAt(i))) return undefined;
    const value = Number(this.text.slice(start, i));
    if (!Number.isFinite(value)) return undefined;
    this.pos = i;
    return this.makeNode({ t: "n", v: value });
  }
  parseIdent(): string {
    const start = this.pos;
    while (this.pos < this.text.length) {
      const code = this.text.charCodeAt(this.pos);
      if (isLetter(code) || isDigit(code) || code === 95 || code === 46) this.pos++;
      else break;
    }
    return this.text.slice(start, this.pos);
  }
  parseQuotedSheet(): string | undefined {
    // At opening "'".
    this.pos++;
    let out = "";
    for (;;) {
      if (this.pos >= this.text.length) return undefined;
      const ch = this.text[this.pos]!;
      if (ch === "'") {
        if (this.text[this.pos + 1] === "'") {
          out += "'";
          this.pos += 2;
          continue;
        }
        this.pos++;
        return out;
      }
      out += ch;
      this.pos++;
      if (out.length > 256) return undefined;
    }
  }
  // Parse [sheet!]ref [: [sheet!]ref] or ident[(args)] or TRUE/FALSE.
  parseRefOrFunc(): Node | undefined {
    const saved = this.pos;
    let sheet: string | undefined;
    if (this.peekChar() === "'") {
      const name = this.parseQuotedSheet();
      if (name === undefined) return undefined;
      this.skip();
      if (this.peekChar() !== "!") {
        this.pos = saved;
        return undefined;
      }
      this.pos++;
      this.skip();
      sheet = name;
    } else {
      // Try unquoted sheet prefix: scan ident then "!" without consuming ref.
      const mark = this.pos;
      // Sheet names may contain spaces when quoted only; unquoted stops at "!".
      let end = mark;
      while (end < this.text.length) {
        const ch = this.text[end]!;
        if (ch === "!") break;
        if (ch === "(" || ch === ")" || ch === "," || ch === ";" || ch === ":" || ch === "+" || ch === "-" || ch === "*" || ch === "/" || ch === "^" || ch === "&" || ch === "=" || ch === "<" || ch === ">" || ch === "%" || ch === '"' || ch === "#" || isSpace(this.text.charCodeAt(end))) break;
        end++;
      }
      if (end < this.text.length && this.text[end] === "!" && end > mark) {
        const candidate = this.text.slice(mark, end);
        // Unquoted sheet names cannot be cell refs themselves when followed by ref? Accept as sheet.
        sheet = candidate;
        this.pos = end + 1;
        this.skip();
      }
    }
    // Function call, boolean, or cell reference starts here.
    if (this.pos < this.text.length && (isLetter(this.text.charCodeAt(this.pos)) || this.text[this.pos] === "$" || this.text[this.pos] === "_")) {
      const identStart = this.pos;
      // Allow $ inside ref parsing below; read raw token up to delimiter.
      const ref = this.parseCellToken();
      if (ref !== undefined) {
        // Function call cannot have $; TRUE/FALSE handled below.
        const after = this.pos;
        this.skip();
        if (sheet === undefined && this.peekChar() === "(") {
          // It was a function name, not a ref: rewind to ident and parse call.
          this.pos = identStart;
          return this.parseFuncCall(undefined);
        }
        this.pos = after;
        // Range?
        this.skip();
        if (this.peekChar() === ":") {
          this.pos++;
          this.skip();
          // Second half may carry its own sheet prefix.
          const secondSaved = this.pos;
          let sheet2: string | undefined;
          if (this.peekChar() === "'") {
            const name2 = this.parseQuotedSheet();
            if (name2 === undefined) return undefined;
            this.skip();
            if (this.peekChar() !== "!") return undefined;
            this.pos++;
            this.skip();
            sheet2 = name2;
          } else {
            const mark2 = this.pos;
            let end2 = mark2;
            while (end2 < this.text.length) {
              const ch = this.text[end2]!;
              if (ch === "!") break;
              if (ch === "(" || ch === ")" || ch === "," || ch === ";" || ch === "+" || ch === "-" || ch === "*" || ch === "/" || ch === "^" || ch === "&" || ch === "=" || ch === "<" || ch === ">" || ch === "%" || ch === '"' || ch === "#" || isSpace(this.text.charCodeAt(end2))) break;
              end2++;
            }
            if (end2 < this.text.length && this.text[end2] === "!" && end2 > mark2) {
              sheet2 = this.text.slice(mark2, end2);
              this.pos = end2 + 1;
              this.skip();
            }
          }
          void secondSaved;
          const ref2 = this.parseCellToken();
          if (ref2 === undefined) return undefined;
          const s1 = sheet;
          const s2 = sheet2 ?? s1;
          if (s1 !== undefined && s2 !== undefined && s1.toUpperCase() !== s2.toUpperCase()) return undefined;
          const effective = s1 ?? s2;
          return this.makeNode({
            t: "range",
            ...(effective === undefined ? {} : { sheet: effective }),
            startRow: ref.row, startCol: ref.col, startAbsRow: ref.absRow, startAbsCol: ref.absCol,
            endRow: ref2.row, endCol: ref2.col, endAbsRow: ref2.absRow, endAbsCol: ref2.absCol,
          });
        }
        return this.makeNode({
          t: "ref",
          ...(sheet === undefined ? {} : { sheet }),
          row: ref.row, col: ref.col, absRow: ref.absRow, absCol: ref.absCol,
        });
      }
      // Not a cell token: function or boolean/name.
      this.pos = identStart;
      const ident = this.parseIdent();
      if (ident === "") return undefined;
      const up = ident.toUpperCase();
      this.skip();
      if (this.peekChar() === "(") {
        if (sheet !== undefined) return undefined;
        return this.parseFuncCall(up);
      }
      if (sheet !== undefined) return undefined;
      if (up === "TRUE") return this.makeNode({ t: "b", v: true });
      if (up === "FALSE") return this.makeNode({ t: "b", v: false });
      // Bare names, table refs, arrays: outside the subset.
      return undefined;
    }
    this.pos = saved;
    return undefined;
  }
  parseCellToken(): { row: number; col: number; absRow: boolean; absCol: boolean } | undefined {
    const start = this.pos;
    let absCol = false;
    if (this.text[this.pos] === "$") {
      absCol = true;
      this.pos++;
    }
    let col = 0;
    let letters = 0;
    while (this.pos < this.text.length && letters < 3) {
      let code = this.text.charCodeAt(this.pos);
      if (code >= 97 && code <= 122) code -= 32;
      if (code < 65 || code > 90) break;
      col = col * 26 + (code - 64);
      this.pos++;
      letters++;
    }
    if (letters < 1 || letters > 3) {
      this.pos = start;
      return undefined;
    }
    let absRow = false;
    if (this.text[this.pos] === "$") {
      absRow = true;
      this.pos++;
    }
    const digitStart = this.pos;
    let row = 0;
    let digits = 0;
    while (this.pos < this.text.length && digits < 7) {
      const code = this.text.charCodeAt(this.pos);
      if (code < 48 || code > 57) break;
      row = row * 10 + (code - 48);
      this.pos++;
      digits++;
    }
    if (digits < 1 || digits > 7) {
      this.pos = start;
      return undefined;
    }
    // Must not continue into a name (A1B) or function.
    if (this.pos < this.text.length) {
      const code = this.text.charCodeAt(this.pos);
      if (isLetter(code) || code === 95 || code === 46) {
        this.pos = start;
        return undefined;
      }
    }
    const zeroRow = row - 1;
    const zeroCol = col - 1;
    if (!isInsideGrid(zeroRow, zeroCol)) {
      // Off-grid stays parsable so evaluation yields #REF! rather than unsupported.
      // Clamp to grid edge; evaluation maps out-of-range to #REF!.
    }
    return { row: zeroRow, col: zeroCol, absRow, absCol };
  }
  parseFuncCall(name: string | undefined): Node | undefined {
    const ident = name ?? this.parseIdent();
    if (ident === "") return undefined;
    const up = ident.toUpperCase();
    if (!SUPPORTED_FUNCS.has(up)) return undefined;
    if (this.peekChar() !== "(") return undefined;
    this.pos++;
    if (this.depth >= MAX_DEPTH) return undefined;
    this.depth++;
    const args: Node[] = [];
    this.skip();
    if (this.peekChar() === ")") {
      this.pos++;
      this.depth--;
      return this.makeNode({ t: "f", name: up, args });
    }
    for (;;) {
      if (args.length >= MAX_ARGS) {
        this.depth--;
        return undefined;
      }
      const arg = this.parseComparison();
      if (arg === undefined) {
        this.depth--;
        return undefined;
      }
      args.push(arg);
      this.skip();
      const c = this.peekChar();
      if (c === "," || c === ";") {
        this.pos++;
        this.skip();
        // Trailing separator without an argument is malformed.
        if (this.peekChar() === ")") {
          this.depth--;
          return undefined;
        }
        continue;
      }
      if (c === ")") {
        this.pos++;
        this.depth--;
        return this.makeNode({ t: "f", name: up, args });
      }
      this.depth--;
      return undefined;
    }
  }
}

export function parseFormula(text: string): Node | undefined {
  if (text.length > MAX_TEXT) return undefined;
  // Reject array constants, braces, and intersection/union operators early:
  // "{", "}", and "!" outside sheet refs are not in the subset. "!" is
  // validated structurally below; a second "!" or brackets fail there.
  if (text.includes("{") || text.includes("}") || text.includes("[") || text.includes("]")) return undefined;
  const parser = new Parser(text);
  return parser.parse();
}

/** Resolve a parsed single-cell reference against anchor/target. */
function resolveRef(
  row: number, col: number, absRow: boolean, absCol: boolean,
  anchorRow: number, anchorCol: number, targetRow: number, targetCol: number,
): { row: number; col: number } {
  const dRow = targetRow - anchorRow;
  const dCol = targetCol - anchorCol;
  return {
    row: absRow ? row : row + dRow,
    col: absCol ? col : col + dCol,
  };
}

function toNumber(value: FormulaScalar): number | FormulaErrorCode {
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (value === null) return 0;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed === "") return 0;
    const num = Number(trimmed);
    return Number.isFinite(num) ? num : "#VALUE!";
  }
  return value.error;
}

function toString(value: FormulaScalar): string | FormulaErrorCode {
  if (typeof value === "string") return value;
  if (typeof value === "number") {
    // Excel stringifies without locale grouping; Infinity cannot arise from cached values.
    return String(value);
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (value === null) return "";
  return value.error;
}

function toBooleanStrict(value: FormulaScalar): boolean | FormulaErrorCode {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  if (value === null) return false;
  if (typeof value === "string") {
    const up = value.trim().toUpperCase();
    if (up === "TRUE") return true;
    if (up === "FALSE") return false;
    if (up === "") return "#VALUE!";
    const num = Number(value.trim());
    if (Number.isFinite(num)) return num !== 0;
    return "#VALUE!";
  }
  return value.error;
}

/** Excel type order for mismatched comparisons: number < text < boolean. Blank behaves as 0/"" for equality only. */
function typeRank(value: FormulaScalar): number {
  if (typeof value === "number") return 0;
  if (typeof value === "string") return 1;
  if (typeof value === "boolean") return 2;
  return -1;
}

function equals(a: FormulaScalar, b: FormulaScalar): boolean | FormulaErrorCode {
  if (isFormulaError(a)) return a.error;
  if (isFormulaError(b)) return b.error;
  // Blank equals blank, 0, "", and FALSE (measured: C2="" / =0 / =FALSE all TRUE).
  if (a === null || b === null) {
    const other = a === null ? b : a;
    if (other === null) return true;
    if (typeof other === "number") return other === 0;
    if (typeof other === "string") return other === "";
    if (typeof other === "boolean") return other === false;
    return false;
  }
  if (typeof a === "number" && typeof b === "number") return a === b;
  if (typeof a === "string" && typeof b === "string") return a.toUpperCase() === b.toUpperCase();
  if (typeof a === "boolean" && typeof b === "boolean") return a === b;
  return false;
}

function compare(a: FormulaScalar, b: FormulaScalar, op: string): boolean | FormulaErrorCode {
  if (isFormulaError(a)) return a.error;
  if (isFormulaError(b)) return b.error;
  if (op === "=") return equals(a, b);
  if (op === "<>") {
    const eq = equals(a, b);
    return typeof eq === "string" ? eq : !eq;
  }
  // Blank in ordering: treat as 0 against numbers, "" against strings, FALSE against booleans.
  const av: FormulaScalar = a === null
    ? (typeof b === "number" ? 0 : typeof b === "string" ? "" : false)
    : a;
  const bv: FormulaScalar = b === null
    ? (typeof a === "number" ? 0 : typeof a === "string" ? "" : false)
    : b;
  if (typeof av === "number" && typeof bv === "number") {
    switch (op) {
      case "<": return av < bv;
      case ">": return av > bv;
      case "<=": return av <= bv;
      case ">=": return av >= bv;
      default: return "#VALUE!";
    }
  }
  if (typeof av === "string" && typeof bv === "string") {
    const l = av.toUpperCase();
    const r = bv.toUpperCase();
    switch (op) {
      case "<": return l < r;
      case ">": return l > r;
      case "<=": return l <= r;
      case ">=": return l >= r;
      default: return "#VALUE!";
    }
  }
  if (typeof av === "boolean" && typeof bv === "boolean") {
    const l = av ? 1 : 0;
    const r = bv ? 1 : 0;
    switch (op) {
      case "<": return l < r;
      case ">": return l > r;
      case "<=": return l <= r;
      case ">=": return l >= r;
      default: return "#VALUE!";
    }
  }
  // Mismatched types follow Excel's number < text < boolean order.
  const order = typeRank(av) - typeRank(bv);
  switch (op) {
    case "<": return order < 0;
    case ">": return order > 0;
    case "<=": return order <= 0;
    case ">=": return order >= 0;
    default: return "#VALUE!";
  }
}

function add(a: FormulaScalar, b: FormulaScalar, op: string): FormulaScalar | FormulaErrorCode {
  const x = toNumber(a);
  if (typeof x === "string") return x;
  const y = toNumber(b);
  if (typeof y === "string") return y;
  switch (op) {
    case "+": return x + y;
    case "-": return x - y;
    case "*": return x * y;
    case "/": return y === 0 ? "#DIV/0!" : x / y;
    case "^": {
      const v = Math.pow(x, y);
      return Number.isFinite(v) ? v : "#NUM!";
    }
    default: return "#VALUE!";
  }
}

function checkBudget(ctx: FormulaCellSource): FormulaResult | undefined {
  ctx.budget.visits++;
  if (ctx.budget.visits > ctx.budget.maxVisits) return fail("budget");
  return undefined;
}

/** Positive modulo: JS `%` keeps the sign, Excel serial math needs 0-6. */
function mod7(n: number): number {
  return ((n % 7) + 7) % 7;
}

/**
 * Calendar date to Excel serial, mirroring `serialToDate` in `numberfmt.ts`
 * (same 1970 epoch, so modern dates round-trip exactly). Month/day overflow
 * rolls over via the calendar (`DATE(2026,13,1)` is next January); years
 * 0-1899 read 1900 higher; the fictitious 1900-02-29 is serial 60; anything
 * below serial 0 is `#NUM!` (all measured).
 */
export function dateToSerial(year: number, month: number, day: number, date1904: boolean): number | FormulaErrorCode {
  const rawYear = Math.trunc(year);
  let m = Math.trunc(month);
  let d = Math.trunc(day);
  if (!Number.isFinite(rawYear) || !Number.isFinite(m) || !Number.isFinite(d)) return "#VALUE!";
  // Excel's two-digit rule spans four digits: `DATE(26,1,1)` is 1926.
  const fullYear = rawYear >= 0 && rawYear <= 1899 ? rawYear + 1900 : rawYear;
  if (!date1904 && fullYear === 1900 && m === 2 && d >= 29) {
    // February 1900 has the fictitious 29th: day 30 spills into March.
    if (d === 29) return 60;
    m = 3;
    d -= 29;
  }
  const base = new Date(0);
  base.setUTCHours(0, 0, 0, 0);
  base.setUTCFullYear(fullYear, m - 1, d);
  if (!Number.isFinite(base.getTime())) return "#NUM!";
  const linear = Math.round(base.getTime() / 86400000) + (date1904 ? 24107 : 25569);
  // Before March 1900 the linear count runs one ahead: serial 1 is
  // 1900-01-01, not the linear 1900-01-02 (the phantom day sits after).
  const serial = !date1904 && base.getTime() < Date.UTC(1900, 2, 1) ? linear - 1 : linear;
  return serial < 0 ? "#NUM!" : serial;
}

/** Decomposed calendar parts of a serial for `YEAR`/`MONTH`/`DAY` (`#NUM!` below serial 0). */
function decomposeSerial(serial: number, date1904: boolean): { year: number; month: number; day: number } | FormulaErrorCode {
  const n = Math.floor(serial);
  if (!Number.isFinite(n) || n < 0) return "#NUM!";
  if (!date1904) {
    // Measured: `DAY(0)` is 0 with `YEAR(0)` 1900, and 60 is the
    // fictitious 1900-02-29; from 61 on the linear count applies again.
    if (n === 0) return { year: 1900, month: 1, day: 0 };
    if (n >= 1 && n <= 31) return { year: 1900, month: 1, day: n };
    if (n >= 32 && n <= 59) return { year: 1900, month: 2, day: n - 31 };
    if (n === 60) return { year: 1900, month: 2, day: 29 };
    const date = new Date((n - 25569) * 86400000);
    return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
  }
  const date = new Date((n - 24107) * 86400000);
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
}

/** `WEEKDAY` of a serial: pure modular arithmetic, so the 1900 quirk comes along (measured `WEEKDAY(0)` is 7). */
function weekdayOf(serial: number, kind: 1 | 2 | 3, date1904: boolean): number | FormulaErrorCode {
  const n = Math.floor(serial);
  if (!Number.isFinite(n) || n < 0) return "#NUM!";
  // Same absolute day, different epochs 1462 apart (`1462 % 7` is 6).
  const t = date1904 ? mod7(n + 6) : mod7(n);
  if (kind === 1) return t === 0 ? 7 : t;
  if (kind === 2) return mod7(t + 5) + 1;
  return mod7(t + 5);
}

/** One regex-escaped character for the wildcard compiler below. */
function escapeRegExpChar(ch: string): string {
  return ".*+?^${}()|[]\\".includes(ch) ? `\\${ch}` : ch;
}

/**
 * An Excel wildcard pattern (`*`, `?`, `~` escape) as a regex body.
 * `undefined` past the caps (long or star-heavy patterns): the caller
 * reports instead of risking catastrophic backtracking.
 */
function wildcardSource(pattern: string): string | undefined {
  let stars = 0;
  for (const ch of pattern) if (ch === "*") stars++;
  if (pattern.length > 1024 || stars > 64) return undefined;
  let out = "";
  const chars = [...pattern];
  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i]!;
    if (ch === "~" && i + 1 < chars.length) {
      out += escapeRegExpChar(chars[i + 1]!);
      i++;
      continue;
    }
    if (ch === "*") {
      out += ".*";
      continue;
    }
    if (ch === "?") {
      out += ".";
      continue;
    }
    out += escapeRegExpChar(ch);
  }
  return out;
}

/** Case-insensitive whole-string wildcard matcher (`undefined` past the caps). */
function wildcardRegExp(pattern: string): RegExp | undefined {
  const source = wildcardSource(pattern);
  return source === undefined ? undefined : new RegExp(`^(?:${source})$`, "iu");
}

/** Case-insensitive wildcard match of a whole cell string (measured `"A*"` hits `Apple`, `apple` and literal `A*`). */
function wildcardMatch(text: string, pattern: string): boolean | undefined {
  const rx = wildcardRegExp(pattern);
  if (rx === undefined) return undefined;
  return rx.test(text);
}

/**
 * A compiled `COUNTIF`-family criteria: the operator split from its operand
 * once, so every cell in the range reuses it. `undefined` only past the
 * wildcard caps (reported, never guessed).
 */
type Criteria =
  | { k: "never" }
  | { k: "error"; code: FormulaErrorCode }
  | { k: "num"; n: number }
  | { k: "bool"; b: boolean }
  | { k: "text"; pattern: string }
  | { k: "emptyStr" }
  | { k: "emptyEq" }
  | { k: "nonEmpty" }
  | { k: "not"; inner: Criteria }
  | { k: "relNum"; op: string; n: number }
  | { k: "relText"; op: string; t: string };

/** Whether a string is one of the seven error literals (compared upper-cased). */
function errorCodeOf(text: string): FormulaErrorCode | undefined {
  const up = text.toUpperCase();
  return (ERROR_CODES as string[]).includes(up) ? (up as FormulaErrorCode) : undefined;
}

/** Classify a bare (operator-free) criteria operand. */
function classifyOperand(text: string): Criteria | undefined {
  const trimmed = text.trim();
  if (trimmed !== "" && Number.isFinite(Number(trimmed))) return { k: "num", n: Number(trimmed) };
  const up = trimmed.toUpperCase();
  if (up === "TRUE") return { k: "bool", b: true };
  if (up === "FALSE") return { k: "bool", b: false };
  const code = errorCodeOf(trimmed);
  if (code !== undefined) return { k: "error", code };
  if (wildcardRegExp(trimmed) === undefined) return undefined;
  return { k: "text", pattern: trimmed };
}

export function compileCriteria(value: FormulaScalar): Criteria | undefined {
  if (value === null) return { k: "never" };
  if (isFormulaError(value)) return { k: "error", code: value.error };
  if (typeof value === "number") return { k: "num", n: value };
  if (typeof value === "boolean") return { k: "bool", b: value };
  const text = value;
  let op = "";
  let rest = text;
  if (rest.startsWith("<>") || rest.startsWith("<=") || rest.startsWith(">=")) {
    op = rest.slice(0, 2);
    rest = rest.slice(2);
  } else if (rest.startsWith("=") || rest.startsWith("<") || rest.startsWith(">")) {
    op = rest.slice(0, 1);
    rest = rest.slice(1);
  }
  // Measured: `""` matches blanks and zero-length strings, `"="` only
  // blanks, `"<>"` everything but blanks.
  if (op === "" || op === "=") {
    if (rest === "") return op === "=" ? { k: "emptyEq" } : { k: "emptyStr" };
    return classifyOperand(rest);
  }
  if (op === "<>") {
    if (rest === "") return { k: "nonEmpty" };
    const inner = classifyOperand(rest);
    return inner === undefined ? undefined : { k: "not", inner };
  }
  // Relational: a numeric operand compares numbers (blanks read as 0), a
  // text operand compares strings; a boolean operand never matches.
  const trimmed = rest.trim();
  if (trimmed !== "" && Number.isFinite(Number(trimmed))) return { k: "relNum", op, n: Number(trimmed) };
  const up = trimmed.toUpperCase();
  if (up === "TRUE" || up === "FALSE") return { k: "never" };
  return { k: "relText", op, t: rest };
}

/** Whether one range cell (blank is `null`) meets compiled criteria. */
export function criteriaMatches(cell: FormulaScalar, crit: Criteria): boolean | undefined {
  switch (crit.k) {
    case "never":
      return false;
    case "error":
      return isFormulaError(cell) && cell.error === crit.code;
    case "num":
      if (typeof cell === "number") return cell === crit.n;
      // Measured: numeric criteria match numeric-string text both ways.
      if (typeof cell === "string") {
        const trimmed = cell.trim();
        return trimmed !== "" && Number.isFinite(Number(trimmed)) && Number(trimmed) === crit.n;
      }
      return false;
    case "bool":
      if (typeof cell === "boolean") return cell === crit.b;
      if (typeof cell === "string") {
        const up = cell.trim().toUpperCase();
        if (up === "TRUE") return crit.b;
        if (up === "FALSE") return !crit.b;
      }
      return false;
    case "text": {
      if (typeof cell !== "string") return false;
      return wildcardMatch(cell, crit.pattern);
    }
    case "emptyStr":
      return cell === null || cell === "";
    case "emptyEq":
      return cell === null;
    case "nonEmpty":
      return cell !== null;
    case "not": {
      const inner = criteriaMatches(cell, crit.inner);
      return inner === undefined ? undefined : !inner;
    }
    case "relNum": {
      if (typeof cell === "number") return compare(cell, crit.n, crit.op) as boolean;
      // Measured `COUNTIF(A2:A10,0)` is 0: blanks sit out numeric equality
      // but read as 0 under ordering, like expression comparisons.
      if (cell === null) return compare(0, crit.n, crit.op) as boolean;
      return false;
    }
    case "relText": {
      if (typeof cell === "string") return compare(cell, crit.t, crit.op) as boolean;
      if (cell === null) return compare("", crit.t, crit.op) as boolean;
      return false;
    }
  }
}

function scalarOfCell(cell: CellModel | undefined): FormulaResult {
  if (cell === undefined) return { ok: true, value: null };
  if (cell.value === null && cell.formula !== undefined && cell.formula !== "") {
    // A formula with no cached result is unknown, never blank.
    return fail("unknown");
  }
  if (cell.type === "error") {
    const code = String(cell.value);
    const known = (ERROR_CODES as string[]).includes(code) ? (code as FormulaErrorCode) : "#VALUE!";
    return { ok: true, value: { error: known } };
  }
  return { ok: true, value: cell.value as FormulaScalar };
}

export function cellResultFor(cell: CellModel | undefined): FormulaResult {
  return scalarOfCell(cell);
}

function resolveRange(
  node: { sheet?: string; startRow: number; startCol: number; startAbsRow: boolean; startAbsCol: boolean; endRow: number; endCol: number; endAbsRow: boolean; endAbsCol: boolean },
  ctx: FormulaCellSource,
): { sheetLower: string | undefined; range: CellRange } | FormulaResult {
  // Cross-sheet ranges resolve here; each consumer reads them through its
  // cross hook, which fails `unsupported` when the caller preloaded nothing.
  const sheetLower = node.sheet === undefined ? undefined : node.sheet.toLowerCase();
  const a = resolveRef(node.startRow, node.startCol, node.startAbsRow, node.startAbsCol, ctx.anchorRow, ctx.anchorCol, ctx.targetRow, ctx.targetCol);
  const b = resolveRef(node.endRow, node.endCol, node.endAbsRow, node.endAbsCol, ctx.anchorRow, ctx.anchorCol, ctx.targetRow, ctx.targetCol);
  const out: { sheetLower: string | undefined; range: CellRange } = {
    sheetLower,
    range: {
      start: { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) },
      end: { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) },
    },
  };
  return out;
}

function numbersFromCollected(
  collected: { ok: true; numbers: number[] } | { ok: true; value: FormulaErrorValue } | { ok: false; reason: FormulaFailReason } | FormulaResult,
): { ok: true; numbers: number[] } | FormulaResult {
  if (!collected.ok) return collected;
  if ("numbers" in collected) return { ok: true, numbers: collected.numbers };
  if ("value" in collected && isFormulaError((collected as { value: FormulaScalar }).value)) {
    return { ok: true, value: (collected as { value: FormulaErrorValue }).value };
  }
  return fail("unsupported");
}

function rangeNumbers(
  arg: { t: "range"; sheet?: string; startRow: number; startCol: number; startAbsRow: boolean; startAbsCol: boolean; endRow: number; endCol: number; endAbsRow: boolean; endAbsCol: boolean },
  ctx: FormulaCellSource,
): { ok: true; numbers: number[] } | FormulaResult {
  const resolved = resolveRange(arg, ctx);
  if ("ok" in resolved) return resolved;
  // An explicit same-sheet prefix (`Data!A1:A3` on `Data`) reads locally.
  const raw = resolved.sheetLower === undefined || resolved.sheetLower === ctx.sheetLower
    ? ctx.numbersIn(resolved.range)
    : ctx.crossNumbersIn === undefined
      ? fail("unsupported")
      : ctx.crossNumbersIn(resolved.sheetLower, resolved.range);
  return numbersFromCollected(raw);
}

/** Authored values (with positions) for `COUNTIF`-family and `CONCAT` ranges; blanks stay a count. */
function rangeScan(
  arg: { t: "range"; sheet?: string; startRow: number; startCol: number; startAbsRow: boolean; startAbsCol: boolean; endRow: number; endCol: number; endAbsRow: boolean; endAbsCol: boolean },
  ctx: FormulaCellSource,
): ScanResult {
  const resolved = resolveRange(arg, ctx);
  if ("ok" in resolved) return fail("unsupported");
  if (resolved.sheetLower === undefined || resolved.sheetLower === ctx.sheetLower) return ctx.scanIn(resolved.range);
  if (ctx.crossScanIn === undefined) return fail("unsupported");
  return ctx.crossScanIn(resolved.sheetLower, resolved.range);
}

/** Lower-cased sheet names one formula reads (for preloading cross-sheet value grids; `evaluate` itself stays sync). */
export function referencedSheets(compiled: CompiledFormula): string[] {
  const out = new Set<string>();
  const walk = (n: CompiledFormula["node"]): void => {
    if (n.t === "ref" || n.t === "range") {
      if (n.sheet !== undefined) out.add(n.sheet.toLowerCase());
    } else if (n.t === "f") {
      for (const arg of n.args) walk(arg);
    } else if (n.t === "u" || n.t === "pct") {
      walk(n.arg);
    } else if (n.t === "op") {
      walk(n.l);
      walk(n.r);
    }
  };
  walk(compiled.node);
  return [...out];
}

function percentileOf(sorted: number[], k: number): number | FormulaErrorCode {
  if (sorted.length === 0) return "#NUM!";
  if (!(k >= 0 && k <= 1)) return "#NUM!";
  const rank = (sorted.length - 1) * k;
  const below = Math.floor(rank);
  const above = Math.ceil(rank);
  return sorted[below]! + (sorted[above]! - sorted[below]!) * (rank - below);
}

function collectArgNumbers(args: Node[], ctx: FormulaCellSource, forCount: boolean): { numbers: number[] } | FormulaResult {  const numbers: number[] = [];
  for (const arg of args) {
    if (arg.t === "range") {
      const got = rangeNumbers(arg, ctx);
      if (!("numbers" in got)) return got;
      numbers.push(...got.numbers);
      continue;
    }
    const value = evaluate(arg, ctx);
    if (!value.ok) return value;
    const scalar = value.value;
    if (isFormulaError(scalar)) return { ok: true, value: scalar };
    if (typeof scalar === "number") {
      numbers.push(scalar);
      continue;
    }
    if (scalar === null) {
      if (forCount) continue;
      // SUM treats blank scalar as 0 (measured: SUM(C2)=0); AVERAGE/MIN/MAX ignore it below.
      if (args.length === 1) continue;
      continue;
    }
    if (typeof scalar === "boolean") {
      numbers.push(scalar ? 1 : 0);
      continue;
    }
    // Direct text: numeric strings coerce (measured: SUM("1")=1, AVERAGE("1",2)=1.5).
    const num = Number(scalar.trim());
    if (scalar.trim() !== "" && Number.isFinite(num)) {
      numbers.push(num);
      continue;
    }
    return errorValue("#VALUE!");
  }
  return { numbers };
}

/** One scanned range with positional values for `COUNTIF`-family alignment. */
interface ScannedRange {
  map: Map<number, FormulaScalar>;
  width: number;
  height: number;
}

/**
 * Resolve a range argument and read its values with positions (offsets from
 * the resolved start, so differently-placed same-size ranges align). Unknown
 * caches fail; errors ride along as values for each consumer to decide.
 */
function resolveAndScan(arg: Node, ctx: FormulaCellSource): ScannedRange | FormulaResult {
  if (arg.t !== "range") return errorValue("#VALUE!");
  const resolved = resolveRange(arg, ctx);
  if ("ok" in resolved) return fail("unsupported");
  const raw = resolved.sheetLower === undefined || resolved.sheetLower === ctx.sheetLower
    ? ctx.scanIn(resolved.range)
    : ctx.crossScanIn === undefined
      ? fail("unsupported")
      : ctx.crossScanIn(resolved.sheetLower, resolved.range);
  if (!("scan" in raw)) return raw;
  const { start, end } = resolved.range;
  const width = end.col - start.col + 1;
  const map = new Map<number, FormulaScalar>();
  for (const entry of raw.scan.entries) {
    map.set((entry.row - start.row) * width + (entry.col - start.col), entry.value);
  }
  return { map, width, height: end.row - start.row + 1 };
}

/** Evaluate one criteria argument and compile it (a bad wildcard reports, never guesses). */
function compiledCriteriaOf(node: Node, ctx: FormulaCellSource): { ok: true; crit: Criteria } | FormulaResult {
  const r = evaluate(node, ctx);
  if (!r.ok) return r;
  const crit = compileCriteria(r.value);
  return crit === undefined ? fail("unsupported") : { ok: true, crit };
}

interface CritPair extends ScannedRange {
  crit: Criteria;
}

/**
 * Range/criteria pairs from `args` at `start` (alternating). Wrong shapes
 * are engine failures (reported); IFS ranges of different shapes are Excel's
 * `#VALUE!` (unpainted, not reported), including the value range when supplied.
 */
function ifsPairs(args: Node[], ctx: FormulaCellSource, start: number, valueRange?: ScannedRange): { pairs: CritPair[] } | FormulaResult {
  if ((args.length - start) % 2 !== 0 || args.length - start < 2) return fail("unsupported");
  const pairs: CritPair[] = [];
  for (let i = start; i < args.length; i += 2) {
    const scanned = resolveAndScan(args[i]!, ctx);
    if (!("map" in scanned)) return scanned;
    const compiled = compiledCriteriaOf(args[i + 1]!, ctx);
    if (!("crit" in compiled)) return compiled;
    pairs.push({ ...scanned, crit: compiled.crit });
  }
  const shape = valueRange ?? pairs[0]!;
  for (const pair of pairs) {
    if (pair.width !== shape.width || pair.height !== shape.height) return errorValue("#VALUE!");
  }
  return { pairs };
}

/** Union of authored offsets across ranges (all-blank positions are counted arithmetically by callers). */
function unionOffsets(ranges: Array<{ map: Map<number, FormulaScalar> }>): Set<number> {
  const union = new Set<number>();
  for (const range of ranges) for (const key of range.map.keys()) union.add(key);
  return union;
}

/** Rows meeting every pair (an `undefined` match reports, never guesses). */
function countPairs(pairs: CritPair[]): number | undefined {
  const first = pairs[0]!;
  const area = first.width * first.height;
  const union = unionOffsets(pairs);
  let count = 0;
  let blankHits = true;
  for (const pair of pairs) {
    const m = criteriaMatches(null, pair.crit);
    if (m === undefined) return undefined;
    if (!m) {
      blankHits = false;
      break;
    }
  }
  if (blankHits) count += area - union.size;
  for (const key of union) {
    let hit = true;
    for (const pair of pairs) {
      const m = criteriaMatches(pair.map.get(key) ?? null, pair.crit);
      if (m === undefined) return undefined;
      if (!m) {
        hit = false;
        break;
      }
    }
    if (hit) count++;
  }
  return count;
}

/** Sum of `sum` where every pair meets (matched errors propagate, like `SUM`). */
function sumPairs(sum: ScannedRange, pairs: CritPair[]): { ok: true; value: number } | FormulaResult {
  // Positions authored nowhere contribute 0, so only the union is visited.
  const union = unionOffsets([...pairs, sum]);
  let total = 0;
  for (const key of union) {
    let hit = true;
    for (const pair of pairs) {
      const m = criteriaMatches(pair.map.get(key) ?? null, pair.crit);
      if (m === undefined) return fail("unsupported");
      if (!m) {
        hit = false;
        break;
      }
    }
    if (!hit) continue;
    const v = sum.map.get(key) ?? null;
    if (typeof v === "number") total += v;
    else if (v === null || typeof v === "string" || typeof v === "boolean") continue;
    else return { ok: true, value: v };
  }
  return { ok: true, value: total };
}

/** Average of matched numbers (`#DIV/0!` with no match, measured). */
function averagePairs(sum: ScannedRange, pairs: CritPair[]): FormulaResult {
  const union = unionOffsets([...pairs, sum]);
  let total = 0;
  let count = 0;
  for (const key of union) {
    let hit = true;
    for (const pair of pairs) {
      const m = criteriaMatches(pair.map.get(key) ?? null, pair.crit);
      if (m === undefined) return fail("unsupported");
      if (!m) {
        hit = false;
        break;
      }
    }
    if (!hit) continue;
    const v = sum.map.get(key) ?? null;
    if (typeof v === "number") {
      total += v;
      count++;
    } else if (v === null || typeof v === "string" || typeof v === "boolean") {
      continue;
    } else {
      return { ok: true, value: v };
    }
  }
  return count === 0 ? errorValue("#DIV/0!") : { ok: true, value: total / count };
}

/** One `CONCAT` piece: blanks vanish, errors propagate (measured `CONCAT` of `#DIV/0!`). */
function concatPiece(value: FormulaScalar): string | FormulaErrorValue {
  if (value === null) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return value;
}

/** `CONCAT` joins scalars and ranges; legacy `CONCATENATE` takes scalars only (a range spills: out of subset). */
function concatArgs(args: Node[], ctx: FormulaCellSource, ranges: boolean): FormulaResult {
  let out = "";
  const push = (text: string): FormulaResult | undefined => {
    out += text;
    return out.length > 32767 ? errorValue("#VALUE!") : undefined;
  };
  for (const arg of args) {
    if (arg.t === "range") {
      if (!ranges) return fail("unsupported");
      const got = rangeScan(arg, ctx);
      if (!("scan" in got)) return got;
      const ordered = [...got.scan.entries].sort((a, b) => a.row - b.row || a.col - b.col);
      for (const entry of ordered) {
        const piece = concatPiece(entry.value);
        if (typeof piece !== "string") return { ok: true, value: piece };
        const over = push(piece);
        if (over !== undefined) return over;
      }
      continue;
    }
    const r = evaluate(arg, ctx);
    if (!r.ok) return r;
    const piece = concatPiece(r.value);
    if (typeof piece !== "string") return { ok: true, value: piece };
    const over = push(piece);
    if (over !== undefined) return over;
  }
  return { ok: true, value: out };
}

/** Case-sensitive literal find over characters (1-based; `FIND` has no wildcards, measured). */
function findLiteral(within: string, find: string, start: number): number | FormulaErrorCode {
  const text = [...within];
  const sought = [...find];
  if (start > text.length + 1) return "#VALUE!";
  if (sought.length === 0) return start;
  for (let i = start - 1; i + sought.length <= text.length; i++) {
    let hit = true;
    for (let j = 0; j < sought.length; j++) {
      if (text[i + j] !== sought[j]) {
        hit = false;
        break;
      }
    }
    if (hit) return i + 1;
  }
  return "#VALUE!";
}

/**
 * Case-insensitive wildcard find (1-based; `SEARCH` wildcards measured).
 * `undefined` past the pattern caps: the caller reports, never guesses.
 */
function findPattern(within: string, pattern: string, start: number): number | FormulaErrorCode | undefined {
  const text = [...within];
  if (start > text.length + 1) return "#VALUE!";
  const source = wildcardSource(pattern);
  if (source === undefined) return undefined;
  const tail = text.slice(start - 1).join("");
  const found = new RegExp(source, "iu").exec(tail);
  if (found === null) return "#VALUE!";
  return start + [...tail.slice(0, found.index)].length;
}

/** Shared `SEARCH`/`FIND` argument handling (case and wildcards differ). */
function searchArgs(name: string, args: Node[], ctx: FormulaCellSource): FormulaResult {
  if (args.length < 2 || args.length > 3) return fail("unsupported");
  const findR = evaluate(args[0]!, ctx);
  if (!findR.ok) return findR;
  if (isFormulaError(findR.value)) return { ok: true, value: findR.value };
  const withinR = evaluate(args[1]!, ctx);
  if (!withinR.ok) return withinR;
  if (isFormulaError(withinR.value)) return { ok: true, value: withinR.value };
  const find = toString(findR.value);
  if (typeof find !== "string") return errorValue(find);
  const within = toString(withinR.value);
  if (typeof within !== "string") return errorValue(within);
  let start = 1;
  if (args.length === 3) {
    const startR = evaluate(args[2]!, ctx);
    if (!startR.ok) return startR;
    if (isFormulaError(startR.value)) return { ok: true, value: startR.value };
    const n = toNumber(startR.value);
    if (typeof n === "string") return errorValue(n);
    if (!Number.isFinite(n)) return errorValue("#VALUE!");
    start = Math.trunc(n);
    // Measured: start 0 is `#VALUE!`.
    if (start < 1) return errorValue("#VALUE!");
  }
  const at = name === "FIND" ? findLiteral(within, find, start) : findPattern(within, find, start);
  if (at === undefined) return fail("unsupported");
  return typeof at === "string" ? errorValue(at) : { ok: true, value: at };
}

function evaluateFunc(node: { name: string; args: Node[] }, ctx: FormulaCellSource): FormulaResult {
  const over = checkBudget(ctx);
  if (over !== undefined) return over;
  const name = node.name;
  const args = node.args;
  const evalScalar = (index: number): FormulaResult => {
    if (index >= args.length) return errorValue("#VALUE!");
    return evaluate(args[index]!, ctx);
  };
  switch (name) {
    case "AND": {
      if (args.length === 0) return errorValue("#VALUE!");
      if (args.length > MAX_ARGS) return fail("budget");
      let saw = false;
      for (const arg of args) {
        if (arg.t === "range") {
          const got = rangeNumbers(arg, ctx);
          if (!("numbers" in got)) return got;
          // AND over a range of numbers: any zero is FALSE; empty is TRUE (measured blank handling).
          saw = saw || got.numbers.length > 0;
          if (got.numbers.some((v: number) => v === 0)) return { ok: true, value: false };
          continue;
        }
        const r = evaluate(arg, ctx);
        if (!r.ok) return r;
        const v = r.value;
        if (isFormulaError(v)) return { ok: true, value: v };
        if (v === null) continue;
        saw = true;
        const b = toBooleanStrict(v);
        if (typeof b === "string") return errorValue(b);
        if (!b) return { ok: true, value: false };
      }
      void saw;
      return { ok: true, value: true };
    }
    case "OR": {
      if (args.length === 0) return errorValue("#VALUE!");
      if (args.length > MAX_ARGS) return fail("budget");
      for (const arg of args) {
        if (arg.t === "range") {
          const got = rangeNumbers(arg, ctx);
          if (!("numbers" in got)) return got;
          if (got.numbers.some((v: number) => v !== 0)) return { ok: true, value: true };
          continue;
        }
        const r = evaluate(arg, ctx);
        if (!r.ok) return r;
        const v = r.value;
        if (isFormulaError(v)) return { ok: true, value: v };
        if (v === null) continue;
        const b = toBooleanStrict(v);
        if (typeof b === "string") return errorValue(b);
        if (b) return { ok: true, value: true };
      }
      return { ok: true, value: false };
    }
    case "NOT": {
      if (args.length !== 1) return fail("unsupported");
      const r = evalScalar(0);
      if (!r.ok) return r;
      if (isFormulaError(r.value)) return { ok: true, value: r.value };
      if (r.value === null) return { ok: true, value: true };
      const b = toBooleanStrict(r.value);
      if (typeof b === "string") return errorValue(b);
      return { ok: true, value: !b };
    }
    case "IF": {
      if (args.length < 2 || args.length > 3) return fail("unsupported");
      const cond = evaluate(args[0]!, ctx);
      if (!cond.ok) return cond;
      if (isFormulaError(cond.value)) return { ok: true, value: cond.value };
      // Measured: IF("",1,2) and IF("abc",1,2) are #VALUE!; blank is FALSE.
      if (typeof cond.value === "string") return errorValue("#VALUE!");
      if (cond.value === null) {
        return args.length === 3 ? evaluate(args[2]!, ctx) : { ok: true, value: false };
      }
      const b = toBooleanStrict(cond.value);
      if (typeof b === "string") return errorValue(b);
      if (b) return evaluate(args[1]!, ctx);
      return args.length === 3 ? evaluate(args[2]!, ctx) : { ok: true, value: false };
    }
    case "IFERROR": {
      if (args.length !== 2) return fail("unsupported");
      const first = evaluate(args[0]!, ctx);
      if (!first.ok) {
        // IFERROR must not conceal unknown/unsupported/budget (audit).
        return first;
      }
      if (isFormulaError(first.value)) return evaluate(args[1]!, ctx);
      return first;
    }
    case "ISBLANK": {
      if (args.length !== 1) return fail("unsupported");
      const target = args[0]!;
      if (target.t === "ref") {
        const at = resolveRef(target.row, target.col, target.absRow, target.absCol, ctx.anchorRow, ctx.anchorCol, ctx.targetRow, ctx.targetCol);
        if (target.sheet !== undefined && target.sheet.toLowerCase() !== ctx.sheetLower) {
          if (ctx.crossScalarAt === undefined) return fail("unsupported");
          const r = ctx.crossScalarAt(target.sheet.toLowerCase(), at.row, at.col);
          if (!r.ok) return r;
          return { ok: true, value: r.value === null };
        }
        const r = ctx.scalarAt(at.row, at.col);
        if (!r.ok) {
          // A missing cache is not blank.
          if (r.reason === "unknown") return { ok: true, value: false };
          return r;
        }
        return { ok: true, value: r.value === null };
      }
      const r = evaluate(target, ctx);
      if (!r.ok) return r;
      return { ok: true, value: r.value === null };
    }
    case "ISNUMBER": {
      if (args.length !== 1) return fail("unsupported");
      const r = evalScalar(0);
      if (!r.ok) return r;
      return { ok: true, value: typeof r.value === "number" };
    }
    case "ISTEXT": {
      if (args.length !== 1) return fail("unsupported");
      const r = evalScalar(0);
      if (!r.ok) return r;
      return { ok: true, value: typeof r.value === "string" };
    }
    case "ISERROR": {
      if (args.length !== 1) return fail("unsupported");
      const r = evalScalar(0);
      if (!r.ok) return r;
      return { ok: true, value: isFormulaError(r.value) };
    }
    case "LEN": {
      if (args.length !== 1) return fail("unsupported");
      const r = evalScalar(0);
      if (!r.ok) return r;
      if (isFormulaError(r.value)) return { ok: true, value: r.value };
      // Measured: LEN(TRUE)=4, LEN(blank)=0.
      const s = toString(r.value === null ? "" : r.value);
      if (typeof s === "string") return { ok: true, value: [...s].length };
      return errorValue(s);
    }
    case "LEFT":
    case "RIGHT": {
      if (args.length < 2 || args.length > 3) return fail("unsupported");
      const textR = evaluate(args[0]!, ctx);
      if (!textR.ok) return textR;
      if (isFormulaError(textR.value)) return { ok: true, value: textR.value };
      const countR = evaluate(args[1]!, ctx);
      if (!countR.ok) return countR;
      if (isFormulaError(countR.value)) return { ok: true, value: countR.value };
      const chars = [...toString(textR.value === null ? "" : textR.value) as string];
      const n = toNumber(countR.value === null ? 0 : countR.value);
      if (typeof n === "string") return errorValue(n);
      if (!Number.isFinite(n) || Math.trunc(n) < 0) return errorValue("#VALUE!");
      const k = Math.trunc(n);
      const out = name === "LEFT" ? chars.slice(0, k).join("") : chars.slice(Math.max(chars.length - k, 0)).join("");
      return { ok: true, value: out };
    }
    case "TRIM": {
      if (args.length !== 1) return fail("unsupported");
      const r = evalScalar(0);
      if (!r.ok) return r;
      if (isFormulaError(r.value)) return { ok: true, value: r.value };
      const s = toString(r.value === null ? "" : r.value);
      if (typeof s === "string") return { ok: true, value: s.replace(/^\s+|\s+$/g, "").replace(/\s+/g, " ") };
      return errorValue(s);
    }
    case "MID": {
      if (args.length !== 3) return fail("unsupported");
      const textR = evaluate(args[0]!, ctx);
      if (!textR.ok) return textR;
      if (isFormulaError(textR.value)) return { ok: true, value: textR.value };
      const startR = evaluate(args[1]!, ctx);
      if (!startR.ok) return startR;
      if (isFormulaError(startR.value)) return { ok: true, value: startR.value };
      const numR = evaluate(args[2]!, ctx);
      if (!numR.ok) return numR;
      if (isFormulaError(numR.value)) return { ok: true, value: numR.value };
      const s = toString(textR.value === null ? "" : textR.value);
      if (typeof s !== "string") return errorValue(s);
      const start = toNumber(startR.value === null ? 0 : startR.value);
      if (typeof start === "string") return errorValue(start);
      const num = toNumber(numR.value === null ? 0 : numR.value);
      if (typeof num === "string") return errorValue(num);
      if (!Number.isFinite(start) || !Number.isFinite(num)) return errorValue("#VALUE!");
      const at = Math.trunc(start);
      const count = Math.trunc(num);
      // Measured: start 0 is `#VALUE!`, length 0 and a start past the end are `""`.
      if (at < 1 || count < 0) return errorValue("#VALUE!");
      const chars = [...s];
      if (count === 0 || at > chars.length) return { ok: true, value: "" };
      return { ok: true, value: chars.slice(at - 1, at - 1 + count).join("") };
    }
    case "SEARCH":
    case "FIND":
      return searchArgs(name, args, ctx);
    case "CONCAT":
      return concatArgs(args, ctx, true);
    case "CONCATENATE":
      return concatArgs(args, ctx, false);
    case "ABS": {
      if (args.length !== 1) return fail("unsupported");
      const r = evalScalar(0);
      if (!r.ok) return r;
      if (isFormulaError(r.value)) return { ok: true, value: r.value };
      const n = toNumber(r.value);
      if (typeof n === "string") return errorValue(n);
      return { ok: true, value: Math.abs(n) };
    }
    case "MOD": {
      if (args.length !== 2) return fail("unsupported");
      const aR = evaluate(args[0]!, ctx);
      if (!aR.ok) return aR;
      if (isFormulaError(aR.value)) return { ok: true, value: aR.value };
      const bR = evaluate(args[1]!, ctx);
      if (!bR.ok) return bR;
      if (isFormulaError(bR.value)) return { ok: true, value: bR.value };
      const a = toNumber(aR.value);
      if (typeof a === "string") return errorValue(a);
      const b = toNumber(bR.value);
      if (typeof b === "string") return errorValue(b);
      if (b === 0) return errorValue("#DIV/0!");
      return { ok: true, value: a - b * Math.floor(a / b) };
    }
    case "TODAY": {
      if (args.length !== 0) return fail("unsupported");
      return { ok: true, value: ctx.todaySerial };
    }
    case "NOW": {
      if (args.length !== 0) return fail("unsupported");
      return { ok: true, value: ctx.nowSerial };
    }
    case "DATE": {
      if (args.length !== 3) return fail("unsupported");
      const yR = evaluate(args[0]!, ctx);
      if (!yR.ok) return yR;
      if (isFormulaError(yR.value)) return { ok: true, value: yR.value };
      const mR = evaluate(args[1]!, ctx);
      if (!mR.ok) return mR;
      if (isFormulaError(mR.value)) return { ok: true, value: mR.value };
      const dR = evaluate(args[2]!, ctx);
      if (!dR.ok) return dR;
      if (isFormulaError(dR.value)) return { ok: true, value: dR.value };
      const y = toNumber(yR.value);
      if (typeof y === "string") return errorValue(y);
      const m = toNumber(mR.value);
      if (typeof m === "string") return errorValue(m);
      const d = toNumber(dR.value);
      if (typeof d === "string") return errorValue(d);
      const v = dateToSerial(y, m, d, ctx.date1904);
      return typeof v === "string" ? errorValue(v) : { ok: true, value: v };
    }
    case "YEAR":
    case "MONTH":
    case "DAY": {
      if (args.length !== 1) return fail("unsupported");
      const r = evalScalar(0);
      if (!r.ok) return r;
      if (isFormulaError(r.value)) return { ok: true, value: r.value };
      const n = toNumber(r.value);
      if (typeof n === "string") return errorValue(n);
      const part = decomposeSerial(n, ctx.date1904);
      if (typeof part === "string") return errorValue(part);
      return { ok: true, value: name === "YEAR" ? part.year : name === "MONTH" ? part.month : part.day };
    }
    case "WEEKDAY": {
      if (args.length < 1 || args.length > 2) return fail("unsupported");
      const sR = evaluate(args[0]!, ctx);
      if (!sR.ok) return sR;
      if (isFormulaError(sR.value)) return { ok: true, value: sR.value };
      const n = toNumber(sR.value);
      if (typeof n === "string") return errorValue(n);
      let kind = 1;
      if (args.length === 2) {
        const kR = evaluate(args[1]!, ctx);
        if (!kR.ok) return kR;
        if (isFormulaError(kR.value)) return { ok: true, value: kR.value };
        const k = toNumber(kR.value);
        if (typeof k === "string") return errorValue(k);
        if (!Number.isFinite(k)) return errorValue("#VALUE!");
        kind = Math.trunc(k);
        // Only Sunday-first/Monday-first shapes are in the subset; the
        // remaining return types report instead of guessing.
        if (kind !== 1 && kind !== 2 && kind !== 3) return errorValue("#NUM!");
      }
      const v = weekdayOf(n, kind as 1 | 2 | 3, ctx.date1904);
      return typeof v === "string" ? errorValue(v) : { ok: true, value: v };
    }
    case "SUM": {
      const collected = collectArgNumbers(args, ctx, false);
      if (!("numbers" in collected)) return collected;
      return { ok: true, value: collected.numbers.reduce((s: number, v: number) => s + v, 0) };
    }
    case "COUNT": {
      const collected = collectArgNumbers(args, ctx, true);
      if (!("numbers" in collected)) return collected;
      return { ok: true, value: collected.numbers.length };
    }
    case "MIN":
    case "MAX": {
      if (args.length === 0) return errorValue("#VALUE!");
      const collected = collectArgNumbers(args, ctx, false);
      if (!("numbers" in collected)) return collected;
      const numbers = collected.numbers;
      if (numbers.length === 0) return { ok: true, value: 0 };
      return { ok: true, value: name === "MIN" ? Math.min(...numbers) : Math.max(...numbers) };
    }
    case "AVERAGE": {
      if (args.length === 0) return errorValue("#DIV/0!");
      const collected = collectArgNumbers(args, ctx, false);
      if (!("numbers" in collected)) return collected;
      const numbers = collected.numbers;
      if (numbers.length === 0) return errorValue("#DIV/0!");
      return { ok: true, value: numbers.reduce((s: number, v: number) => s + v, 0) / numbers.length };
    }
    case "COUNTIF":
    case "COUNTIFS": {
      // One uniform path: `COUNTIF` is a single pair. Arity slips report;
      // `COUNTIFS` ranges of different sizes are Excel's `#VALUE!`.
      const got = ifsPairs(args, ctx, 0);
      if (!("pairs" in got)) return got;
      const n = countPairs(got.pairs);
      return n === undefined ? fail("unsupported") : { ok: true, value: n };
    }
    case "SUMIF": {
      if (args.length < 2 || args.length > 3) return fail("unsupported");
      const sum = resolveAndScan(args.length === 3 ? args[2]! : args[0]!, ctx);
      if (!("map" in sum)) return sum;
      const got = ifsPairs([args[0]!, args[1]!], ctx, 0);
      if (!("pairs" in got)) return got;
      // A short sum range aligns at the top-left (measured); missing cells read blank.
      return sumPairs(sum, got.pairs);
    }
    case "SUMIFS": {
      if (args.length < 3) return fail("unsupported");
      const sum = resolveAndScan(args[0]!, ctx);
      if (!("map" in sum)) return sum;
      // Native Excel evaluation of the read-only Shot Log confirmed #VALUE!
      // for both row/column mismatches and equal-area 3x1 versus 1x3 ranges.
      // SUMIF's top-left overlay is a different rule.
      const got = ifsPairs(args, ctx, 1, sum);
      if (!("pairs" in got)) return got;
      return sumPairs(sum, got.pairs);
    }
    case "AVERAGEIF": {
      if (args.length < 2 || args.length > 3) return fail("unsupported");
      const sum = resolveAndScan(args.length === 3 ? args[2]! : args[0]!, ctx);
      if (!("map" in sum)) return sum;
      const got = ifsPairs([args[0]!, args[1]!], ctx, 0);
      if (!("pairs" in got)) return got;
      return averagePairs(sum, got.pairs);
    }
    case "AVERAGEIFS": {
      if (args.length < 3) return fail("unsupported");
      const sum = resolveAndScan(args[0]!, ctx);
      if (!("map" in sum)) return sum;
      const got = ifsPairs(args, ctx, 1, sum);
      if (!("pairs" in got)) return got;
      return averagePairs(sum, got.pairs);
    }
    case "PERCENTILE":
    case "PERCENTILE.INC": {
      if (args.length !== 2) return fail("unsupported");
      const arrayArg = args[0]!;
      let numbers: number[];
      if (arrayArg.t === "range") {
        const got = rangeNumbers(arrayArg, ctx);
        if (!("numbers" in got)) return got;
        numbers = [...got.numbers].sort((x: number, y: number) => x - y);
      } else {
        const collected = collectArgNumbers([arrayArg], ctx, false);
        if (!("numbers" in collected)) return collected;
        numbers = [...collected.numbers].sort((x: number, y: number) => x - y);
      }
      const kR = evaluate(args[1]!, ctx);
      if (!kR.ok) return kR;
      if (isFormulaError(kR.value)) return { ok: true, value: kR.value };
      const k = toNumber(kR.value === null ? 0 : kR.value);
      if (typeof k === "string") return errorValue(k);
      // PERCENTILE legacy k is 0-1 like PERCENTILE.INC (measured 0.9 over F2:F6 is 9).
      const v = percentileOf(numbers, k);
      return typeof v === "string" ? errorValue(v) : { ok: true, value: v };
    }
    default:
      return fail("unsupported");
  }
}

export function evaluate(node: Node, ctx: FormulaCellSource): FormulaResult {
  const over = checkBudget(ctx);
  if (over !== undefined) return over;
  switch (node.t) {
    case "n": return { ok: true, value: node.v };
    case "s": return { ok: true, value: node.v };
    case "b": return { ok: true, value: node.v };
    case "e": return { ok: true, value: { error: node.code } };
    case "u": {
      const r = evaluate(node.arg, ctx);
      if (!r.ok) return r;
      if (isFormulaError(r.value)) return { ok: true, value: r.value };
      const n = toNumber(r.value);
      if (typeof n === "string") return errorValue(n);
      return { ok: true, value: node.op === "-" ? -n : n };
    }
    case "pct": {
      const r = evaluate(node.arg, ctx);
      if (!r.ok) return r;
      if (isFormulaError(r.value)) return { ok: true, value: r.value };
      const n = toNumber(r.value);
      if (typeof n === "string") return errorValue(n);
      return { ok: true, value: n / 100 };
    }
    case "op": {
      const l = evaluate(node.l, ctx);
      if (!l.ok) return l;
      const r = evaluate(node.r, ctx);
      if (!r.ok) return r;
      if (node.op === "&") {
        if (isFormulaError(l.value)) return { ok: true, value: l.value };
        if (isFormulaError(r.value)) return { ok: true, value: r.value };
        const a = toString(l.value === null ? "" : l.value);
        if (typeof a !== "string") return errorValue(a);
        const b = toString(r.value === null ? "" : r.value);
        if (typeof b !== "string") return errorValue(b);
        return { ok: true, value: a + b };
      }
      if (node.op === "=" || node.op === "<>" || node.op === "<" || node.op === ">" || node.op === "<=" || node.op === ">=") {
        const v = compare(l.value, r.value, node.op);
        return typeof v === "string" ? errorValue(v as FormulaErrorCode) : { ok: true, value: v };
      }
      if (isFormulaError(l.value)) return { ok: true, value: l.value };
      if (isFormulaError(r.value)) return { ok: true, value: r.value };
      const v = add(l.value, r.value, node.op);
      return typeof v === "string" ? errorValue(v as FormulaErrorCode) : { ok: true, value: v };
    }
    case "ref": {
      const at = resolveRef(node.row, node.col, node.absRow, node.absCol, ctx.anchorRow, ctx.anchorCol, ctx.targetRow, ctx.targetCol);
      if (!isInsideGrid(at.row, at.col)) return errorValue("#REF!");
      if (node.sheet !== undefined && node.sheet.toLowerCase() !== ctx.sheetLower) {
        if (ctx.crossScalarAt === undefined) return fail("unsupported");
        return ctx.crossScalarAt(node.sheet.toLowerCase(), at.row, at.col);
      }
      return ctx.scalarAt(at.row, at.col);
    }
    case "range": {
      // A bare range outside aggregates has no scalar meaning in this subset.
      return fail("unsupported");
    }
    case "f":
      return evaluateFunc(node, ctx);
  }
}

export function compareScalars(a: FormulaScalar, b: FormulaScalar, op: string): boolean | FormulaErrorCode {
  return compare(a, b, op);
}

/** Truth for `expression` rules: errors never match; numbers follow 0=FALSE. */
export function isTruthyForCF(value: FormulaScalar): boolean | FormulaErrorCode {
  if (isFormulaError(value)) return value.error;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  if (value === null) return false;
  // Measured: IF("",1,2) is #VALUE!, so a bare "" condition is an error, not FALSE.
  return "#VALUE!" as FormulaErrorCode;
}

export type CompiledFormula = { node: Node; source: string };

export function compileFormula(text: string): CompiledFormula | undefined {
  const node = parseFormula(text);
  return node === undefined ? undefined : { node, source: text };
}
