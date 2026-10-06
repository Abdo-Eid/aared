/**
 * A1 references (`"A1:M200"`, `"D14"`): the scanner every reader shares, and
 * the range parser behind `SheetHandle.read()`. Pure and headless: the range
 * parser runs on the main thread, validated before anything is sent to the
 * Worker. Zero-based inclusive output in logical order.
 */

import { AaredError } from "../errors.js";
import type { CellRange } from "../model.js";
import { isInsideGrid } from "./limits.js";

export type A1Range = string;

interface Cell {
  row: number;
  col: number;
}

/**
 * Zero-based cell from an A1 token (one to three letters, then one to seven
 * digits, no `$`); `undefined` when malformed or off the worksheet grid.
 * Hand-rolled instead of a regex because the sheet parser runs it once per cell.
 */
export function parseCell(token: string): Cell | undefined {
  const length = token.length;
  let i = 0;
  let col = 0;
  for (; i < length && i < 4; i++) {
    let code = token.charCodeAt(i);
    if (code >= 97 && code <= 122) code -= 32;
    if (code < 65 || code > 90) break;
    col = col * 26 + (code - 64);
  }
  const digits = length - i;
  if (i < 1 || i > 3 || digits < 1 || digits > 7) return undefined;
  let row = 0;
  for (; i < length; i++) {
    const code = token.charCodeAt(i);
    if (code < 48 || code > 57) return undefined;
    row = row * 10 + (code - 48);
  }
  return isInsideGrid(row - 1, col - 1) ? { row: row - 1, col: col - 1 } : undefined;
}

/** The inclusive range two corners span, in logical order however they were written. */
export function spanRange(first: Cell, second: Cell): CellRange {
  return {
    start: { row: Math.min(first.row, second.row), col: Math.min(first.col, second.col) },
    end: { row: Math.max(first.row, second.row), col: Math.max(first.col, second.col) },
  };
}

/**
 * An A1 reference as a zero-based inclusive `CellRange`: a single cell
 * (`"D14"`) or a bounded range (`"A1:M200"`), case-insensitive, `$` markers
 * and surrounding whitespace allowed, reversed corners put in logical order.
 * `undefined` for anything else (whole column `"A:A"`, whole row `"1:3"`,
 * off the grid, malformed): the worksheet readers skip such a reference.
 */
export function parseRef(reference: string): CellRange | undefined {
  const parts = reference.trim().split(":");
  if (parts.length > 2) return undefined;
  const first = parseCell(parts[0]!.trim().replaceAll("$", ""));
  const second = parts.length === 2 ? parseCell(parts[1]!.trim().replaceAll("$", "")) : first;
  return first === undefined || second === undefined ? undefined : spanRange(first, second);
}

/** `parseRef` for the public API: anything it rejects is an `invalid-range` `AaredError`. */
export function parseA1Range(reference: A1Range): CellRange {
  const range = typeof reference === "string" ? parseRef(reference) : undefined;
  if (range === undefined) {
    throw new AaredError(
      "invalid-range",
      `aared/xlsx: invalid cell range ${JSON.stringify(reference)}; use "A1:M200" or a single cell like "D14"`,
    );
  }
  return range;
}

/** Zero-based column index to Excel letters: 0 → A … 25 → Z, 26 → AA. */
export function columnName(index: number): string {
  let name = "";
  let i = index;
  do {
    name = String.fromCharCode(65 + (i % 26)) + name;
    i = Math.floor(i / 26) - 1;
  } while (i >= 0);
  return name;
}

/** Zero-based cell to A1 text, e.g. `{ row: 13, col: 3 }` → `"D14"`. */
export function cellAddress(row: number, col: number): string {
  return `${columnName(col)}${row + 1}`;
}
