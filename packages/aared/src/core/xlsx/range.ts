/**
 * A1 range references (`"A1:M200"`, `"D14"`) for `SheetHandle.read()`.
 * Pure and headless: runs on the main thread, validated before anything
 * is sent to the Worker. Zero-based inclusive output in logical order.
 */

import { AaredError } from "../errors.js";
import type { CellRange } from "../model.js";

export type A1Range = string;

function columnLettersToIndex(letters: string): number | undefined {
  if (!/^[A-Za-z]{1,3}$/.test(letters)) return undefined;
  let index = 0;
  for (const char of letters.toUpperCase()) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  const zeroBased = index - 1;
  // Three letters max out at XFD (16383), Excel's last column.
  if (zeroBased < 0 || zeroBased > 16383) return undefined;
  return zeroBased;
}

function parseCell(token: string): { row: number; col: number } | undefined {
  const cell = token.replaceAll("$", "");
  const match = /^([A-Za-z]{1,3})([0-9]{1,7})$/.exec(cell);
  if (match === null) return undefined;
  const col = columnLettersToIndex(match[1]!);
  const row = Number(match[2]!) - 1;
  if (col === undefined || row < 0 || row > 1048575) return undefined;
  return { row, col };
}

/**
 * Parse an A1 reference into a zero-based inclusive `CellRange`.
 * Accepts a single cell (`"D14"`) or a bounded range (`"A1:M200"`).
 * Reversed corners are normalized to logical order. Anything else —
 * whole-column (`"A:A"`), whole-row (`"1:3"`), unbounded, or malformed
 * input — rejects with an `invalid-range` `AaredError`.
 */
export function parseA1Range(reference: A1Range): CellRange {
  const invalid = new AaredError(
    "invalid-range",
    `aared/xlsx: invalid cell range ${JSON.stringify(reference)}; use "A1:M200" or a single cell like "D14"`,
  );
  if (typeof reference !== "string") throw invalid;
  const parts = reference.trim().split(":");
  if (parts.length === 1) {
    const cell = parseCell(parts[0]!.trim());
    if (cell === undefined) throw invalid;
    return { start: cell, end: { row: cell.row, col: cell.col } };
  }
  if (parts.length === 2) {
    const first = parseCell(parts[0]!.trim());
    const second = parseCell(parts[1]!.trim());
    if (first === undefined || second === undefined) throw invalid;
    return {
      start: { row: Math.min(first.row, second.row), col: Math.min(first.col, second.col) },
      end: { row: Math.max(first.row, second.row), col: Math.max(first.col, second.col) },
    };
  }
  throw invalid;
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
