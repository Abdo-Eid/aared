import { expect, test } from "bun:test";
import { AaredError } from "../src/core/index.js";
import { MAX_COLUMNS, MAX_ROWS, isInsideGrid } from "../src/core/xlsx/limits.js";
import { cellAddress, columnName, parseA1Range, parseCell, parseRef, spanRange } from "../src/core/xlsx/range.js";

function invalidRange(reference: string): AaredError {
  try {
    parseA1Range(reference);
  } catch (error) {
    if (error instanceof AaredError && error.code === "invalid-range") return error;
    throw new Error(`expected an invalid-range AaredError for ${JSON.stringify(reference)}, got ${String(error)}`);
  }
  throw new Error(`expected parseA1Range(${JSON.stringify(reference)}) to throw`);
}

test("parses one cell, a bounded range, and the loose corners Excel accepts", () => {
  expect(parseA1Range("D14")).toEqual({ start: { row: 13, col: 3 }, end: { row: 13, col: 3 } });
  expect(parseA1Range("A1:M200")).toEqual({ start: { row: 0, col: 0 }, end: { row: 199, col: 12 } });
  // Lowercase, surrounding whitespace, absolute markers and reversed corners
  // all name the same window Excel's own reference box would.
  expect(parseA1Range("  a1:m200 ")).toEqual(parseA1Range("A1:M200"));
  expect(parseA1Range("$D$14")).toEqual(parseA1Range("D14"));
  expect(parseA1Range("M200:A1")).toEqual(parseA1Range("A1:M200"));
});

test("rejects whole-column, whole-row, and malformed references", () => {
  for (const reference of ["", "A", "1", "A:A", "1:3", "A1:M200:Z9", "XFE1", "A1048577", "A0", "!!!"]) {
    const error = invalidRange(reference);
    expect(error.message).toContain("aared/xlsx:");
  }
});

test("the worksheet grid is XFD1048576, and its edges are the only cells that fit", () => {
  expect(MAX_ROWS).toBe(1048576);
  expect(MAX_COLUMNS).toBe(16384);
  expect(isInsideGrid(0, 0)).toBe(true);
  expect(isInsideGrid(MAX_ROWS - 1, MAX_COLUMNS - 1)).toBe(true);
  for (const [row, col] of [[-1, 0], [0, -1], [MAX_ROWS, 0], [0, MAX_COLUMNS], [1.5, 0], [0, Number.NaN]] as const) {
    expect(isInsideGrid(row, col)).toBe(false);
  }
  expect(parseCell("XFD1048576")).toEqual({ row: MAX_ROWS - 1, col: MAX_COLUMNS - 1 });
});

test("the cell scanner is the one grammar behind both readers", () => {
  // Case-insensitive; 1-3 letters then 1-7 digits; `$` is the range parser's business, not the scanner's.
  expect(parseCell("aa10")).toEqual({ row: 9, col: 26 });
  for (const token of ["", "A", "1", "ABCD1", "A12345678", "A0", "XFE1", "A1048577", "A1B", "$A$1", "A1 "]) {
    expect(parseCell(token)).toBeUndefined();
  }
});

test("the worksheet readers' parseRef is the grammar behind the API's parseA1Range", () => {
  expect(parseRef("H34")).toEqual({ start: { row: 33, col: 7 }, end: { row: 33, col: 7 } });
  expect(parseRef("A1:H1")).toEqual({ start: { row: 0, col: 0 }, end: { row: 0, col: 7 } });
  expect(parseRef("H34:A1")).toEqual(parseA1Range("A1:H34"));
  for (const reference of ["", "nope", "XFE1", "A:A", "A1:M200:Z9"]) expect(parseRef(reference)).toBeUndefined();
  expect(spanRange({ row: 5, col: 1 }, { row: 2, col: 4 })).toEqual({ start: { row: 2, col: 1 }, end: { row: 5, col: 4 } });
});

test("column letters and cell addresses round-trip", () => {
  expect(columnName(0)).toBe("A");
  expect(columnName(25)).toBe("Z");
  expect(columnName(26)).toBe("AA");
  expect(cellAddress(13, 3)).toBe("D14");
  const round = parseA1Range(cellAddress(199, 12));
  expect(round).toEqual({ start: { row: 199, col: 12 }, end: { row: 199, col: 12 } });
});
