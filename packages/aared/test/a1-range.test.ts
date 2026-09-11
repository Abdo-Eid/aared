import { expect, test } from "bun:test";
import { AaredError } from "../src/core/index.js";
import { cellAddress, columnName, parseA1Range } from "../src/core/xlsx/range.js";

function invalidRange(reference: string): AaredError {
  try {
    parseA1Range(reference);
  } catch (error) {
    if (error instanceof AaredError && error.code === "invalid-range") return error;
    throw new Error(`expected an invalid-range AaredError for ${JSON.stringify(reference)}, got ${String(error)}`);
  }
  throw new Error(`expected parseA1Range(${JSON.stringify(reference)}) to throw`);
}

test("parses a single cell to a one-cell window", () => {
  expect(parseA1Range("D14")).toEqual({ start: { row: 13, col: 3 }, end: { row: 13, col: 3 } });
  expect(parseA1Range("A1")).toEqual({ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
});

test("parses a bounded range", () => {
  expect(parseA1Range("A1:M200")).toEqual({ start: { row: 0, col: 0 }, end: { row: 199, col: 12 } });
});

test("tolerates lowercase, whitespace, and absolute markers", () => {
  expect(parseA1Range("  a1:m200 ")).toEqual(parseA1Range("A1:M200"));
  expect(parseA1Range("$D$14")).toEqual(parseA1Range("D14"));
});

test("normalizes reversed corners to logical order", () => {
  expect(parseA1Range("M200:A1")).toEqual(parseA1Range("A1:M200"));
});

test("rejects whole-column, whole-row, and malformed references", () => {
  for (const reference of ["", "A", "1", "A:A", "1:3", "A1:M200:Z9", "XFE1", "A1048577", "A0", "!!!"]) {
    const error = invalidRange(reference);
    expect(error.message).toContain("aared/xlsx:");
  }
});

test("column letters and cell addresses round-trip", () => {
  expect(columnName(0)).toBe("A");
  expect(columnName(25)).toBe("Z");
  expect(columnName(26)).toBe("AA");
  expect(cellAddress(13, 3)).toBe("D14");
  const round = parseA1Range(cellAddress(199, 12));
  expect(round).toEqual({ start: { row: 199, col: 12 }, end: { row: 199, col: 12 } });
});
