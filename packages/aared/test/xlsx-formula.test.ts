import { describe, expect, test } from "bun:test";
import type { CellRange } from "../src/core/model.js";
import {
  dateToSerial,
  evaluate,
  makeFormulaBudget,
  parseFormula,
  type FormulaCellSource,
  type FormulaResult,
  type FormulaScalar,
} from "../src/core/xlsx/formula.js";

// The conditional-format formula evaluator at the rule level: exact values,
// not "which cells painted". The paint outcomes are read from a real workbook
// in xlsx-conditional.test.ts; these cases pin what each supported function
// actually returns, so a bound chosen loosely there cannot hide a wrong value
// here. The 1900 date system, Excel's two-digit-year rule and the serial
// offsets are the measured behaviour (Notes §§105-111).

/** A source over one sheet's cached values; an unlisted cell is blank. */
function source(cells: Record<string, FormulaScalar>, todaySerial = 46300): FormulaCellSource {
  const at = (row: number, col: number): FormulaScalar => cells[`${row}:${col}`] ?? null;
  const each = (box: CellRange, visit: (value: FormulaScalar) => void): void => {
    for (let row = box.start.row; row <= box.end.row; row++) {
      for (let col = box.start.col; col <= box.end.col; col++) visit(at(row, col));
    }
  };
  return {
    sheetLower: "shot log",
    anchorRow: 4,
    anchorCol: 5,
    targetRow: 4,
    targetCol: 5,
    todaySerial,
    nowSerial: todaySerial + 0.5,
    date1904: false,
    scalarAt: (row, col) => ({ ok: true, value: at(row, col) }),
    numbersIn: (box) => {
      const numbers: number[] = [];
      each(box, (value) => { if (typeof value === "number") numbers.push(value); });
      return { ok: true, numbers };
    },
    scanIn: (box) => {
      const entries: Array<{ row: number; col: number; value: FormulaScalar }> = [];
      let blanks = 0;
      for (let row = box.start.row; row <= box.end.row; row++) {
        for (let col = box.start.col; col <= box.end.col; col++) {
          const value = at(row, col);
          if (value === null) blanks++;
          else entries.push({ row, col, value });
        }
      }
      return { ok: true, scan: { entries, blanks } };
    },
    budget: makeFormulaBudget(),
  };
}

/** Evaluate a formula's value, its error code, or the reason it reported. */
function value(formula: string, cells: Record<string, FormulaScalar> = {}, todaySerial?: number): unknown {
  const node = parseFormula(formula);
  if (node === undefined) return "unparsed";
  const result: FormulaResult = evaluate(node, source(cells, todaySerial));
  if (!result.ok) return result.reason;
  // A scalar error is itself a value: `#DIV/0!` and friends travel as data.
  return typeof result.value === "object" && result.value !== null ? result.value.error : result.value;
}

describe("XLSX date functions", () => {
  test("DATE builds Excel's own serials, including the leap day that never existed", () => {
    expect(value("DATE(2026,10,11)")).toBe(46306);
    // Month 13 rolls into the next January, as Excel's DATE does.
    expect(value("DATE(2026,13,1)")).toBe(value("DATE(2027,1,1)"));
    // Day 0 is the last day of the previous month.
    expect(value("DATE(2026,3,0)")).toBe(value("DATE(2026,2,28)"));
    // The fictitious 1900-02-29 is serial 60; anything past it spills into March.
    expect(value("DATE(1900,2,29)")).toBe(60);
    expect(value("DATE(1900,3,1)")).toBe(61);
    // A serial below 1 is `#NUM!`, not a wrapped date.
    expect(value("DATE(-5,1,1)")).toBe("#NUM!");
    // Excel's two-digit-year rule spans four digits.
    expect(value("DATE(26,1,1)")).toBe(value("DATE(1926,1,1)"));
  });

  test("dateToSerial is the date system behind those values", () => {
    expect(dateToSerial(2026, 10, 11, false)).toBe(46306);
    expect(dateToSerial(1900, 2, 29, false)).toBe(60);
    expect(dateToSerial(1900, 2, 30, false)).toBe(61);
    // The Mac date system runs 1462 days behind and has no fictitious day.
    expect(dateToSerial(2026, 10, 11, true)).toBe(46306 - 1462);
    expect(dateToSerial(1900, 2, 29, true)).toBe(dateToSerial(1900, 3, 1, true));
    expect(dateToSerial(Number.NaN, 1, 1, false)).toBe("#VALUE!");
  });

  test("YEAR, MONTH and DAY decompose a serial", () => {
    expect(value("YEAR(DATE(2026,10,11))")).toBe(2026);
    expect(value("MONTH(DATE(2026,10,11))")).toBe(10);
    expect(value("DAY(DATE(2026,10,11))")).toBe(11);
  });

  test("WEEKDAY is Sunday-first by default and refuses a kind it does not implement", () => {
    // 2026-10-11 is a Sunday, which Excel's default type 1 numbers 1.
    expect(value("WEEKDAY(DATE(2026,10,11))")).toBe(1);
    expect(value("WEEKDAY(DATE(2026,10,11),2)")).toBe(7);
    expect(value("WEEKDAY(DATE(2026,10,11),3)")).toBe(6);
    expect(value("WEEKDAY(DATE(2026,10,11),99)")).toBe("#NUM!");
  });

  test("TODAY and NOW read the parse-time snapshot, never the clock", () => {
    // A cell far below the snapshot is in the past; one far above is not.
    expect(value("1000<TODAY()", {}, 46300)).toBe(true);
    expect(value("99999<TODAY()", {}, 46300)).toBe(false);
    expect(value("1000<NOW()", {}, 46300)).toBe(true);
    // NOW carries the parse-time day fraction, so it is past TODAY's midnight.
    expect(value("NOW()>TODAY()", {}, 46300)).toBe(true);
  });
});

describe("XLSX text functions", () => {
  test("SEARCH is case-insensitive with wildcards; FIND is literal and case-sensitive", () => {
    expect(value(`SEARCH("in*iew","In Review")`)).toBe(1);
    expect(value(`SEARCH("IN REVIEW","In Review")`)).toBe(1);
    expect(value(`FIND("In Review","In Review")`)).toBe(1);
    // FIND cannot match a different case or a wildcard.
    expect(value(`FIND("IN REVIEW","In Review")`)).toBe("#VALUE!");
    expect(value(`FIND("in*iew","In Review")`)).toBe("#VALUE!");
    expect(value(`FIND("missing","In Review")`)).toBe("#VALUE!");
  });

  test("MID is 1-based, as in Excel", () => {
    expect(value(`MID("In Review",1,2)`)).toBe("In");
    expect(value(`MID("In Review",4,6)`)).toBe("Review");
    expect(value(`MID("In Review",1,0)`)).toBe("");
  });

  test("CONCAT joins its arguments", () => {
    expect(value(`CONCAT("a","b","c")`)).toBe("abc");
    expect(value(`LEFT(CONCAT("n=","42"),2)`)).toBe("n=");
  });
});

describe("XLSX aggregate functions", () => {
  // A synthetic status/value grid for pure evaluator cases, not fixture data.
  const data: Record<string, FormulaScalar> = {
    "4:4": "In review", "5:4": "In review", "6:4": "Done",
    "4:5": 10, "5:5": 20, "6:5": 30,
  };
  const counted = (formula: string) => value(formula, data);

  test("COUNTIF applies Excel's criteria over a range", () => {
    expect(counted(`COUNTIF(E5:E6,"In review")`)).toBe(2);
    expect(counted(`COUNTIF(F5:F7,">15")`)).toBe(2);
    expect(counted(`COUNTIF(E5:E6,"In*")`)).toBe(2);
    // A blank is a blank, not an empty string.
    expect(counted(`COUNTIF(E5:E6,"")`)).toBe(0);
  });

  test("COUNTIFS pairs its ranges", () => {
    expect(counted(`COUNTIFS(E5:E7,"In review",F5:F7,">15")`)).toBe(1);
    // Excel refuses ranges of different sizes, as aared does.
    expect(counted(`COUNTIFS(E5:E6,"In review",F5:F7,">15")`)).toBe("#VALUE!");
  });

  test("SUMIF and SUMIFS sum the matching rows", () => {
    // SUMIF lets Excel expand the sum range to the criteria range's shape;
    // SUMIFS requires every range to be the same shape as the criteria pair.
    expect(counted(`SUMIF(E5:E6,"In review",F5:F7)`)).toBe(30);
    expect(counted(`SUMIFS(F5:F7,E5:E7,"In review")`)).toBe(30);
  });

  test("SUMIFS and AVERAGEIFS require the value range and every criteria range to have the same shape", () => {
    for (const name of ["SUMIFS", "AVERAGEIFS"]) {
      expect(counted(`${name}(F5:F7,E5:E6,"In review")`)).toBe("#VALUE!");
      expect(counted(`${name}(F5:F6,E5:E7,"In review")`)).toBe("#VALUE!");
      expect(counted(`${name}(F5:G7,E5:E7,"In review")`)).toBe("#VALUE!");
      // Equal cell counts are not equal shapes: 3x1 versus 1x3.
      expect(counted(`${name}(F5:H5,E5:E7,"In review")`)).toBe("#VALUE!");
      expect(counted(`${name}(F5:F7,E5:E7,"In review",F5:F6,">0")`)).toBe("#VALUE!");
    }
    expect(counted(`SUMIFS(F5:F7,E5:E7,"In review",F5:F7,">15")`)).toBe(20);
    expect(counted(`AVERAGEIFS(F5:F7,E5:E7,"In review")`)).toBe(15);
    // Same-sized ranges align by position, not absolute coordinates.
    expect(counted(`SUMIFS(F5:F6,E6:E7,"In review")`)).toBe(10);
    expect(counted(`AVERAGEIFS(F5:F6,E6:E7,"In review")`)).toBe(10);
  });

  test("AVERAGEIF over no matching row is #DIV/0!, as Excel reports it", () => {
    expect(counted(`AVERAGEIF(E5:E6,"Shipped",F5:F7)`)).toBe("#DIV/0!");
    expect(counted(`AVERAGEIF(E5:E7,"Done",F5:F7)`)).toBe(30);
  });
});

describe("XLSX formula parsing", () => {
  test("a comparison, a reference and a bare number all parse; an empty cell does not", () => {
    expect(parseFormula("F5>10")).toBeDefined();
    expect(parseFormula("$F$5>=5")).toBeDefined();
    expect(parseFormula('AND($F5>5,E5="In review")')).toBeDefined();
    expect(parseFormula("")).toBeUndefined();
    expect(parseFormula("   ")).toBeUndefined();
  });

  test("a relative reference shifts from the anchor to each target cell", () => {
    // The same formula text, read against two different target cells.
    const node = parseFormula("F5>10");
    if (node === undefined) throw new Error("expected the formula to parse");
    const at = (targetRow: number): unknown => {
      const ctx = { ...source({ "4:5": 10, "5:5": 20 }), targetRow, targetCol: 5 };
      const result: FormulaResult = evaluate(node, ctx);
      if (!result.ok) return result.reason;
      return typeof result.value === "object" && result.value !== null ? result.value.error : result.value;
    };
    // At row 5 the reference stays on itself (20 > 10); at row 6 it shifts to
    // the cell below, which the sheet leaves blank, so the rule is false.
    expect(at(5)).toBe(true);
    expect(at(6)).toBe(false);
  });
});
