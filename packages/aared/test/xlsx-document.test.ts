import { expect, test } from "bun:test";
import { AaredError } from "../src/core/index.js";
import { open } from "../src/xlsx/index.js";
import { rejection } from "./harness.js";

const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;
const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

test("opens a workbook with an addressable sheet collection", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    expect(workbook.info.format).toBe("xlsx");
    const sheets = workbook.sheets.list();
    expect(sheets.map((sheet) => sheet.info.title)).toEqual(["الميزانية", "Summary", "Shot Log", "Notes"]);
    for (const sheet of sheets) expect(sheet.info.state).toBe("visible");
    expect(workbook.sheets.get(1).info.title).toBe("Summary");
  } finally {
    workbook.close();
  }
});

test("looks sheets up by title — exact, then ignoring case, never a substring", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    expect(workbook.sheets.get("Summary").info.index).toBe(1);
    expect(workbook.sheets.get("summary").info.index).toBe(1);
    expect(() => workbook.sheets.get("Revenue")).toThrow('no sheet titled "Revenue"');
    expect(() => workbook.sheets.get("Summ")).toThrow("no sheet titled");
  } finally {
    workbook.close();
  }
});

test("reads a whole sheet as a full-extent range model", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const sheet = workbook.sheets.get("Summary");
    const model = await sheet.read();
    expect(model.name).toBe("Summary");
    expect(model.range).toEqual({
      start: { row: 0, col: 0 },
      end: { row: model.extent.rows - 1, col: model.extent.cols - 1 },
    });
    expect(model.rows.length).toBeGreaterThan(0);
  } finally {
    workbook.close();
  }
});

test("reads A1 windows with clipped merges and an echoed range", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const sheet = workbook.sheets.get("Summary");
    const window = await sheet.read({ range: "A1:C3" });
    expect(window.range).toEqual({ start: { row: 0, col: 0 }, end: { row: 2, col: 2 } });
    for (const row of window.rows) {
      expect(row.index).toBeLessThan(3);
      for (const { col } of row.cells) expect(col).toBeLessThan(3);
    }
    for (const merge of window.merges) {
      expect(merge.start.row).toBeGreaterThanOrEqual(0);
      expect(merge.end.row).toBeLessThanOrEqual(2);
      expect(merge.start.col).toBeGreaterThanOrEqual(0);
      expect(merge.end.col).toBeLessThanOrEqual(2);
    }
    // A repeated windowed read serves the cached sheet without re-parsing.
    const again = await sheet.read({ range: "A1:C3" });
    expect(again).toEqual(window);
  } finally {
    workbook.close();
  }
});

test("reads single cells and structured ranges", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const sheet = workbook.sheets.get("Summary");
    const single = await sheet.read({ range: "B2" });
    expect(single.range).toEqual({ start: { row: 1, col: 1 }, end: { row: 1, col: 1 } });
    const structured = await sheet.read({ range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } } });
    expect(structured.range).toEqual({ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } });
    const reversed = await sheet.read({ range: "C3:A1" });
    expect(reversed.range).toEqual({ start: { row: 0, col: 0 }, end: { row: 2, col: 2 } });
  } finally {
    workbook.close();
  }
});

test("rejects malformed ranges with invalid-range errors", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const sheet = workbook.sheets.get("Summary");
    for (const bad of ["A:A", "XFE1", "A1:M200:Z9", ""]) {
      const error = await rejection(sheet.read({ range: bad }));
      expect((error as AaredError).code).toBe("invalid-range");
    }
    const numeric = await rejection(
      sheet.read({ range: { start: { row: 5, col: 5 }, end: { row: 1, col: 1 } } }),
    );
    expect((numeric as AaredError).code).toBe("invalid-range");
  } finally {
    workbook.close();
  }
});

test("searches formatted cell text across sheets", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    expect(await workbook.search("")).toEqual([]);
    const matches = await workbook.search("budget");
    expect(matches.length).toBeGreaterThan(0);
    expect(matches.some((match) => match.sheet.title === "Summary")).toBe(true);
    for (const match of matches) {
      // A match names its sheet as { index, title } and its cell as { address, row, col }.
      expect(workbook.sheets.get(match.sheet.index).info.title).toBe(match.sheet.title);
      expect(match.address).toMatch(/^[A-Z]+[0-9]+$/);
      expect(Number.isInteger(match.row) && Number.isInteger(match.col)).toBe(true);
      expect(match.text.length).toBeGreaterThan(0);
    }
  } finally {
    workbook.close();
  }
});

test("rejects non-XLSX files", async () => {
  const error = await rejection(open(await Bun.file(PPTX_FIXTURE).arrayBuffer()));
  expect((error as AaredError).code).toBe("unsupported-format");
  expect(error.message).toContain("aared/pptx");
});
