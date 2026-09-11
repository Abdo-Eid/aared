import { expect, test } from "bun:test";
import { AaredError } from "../src/core/index.js";
import { open } from "../src/xlsx/index.js";

const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-xlsx-rtl-coverage.xlsx`;
const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-pptx-rtl-coverage.pptx`;

async function rejection(promise: Promise<unknown>): Promise<Error> {
  try {
    await promise;
  } catch (error) {
    return error instanceof Error ? error : new Error(String(error));
  }
  throw new Error("expected the request to reject");
}

test("opens a workbook with an addressable sheet collection", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    expect(workbook.info.format).toBe("xlsx");
    const sheets = workbook.sheets.list();
    expect(sheets.map((sheet) => sheet.info.name)).toEqual(["الميزانية", "Summary", "Shot Log", "Notes"]);
    for (const sheet of sheets) expect(sheet.info.state).toBe("visible");
    expect(workbook.sheets.get(1).info.name).toBe("Summary");
  } finally {
    workbook.close();
  }
});

test("looks sheets up by name, case-insensitively as a fallback", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    expect(workbook.sheets.get("Summary").info.index).toBe(1);
    expect(workbook.sheets.get("summary").info.index).toBe(1);
    expect(() => workbook.sheets.get("Revenue")).toThrow('no sheet named "Revenue"');
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
    const window = await sheet.read("A1:C3");
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
    const again = await sheet.read("A1:C3");
    expect(again).toEqual(window);
  } finally {
    workbook.close();
  }
});

test("reads single cells and structured ranges", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const sheet = workbook.sheets.get("Summary");
    const single = await sheet.read("B2");
    expect(single.range).toEqual({ start: { row: 1, col: 1 }, end: { row: 1, col: 1 } });
    const structured = await sheet.read({ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } });
    expect(structured.range).toEqual({ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } });
    const reversed = await sheet.read("C3:A1");
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
      const error = await rejection(sheet.read(bad));
      expect((error as AaredError).code).toBe("invalid-range");
    }
    const numeric = await rejection(
      sheet.read({ start: { row: 5, col: 5 }, end: { row: 1, col: 1 } }),
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
    expect(matches.some((match) => match.sheet === "Summary")).toBe(true);
    for (const match of matches) {
      expect(match.cell).toMatch(/^[A-Z]+[0-9]+$/);
      expect(match.text.length).toBeGreaterThan(0);
    }
  } finally {
    workbook.close();
  }
});

test("rejects non-XLSX files and exposes spreadsheet capabilities", async () => {
  const error = await rejection(open(await Bun.file(PPTX_FIXTURE).arrayBuffer()));
  expect((error as AaredError).code).toBe("unsupported-format");
  expect(error.message).toContain("aared/pptx");

  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    expect(workbook.capabilities).toEqual({
      textSelection: true,
      media: true,
      spreadsheet: { rangeRead: true, cellNavigation: true, selection: true, formulas: "cached-only" },
      editing: false,
    });
  } finally {
    workbook.close();
  }
});
