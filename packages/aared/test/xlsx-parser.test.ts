import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexXlsx, readSharedStrings, readWorkbook } from "../src/core/xlsx/workbook.js";
import { parseCellRange, parseCellReference, parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { findCoverageSheet, xlsxCoverageManifest } from "./xlsx-fixture.js";
import type { SheetModel } from "../src/core/model.js";

// XLSX parsing against real workbooks: the long-lived production tracker and
// the Arabic RTL coverage fixture. Only the A1-notation helpers below are
// synthetic (pure functions, allowed by AGENTS.md § Testing rule).

const TRACKER_PATH = `${import.meta.dir}/../../../fixtures/post-production-tracker.xlsx`;
const COVERAGE_PATH = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;

async function tracker() {
  const pkg = await OpcPackage.open(await Bun.file(TRACKER_PATH).arrayBuffer());
  return readWorkbook(pkg).then((workbook) => ({ pkg, workbook }));
}

async function sheetOf(workbook: Awaited<ReturnType<typeof readWorkbook>>, pkg: OpcPackage, index: number): Promise<SheetModel> {
  return parseXlsxSheet(pkg, workbook, index);
}

function cellAt(sheet: SheetModel, row: number, col: number) {
  const line = sheet.rows.find((entry) => entry.index === row);
  const found = line?.cells.find((entry) => entry.col === col)?.cell;
  if (found === undefined) throw new Error(`fixture structure changed: expected a cell at ${row}:${col} in "${sheet.name}"`);
  return found;
}

describe("XLSX workbook index", () => {
  test("lists sheets cheaply without inflating them", async () => {
    const pkg = await OpcPackage.open(await Bun.file(TRACKER_PATH).arrayBuffer());
    const index = await indexXlsx(pkg);
    expect(index.sheets.map((sheet) => sheet.name)).toEqual(["Dashboard", "Shot Log", "الميزانية", "Notes"]);
    expect(index.sheets.every((sheet) => sheet.state === "visible")).toBe(true);
    expect(index.sheets.map((sheet) => sheet.part)).toEqual([
      "xl/worksheets/sheet1.xml",
      "xl/worksheets/sheet2.xml",
      "xl/worksheets/sheet3.xml",
      "xl/worksheets/sheet4.xml",
    ]);
    expect(index.date1904).toBe(false);
    expect(index.theme.colors["dk1"]).toBe("#000000");
  });

  test("reads shared strings including rich-text runs", async () => {
    const pkg = await OpcPackage.open(await Bun.file(TRACKER_PATH).arrayBuffer());
    const index = await indexXlsx(pkg);
    const strings = await readSharedStrings(pkg, index.workbookPart);
    expect(strings[0]).toBe("Post-Production Tracker");
    expect(strings[1]).toBe("متتبع أعمال ما بعد الإنتاج");
    expect(strings.some((text) => text.includes("غيّر هذه الخلية فقط"))).toBe(true);
  });

  test("indexes the Arabic coverage fixture", async () => {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE_PATH).arrayBuffer());
    const index = await indexXlsx(pkg);
    expect(index.sheets.map((sheet) => sheet.name)).toEqual(["الميزانية", "Summary", "Shot Log", "Notes"]);
  });

  test("rejects out-of-range sheet indexes with an actionable error", async () => {
    const { pkg, workbook } = await tracker();
    await expect(sheetOf(workbook, pkg, 9)).rejects.toThrow("sheet index 9 is outside this 4-sheet workbook");
  });

  test("rejects an already-aborted parser request with AbortError", async () => {
    const pkg = await OpcPackage.open(await Bun.file(TRACKER_PATH).arrayBuffer());
    const workbook = await readWorkbook(pkg);
    const controller = new AbortController();
    controller.abort();
    await expect(parseXlsxSheet(pkg, workbook, 0, controller.signal)).rejects.toMatchObject({
      name: "AbortError",
      message: "aared/core: request was cancelled",
    });
  });
});

describe("XLSX styles, values, and formats on the real tracker", () => {
  test("resolves fills, fonts, alignment, and borders on real cells", async () => {
    const { pkg, workbook } = await tracker();
    const dashboard = await sheetOf(workbook, pkg, 0);
    const header = cellAt(dashboard, 10, 0);
    expect(header.formatted).toBe("Episode");
    expect(header.style).toMatchObject({ fill: "#12171F", bold: true, align: "center" });
    expect(header.style?.border?.bottom?.style).toBe("thin");
    expect(header.style?.border?.bottom?.color).toBe("#D8D3C9");
    const kpi = cellAt(dashboard, 4, 0);
    expect(kpi.formatted).toBe("Shots tracked · عدد اللقطات");
    expect(kpi.style).toMatchObject({ fill: "#F4F1EC" });
  });

  test("shows cached formula results and keeps their source, never evaluating", async () => {
    const { pkg, workbook } = await tracker();
    const dashboard = await sheetOf(workbook, pkg, 0);
    const shots = cellAt(dashboard, 5, 0);
    expect(shots.value).toBe(12);
    expect(shots.formatted).toBe("12");
    expect(shots.formula).toContain("COUNTA");
    const hours = cellAt(dashboard, 5, 2);
    expect(hours.value).toBe(74);
    expect(hours.formatted).toBe("74.0");
    expect(hours.formula).toContain("SUM");
  });

  test("applies percent sections with parentheses and zero dashes", async () => {
    const { pkg, workbook } = await tracker();
    const budget = await sheetOf(workbook, pkg, 2);
    expect(cellAt(budget, 4, 3).formatted).toBe("(81.5)%");
    expect(cellAt(budget, 8, 3).formatted).toBe("24.4%");
    expect(cellAt(budget, 9, 3).formatted).toBe("-");
    expect(cellAt(budget, 4, 1).formatted).toBe("180,000");
  });

  test("reads merges, frozen panes, and sheet RTL", async () => {
    const { pkg, workbook } = await tracker();
    const dashboard = await sheetOf(workbook, pkg, 0);
    expect(dashboard.merges).toHaveLength(11);
    expect(dashboard.merges[0]).toEqual({ start: { row: 0, col: 0 }, end: { row: 0, col: 7 } });
    expect(dashboard.frozen).toEqual({ rows: 0, cols: 0 });
    const log = await sheetOf(workbook, pkg, 1);
    expect(log.frozen).toEqual({ rows: 4, cols: 2 });
    expect(log.merges).toHaveLength(2);
    const budget = await sheetOf(workbook, pkg, 2);
    expect(budget.rightToLeft).toBe(true);
    expect(budget.frozen).toEqual({ rows: 4, cols: 0 });
    expect(dashboard.rightToLeft).toBe(false);
  });

  test("snapshots conditional formatting from cached values", async () => {
    const { pkg, workbook } = await tracker();
    const log = await sheetOf(workbook, pkg, 1);
    const locked = log.rows.flatMap((row) => row.cells).find(({ cell }) => cell.formatted === "Locked");
    expect(locked?.cell.conditional).toMatchObject({ fill: "#DDEBE8", color: "#2E7D74", bold: true });
    const bar = log.rows.flatMap((row) => row.cells).find(({ cell }) => cell.conditional?.dataBar !== undefined);
    expect(bar?.cell.conditional?.dataBar).toMatchObject({ ratio: 0.325, color: "#2E7D74" });
  });

  test("resolves hyperlinks and keeps rich text concatenated", async () => {
    const { pkg, workbook } = await tracker();
    const notes = await sheetOf(workbook, pkg, 3);
    const guide = cellAt(notes, 19, 1);
    expect(guide.hyperlink).toEqual({ target: "https://example.com/nileframe/handbook", display: "example.com" });
    const rich = notes.rows.flatMap((row) => row.cells).find(({ cell }) => cell.formatted.includes("غيّر هذه الخلية فقط"));
    expect(rich?.cell.formatted).toContain("سعر ساعة المونتاج");
  });

  test("adapts embedded charts through the shared primitive vocabulary", async () => {
    const { pkg, workbook } = await tracker();
    const dashboard = await sheetOf(workbook, pkg, 0);
    expect(dashboard.drawings).toHaveLength(1);
    const chart = dashboard.drawings[0];
    if (chart?.kind !== "chart") throw new Error("fixture structure changed: expected a dashboard chart");
    expect(chart.sourceChartId).toBe("xl/charts/chart1.xml");
    expect(chart.graphic?.viewBox).toEqual({ width: 1000, height: 600 });
    expect(chart.graphic?.primitives.length).toBeGreaterThan(0);
    expect(chart.graphic?.primitives.some((primitive) => primitive.kind === "rect")).toBe(true);
    expect(chart.graphic?.primitives.some((primitive) => primitive.kind === "text" && primitive.text.includes("Hours logged by episode"))).toBe(true);
    expect(chart.bounds.width).toBeGreaterThan(0);
    expect(chart.bounds.height).toBeGreaterThan(0);
    expect(chart.unsupportedReason).toBeUndefined();
  });

  test("records non-goals visibly instead of dropping them", async () => {
    const { pkg, workbook } = await tracker();
    const log = await sheetOf(workbook, pkg, 1);
    expect(log.unsupported.map((entry) => entry.kind)).toContain("data-validation");
    expect(log.unsupported.map((entry) => entry.kind)).toContain("comment");
  });

  test("returns structured-cloneable sheet models", async () => {
    const { pkg, workbook } = await tracker();
    const budget = await sheetOf(workbook, pkg, 2);
    expect(structuredClone(budget)).toEqual(budget);
  });
});

describe("XLSX Arabic RTL coverage fixture", () => {
  // Coverage sheets resolve through the manifest by unique authored text,
  // never by sheet index — a rebuilt workbook that reorders sheets fails
  // loud in findCoverageSheet instead of silently testing the wrong cells.
  async function coverage(): Promise<SheetModel[]> {
    const pkg = await OpcPackage.open(await Bun.file(COVERAGE_PATH).arrayBuffer());
    const workbook = await readWorkbook(pkg);
    const sheets: SheetModel[] = [];
    for (let index = 0; index < workbook.sheets.length; index++) {
      sheets.push(await parseXlsxSheet(pkg, workbook, index));
    }
    return sheets;
  }

  test("every manifest selector resolves exactly one sheet", async () => {
    const sheets = await coverage();
    const manifest = await xlsxCoverageManifest();
    for (const id of Object.keys(manifest.sheets)) {
      const found = await findCoverageSheet(sheets, id);
      expect(found.sheet).toBe(sheets[found.index]!);
    }
    await expect(findCoverageSheet(sheets, "no-such-sheet")).rejects.toThrow('unknown sheet selector "no-such-sheet"');
  });

  test("parses the RTL sheet with merges, frozen panes, and hidden rows/cols", async () => {
    const sheet = (await findCoverageSheet(await coverage(), "rtl-grid")).sheet;
    expect(sheet.name).toBe("الميزانية");
    expect(sheet.rightToLeft).toBe(true);
    expect(sheet.extent).toEqual({ cols: 5, rows: 11 });
    expect(sheet.merges).toEqual([
      { start: { row: 0, col: 0 }, end: { row: 0, col: 4 } },
      { start: { row: 1, col: 0 }, end: { row: 1, col: 4 } },
    ]);
    expect(sheet.frozen).toEqual({ rows: 3, cols: 1 });
    expect(sheet.hiddenColumns).toEqual([3]);
    expect(sheet.hiddenRows).toEqual([6]);
    expect(sheet.rows.find((row) => row.index === 6)?.hidden).toBe(true);
    expect(structuredClone(sheet)).toEqual(sheet);
  });

  test("formats EGP currency, percents, and dates from cached values", async () => {
    const sheet = (await findCoverageSheet(await coverage(), "rtl-grid")).sheet;
    expect(cellAt(sheet, 0, 0).formatted).toBe("ميزانية الإنتاج الشهرية");
    expect(cellAt(sheet, 3, 1).formatted).toBe("180,000 EGP");
    expect(cellAt(sheet, 3, 3).formatted).toBe("-19.3%");
    expect(cellAt(sheet, 9, 1).formatted).toBe("13 Sep 2026");
    const total = cellAt(sheet, 8, 1);
    expect(total.value).toBe(430000);
    expect(total.formatted).toBe("430,000 EGP");
    expect(total.formula).toBe("SUM(B4:B8)");
    const variance = cellAt(sheet, 3, 3);
    expect(variance.conditional).toMatchObject({ fill: "#DDEBE8", color: "#2E7D74", bold: true });
    const overrun = cellAt(sheet, 7, 3);
    expect(overrun.formatted).toBe("24.4%");
    expect(overrun.conditional).toMatchObject({ fill: "#F6D9D2", color: "#B24830", bold: true });
    expect(cellAt(sheet, 10, 1).hyperlink).toEqual({ target: "https://example.com/nileframe/handbook", display: "example.com" });
  });

  test("keeps title and header styling with RTL reading order", async () => {
    const sheet = (await findCoverageSheet(await coverage(), "rtl-grid")).sheet;
    expect(cellAt(sheet, 0, 0).style).toMatchObject({ fill: "#12171F", bold: true, align: "center", readingOrder: 2 });
    expect(cellAt(sheet, 2, 0).style).toMatchObject({ fill: "#156082", bold: true, align: "center" });
    expect(cellAt(sheet, 2, 0).style?.border?.bottom?.style).toBe("thin");
    expect(cellAt(sheet, 8, 0).style).toMatchObject({ fill: "#F4F1EC", bold: true });
  });

  test("parses the LTR summary sheet without frozen columns", async () => {
    const sheet = (await findCoverageSheet(await coverage(), "summary")).sheet;
    expect(sheet.rightToLeft).toBe(false);
    expect(sheet.frozen).toEqual({ rows: 2, cols: 0 });
    expect(sheet.merges).toEqual([{ start: { row: 0, col: 0 }, end: { row: 0, col: 2 } }]);
    expect(cellAt(sheet, 4, 1).formatted).toBe("430,000 EGP");
  });

  test("keeps the deep shot log with live cost formulas", async () => {
    const sheet = (await findCoverageSheet(await coverage(), "deep-log")).sheet;
    expect(sheet.name).toBe("Shot Log");
    expect(sheet.extent).toEqual({ cols: 9, rows: 97 });
    expect(cellAt(sheet, 69, 0).formatted).toBe("S-041");
    // The Cost column still computes from the Notes rate assumption that
    // traveled with the sheet — never an external link, never a frozen value.
    const cost = cellAt(sheet, 4, 8);
    expect(cost.formula).toBe("$H5*Rate_per_hour");
    expect(cost.formatted).toBe("2,925 EGP");
    const notes = (await findCoverageSheet(await coverage(), "notes")).sheet;
    expect(cellAt(notes, 5, 1).formatted).toBe("450");
  });
});

describe("A1 notation helpers", () => {
  test("parses cell references and ranges", () => {
    expect(parseCellReference("A1")).toEqual({ row: 0, col: 0 });
    expect(parseCellReference("H34")).toEqual({ row: 33, col: 7 });
    expect(parseCellRange("A1:H1")).toEqual({ start: { row: 0, col: 0 }, end: { row: 0, col: 7 } });
    expect(() => parseCellReference("nope")).toThrow(/invalid cell reference/);
  });
});
