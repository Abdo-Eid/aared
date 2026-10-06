import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexXlsx, readSharedStrings, readWorkbook } from "../src/core/xlsx/workbook.js";
import { columnCharsToInches, parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { layoutChart } from "../src/core/chart/draw.js";
import { applySpreadsheetTint } from "../src/core/theme.js";
import { findCoverageSheet, xlsxCoverageManifest } from "./xlsx-fixture.js";
import type { SheetModel } from "../src/core/model.js";
import { withPart } from "./mutate.js";

// XLSX parsing against real workbooks: the long-lived production tracker and
// the Arabic RTL coverage fixture (variants are derived in memory by `withPart`).

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

describe("XLSX column widths", () => {
  // Read back from Excel (COM, `Columns(n).Width` in pixels) on copies of the
  // fixtures: the stored width already holds the cell padding.
  test("a stored width becomes pixels as Excel draws it, with no padding added again", () => {
    const pixels = (width: number) => Math.round(columnCharsToInches(width) * 96);
    expect(pixels(26)).toBe(182);
    expect(pixels(16)).toBe(112);
    expect(pixels(8.4258)).toBe(59);
    expect(pixels(9.140625)).toBe(64);
  });

  test("every column of a real workbook is a whole number of pixels", async () => {
    const { pkg, workbook } = await tracker();
    for (let index = 0; index < workbook.sheets.length; index++) {
      const sheet = await sheetOf(workbook, pkg, index);
      for (const inches of sheet.columnWidths) expect(Math.abs(inches * 96 - Math.round(inches * 96))).toBeLessThan(1e-9);
    }
  });
});

describe("XLSX theme tint", () => {
  // The first three are Excel's own `.Color` for a theme colour with a tint
  // (read back over COM); the last three are the Office 2013 palette's
  // "Lighter 80/60/40%" of the same accent.
  test("moves lightness as Excel does", () => {
    expect(applySpreadsheetTint("#5B9BD5", -0.25)).toBe("#2E75B6");
    expect(applySpreadsheetTint("#5B9BD5", -0.5)).toBe("#1F4E79");
    expect(applySpreadsheetTint("#ED7D31", -0.25)).toBe("#C55A11");
    expect(applySpreadsheetTint("#5B9BD5", 0.8)).toBe("#DEEBF7");
    expect(applySpreadsheetTint("#5B9BD5", 0.6)).toBe("#BDD7EE");
    expect(applySpreadsheetTint("#5B9BD5", 0.4)).toBe("#9DC3E6");
    expect(applySpreadsheetTint("#5B9BD5", 0)).toBe("#5B9BD5");
  });
});

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

  test("reads embedded charts with the PPTX chart reader, laid out as Excel draws them", async () => {
    const { pkg, workbook } = await tracker();
    const dashboard = await sheetOf(workbook, pkg, 0);
    expect(dashboard.drawings).toHaveLength(1);
    const chart = dashboard.drawings[0];
    if (chart?.kind !== "chart") throw new Error("fixture structure changed: expected a dashboard chart");
    expect(chart.sourceChartId).toBe("xl/charts/chart1.xml");
    expect(chart.unsupportedReason).toBeUndefined();
    expect(chart.spec!.groups[0]!.series[0]!.values).toEqual([39, 27, 8]);
    expect(chart.bounds.width).toBeGreaterThan(0);
    expect(chart.bounds.height).toBeGreaterThan(0);
    // Excel draws this 5.1 x 2.95in chart with an axis of 0.0-45.0 by 5.
    const texts = layoutChart(chart.spec!, chart.bounds).primitives.flatMap((p) => (p.kind === "text" ? [p.text] : []));
    expect(texts).toEqual(expect.arrayContaining(["Hours logged by episode", "Hours", "Episode 4", "0.0", "5.0", "45.0"]));
    expect(texts).not.toContain("50.0");
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

describe("XLSX worksheet structure the file states loosely", () => {
  const SHOT_LOG = "xl/worksheets/sheet3.xml";
  const shotLog = async (rewrite: (xml: string) => string): Promise<SheetModel> => {
    const pkg = await withPart("xlsx-coverage.xlsx", SHOT_LOG, rewrite);
    return parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
  };

  test("a stale <dimension> does not cut away cells the sheet really holds", async () => {
    const sheet = await shotLog((xml) => xml.replace(/<dimension[^>]*\/>/, '<dimension ref="A1"/>'));
    expect(sheet.extent).toEqual({ cols: 9, rows: 97 });
    expect(cellAt(sheet, 4, 8).formatted).toBe("2,925 EGP");
  });

  test("a merge wider than the cells still counts when the <dimension> is stale", async () => {
    const stale = await shotLog((xml) => xml
      .replace(/<dimension[^>]*\/>/, '<dimension ref="A1"/>')
      .replace("</sheetData>", '</sheetData><mergeCells count="1"><mergeCell ref="A1:L1"/></mergeCells>'));
    // Sheet 3 holds nine columns of cells; the merge reaches column L, and is not clipped to the dimension or the cells.
    expect(stale.extent.cols).toBe(12);
    expect(stale.merges).toContainEqual({ start: { row: 0, col: 0 }, end: { row: 0, col: 11 } });
  });

  test("a row with its own height but no cells keeps that height", async () => {
    const baseline = await shotLog((xml) => xml);
    const present = new Set(baseline.rows.map((row) => row.index));
    const missing = [...Array(baseline.extent.rows).keys()].find((index) => !present.has(index))!;
    const sheet = await shotLog((xml) => xml.replace("</sheetData>", `<row r="${missing + 1}" ht="45" customHeight="1"/></sheetData>`));
    expect(sheet.rows.find((row) => row.index === missing)).toEqual({ index: missing, cells: [], height: 45 / 72 });
    // A cell-less row at the default height says nothing and adds no row.
    const plain = await shotLog((xml) => xml.replace("</sheetData>", `<row r="${missing + 1}" ht="15"/></sheetData>`));
    expect(plain.rows.some((row) => row.index === missing)).toBe(false);
  });

  test("a sparkline extension is reported as not drawn", async () => {
    const uri = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
    // The sheet's own worksheet-level <extLst> (conditional-format extensions) is the last one.
    const sheet = await shotLog((xml) => `${xml.slice(0, xml.lastIndexOf("<extLst>"))}` +
      `<extLst><ext uri="${uri}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups><x14:sparklineGroup/></x14:sparklineGroups></ext></extLst></worksheet>`);
    expect(sheet.unsupported.some((entry) => entry.kind === "sparkline")).toBe(true);
    const plain = await shotLog((xml) => xml);
    expect(plain.unsupported.some((entry) => entry.kind === "sparkline")).toBe(false);
  });

  test("an 8-digit rgb is read without its alpha byte, as the worksheet's own colours are", async () => {
    const pkg = await withPart("xlsx-coverage.xlsx", "xl/styles.xml", (xml) => xml.replace(/(<fgColor rgb=")FF/, "$100"));
    const workbook = await readWorkbook(pkg);
    const fills = [...Array(40).keys()].map((index) => workbook.styles.cellStyle(index).fill).filter((fill): fill is string => fill !== undefined);
    expect(fills.length).toBeGreaterThan(0);
    expect(fills.every((fill) => /^#[0-9A-F]{6}$/.test(fill))).toBe(true);
  });
});

describe("XLSX drawings with no id", () => {
  test("a chart with no xdr:cNvPr/@id reads exactly as one with it: XLSX never uses a drawing's id", async () => {
    const { pkg, workbook } = await tracker();
    const original = await sheetOf(workbook, pkg, 0);
    expect(original.drawings.some((drawing) => drawing.kind === "chart")).toBe(true);
    const stripped = await withPart("post-production-tracker.xlsx", "xl/drawings/drawing1.xml", (xml) => {
      expect(xml).toContain("<xdr:cNvPr id=");
      return xml.replace(/(<xdr:cNvPr) id="[^"]*"/g, "$1");
    });
    const sheet = await parseXlsxSheet(stripped, await readWorkbook(stripped), 0);
    // Only the count differs (it drives the renderer's `xlsx-missing-id` warning); a sheet with every id has none.
    const { missingIds, ...rest } = sheet;
    expect(missingIds).toBe(1);
    expect(original.missingIds).toBeUndefined();
    expect(JSON.stringify(rest)).toBe(JSON.stringify(original));
  });
});

describe("XLSX drawings past the used range", () => {
  test("a chart below and beside the data keeps its full frame, and the grid grows to hold it", async () => {
    const { pkg, workbook } = await tracker();
    const original = await sheetOf(workbook, pkg, 0);
    const chart = original.drawings.find((drawing) => drawing.kind === "chart");
    if (chart === undefined) throw new Error("fixture structure changed: the Dashboard lost its chart");
    // The same workbook with the Dashboard's used range cut to A1:B3: the chart (A19 to E33) now lies wholly past it.
    const cut = await withPart("post-production-tracker.xlsx", "xl/worksheets/sheet1.xml", (xml) => xml.replace('<dimension ref="A1:H34"/>', '<dimension ref="A1:B3"/>'));
    const sheet = await parseXlsxSheet(cut, await readWorkbook(cut), 0);
    const moved = sheet.drawings.find((drawing) => drawing.kind === "chart")!;
    expect(moved.bounds.x).toBeCloseTo(chart.bounds.x, 6);
    expect(moved.bounds.y).toBeCloseTo(chart.bounds.y, 6);
    expect(moved.bounds.width).toBeCloseTo(chart.bounds.width, 6);
    expect(moved.bounds.height).toBeCloseTo(chart.bounds.height, 6);
    expect(chart.bounds.width).toBeGreaterThan(3);
    // The grid reaches the chart's last cell (row 33, column E).
    expect(sheet.extent.rows).toBeGreaterThanOrEqual(33);
    expect(sheet.extent.cols).toBeGreaterThanOrEqual(5);
    expect(sheet.columnWidths).toHaveLength(sheet.extent.cols);
  });
});
