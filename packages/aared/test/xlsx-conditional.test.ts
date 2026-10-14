import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import type { CellModel } from "../src/core/model.js";
import { compileCriteria, criteriaMatches } from "../src/core/xlsx/formula.js";
import type { FormulaScalar } from "../src/core/xlsx/formula.js";
import { parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { readWorkbook } from "../src/core/xlsx/workbook.js";
import { withPart } from "./mutate.js";
import { conditionalFormulaFixture } from "./xlsx-conditional-fixture.js";
import { storedZip } from "./zip.js";
import { open as openXlsx, readSheetModel } from "../src/xlsx/renderer.js";

// Conditional formats on the Shot Log sheet of the Arabic coverage workbook:
// its Excel-authored rules are replaced by rules in the shapes Excel writes
// (ECMA-376 18.3.1.8), over the Hours column (F5:F97, 93 numbers). Rule
// bodies are written here, not by Excel, so they pin additional rule shapes.
// The CF Rules sheet below supplies native-authored icon, date, stdDev and
// three-colour-scale cases, also covered by the paired A4 capture suite.

const SHOT_LOG = "xl/worksheets/sheet3.xml";
const HOURS = 5;

async function shotLog(rules: string): Promise<{ cells: CellModel[]; unsupported: string[] }> {
  const pkg: OpcPackage = await withPart("xlsx-coverage.xlsx", SHOT_LOG, (xml) =>
    xml.replace(/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/, "").replace("</sheetData>", `</sheetData>${rules}`),
  );
  const sheet = await parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
  const cells = sheet.rows.flatMap((row) => row.cells.filter((entry) => entry.col === HOURS && typeof entry.cell.value === "number").map((entry) => entry.cell));
  return { cells, unsupported: sheet.unsupported.map((entry) => entry.detail) };
}

// The synthetic rules write dxfId="0" for the fixture's red fill and
// dxfId="1" for its tan fill. Excel puts the Summary sheet's own differential
// styles first whenever it saves this workbook, so those two move with every
// Summary rule: they are found by colour, never by position.
const fixtureDxfs = [...(await (await OpcPackage.open(await Bun.file(`${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`).arrayBuffer()))
  .readText("xl/styles.xml")).matchAll(/<dxf>[\s\S]*?<\/dxf>/g)].map((match) => match[0]);
const dxfWithFill = (argb: string) => {
  const index = fixtureDxfs.findIndex((dxf) => dxf.includes(`<bgColor rgb="${argb}"/>`));
  if (index < 0) throw new Error(`xlsx-coverage.xlsx has no differential style filled ${argb}`);
  return index;
};
const SYNTHETIC_DXF = [dxfWithFill("FFFF0000"), dxfWithFill("FFFFDCB4")];
const fixtureDxfIds = (xml: string) => xml.replace(/dxfId="([01])"/g, (_match, id: string) => `dxfId="${SYNTHETIC_DXF[Number(id)]}"`);
const rule = (body: string, sqref = "F5:F97") => `<conditionalFormatting sqref="${sqref}">${fixtureDxfIds(body)}</conditionalFormatting>`;
const numbers = (cells: CellModel[]) => cells.map((cell) => cell.value as number).sort((a, b) => a - b);

/** Status column (E5:E97) with the base rules replaced: text, blanks and all. */
const statusCells = async (body: string) => {
  const pkg: OpcPackage = await withPart("xlsx-coverage.xlsx", SHOT_LOG, (xml) =>
    xml.replace(/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/, "").replace("</sheetData>", `</sheetData>${body}`),
  );
  const sheet = await parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
  return {
    sheet,
    cells: sheet.rows.flatMap((row) => row.cells.filter((entry) => entry.col === 4 && row.index >= 4).map((entry) => entry.cell)),
    unsupported: sheet.unsupported.map((entry) => entry.detail),
  };
};

describe("XLSX conditional formats", () => {
  test("a three-colour scale blends through its middle stop (Excel's red-yellow-green)", async () => {
    const { cells } = await shotLog(
      rule(`<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>` +
        `<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule>`),
    );
    const sorted = numbers(cells);
    const fillOf = (value: number) => cells.find((cell) => cell.value === value)!.conditional?.fill;
    expect(fillOf(sorted[0]!)).toBe("#F8696B");
    expect(fillOf(sorted[sorted.length - 1]!)).toBe("#63BE7B");
    // The 50th percentile of an odd count of values is the middle value itself, so
    // it takes the middle colour exactly. Derived from the read, so a fixture
    // rebuild that adds rows changes the index rather than the expectation.
    const middle = Math.floor((sorted.length - 1) / 2);
    expect(fillOf(sorted[middle]!)).toBe("#FFEB84");
    expect(new Set(cells.map((cell) => cell.conditional?.fill)).size).toBeGreaterThan(3);
  });

  test("percent bounds are shares of the min-to-max span, and values past them keep the end colour", async () => {
    const { cells } = await shotLog(
      rule(`<cfRule type="colorScale" priority="1"><colorScale><cfvo type="percent" val="10"/><cfvo type="percent" val="90"/>` +
        `<color rgb="FF000000"/><color rgb="FFFFFFFF"/></colorScale></cfRule>`),
    );
    const sorted = numbers(cells);
    const min = sorted[0]!;
    const max = sorted[sorted.length - 1]!;
    for (const cell of cells) {
      const share = ((cell.value as number) - min) / (max - min);
      if (share <= 0.1) expect(cell.conditional?.fill).toBe("#000000");
      if (share >= 0.9) expect(cell.conditional?.fill).toBe("#FFFFFF");
    }
    expect(cells.some((cell) => cell.conditional?.fill === "#000000")).toBe(true);
    expect(cells.some((cell) => cell.conditional?.fill === "#FFFFFF")).toBe(true);
  });

  test("stopIfTrue keeps lower-priority rules off the cells it matched", async () => {
    const both = (stop: string) =>
      rule(
        `<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"${stop}><formula>0</formula></cfRule>` +
          `<cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF2E7D74"/></dataBar></cfRule>`,
      );
    const open = await shotLog(both(""));
    const stopped = await shotLog(both(` stopIfTrue="1"`));
    const positive = (cells: CellModel[]) => cells.filter((cell) => (cell.value as number) > 0);
    expect(positive(open.cells).length).toBeGreaterThan(0);
    expect(positive(open.cells).every((cell) => cell.conditional?.fill !== undefined && cell.conditional.dataBar !== undefined)).toBe(true);
    expect(positive(stopped.cells).every((cell) => cell.conditional?.fill !== undefined && cell.conditional.dataBar === undefined)).toBe(true);
    // Cells the first rule did not match still reach the second.
    const rest = stopped.cells.filter((cell) => !((cell.value as number) > 0));
    expect(rest.every((cell) => cell.conditional?.dataBar !== undefined)).toBe(true);
  });

  test("Duplicate Values and Unique Values split the cells by how often their value repeats", async () => {
    const duplicate = await shotLog(rule(`<cfRule type="duplicateValues" dxfId="1" priority="1"/>`));
    const unique = await shotLog(rule(`<cfRule type="uniqueValues" dxfId="1" priority="1"/>`));
    const counts = new Map<number, number>();
    for (const cell of duplicate.cells) counts.set(cell.value as number, (counts.get(cell.value as number) ?? 0) + 1);
    const repeated = duplicate.cells.filter((cell) => counts.get(cell.value as number)! > 1);
    const single = duplicate.cells.filter((cell) => counts.get(cell.value as number) === 1);
    expect(repeated.length).toBeGreaterThan(0);
    expect(single.length).toBeGreaterThan(0);
    const painted = (cells: CellModel[]) => cells.filter((cell) => cell.conditional?.fill !== undefined).map((cell) => cell.value);
    expect(painted(duplicate.cells)).toEqual(repeated.map((cell) => cell.value));
    expect(painted(unique.cells)).toEqual(single.map((cell) => cell.value));
  });

  test("the text rules read the cfRule's own text and ignore case, as Excel's SEARCH does", async () => {
    const status = async (type: string, text: string, operator = type) => {
      const pkg: OpcPackage = await withPart("xlsx-coverage.xlsx", SHOT_LOG, (xml) =>
        xml.replace(/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/, "").replace("</sheetData>",
          `</sheetData>${rule(`<cfRule type="${type}" operator="${operator}" text="${text}" dxfId="1" priority="1"><formula>NOT(ISERROR(SEARCH("${text}",E5)))</formula></cfRule>`, "E5:E97")}`),
      );
      const sheet = await parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
      // Column E under the header row (the rules cover E5:E97).
      const cells = sheet.rows.filter((row) => row.index >= 4).flatMap((row) => row.cells.filter((entry) => entry.col === 4).map((entry) => entry.cell));
      const painted = new Set(cells.filter((cell) => cell.conditional?.fill !== undefined).map((cell) => cell.formatted));
      return { painted, all: new Set(cells.map((cell) => cell.formatted)) };
    };
    const contains = await status("containsText", "REVIEW");
    expect([...contains.painted]).toEqual(["In review"]);
    expect([...(await status("beginsWith", "in ")).painted].sort()).toEqual(["In progress", "In review"]);
    expect([...(await status("endsWith", "ed")).painted].sort()).toEqual(["Blocked", "Locked", "Not started"]);
    const not = await status("notContainsText", "review", "notContains");
    expect([...not.painted].sort()).toEqual([...contains.all].filter((text) => text !== "In review").sort());
  });

  test("an expression rule evaluates relative references per cell", async () => {
    const { cells, unsupported } = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>F5&gt;10</formula></cfRule>`));
    expect(unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    for (const cell of cells) {
      const shouldPaint = (cell.value as number) > 10;
      expect((cell.conditional?.fill !== undefined)).toBe(shouldPaint);
    }
    expect(cells.some((cell) => cell.conditional?.fill !== undefined)).toBe(true);
  });

  test("an unsupported expression is reported, not guessed", async () => {
    const { cells, unsupported } = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>NOSUCHFUNC(F5)&gt;10</formula></cfRule>`));
    expect(cells.every((cell) => cell.conditional === undefined)).toBe(true);
    expect(unsupported.some((detail) => detail.includes("expression"))).toBe(true);
  });

  test("expression AND with mixed references paints per-row truth", async () => {
    // Hours >5 AND Status "In review"? Use Hours (F) and Status (E) on Shot Log: E5:E97 hold status text.
    const pkg: OpcPackage = await withPart("xlsx-coverage.xlsx", SHOT_LOG, (xml) =>
      xml.replace(/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/, "").replace("</sheetData>",
        `</sheetData>${rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>AND($F5&gt;5,E5="In review")</formula></cfRule>`)}`),
    );
    const sheet = await parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
    const rows = sheet.rows.filter((row) => row.index >= 4);
    let painted = 0;
    for (const row of rows) {
      const hours = row.cells.find((entry) => entry.col === HOURS)?.cell;
      const status = row.cells.find((entry) => entry.col === 4)?.cell;
      if (hours === undefined || status === undefined) continue;
      const should = typeof hours.value === "number" && hours.value > 5 && status.formatted === "In review";
      const did = hours.conditional?.fill !== undefined || status.conditional?.fill !== undefined;
      // Expression applies to whole row range? sqref is F5:F97, so only Hours cells paint, but condition reads E.
      const hoursPainted = hours.conditional?.fill !== undefined;
      expect(hoursPainted).toBe(should);
      if (hoursPainted) painted++;
    }
    expect(painted).toBeGreaterThan(0);
    expect(sheet.unsupported.some((entry) => entry.detail.includes("expression"))).toBe(false);
  });

  test("cellIs with a relative operand shifts per row", async () => {
    // Each Hours cell (F) compares against column H in the same row ("Hours
    // logged"), with the operand written as $H5 so it shifts down the range. A
    // fixed subset would paint, but not these cells.
    const pkg: OpcPackage = await withPart("xlsx-coverage.xlsx", SHOT_LOG, (xml) =>
      xml.replace(/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/, "").replace("</sheetData>",
        `</sheetData>${rule(`<cfRule type="cellIs" dxfId="1" priority="1" operator="greaterThan"><formula>$H5</formula></cfRule>`)}`),
    );
    const sheet = await parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
    let checked = 0;
    for (const row of sheet.rows.filter((entry) => entry.index >= 4)) {
      const hours = row.cells.find((entry) => entry.col === HOURS)?.cell;
      const cost = row.cells.find((entry) => entry.col === 7)?.cell;
      if (typeof hours?.value !== "number" || typeof cost?.value !== "number") continue;
      checked++;
      expect(hours.conditional?.fill !== undefined, `row ${row.index + 1}`).toBe(hours.value > cost.value);
    }
    expect(checked).toBeGreaterThan(0);
    expect(sheet.unsupported.some((entry) => entry.detail.includes("cellIs"))).toBe(false);
  });

  test("ISBLANK paints no authored number, and adds no row the file did not have", async () => {
    // F5:F10 are all authored numbers in the fixture, so ISBLANK is false for
    // every one of them and nothing paints.
    const before = (await parseXlsxSheet(await withPart("xlsx-coverage.xlsx", SHOT_LOG, (xml) => xml), await readWorkbook(await withPart("xlsx-coverage.xlsx", SHOT_LOG, (xml) => xml)), 2)).rows.length;
    const pkg: OpcPackage = await withPart("xlsx-coverage.xlsx", SHOT_LOG, (xml) =>
      xml.replace(/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/, "").replace("</sheetData>",
        `</sheetData>${rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>ISBLANK(F5)</formula></cfRule>`, "F5:F10")}`),
    );
    const sheet = await parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
    const cells = sheet.rows.flatMap((row) => row.cells.filter((entry) => entry.col === HOURS && row.index >= 4 && row.index <= 9).map((entry) => entry.cell));
    expect(cells.length).toBeGreaterThan(0);
    expect(cells.every((cell) => cell.conditional === undefined)).toBe(true);
    // A blank-sensitive rule must not invent rows for a range the sheet has no
    // cells in: the row count is the file's, unchanged.
    expect(sheet.rows.length).toBe(before);
  });

  test("ISBLANK materializes only matching missing cells, without inventing values or unrelated rows", async () => {
    const parse = async (formula: string, sqref = "F5:F7") => {
      const pkg = await conditionalFormulaFixture({ formula, sqref, blankCells: ["F5", "F7"] });
      return parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
    };
    const control = await parse("FALSE");
    const matched = await parse("ISBLANK(F5)");
    for (const row of [4, 6]) {
      expect(control.rows.find((entry) => entry.index === row)?.cells.some((entry) => entry.col === 5)).toBe(false);
      const blank = matched.rows.find((entry) => entry.index === row)?.cells.find((entry) => entry.col === 5)?.cell;
      expect(blank).toMatchObject({ value: null, formatted: "", type: "empty", conditional: { fill: "#FFDCB4" } });
    }
    expect(matched.rows.map((row) => row.index)).toEqual(control.rows.map((row) => row.index));
    // Every pre-existing cell, including numeric F6 and the cells outside the
    // rule's range, keeps its value/style; only the two recipients are added.
    for (const row of control.rows) {
      for (const entry of row.cells) expect(matched.rows.find((candidate) => candidate.index === row.index)?.cells.find((candidate) => candidate.col === entry.col)?.cell).toEqual(entry.cell);
    }
    expect(matched.unsupported).toEqual(control.unsupported);
    expect(control.rows.some((row) => row.index === 99)).toBe(false);
    const newRow = await parse("ISBLANK(K100)", "K100");
    const beforeRows = new Set(control.rows.map((row) => row.index));
    expect(newRow.rows.filter((row) => !beforeRows.has(row.index))).toEqual([{
      index: 99,
      cells: [{ col: 10, cell: { value: null, formatted: "", type: "empty", conditional: { fill: "#FFDCB4" } } }],
    }]);
    // 50,001 recipients exceeds the cap: missing blanks stay absent rather
    // than allocating thousands of cells or extending the sheet's extent.
    const bounded = await parse("ISBLANK(F5)", "F5:F50005");
    expect(bounded.rows).toEqual(control.rows);
    expect(bounded.extent).toEqual(control.extent);
    expect(bounded.unsupported).toEqual(control.unsupported);
  });

  test("IFS shape errors are Excel values: unpainted normally and matchable with ISERROR", async () => {
    for (const name of ["SUMIFS", "AVERAGEIFS"]) {
      const parse = async (formula: string) => {
        const pkg = await conditionalFormulaFixture({ formula, sqref: "F5:F7" });
        const sheet = await parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
        return { sheet, recipients: sheet.rows.filter((row) => row.index >= 4 && row.index <= 6).flatMap((row) => row.cells.filter((cell) => cell.col === 5).map((entry) => entry.cell)) };
      };
      const formula = `${name}($F$5:$F$7,$E$5:$E$6,"Locked")`;
      const wrong = await parse(`${formula}>0`);
      expect(wrong.recipients).toHaveLength(3);
      expect(wrong.recipients.every((cell) => cell.conditional === undefined)).toBe(true);
      expect(wrong.sheet.unsupported.filter((entry) => entry.kind === "conditional-format")).toEqual([]);
      const errors = await parse(`ISERROR(${formula})`);
      expect(errors.recipients.map((cell) => cell.conditional?.fill)).toEqual(["#FFDCB4", "#FFDCB4", "#FFDCB4"]);
      const valid = await parse(`${name}($F$5:$F$7,$E$5:$E$7,"Locked")>0`);
      expect(valid.recipients.map((cell) => cell.conditional?.fill)).toEqual(["#FFDCB4", "#FFDCB4", "#FFDCB4"]);
    }
  });

  test("IFS shape errors survive the shipped XLSX Worker route", async () => {
    for (const name of ["SUMIFS", "AVERAGEIFS"]) {
      const pkg = await conditionalFormulaFixture({ formula: `ISERROR(${name}($F$5:$F$7,$E$5:$E$6,"Locked"))`, sqref: "F5:F7" });
      const parts: Record<string, Uint8Array> = {};
      for (const part of pkg.partNames()) parts[part] = await pkg.readPart(part);
      const workbook = await openXlsx(storedZip(parts).slice().buffer as ArrayBuffer, {
        worker: { url: new URL("../dist/core/worker.mjs", import.meta.url) },
      });
      try {
        const window = await readSheetModel(workbook.sheets.get(2), { range: "F5:F7" });
        expect(window.rows.flatMap((row) => row.cells.map((entry) => entry.cell.conditional?.fill))).toEqual(["#FFDCB4", "#FFDCB4", "#FFDCB4"]);
      } finally { workbook.close(); }
    }
  });

  test("cross-sheet and missing-cache expressions report instead of guessing", async () => {
    const cross = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>Other!$A$1&gt;50</formula></cfRule>`));
    expect(cross.cells.every((cell) => cell.conditional === undefined)).toBe(true);
    expect(cross.unsupported.some((detail) => detail.includes("expression"))).toBe(true);
  });

  test("TODAY and NOW paint as a parse-time snapshot", async () => {
    // Hours are all far below the current serial (~46300), so `<` paints all and `>` none.
    const past = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>F5&lt;TODAY()</formula></cfRule>`));
    expect(past.unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    expect(past.cells.length).toBeGreaterThan(0);
    expect(past.cells.every((cell) => cell.conditional?.fill !== undefined)).toBe(true);
    const future = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>F5&gt;NOW()</formula></cfRule>`));
    expect(future.unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    expect(future.cells.every((cell) => cell.conditional === undefined)).toBe(true);
    const year = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>YEAR(TODAY())&gt;2000</formula></cfRule>`));
    expect(year.cells.every((cell) => cell.conditional?.fill !== undefined)).toBe(true);
  });

  test("DATE rolls overflow, keeps the 1900 leap day, and fails closed below serial 0", async () => {
    // Month 13 rolls into next January (measured `DATE(2026,13,1)` is 46388).
    const rolled = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>F5&lt;DATE(2026,13,1)</formula></cfRule>`));
    expect(rolled.unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    expect(rolled.cells.every((cell) => cell.conditional?.fill !== undefined)).toBe(true);
    // The fictitious 1900-02-29 is serial 60: every Hour sits below it.
    const leap = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>F5&gt;DATE(1900,2,29)</formula></cfRule>`));
    for (const cell of leap.cells) {
      expect(cell.conditional?.fill !== undefined).toBe((cell.value as number) > 60);
    }
    // A negative result is Excel's `#NUM!`: unpainted, not reported.
    const negative = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>DATE(-5,1,1)&gt;0</formula></cfRule>`));
    expect(negative.cells.every((cell) => cell.conditional === undefined)).toBe(true);
    expect(negative.unsupported.some((detail) => detail.includes("outside the supported subset"))).toBe(false);
  });

  test("YEAR MONTH DAY WEEKDAY decompose serials like Excel", async () => {
    const ymd = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>AND(YEAR(DATE(2026,10,11))=2026,MONTH(DATE(2026,10,11))=10,DAY(DATE(2026,10,11))=11)</formula></cfRule>`));
    expect(ymd.unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    expect(ymd.cells.every((cell) => cell.conditional?.fill !== undefined)).toBe(true);
    // 2026-10-11 is a Sunday: 1 under the default Sunday-first shape.
    const sunday = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>WEEKDAY(DATE(2026,10,11))=1</formula></cfRule>`));
    expect(sunday.cells.every((cell) => cell.conditional?.fill !== undefined)).toBe(true);
    // An unknown return type is `#NUM!`: unpainted, not reported.
    const badKind = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>WEEKDAY(DATE(2026,10,11),99)&gt;0</formula></cfRule>`));
    expect(badKind.cells.every((cell) => cell.conditional === undefined)).toBe(true);
    expect(badKind.unsupported.some((detail) => detail.includes("outside the supported subset"))).toBe(false);
  });

  test("SEARCH is case-insensitive with wildcards; FIND is literal and case-sensitive", async () => {
    const search = await statusCells(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>ISNUMBER(SEARCH("REVIEW",E5))</formula></cfRule>`, "E5:E97"));
    expect(search.unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    let painted = 0;
    for (const cell of search.cells) {
      const should = typeof cell.value === "string" && cell.value.toLowerCase().includes("review");
      expect(cell.conditional?.fill !== undefined).toBe(should);
      if (should) painted++;
    }
    expect(painted).toBeGreaterThan(0);
    // The statuses are lowercase "In review": a case-sensitive FIND for capitals misses everything.
    const find = await statusCells(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>ISNUMBER(FIND("REVIEW",E5))</formula></cfRule>`, "E5:E97"));
    expect(find.unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    expect(find.cells.every((cell) => cell.conditional === undefined)).toBe(true);
    // A wildcard still matches through SEARCH.
    const wild = await statusCells(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>ISNUMBER(SEARCH("In*iew",E5))</formula></cfRule>`, "E5:E97"));
    for (const cell of wild.cells) {
      const should = typeof cell.value === "string" && cell.value.toLowerCase().includes("review");
      expect(cell.conditional?.fill !== undefined).toBe(should);
    }
  });

  test("MID is 1-based with Excel edge behaviour", async () => {
    const mid = await statusCells(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>MID(E5,1,2)="In"</formula></cfRule>`, "E5:E97"));
    expect(mid.unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    let painted = 0;
    for (const cell of mid.cells) {
      const should = typeof cell.value === "string" && cell.value.startsWith("In");
      expect(cell.conditional?.fill !== undefined).toBe(should);
      if (should) painted++;
    }
    // "In review" and "In progress" both qualify.
    expect(painted).toBeGreaterThan(1);
  });

  test("CONCAT joins ranges; CONCATENATE with a range reports", async () => {
    const concat = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>LEFT(CONCAT("n=",F5),2)="n="</formula></cfRule>`));
    expect(concat.unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    expect(concat.cells.every((cell) => cell.conditional?.fill !== undefined)).toBe(true);
    const legacy = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>CONCATENATE(F5:F6)="x"</formula></cfRule>`));
    expect(legacy.cells.every((cell) => cell.conditional === undefined)).toBe(true);
    expect(legacy.unsupported.some((detail) => detail.includes("expression"))).toBe(true);
  });

  test("COUNTIF criteria match Excel's equality, blanks and wildcards", () => {
    const match = (criteria: FormulaScalar, cell: FormulaScalar): boolean | undefined => {
      const compiled = compileCriteria(criteria);
      expect(compiled).not.toBeUndefined();
      return criteriaMatches(cell, compiled!);
    };
    // Numeric criteria match numbers and numeric text, never blanks or booleans.
    expect(match(5, 5)).toBe(true);
    expect(match(5, "5")).toBe(true);
    expect(match("5", 5)).toBe(true);
    expect(match(5, "6")).toBe(false);
    expect(match(5, null)).toBe(false);
    expect(match(5, true)).toBe(false);
    // Text criteria are case-insensitive; `*` spans.
    expect(match("apple", "Apple")).toBe(true);
    expect(match("apple", "apples")).toBe(false);
    expect(match("A*", "apple")).toBe(true);
    expect(match("A*", "A*")).toBe(true);
    expect(match("A*", "banana")).toBe(false);
    expect(match("apple", 5)).toBe(false);
    // `""` takes blanks and zero-length strings, `"="` only blanks, `"<>"` everything but blanks.
    expect(match("", null)).toBe(true);
    expect(match("", "")).toBe(true);
    expect(match("", "x")).toBe(false);
    expect(match("=", null)).toBe(true);
    expect(match("=", "")).toBe(false);
    expect(match("<>", "x")).toBe(true);
    expect(match("<>", 5)).toBe(true);
    expect(match("<>", null)).toBe(false);
    expect(match("<>5", null)).toBe(true);
    expect(match("<>5", 5)).toBe(false);
    // Relational numbers ignore text; blanks read as 0 under ordering but sit out `=0`.
    expect(match(">=5", 5)).toBe(true);
    expect(match(">=5", 4)).toBe(false);
    expect(match(">=5", "7")).toBe(false);
    expect(match(">=5", null)).toBe(false);
    expect(match("<0", null)).toBe(false);
    expect(match(0, null)).toBe(false);
    // Booleans and errors compare within their kind.
    expect(match(true, true)).toBe(true);
    expect(match("TRUE", true)).toBe(true);
    expect(match(true, "TRUE")).toBe(true);
    expect(match(true, false)).toBe(false);
    expect(match("#DIV/0!", { error: "#DIV/0!" })).toBe(true);
    expect(match("#DIV/0!", 5)).toBe(false);
    expect(match({ error: "#DIV/0!" }, { error: "#DIV/0!" })).toBe(true);
    // A blank criteria matches nothing at all.
    expect(match(null, null)).toBe(false);
    expect(match(null, 0)).toBe(false);
  });

  test("COUNTIF duplicates and operators paint per cell", async () => {
    const dupes = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>COUNTIF($F$5:$F$97,F5)&gt;1</formula></cfRule>`));
    expect(dupes.unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    const counts = new Map<number, number>();
    for (const cell of dupes.cells) counts.set(cell.value as number, (counts.get(cell.value as number) ?? 0) + 1);
    let painted = 0;
    for (const cell of dupes.cells) {
      const should = counts.get(cell.value as number)! > 1;
      expect(cell.conditional?.fill !== undefined).toBe(should);
      if (should) painted++;
    }
    expect(painted).toBeGreaterThan(0);
    // Every Hour matches itself under `>=`, so the self-comparison paints all.
    const self = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>COUNTIF($F$5:$F$97,"&gt;="&amp;F5)&gt;=1</formula></cfRule>`));
    expect(self.cells.every((cell) => cell.conditional?.fill !== undefined)).toBe(true);
  });

  test("COUNTIFS ANDs pairs and stays silent on mismatched sizes", async () => {
    // A blank criteria matches nothing (measured), so blank-status rows stay
    // cold while every other row matches itself on both pairs.
    const pkg: OpcPackage = await withPart("xlsx-coverage.xlsx", SHOT_LOG, (xml) =>
      xml.replace(/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/, "").replace("</sheetData>",
        `</sheetData>${rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>COUNTIFS($F$5:$F$97,F5,$E$5:$E$97,E5)&gt;0</formula></cfRule>`)}`),
    );
    const sheet = await parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
    expect(sheet.unsupported.some((entry) => entry.detail.includes("expression"))).toBe(false);
    let painted = 0;
    let cold = 0;
    for (const row of sheet.rows.filter((entry) => entry.index >= 4)) {
      const hours = row.cells.find((entry) => entry.col === HOURS)?.cell;
      if (hours === undefined || typeof hours.value !== "number") continue;
      const status = row.cells.find((entry) => entry.col === 4)?.cell;
      const should = status !== undefined && status.value !== null;
      expect(hours.conditional?.fill !== undefined).toBe(should);
      if (should) painted++;
      else cold++;
    }
    expect(painted).toBeGreaterThan(0);
    expect(cold).toBeGreaterThan(0);
    // Different-size ranges are Excel's `#VALUE!`: unpainted, not reported.
    const skewed = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>COUNTIFS($F$5:$F$6,F5,$F$5:$F$7,F5)&gt;0</formula></cfRule>`));
    expect(skewed.cells.every((cell) => cell.conditional === undefined)).toBe(true);
    expect(skewed.unsupported.some((detail) => detail.includes("outside the supported subset"))).toBe(false);
  });

  test("SUMIF aligns its sum range; AVERAGEIF with no match is #DIV/0!", async () => {
    // Each Hour's peers sum to at least itself (all values are positive).
    const sums = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>SUMIF($F$5:$F$97,F5)&gt;=F5</formula></cfRule>`));
    expect(sums.unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    expect(sums.cells.every((cell) => cell.conditional?.fill !== undefined)).toBe(true);
    // An explicit shorter sum range aligns at the top-left and still evaluates.
    const short = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>SUMIF($F$5:$F$97,F5,$F$5:$F$6)&gt;=0</formula></cfRule>`));
    expect(short.unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    expect(short.cells.every((cell) => cell.conditional?.fill !== undefined)).toBe(true);
    // No match averages nothing: `#DIV/0!` paints nothing and reports nothing.
    const empty = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>AVERAGEIF($F$5:$F$97,"&gt;=1000000")&gt;0</formula></cfRule>`));
    expect(empty.cells.every((cell) => cell.conditional === undefined)).toBe(true);
    expect(empty.unsupported.some((detail) => detail.includes("outside the supported subset"))).toBe(false);
  });

  test("SUMIFS and AVERAGEIFS pair multiple ranges", async () => {
    const sums = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>SUMIFS($F$5:$F$97,$F$5:$F$97,F5,$F$5:$F$97,"&gt;="&amp;F5)&gt;=F5</formula></cfRule>`));
    expect(sums.unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    expect(sums.cells.every((cell) => cell.conditional?.fill !== undefined)).toBe(true);
    const avgs = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>AVERAGEIFS($F$5:$F$97,$F$5:$F$97,F5)&gt;0</formula></cfRule>`));
    expect(avgs.unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    expect(avgs.cells.every((cell) => cell.conditional?.fill !== undefined)).toBe(true);
  });

  test("a real cross-sheet reference evaluates through preloaded values", async () => {
    // Notes!B6 is the 450 EGP hourly rate the Shot Log costs already live against.
    const rate = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>F5&lt;Notes!$B$6</formula></cfRule>`));
    expect(rate.unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    expect(rate.cells.every((cell) => cell.conditional?.fill !== undefined)).toBe(true);
    const agg = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>COUNTIF(Notes!$B$6:$B$6,"&gt;100")&gt;0</formula></cfRule>`));
    expect(agg.unsupported.some((detail) => detail.includes("expression"))).toBe(false);
    expect(agg.cells.every((cell) => cell.conditional?.fill !== undefined)).toBe(true);
    const sums = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>SUM(Notes!$B$6:$B$6)&gt;100</formula></cfRule>`));
    expect(sums.cells.every((cell) => cell.conditional?.fill !== undefined)).toBe(true);
  });

  test("an x14 expression with inline dxf paints like a base rule", async () => {
    const ext =
      `<extLst><ext uri="{B025F937-C7B1-47D3-BEBA-786C5E5F2C56}"><x14:conditionalFormattings>` +
      `<x14:conditionalFormatting><xm:sqref>F5:F97</xm:sqref>` +
      `<x14:cfRule type="expression" priority="1"><xm:f>F5&gt;10</xm:f>` +
      `<x14:dxf><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></x14:dxf>` +
      `</x14:cfRule></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`;
    const pkg: OpcPackage = await withPart("xlsx-coverage.xlsx", SHOT_LOG, (xml) =>
      xml.replace(/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/, "").replace("</sheetData>", `</sheetData>${ext}`),
    );
    const sheet = await parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
    const cells = sheet.rows.flatMap((row) => row.cells.filter((entry) => entry.col === HOURS && typeof entry.cell.value === "number").map((entry) => entry.cell));
    expect(sheet.unsupported.some((entry) => entry.detail.includes("conditional-format"))).toBe(false);
    for (const cell of cells) {
      if ((cell.value as number) > 10) expect(cell.conditional?.fill).toBe("#FF0000");
      else expect(cell.conditional).toBeUndefined();
    }
    // The same x14 path carries cross-sheet formulas with their inline style.
    const crossExt = ext.replace("F5&lt;10", "F5&lt;Notes!$B$6").replace("F5&gt;10", "F5&lt;Notes!$B$6");
    const crossPkg: OpcPackage = await withPart("xlsx-coverage.xlsx", SHOT_LOG, (xml) =>
      xml.replace(/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/, "").replace("</sheetData>", `</sheetData>${crossExt}`),
    );
    const crossSheet = await parseXlsxSheet(crossPkg, await readWorkbook(crossPkg), 2);
    const crossCells = crossSheet.rows.flatMap((row) => row.cells.filter((entry) => entry.col === HOURS && typeof entry.cell.value === "number").map((entry) => entry.cell));
    expect(crossCells.every((cell) => cell.conditional?.fill === "#FF0000")).toBe(true);
  });

  test("a PERCENTILE formula bound paints through the evaluator", async () => {
    const { cells, unsupported } = await shotLog(
      rule(`<cfRule type="colorScale" priority="1"><colorScale><cfvo type="formula" val="PERCENTILE($F$5:$F$97,0.9)"/><cfvo type="max"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/></colorScale></cfRule>`),
    );
    expect(unsupported.some((detail) => detail.includes("formula bound"))).toBe(false);
    const sorted = numbers(cells);
    const rank = (sorted.length - 1) * 0.9;
    const threshold = sorted[Math.floor(rank)]! + (sorted[Math.ceil(rank)]! - sorted[Math.floor(rank)]!) * (rank - Math.floor(rank));
    for (const cell of cells) {
      const value = cell.value as number;
      if (value <= threshold) expect(cell.conditional?.fill).toBe("#000000");
      if (value >= sorted[sorted.length - 1]!) expect(cell.conditional?.fill).toBe("#FFFFFF");
    }
    expect(cells.some((cell) => cell.conditional?.fill !== undefined)).toBe(true);
  });

  test("a formula bound naming one cell takes that cell's cached number", async () => {
    const pkg: OpcPackage = await withPart("xlsx-coverage.xlsx", SHOT_LOG, (xml) =>
      xml.replace(/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/, "").replace("</sheetData>",
        `</sheetData>${rule(`<cfRule type="colorScale" priority="1"><colorScale><cfvo type="formula" val="$F$5"/><cfvo type="max"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/></colorScale></cfRule>`)}`),
    );
    const sheet = await parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
    const f5 = sheet.rows.find((row) => row.index === 4)!.cells.find((entry) => entry.col === HOURS)!.cell.value as number;
    const { cells } = await shotLog(rule(`<cfRule type="colorScale" priority="1"><colorScale><cfvo type="formula" val="$F$5"/><cfvo type="max"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/></colorScale></cfRule>`));
    // Everything at or below F5 is the start colour, the maximum the end colour.
    for (const cell of cells) {
      if ((cell.value as number) <= f5) expect(cell.conditional?.fill).toBe("#000000");
    }
    expect(cells.some((cell) => (cell.value as number) > f5)).toBe(true);
    expect(cells.find((cell) => cell.value === numbers(cells).at(-1))!.conditional?.fill).toBe("#FFFFFF");
  });

  test("the Excel-authored Hours data bar resolves its linked x14 defaults", async () => {
    const pkg = await OpcPackage.open(await Bun.file(`${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`).arrayBuffer());
    const sheet = await parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
    const cell = sheet.rows.find((row) => row.index === 4)!.cells.find((entry) => entry.col === 7)!.cell;
    const bar = cell.conditional?.dataBar;
    expect(bar).toEqual({
      ratio: 0.36,
      color: "#2E7D74",
      gradient: true,
      showValue: true,
      direction: "context",
      border: false,
      axisPosition: "none",
      negativeFillColor: "#2E7D74",
    });
    // A zero value uses the Office x14 default minimum length rather than disappearing.
    const zero = sheet.rows.find((row) => row.index === 11)!.cells.find((entry) => entry.col === 7)!.cell;
    expect(zero.value).toBe(0);
    expect(zero.conditional?.dataBar?.ratio).toBe(0.1);
  });

  test("an explicit x14 solid-fill setting overrides the Office gradient default", async () => {
    const pkg: OpcPackage = await withPart("xlsx-coverage.xlsx", SHOT_LOG, (xml) =>
      xml.replace('<x14:dataBar axisPosition="none">', '<x14:dataBar axisPosition="none" gradient="0">'),
    );
    const sheet = await parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
    const cell = sheet.rows.find((row) => row.index === 4)!.cells.find((entry) => entry.col === 7)!.cell;
    expect(cell.conditional?.dataBar?.gradient).toBe(false);
  });

  test("a negative data-bar range is reported instead of shown with the positive-bar layout", async () => {
    const pkg = await withPart("xlsx-coverage.xlsx", SHOT_LOG, (xml) =>
      xml.replace(/(<c\b[^>]*r="H5"[^>]*>[\s\S]*?<v>)6\.5(<\/v>)/, "$1-6.5$2"),
    );
    const sheet = await parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
    const cells = sheet.rows.flatMap((row) => row.cells.filter((entry) => entry.col === 7).map((entry) => entry.cell));
    expect(cells.find((cell) => cell.value === -6.5)?.conditional?.dataBar).toBeUndefined();
    expect(sheet.unsupported.some((entry) => entry.detail.includes("negative values"))).toBe(true);
  });

  test("top10 paints the N-th largest value and every cell tied with it, as Excel does", async () => {
    const { cells } = await shotLog(rule(`<cfRule type="top10" dxfId="1" priority="1" rank="3"/>`));
    const sorted = numbers(cells);
    const threshold = sorted[sorted.length - 3]!;
    const painted = cells.filter((cell) => cell.conditional?.fill !== undefined).map((cell) => cell.value as number);
    expect(painted.sort((a, b) => a - b)).toEqual(sorted.filter((value) => value >= threshold));
    expect(painted.length).toBeGreaterThanOrEqual(3);
  });

  test("top10 bottom rules paint the N-th smallest value and its ties", async () => {
    const { cells } = await shotLog(rule(`<cfRule type="top10" dxfId="1" priority="1" bottom="1" rank="3"/>`));
    const sorted = numbers(cells);
    const threshold = sorted[2]!;
    const painted = cells.filter((cell) => cell.conditional?.fill !== undefined).map((cell) => cell.value as number);
    expect(painted.sort((a, b) => a - b)).toEqual(sorted.filter((value) => value <= threshold));
  });

  test("a percent top rule paints floor(share of the numbers) cells, at least one", async () => {
    // Measured in Excel (Notes §105): count = max(1, floor(numeric count x percent / 100)); blanks and text are not counted.
    for (const percent of [10, 1]) {
      const { cells } = await shotLog(rule(`<cfRule type="top10" dxfId="1" priority="1" percent="1" rank="${percent}"/>`));
      const sorted = numbers(cells);
      const count = Math.max(1, Math.floor((sorted.length * percent) / 100));
      const threshold = sorted[sorted.length - count]!;
      const painted = cells.filter((cell) => cell.conditional?.fill !== undefined).map((cell) => cell.value as number);
      expect(painted.sort((a, b) => a - b)).toEqual(sorted.filter((value) => value >= threshold));
    }
  });

  test("aboveAverage paints values over the mean of the numbers, with equal and below variants", async () => {
    const { cells } = await shotLog(rule(`<cfRule type="aboveAverage" dxfId="1" priority="1"/>`));
    const values = numbers(cells);
    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
    const painted = (body: string) =>
      shotLog(body).then(({ cells: out }) => out.filter((cell) => cell.conditional?.fill !== undefined).map((cell) => cell.value as number).sort((a, b) => a - b));
    expect(await painted(rule(`<cfRule type="aboveAverage" dxfId="1" priority="1"/>`))).toEqual(values.filter((value) => value > mean));
    expect(await painted(rule(`<cfRule type="aboveAverage" dxfId="1" priority="1" aboveAverage="0"/>`))).toEqual(values.filter((value) => value < mean));
    expect(await painted(rule(`<cfRule type="aboveAverage" dxfId="1" priority="1" equalAverage="1"/>`))).toEqual(values.filter((value) => value >= mean));
    expect(await painted(rule(`<cfRule type="aboveAverage" dxfId="1" priority="1" aboveAverage="0" equalAverage="1"/>`))).toEqual(values.filter((value) => value <= mean));
  });

  test("an aboveAverage rule with a standard deviation paints mean +/- population stdev", async () => {
    // Measured in Excel 16 (tools/probe/xlsx-conditional-path/probe_std.py):
    // 10,9,8,8,8,7,6,5,4,3 with k=1 paints 10,9 above and 4,3 below, which is
    // mean +/- pstdev (8.94), not sample stdev (9.05).
    const painted = async (body: string) =>
      shotLog(body).then(({ cells: out }) => out.filter((cell) => cell.conditional?.fill !== undefined).map((cell) => cell.value as number).sort((a, b) => a - b));
    const { cells } = await shotLog(rule(`<cfRule type="aboveAverage" dxfId="1" priority="1" stdDev="1"/>`));
    const values = numbers(cells);
    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
    const pstdev = Math.sqrt(values.reduce((sum, value) => sum + (value - mean) * (value - mean), 0) / values.length);
    expect(await painted(rule(`<cfRule type="aboveAverage" dxfId="1" priority="1" stdDev="1"/>`))).toEqual(values.filter((value) => value > mean + pstdev));
    expect(await painted(rule(`<cfRule type="aboveAverage" dxfId="1" priority="1" aboveAverage="0" stdDev="1"/>`))).toEqual(values.filter((value) => value < mean - pstdev));
    expect(await painted(rule(`<cfRule type="aboveAverage" dxfId="1" priority="1" stdDev="1" equalAverage="1"/>`))).toEqual(values.filter((value) => value >= mean + pstdev));
  });

  test("timePeriod paints date-formatted serials around a pinned today", async () => {
    // RTL sheet B10 holds serial 46278 (13 Sep 2026, date-formatted); Hours are numbers, never dates.
    const rtl = async (body: string, today: { year: number; month: number; day: number }) => {
      const pkg: OpcPackage = await withPart("xlsx-coverage.xlsx", "xl/worksheets/sheet1.xml", (xml) =>
        xml.replace(/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/, "").replace("</sheetData>", `</sheetData>${fixtureDxfIds(body)}`));
      return parseXlsxSheet(pkg, await readWorkbook(pkg), 0, undefined, today);
    };
    const today = { year: 2026, month: 9, day: 13 };
    const sheet = await rtl(`<conditionalFormatting sqref="B10"><cfRule type="timePeriod" dxfId="1" priority="1" timePeriod="today"/></conditionalFormatting>`, today);
    const dateCell = sheet.rows.find((row) => row.index === 9)!.cells.find((entry) => entry.col === 1)!.cell;
    expect(dateCell.value).toBe(46278);
    expect(dateCell.conditional?.fill).toBeDefined();
    const otherDay = await rtl(`<conditionalFormatting sqref="B10"><cfRule type="timePeriod" dxfId="1" priority="1" timePeriod="today"/></conditionalFormatting>`, { year: 2026, month: 9, day: 14 });
    const cold = otherDay.rows.find((row) => row.index === 9)!.cells.find((entry) => entry.col === 1)!.cell;
    expect(cold.conditional).toBeUndefined();
    // last7Days includes today; numbers never match, whatever today is.
    const week = await rtl(`<conditionalFormatting sqref="B10"><cfRule type="timePeriod" dxfId="1" priority="1" timePeriod="last7Days"/></conditionalFormatting>`, today);
    expect(week.rows.find((row) => row.index === 9)!.cells.find((entry) => entry.col === 1)!.cell.conditional?.fill).toBeDefined();
    const hours = await shotLog(rule(`<cfRule type="timePeriod" dxfId="1" priority="1" timePeriod="today"/>`));
    expect(hours.cells.every((cell) => cell.conditional === undefined)).toBe(true);
  });

  test("icon sets resolve tiers with the first cfvo ignored and inclusive bounds", async () => {
    const { cells } = await shotLog(
      rule(`<cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>`),
    );
    const sorted = numbers(cells);
    const min = sorted[0]!;
    const max = sorted[sorted.length - 1]!;
    const b1 = min + ((max - min) * 33) / 100;
    const b2 = min + ((max - min) * 67) / 100;
    for (const cell of cells) {
      const value = cell.value as number;
      const tier = value >= b2 ? 2 : value >= b1 ? 1 : 0;
      expect(cell.conditional?.icon).toEqual({ family: "3Arrows", icon: tier, showValue: true });
    }
  });

  test("icon sets honour reverse and icon-only display", async () => {
    const { cells } = await shotLog(
      rule(`<cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows" reverse="1" showValue="0"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>`),
    );
    const sorted = numbers(cells);
    const min = sorted[0]!;
    const max = sorted[sorted.length - 1]!;
    const b1 = min + ((max - min) * 33) / 100;
    const b2 = min + ((max - min) * 67) / 100;
    for (const cell of cells) {
      const value = cell.value as number;
      const tier = value >= b2 ? 2 : value >= b1 ? 1 : 0;
      expect(cell.conditional?.icon).toEqual({ family: "3Arrows", icon: 2 - tier, showValue: false });
    }
  });

  test("custom icon sets name each tier's icon, honour NoIcons tiers and ignore reverse", async () => {
    // `custom="1"` picks per-tier `cfIcon`s (here from other families) rather
    // than the named set's own order, and MS-XLSX ignores `reverse` then.
    const { cells } = await shotLog(
      rule(`<cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows" custom="1" reverse="1">` +
        `<cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/>` +
        `<cfIcon iconSet="3TrafficLights1" iconId="0"/><cfIcon iconSet="NoIcons" iconId="0"/><cfIcon iconSet="4Arrows" iconId="3"/>` +
        `</iconSet></cfRule>`),
    );
    const sorted = numbers(cells);
    const min = sorted[0]!;
    const max = sorted[sorted.length - 1]!;
    const b1 = min + ((max - min) * 33) / 100;
    const b2 = min + ((max - min) * 67) / 100;
    for (const cell of cells) {
      const value = cell.value as number;
      const tier = value >= b2 ? 2 : value >= b1 ? 1 : 0;
      if (tier === 1) {
        // A per-tier `NoIcons` paints nothing for that tier, by design.
        expect(cell.conditional?.icon).toBeUndefined();
      } else if (tier === 0) {
        expect(cell.conditional?.icon).toEqual({ family: "3TrafficLights1", icon: 0, showValue: true });
      } else {
        expect(cell.conditional?.icon).toEqual({ family: "4Arrows", icon: 3, showValue: true });
      }
    }
  });

  test("an icon-set rule with a PERCENTILE bound paints through the evaluator", async () => {
    const { cells, unsupported } = await shotLog(
      rule(`<cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="formula" val="PERCENTILE($F$5:$F$97,0.5)"/><cfvo type="percent" val="67"/></iconSet></cfRule>`),
    );
    expect(unsupported.some((detail) => detail.includes("icon-set"))).toBe(false);
    expect(cells.some((cell) => cell.conditional?.icon !== undefined)).toBe(true);
    for (const cell of cells) {
      const icon = cell.conditional?.icon;
      if (icon !== undefined) {
        expect(icon.family).toBe("3Arrows");
        expect(icon.icon).toBeGreaterThanOrEqual(0);
        expect(icon.icon).toBeLessThanOrEqual(2);
      }
    }
  });

  test("an icon-set rule with an unsupported formula bound is reported", async () => {
    const { cells, unsupported } = await shotLog(
      rule(`<cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="formula" val="NOSUCHFUNC($F$5)"/><cfvo type="percent" val="67"/></iconSet></cfRule>`),
    );
    expect(cells.every((cell) => cell.conditional?.icon === undefined)).toBe(true);
    expect(unsupported.some((detail) => detail.includes("icon-set"))).toBe(true);
  });

  test("a malformed standalone x14 icon set is reported, not silently dropped", async () => {
    const ext = (set: string) =>
      `<extLst><ext uri="{B025F937-C7B1-47D3-BEBA-786C5E5F2C56}"><x14:conditionalFormattings>` +
      `<x14:conditionalFormatting><xm:sqref>F5:F97</xm:sqref>` +
      `<x14:cfRule type="iconSet" priority="1"><x14:iconSet ${set}>` +
      `<x14:cfvo type="percent" val="0"/><x14:cfvo type="percent" val="33"/><x14:cfvo type="percent" val="67"/>` +
      `</x14:iconSet></x14:cfRule></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`;
    const withExt = async (inner: string) => {
      const pkg: OpcPackage = await withPart("xlsx-coverage.xlsx", SHOT_LOG, (xml) =>
        xml.replace(/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/, "").replace("</sheetData>", `</sheetData>${inner}`),
      );
      return parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
    };
    const bad = await withExt(ext(`iconSet="NoSuchFamily"`));
    const badCells = bad.rows.flatMap((row) => row.cells.filter((entry) => entry.col === HOURS).map((entry) => entry.cell));
    expect(badCells.every((cell) => cell.conditional?.icon === undefined)).toBe(true);
    expect(bad.unsupported.some((entry) => entry.detail.includes('"iconSet" conditional rule is not evaluated'))).toBe(true);
    // `NoIcons` paints nothing by design and is never reported.
    const none = await withExt(ext(`iconSet="NoIcons"`));
    const noneCells = none.rows.flatMap((row) => row.cells.filter((entry) => entry.col === HOURS).map((entry) => entry.cell));
    expect(noneCells.every((cell) => cell.conditional?.icon === undefined)).toBe(true);
    expect(none.unsupported.some((entry) => entry.detail.includes("iconSet"))).toBe(false);
  });

  test("the Excel-authored CF Rules sheet paints icons and stdDev, and evaluates its date rule", async () => {
    const pkg = await OpcPackage.open(await Bun.file(`${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`).arrayBuffer());
    const workbook = await readWorkbook(pkg);
    const sheet = await parseXlsxSheet(pkg, workbook, 4);
    expect(sheet.name).toBe("CF Rules");
    // Icons on A2:A11: 1..10 with Excel's default 0/33/67 percent thresholds.
    const icons = sheet.rows.flatMap((row) => row.cells.filter((entry) => entry.col === 0 && typeof entry.cell.value === "number"));
    expect(icons).toHaveLength(10);
    expect(icons.every(({ cell }) => cell.conditional?.icon?.family === "3Arrows")).toBe(true);
    const tiers = icons.map(({ cell }) => cell.conditional?.icon?.icon);
    expect(Math.min(...tiers as number[])).toBe(0);
    expect(Math.max(...tiers as number[])).toBe(2);
    // StdDev on C2:C11 (10,9,8,8,8,7,6,5,4,3, k=1) paints 10 and 9.
    const stds = sheet.rows.flatMap((row) => row.cells.filter((entry) => entry.col === 2).map((entry) => entry.cell));
    const painted = stds.filter((cell) => cell.conditional?.fill !== undefined).map((cell) => cell.value as number).sort((a, b) => a - b);
    expect(painted).toEqual([9, 10]);
    // The date rule is evaluated (not reported) but its 2020 dates miss today.
    expect(sheet.unsupported.some((entry) => entry.detail.includes("timePeriod"))).toBe(false);
    expect(sheet.unsupported.some((entry) => entry.detail.includes("icon-set"))).toBe(false);
    expect(sheet.unsupported.some((entry) => entry.detail.includes("standard deviation"))).toBe(false);
  });

  test("the Excel-authored expression paints even values via relative MOD", async () => {
    // COM-authored `=MOD($F2,2)=0` on F2:F12 (1..11) paints rows 3,5,7,9,11 (values 2,4,6,8,10).
    const pkg = await OpcPackage.open(await Bun.file(`${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`).arrayBuffer());
    const workbook = await readWorkbook(pkg);
    const index = workbook.sheets.findIndex((sheet) => sheet.name === "CF Rules");
    const sheet = await parseXlsxSheet(pkg, workbook, index);
    const cells = sheet.rows.flatMap((row) => row.cells.filter((entry) => entry.col === 5 && typeof entry.cell.value === "number").map((entry) => ({ row: row.index, cell: entry.cell })));
    expect(cells.map(({ cell }) => cell.value)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]);
    const painted = cells.filter(({ cell }) => cell.conditional?.fill !== undefined).map(({ cell }) => cell.value as number).sort((a, b) => a - b);
    expect(painted).toEqual([2, 4, 6, 8, 10]);
    expect(sheet.unsupported.some((entry) => entry.detail.includes("expression"))).toBe(false);
  });

  test("the Excel-authored TODAY, COUNTIF, SEARCH and cross-sheet rules paint their native rows", async () => {
    // COM-authored (§111, native DisplayFormat readback G:[2-7] H:[2,3,4,6,8,11,13] I:[2,3,5,7,9] J:[6-13]).
    const pkg = await OpcPackage.open(await Bun.file(`${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`).arrayBuffer());
    const workbook = await readWorkbook(pkg);
    const index = workbook.sheets.findIndex((sheet) => sheet.name === "CF Rules");
    const sheet = await parseXlsxSheet(pkg, workbook, index);
    const col = (n: number) =>
      sheet.rows.flatMap((row) => row.cells.filter((entry) => entry.col === n).map((entry) => ({ row: row.index + 1, cell: entry.cell })));
    // G2:G13 are frozen 2020/2099 serials under `=$G2<TODAY()`: the past paints.
    const due = col(6).filter(({ cell }) => typeof cell.value === "number");
    expect(due.map(({ cell }) => cell.value)).toEqual([43831, 43832, 43833, 43834, 43835, 43836, 72686, 72687, 72688, 72689, 72690, 72691]);
    expect(due.filter(({ cell }) => cell.conditional?.fill !== undefined).map(({ row }) => row)).toEqual([2, 3, 4, 5, 6, 7]);
    // H2:H13 duplicates under `=COUNTIF($H$2:$H$13,H2)>1`: apple x4 and banana x3.
    const dup = col(7);
    expect(dup.filter(({ cell }) => cell.conditional?.fill !== undefined).map(({ row }) => row)).toEqual([2, 3, 4, 6, 8, 11, 13]);
    // I2:I9 text under `=ISNUMBER(SEARCH("apple",$I2))`.
    const found = col(8);
    expect(found.filter(({ cell }) => cell.conditional?.fill !== undefined).map(({ row }) => row)).toEqual([2, 3, 5, 7, 9]);
    // J2:J13 hundreds under the x14 `=$J2>Notes!$B$6` (rate 450): 500 and up.
    const cross = col(9).filter(({ cell }) => typeof cell.value === "number");
    expect(cross.filter(({ cell }) => cell.conditional?.fill !== undefined).map(({ row }) => row)).toEqual([6, 7, 8, 9, 10, 11, 12, 13]);
    expect(cross.filter(({ cell }) => cell.conditional?.fill !== undefined).every(({ cell }) => cell.conditional?.fill === "#FF0000")).toBe(true);
    expect(sheet.unsupported.some((entry) => entry.kind === "conditional-format")).toBe(false);
  });

  test("the Excel-authored three-colour scale agrees with native painted colours", async () => {
    const pkg = await OpcPackage.open(await Bun.file(`${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`).arrayBuffer());
    const workbook = await readWorkbook(pkg);
    const index = workbook.sheets.findIndex((sheet) => sheet.name === "CF Rules");
    const sheet = await parseXlsxSheet(pkg, workbook, index);
    const cells = sheet.rows.flatMap((row) => row.cells.filter((entry) => entry.col === 4 && typeof entry.cell.value === "number").map((entry) => entry.cell));
    // Excel 16.0 DisplayFormat.Interior.Color, read after AddColorScale(3)
    // authored min / percentile 50 / max on E2:E12. Interpolated channels
    // differ by at most one level; the three authored stops match exactly.
    const native = ["F8696B", "F98370", "FA9D75", "FCB77A", "FDD17F", "FFEB84", "E0E383", "C1DA81", "A2D07F", "83C77D", "63BE7B"];
    expect(cells.map((cell) => cell.value)).toEqual([0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100]);
    for (const [i, cell] of cells.entries()) {
      const fill = cell.conditional?.fill;
      expect(fill).toMatch(/^#[0-9A-F]{6}$/);
      for (const channel of [0, 2, 4]) {
        const painted = Number.parseInt(fill!.slice(channel + 1, channel + 3), 16);
        const excel = Number.parseInt(native[i]!.slice(channel, channel + 2), 16);
        expect(Math.abs(painted - excel)).toBeLessThanOrEqual(1);
      }
    }
    expect(cells[0]!.conditional?.fill).toBe("#F8696B");
    expect(cells[5]!.conditional?.fill).toBe("#FFEB84");
    expect(cells[10]!.conditional?.fill).toBe("#63BE7B");
    expect(sheet.unsupported.some((entry) => entry.kind === "conditional-format")).toBe(false);
  });
});
