import { describe, expect, test } from "bun:test";
import type { OpcPackage } from "../src/core/opc.js";
import type { CellModel } from "../src/core/model.js";
import { parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { readWorkbook } from "../src/core/xlsx/workbook.js";
import { withPart } from "./mutate.js";

// Conditional formats on the Shot Log sheet of the Arabic coverage workbook:
// its Excel-authored rules are replaced by rules in the shapes Excel writes
// (ECMA-376 18.3.1.10), over the Hours column (F5:F97, 93 numbers). Rule
// bodies are written here, not by Excel, so the shapes are pinned and Excel's
// own rendering of them is not (docs/ROADMAP.md, Open gaps).

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

const rule = (body: string, sqref = "F5:F97") => `<conditionalFormatting sqref="${sqref}">${body}</conditionalFormatting>`;
const numbers = (cells: CellModel[]) => cells.map((cell) => cell.value as number).sort((a, b) => a - b);

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
    // 93 values: the 50th percentile is the 47th value itself, so it takes the middle colour exactly.
    expect(fillOf(sorted[46]!)).toBe("#FFEB84");
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

  test("a rule type that needs a formula or ranking is reported, not guessed", async () => {
    const { cells, unsupported } = await shotLog(rule(`<cfRule type="expression" dxfId="1" priority="1"><formula>F5&gt;10</formula></cfRule>`));
    expect(cells.every((cell) => cell.conditional === undefined)).toBe(true);
    expect(unsupported.some((detail) => detail.includes('"expression" conditional rule is not evaluated'))).toBe(true);
  });

  test("a formula bound leaves the scale unpainted and says so", async () => {
    const { cells, unsupported } = await shotLog(
      rule(`<cfRule type="colorScale" priority="1"><colorScale><cfvo type="formula" val="$A$1"/><cfvo type="max"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/></colorScale></cfRule>`),
    );
    expect(cells.every((cell) => cell.conditional === undefined)).toBe(true);
    expect(unsupported.some((detail) => detail.includes("formula bound"))).toBe(true);
  });
});
