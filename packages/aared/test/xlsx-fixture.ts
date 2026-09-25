import type { SheetModel } from "../src/core/model.js";

// XLSX counterpart of fixture.ts / the DOCX manifest lookup: coverage
// sheets are located by unique authored text, never by sheet index, so a
// rebuilt workbook that reorders sheets fails loud instead of silently
// testing the wrong cells.

const MANIFEST_PATH = `${import.meta.dir}/../../../fixtures/xlsx-coverage.json`;

interface TextSelector {
  hasText: string;
}

interface CoverageManifest {
  id: string;
  sheets: Record<string, TextSelector>;
}

function sheetText(sheet: SheetModel): string {
  return sheet.rows
    .map((row) => row.cells.map(({ cell }) => cell.formatted).join("\t"))
    .join("\n");
}

function sheetSummary(sheet: SheetModel): string {
  const head = sheet.rows
    .slice(0, 3)
    .map((row) => row.cells.map(({ cell }) => cell.formatted).join(" | ").slice(0, 80))
    .join(" / ");
  return `${sheet.name} (rtl=${sheet.rightToLeft}): ${head}`;
}

export async function xlsxCoverageManifest(): Promise<CoverageManifest> {
  return JSON.parse(await Bun.file(MANIFEST_PATH).text()) as CoverageManifest;
}

export async function findCoverageSheet(
  sheets: SheetModel[],
  id: string,
): Promise<{ index: number; sheet: SheetModel }> {
  const manifest = await xlsxCoverageManifest();
  const selector = manifest.sheets[id];
  if (selector === undefined) throw new Error(`fixture ${manifest.id}: unknown sheet selector "${id}"`);
  const matches = sheets
    .map((sheet, index) => ({ index, sheet }))
    .filter(({ sheet }) => sheetText(sheet).includes(selector.hasText));
  if (matches.length === 1) return matches[0]!;
  const candidates = sheets.map((sheet, index) => `${index}: ${sheetSummary(sheet)}`).join("\n");
  throw new Error(`fixture ${manifest.id}: selector "${id}" expected one sheet containing "${selector.hasText}", found ${matches.length}\n${candidates}`);
}
