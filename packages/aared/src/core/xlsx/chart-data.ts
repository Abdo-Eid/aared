/**
 * The cell values behind an Office 2016 chart (chartex) in a workbook.
 *
 * Excel writes such a chart's part with no data cache: each `cx:f` names a
 * hidden `_xlchart.v1.N` defined name, and the values live in the cells the
 * name points at (`tools/chart-oracle/chartex-oracle.xlsx`). So the sheet
 * reader resolves those formulas before `readChartModel` reads the part. Only
 * what the chart needs is read: the streamed rows of each sheet a range sits
 * on, kept only inside the range, bounded like every chart read.
 *
 * A range it cannot resolve (a union of areas, a whole column, a missing
 * sheet, more points than a chart reads) is left out; the chart then says it
 * has no values to preview.
 */

import { attr, childElements, firstChild, textContent } from "../xml.js";
import type { XmlElement } from "../xml.js";
import type { OpcPackage } from "../opc.js";
import { throwIfAborted } from "../errors.js";
import type { ChartexData, ChartexGrid } from "../chart.js";
import { MAX_CATEGORY_LEVELS, MAX_SERIES_POINTS } from "../chart/read.js";
import { parseCell, parseA1Range } from "./range.js";
import type { XlsxWorkbook } from "./workbook.js";

interface Target {
  formula: string;
  part: string;
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

/** `Sheet1!$A$2:$B$7`, `'My Sheet'!A1:A9` -> its sheet name and A1 range; `undefined` for anything else. */
function splitReference(text: string): { sheet: string; range: string } | undefined {
  const t = text.trim();
  if (t.includes(",") || t.startsWith("(")) return undefined;
  const bang = t.lastIndexOf("!");
  if (bang <= 0) return undefined;
  let sheet = t.slice(0, bang);
  if (sheet.startsWith("'") && sheet.endsWith("'") && sheet.length >= 2) sheet = sheet.slice(1, -1).replaceAll("''", "'");
  return { sheet, range: t.slice(bang + 1) };
}

function cellValue(cell: XmlElement, sharedStrings: readonly string[]): string | number | null {
  const kind = attr(cell, "t") ?? "n";
  const v = firstChild(cell, "v");
  const raw = v === undefined ? undefined : textContent(v);
  switch (kind) {
    case "s": {
      const index = raw === undefined ? NaN : Number(raw);
      return Number.isInteger(index) && index >= 0 && index < sharedStrings.length ? sharedStrings[index]! : null;
    }
    case "inlineStr": {
      const inline = firstChild(cell, "is");
      return inline === undefined ? null : textContent(inline);
    }
    case "str":
      return raw ?? null;
    case "b":
    case "e":
      return null;
    default: {
      if (raw === undefined || raw === "") return null;
      const n = Number(raw);
      return Number.isFinite(n) ? n : null;
    }
  }
}

/**
 * The grids for `formulas` (what `chartexFormulas` lists), by formula text.
 * Never throws for a range it cannot read: that formula is simply absent.
 */
export async function resolveChartexData(
  pkg: OpcPackage,
  workbook: XlsxWorkbook,
  formulas: readonly string[],
  signal?: AbortSignal,
): Promise<ChartexData> {
  const out: ChartexData = {};
  if (formulas.length === 0) return out;
  const root = await pkg.readXml(workbook.workbookPart, signal);
  const names = new Map<string, string>();
  const defined = firstChild(root, "definedNames");
  for (const el of defined === undefined ? [] : childElements(defined, "definedName")) {
    const name = attr(el, "name");
    if (name !== undefined && !names.has(name)) names.set(name, textContent(el));
  }
  const byPart = new Map<string, Target[]>();
  for (const formula of formulas) {
    const ref = splitReference(names.get(formula) ?? formula);
    if (ref === undefined) continue;
    const sheet = workbook.sheets.find((s) => s.name.toLowerCase() === ref.sheet.toLowerCase());
    if (sheet === undefined) continue;
    let range;
    try {
      range = parseA1Range(ref.range);
    } catch {
      continue;
    }
    const rows = range.end.row - range.start.row + 1;
    const cols = range.end.col - range.start.col + 1;
    if (rows > MAX_SERIES_POINTS || cols > MAX_CATEGORY_LEVELS) continue;
    const target: Target = { formula, part: sheet.part, r0: range.start.row, r1: range.end.row, c0: range.start.col, c1: range.end.col };
    const list = byPart.get(sheet.part);
    if (list === undefined) byPart.set(sheet.part, [target]);
    else list.push(target);
    const grid: ChartexGrid = [];
    for (let r = 0; r < rows; r++) grid.push(new Array<string | number | null>(cols).fill(null));
    out[formula] = grid;
  }
  for (const [part, targets] of byPart) {
    throwIfAborted(signal);
    await pkg.readXml(part, signal, {
      streamChildrenOf: "sheetData",
      onChild: (row) => {
        if (row.local !== "row") return;
        const rowNumber = Number(attr(row, "r"));
        if (!Number.isInteger(rowNumber)) return;
        const wanted = targets.filter((t) => rowNumber - 1 >= t.r0 && rowNumber - 1 <= t.r1);
        if (wanted.length === 0) return;
        for (const cell of row.children) {
          if (cell.local !== "c") continue;
          const ref = attr(cell, "r");
          const at = ref === undefined ? undefined : parseCell(ref);
          if (at === undefined) continue;
          for (const t of wanted) {
            if (at.row < t.r0 || at.row > t.r1 || at.col < t.c0 || at.col > t.c1) continue;
            out[t.formula]![at.row - t.r0]![at.col - t.c0] = cellValue(cell, workbook.sharedStrings);
          }
        }
      },
    });
  }
  return out;
}
