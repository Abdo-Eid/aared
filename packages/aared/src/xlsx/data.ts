import type { CellModel, CellRange, SheetRangeModel } from "../core/model.js";

export type { CellRange } from "../core/model.js";

/** Cached spreadsheet data, independent of rendering styles and storage kinds. */
export interface CellData {
  value: string | number | boolean | null;
  formatted: string;
  type: "string" | "number" | "boolean" | "error" | "empty";
  formula?: string;
  hyperlink?: { target: string; display?: string };
}

/** Sparse cells in ascending logical column order. */
export interface RowData {
  index: number;
  cells: Array<{ col: number; cell: CellData }>;
}

/** A read window: whole-sheet extent/frozen panes, with merges clipped to the window. */
export interface SheetData {
  name: string;
  rightToLeft: boolean;
  extent: { cols: number; rows: number };
  range: CellRange;
  merges: CellRange[];
  frozen: { rows: number; cols: number };
  rows: RowData[];
}

/** Project, never cast: nested containers must not retain renderer-owned data. */
export function cellData(cell: CellModel): CellData {
  const type = cell.type === "shared" || cell.type === "inline" || cell.type === "formula-string" ? "string" : cell.type;
  return {
    value: cell.value, formatted: cell.formatted, type,
    ...(cell.formula === undefined ? {} : { formula: cell.formula }),
    ...(cell.hyperlink === undefined ? {} : { hyperlink: {
      target: cell.hyperlink.target,
      ...(cell.hyperlink.display === undefined ? {} : { display: cell.hyperlink.display }),
    } }),
  };
}

function copyRange(range: CellRange): CellRange {
  return { start: { row: range.start.row, col: range.start.col }, end: { row: range.end.row, col: range.end.col } };
}

export function sheetData(model: SheetRangeModel): SheetData {
  return {
    name: model.name, rightToLeft: model.rightToLeft,
    extent: { rows: model.extent.rows, cols: model.extent.cols }, range: copyRange(model.range),
    merges: model.merges.map(copyRange), frozen: { rows: model.frozen.rows, cols: model.frozen.cols },
    rows: model.rows.map(row => ({ index: row.index, cells: row.cells.map(({ col, cell }) => ({ col, cell: cellData(cell) })) })),
  };
}
