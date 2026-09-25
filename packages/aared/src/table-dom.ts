/**
 * The table shell PPTX and DOCX both paint: a CSS grid in the table's own
 * direction, with cells that span and align vertically. Renderer-side. What
 * differs per format stays with the format: PPTX pads in physical left/right
 * and falls back to a faint grid; DOCX pads in inline-start/end and collapses
 * shared edges.
 */

import type { TextDirection } from "./core/model.js";

export type CellVerticalAlign = "top" | "middle" | "bottom";

/** Grid tracks for column widths: each a percentage of the total, or equal `1fr` tracks when no width is authored. */
function columnTracks(columns: readonly number[]): string {
  const total = columns.reduce((sum, width) => sum + width, 0);
  return total > 0
    ? columns.map((width) => `${(width / total) * 100}%`).join(" ")
    : columns.map(() => "1fr").join(" ");
}

/** The `<div role="table">` shell: a grid in the table's direction, columns proportioned to their widths. */
export function createTableGrid(direction: TextDirection, columns: readonly number[]): HTMLDivElement {
  const grid = document.createElement("div");
  grid.setAttribute("role", "table");
  grid.dataset.aaredElement = "table";
  grid.dataset.aaredTableDirection = direction;
  grid.style.display = "grid";
  grid.style.direction = direction;
  if (columns.length > 0) grid.style.gridTemplateColumns = columnTracks(columns);
  return grid;
}

/** How a cell's content sits in a taller cell: the flex `justify-content` for a vertical alignment. */
export function tableCellJustifyContent(verticalAlign: CellVerticalAlign): "flex-start" | "center" | "flex-end" {
  switch (verticalAlign) {
    case "top":
      return "flex-start";
    case "middle":
      return "center";
    case "bottom":
      return "flex-end";
  }
}

/** One `role="cell"` shell: it may shrink below its content, and carries its spans and vertical alignment. */
export function createTableCell(cell: { columnSpan?: number; rowSpan?: number; verticalAlign?: CellVerticalAlign }): HTMLDivElement {
  const element = document.createElement("div");
  element.setAttribute("role", "cell");
  element.dataset.aaredElement = "table-cell";
  element.style.minWidth = "0";
  element.style.boxSizing = "border-box";
  if (cell.columnSpan !== undefined) {
    element.style.gridColumn = `span ${cell.columnSpan}`;
    element.dataset.aaredColumnSpan = String(cell.columnSpan);
  }
  if (cell.rowSpan !== undefined) {
    element.style.gridRow = `span ${cell.rowSpan}`;
    element.dataset.aaredRowSpan = String(cell.rowSpan);
  }
  if (cell.verticalAlign !== undefined) {
    element.style.display = "flex";
    element.style.flexDirection = "column";
    element.style.justifyContent = tableCellJustifyContent(cell.verticalAlign);
  }
  return element;
}
