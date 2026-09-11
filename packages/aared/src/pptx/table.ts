import type { TableCellModel } from "../core/model.js";

export function tableCellJustifyContent(
  verticalAlign: NonNullable<TableCellModel["verticalAlign"]>,
): "flex-start" | "center" | "flex-end" {
  switch (verticalAlign) {
    case "top":
      return "flex-start";
    case "middle":
      return "center";
    case "bottom":
      return "flex-end";
  }
}
