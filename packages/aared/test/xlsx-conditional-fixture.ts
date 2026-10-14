import { OpcPackage } from "../src/core/opc.js";
import { withPart } from "./mutate.js";

/** Replace Shot Log's rules using its real tan differential style and optional missing cells. */
export async function conditionalFormulaFixture(options: {
  formula: string;
  sqref: string;
  blankCells?: readonly string[];
}): Promise<OpcPackage> {
  const original = await OpcPackage.open(await Bun.file(`${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`).arrayBuffer());
  const styles = [...(await original.readText("xl/styles.xml")).matchAll(/<dxf>[\s\S]*?<\/dxf>/g)];
  const dxf = styles.findIndex((entry) => entry[0].includes('<bgColor rgb="FFFFDCB4"/>'));
  if (dxf < 0) throw new Error("XLSX coverage fixture has no tan differential fill");
  const blanks = new Set(options.blankCells ?? []);
  let removed = 0;
  const formula = options.formula.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
  const pkg = await withPart("xlsx-coverage.xlsx", "xl/worksheets/sheet3.xml", (xml) =>
    xml.replace(/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/, "")
      .replace(/<c\b[^>]*\br="([^"]+)"[^>]*(?:\/>|>[\s\S]*?<\/c>)/g, (cell, ref: string) => {
        if (!blanks.has(ref)) return cell;
        removed++;
        return "";
      })
      .replace("</sheetData>", `</sheetData><conditionalFormatting sqref="${options.sqref}"><cfRule type="expression" dxfId="${dxf}" priority="1"><formula>${formula}</formula></cfRule></conditionalFormatting>`),
  );
  if (removed !== blanks.size) throw new Error("XLSX coverage mutation did not remove every requested cell");
  return pkg;
}
