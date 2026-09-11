import { open, render } from "aared/xlsx";
import type { DemoOutcome } from "../lib/demo-frame.js";

/**
 * Headless XLSX: open the workbook, render the first sheet.
 * `sheet.read("A1:M200")` windows a range without rendering for data access.
 */
export async function runHeadlessXlsx(stage: HTMLElement, source: string | File): Promise<DemoOutcome> {
  const workbook = await open(source);
  try {
    const viewer = await render({ sheet: workbook.sheets.get(0), target: stage });
    return {
      summary: { format: workbook.info.format, units: workbook.info.count },
      dispose: () => {
        viewer.destroy();
        workbook.close();
      },
    };
  } catch (error) {
    workbook.close();
    throw error;
  }
}
