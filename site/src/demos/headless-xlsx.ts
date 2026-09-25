import { open, render } from "aared/xlsx";
import type { XlsxRenderHandle } from "aared/xlsx";
import type { ZoomValue } from "aared/core";
import { mountDemoControls, type DemoController } from "../lib/demo-controls.js";
import type { DemoOutcome } from "../lib/demo-frame.js";

/**
 * Headless XLSX: open the workbook, render the first sheet, with the
 * built-in sheet tabs so people can move between sheets (or drive it
 * yourself with `viewer.goToSheet(titleOrIndex)`, as the paging bar does).
 * `sheet.read({ range: "A1:M200" })` windows a range without rendering for
 * data access.
 */
export async function runHeadlessXlsx(stage: HTMLElement, source: string | File): Promise<DemoOutcome> {
  const workbook = await open(source);
  try {
    const viewer = await render({
      sheet: workbook.sheets.get(0),
      target: stage,
      appearance: { spreadsheet: { sheetTabs: true } },
    });
    const unmountControls = mountDemoControls(stage, sheetController(viewer), workbook.info.count);
    return {
      summary: { format: workbook.info.format, units: workbook.info.count },
      dispose: () => {
        unmountControls();
        viewer.destroy();
        workbook.close();
      },
    };
  } catch (error) {
    workbook.close();
    throw error;
  }
}

/** The one render handle already moves between sheets: `goToSheet()` and `sheetchange`. */
function sheetController(viewer: XlsxRenderHandle): DemoController {
  const controller = new EventTarget() as DemoController;
  viewer.addEventListener("sheetchange", () => controller.dispatchEvent(new Event("indexchange")));
  viewer.addEventListener("zoomchange", () => controller.dispatchEvent(new Event("zoomchange")));
  // Live getters: the handle's sheet and zoom change after this is built.
  return Object.defineProperties(controller, {
    index: { get: () => viewer.sheet.index },
    zoom: { get: () => viewer.zoom },
    scale: { get: () => viewer.scale },
    goTo: { value: (index: number) => viewer.goToSheet(index) },
    setZoom: { value: (zoom: ZoomValue) => viewer.setZoom(zoom) },
  });
}
