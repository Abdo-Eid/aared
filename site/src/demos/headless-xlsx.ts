import { open, render } from "aared/xlsx";
import { button, label, mountBar, stepZoom, zoomLabel } from "../lib/demo-controls.js";
import type { DemoOutcome } from "../lib/demo-frame.js";

/**
 * Headless XLSX: open the workbook, render the first sheet with the built-in
 * sheet tabs. One render handle moves between sheets itself: the bar calls
 * `goToSheet()` (the newest call wins) and follows `sheetchange` (a tab click
 * fires it too).
 * `sheet.read({ range: "A1:M200" })` windows a range without rendering for
 * data access.
 */
export async function runHeadlessXlsx(stage: HTMLElement, source: string | File): Promise<DemoOutcome> {
  const workbook = await open(source);
  try {
    const count = workbook.info.count;
    const viewer = await render({
      sheet: workbook.sheets.get(0),
      target: stage,
      appearance: { spreadsheet: { sheetTabs: true } },
    });

    // Where the newest click is heading, so quick clicks add up; a sheet tab moves it too.
    let heading = 0;
    const go = (delta: number): void => {
      const to = Math.min(Math.max(heading + delta, 0), count - 1);
      if (to === heading) return;
      heading = to;
      void viewer.goToSheet(to).catch(() => {
        if (heading === to) heading = viewer.sheet.index; // this sheet failed: step on from the one on screen
      });
    };
    const previous = button("Previous", () => go(-1));
    const position = label();
    const next = button("Next", () => go(1));
    const zoom = label();
    const update = (): void => {
      position.textContent = `${viewer.sheet.index + 1} / ${count}`;
      previous.disabled = viewer.sheet.index <= 0;
      next.disabled = viewer.sheet.index >= count - 1;
      zoom.textContent = zoomLabel(viewer.zoom);
    };
    viewer.addEventListener("sheetchange", () => {
      heading = viewer.sheet.index;
      update();
    });
    viewer.addEventListener("zoomchange", update);
    update();
    const paging = count > 1 ? [previous, position, next] : [];
    const unmountBar = mountBar(stage, ...paging, button("-", () => stepZoom(viewer, -0.1), "Zoom out"), zoom, button("+", () => stepZoom(viewer, 0.1), "Zoom in"));
    return {
      summary: { format: workbook.info.format, units: count },
      dispose: () => {
        unmountBar();
        viewer.destroy();
        workbook.close();
      },
    };
  } catch (error) {
    workbook.close();
    throw error;
  }
}
