import { open, render } from "aared/docx";
import { UnitController, mountDemoControls } from "../lib/demo-controls.js";
import type { DemoOutcome } from "../lib/demo-frame.js";

/**
 * Headless DOCX: open the document, render it in flow mode. Pass
 * `pagination: { mode: "layout" }` for physical pages instead. A document
 * is one unit, so its bar only zooms.
 */
export async function runHeadlessDocx(stage: HTMLElement, source: string | File): Promise<DemoOutcome> {
  const document = await open(source);
  try {
    const viewer = await render({ document, target: stage });
    const controller = new UnitController(viewer, () => Promise.resolve(viewer));
    const unmountControls = mountDemoControls(stage, controller, document.info.count);
    return {
      summary: { format: document.info.format, units: document.info.count },
      dispose: () => {
        unmountControls();
        viewer.destroy();
        document.close();
      },
    };
  } catch (error) {
    document.close();
    throw error;
  }
}
