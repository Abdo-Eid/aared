import { open, render } from "@aared/viewer/docx";
import { button, label, mountBar, stepZoom, zoomLabel } from "../lib/demo-controls.js";
import type { DemoOutcome } from "../lib/demo-frame.js";

/**
 * Headless DOCX: open the document, render it in flow mode. Pass
 * `pagination: { mode: "layout" }` for physical pages instead. A document
 * is one unit, so its bar only zooms.
 */
export async function runHeadlessDocx(stage: HTMLElement, source: string | File, signal?: AbortSignal): Promise<DemoOutcome> {
  // `attachDemo` always supplies the frame's signal; tolerate a direct caller
  // that does not, rather than aborting on a non-signal.
  const run = signal ?? new AbortController().signal;
  const doc = await open(source, { signal: run });
  try {
    const viewer = await render({ document: doc, target: stage, signal: run });
    const zoom = label();
    const update = (): void => {
      zoom.textContent = zoomLabel(viewer.zoom);
    };
    viewer.addEventListener("zoomchange", update);
    update();
    const unmountBar = mountBar(stage, button("-", () => stepZoom(viewer, -0.1), "Zoom out"), zoom, button("+", () => stepZoom(viewer, 0.1), "Zoom in"));
    return {
      summary: { format: doc.info.format, units: doc.info.count },
      dispose: () => {
        unmountBar();
        viewer.destroy();
        doc.close();
      },
    };
  } catch (error) {
    doc.close();
    throw error;
  }
}
