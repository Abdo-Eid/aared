import { preview } from "@aared/viewer/preview";
import { viewerToolbar } from "../lib/demo-controls.js";
import type { DemoOutcome } from "../lib/demo-frame.js";

/**
 * One call for an arbitrary file: detection runs in the Worker, only the
 * detected format loads, and the handle owns navigation and teardown. A
 * workbook shows its sheet tabs by default. The bar is the handle's own
 * `next()`, `zoomIn()`, … behind buttons.
 */
export async function runFacadeDemo(stage: HTMLElement, source: string | File, signal: AbortSignal): Promise<DemoOutcome> {
  // `signal` is page plumbing: a newer run in this stage cancels this one.
  const viewer = await preview({ source, target: stage, signal });
  const unmountToolbar = viewerToolbar(stage, viewer);
  return {
    summary: { format: viewer.format, units: viewer.info.count },
    dispose: () => {
      unmountToolbar();
      viewer.destroy();
    },
  };
}
