import { preview } from "aared/preview";
import type { DemoOutcome } from "../lib/demo-frame.js";

/**
 * One call for an arbitrary file: detection runs in the Worker, only the
 * detected format loads, and the handle owns navigation and teardown.
 */
export async function runFacadeDemo(stage: HTMLElement, source: string | File): Promise<DemoOutcome> {
  const viewer = await preview(source, stage, {
    appearance: { theme: "light" },
  });
  return {
    summary: { format: viewer.format, units: viewer.info.count },
    dispose: () => viewer.destroy(),
  };
}
