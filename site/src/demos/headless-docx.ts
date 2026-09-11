import { open, render } from "aared/docx";
import type { DemoOutcome } from "../lib/demo-frame.js";

/**
 * Headless DOCX: open the document, render it in flow mode. Pass
 * `pagination: { mode: "layout" }` for physical pages instead.
 */
export async function runHeadlessDocx(stage: HTMLElement, source: string | File): Promise<DemoOutcome> {
  const document = await open(source);
  try {
    const viewer = await render({ document, target: stage });
    return {
      summary: { format: document.info.format, units: document.info.count },
      dispose: () => {
        viewer.destroy();
        document.close();
      },
    };
  } catch (error) {
    document.close();
    throw error;
  }
}
