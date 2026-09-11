import { open, render } from "aared/pptx";
import type { DemoOutcome } from "../lib/demo-frame.js";

/**
 * Headless PPTX: open the deck, render the first slide, own the teardown.
 * A non-presentation rejects with a typed `unsupported-format` error.
 */
export async function runHeadlessPptx(stage: HTMLElement, source: string | File): Promise<DemoOutcome> {
  const deck = await open(source);
  try {
    const viewer = await render({ slide: deck.slides.get(0), target: stage });
    return {
      summary: { format: deck.info.format, units: deck.info.count },
      dispose: () => {
        viewer.destroy();
        deck.close();
      },
    };
  } catch (error) {
    deck.close();
    throw error;
  }
}
