import { open, render } from "aared/pptx";
import { UnitController, mountDemoControls } from "../lib/demo-controls.js";
import type { DemoOutcome } from "../lib/demo-frame.js";

/**
 * Headless PPTX: open the deck, render the first slide, own the teardown.
 * Moving to another slide is another `render()` (the paging bar does that,
 * carrying the zoom along). A non-presentation rejects with a typed
 * `unsupported-format` error.
 */
export async function runHeadlessPptx(stage: HTMLElement, source: string | File): Promise<DemoOutcome> {
  const deck = await open(source);
  try {
    const first = await render({ slide: deck.slides.get(0), target: stage });
    const slides = new UnitController(first, (index, zoom) => render({ slide: deck.slides.get(index), target: stage, zoom }));
    const unmountControls = mountDemoControls(stage, slides, deck.info.count);
    return {
      summary: { format: deck.info.format, units: deck.info.count },
      dispose: () => {
        unmountControls();
        slides.destroy();
        deck.close();
      },
    };
  } catch (error) {
    deck.close();
    throw error;
  }
}
