import { isAbort, messageFor } from "@aared/viewer/core";
import { open, render } from "@aared/viewer/pptx";
import { button, label, mountBar, stepZoom, zoomLabel } from "../lib/demo-controls.js";
import type { DemoOutcome } from "../lib/demo-frame.js";

/**
 * Headless PPTX: open the deck, render the first slide, own the teardown. A
 * render handle is one slide, so paging is another `render()` into the same
 * stage, carrying the zoom along; each click cancels a render still on its
 * way (the API page's paging pattern). A non-presentation rejects with a
 * typed `unsupported-format` error.
 */
export async function runHeadlessPptx(stage: HTMLElement, source: string | File, signal?: AbortSignal): Promise<DemoOutcome> {
  // `attachDemo` always supplies the frame's signal; tolerate a direct caller
  // that does not, rather than aborting on a non-signal.
  const run = signal ?? new AbortController().signal;
  const deck = await open(source, { signal: run });
  try {
    const count = deck.info.count;
    let shown = await render({ slide: deck.slides.get(0), target: stage, signal: run });
    let index = 0;
    // Where the newest click is heading, so quick clicks add up.
    let heading = 0;
    let latest: AbortController | undefined;
    const failure = label();
    failure.setAttribute("role", "status");
    const report = (error: unknown): void => {
      if (!isAbort(error)) failure.textContent = messageFor(error);
    };

    const go = async (delta: number): Promise<void> => {
      const to = Math.min(Math.max(heading + delta, 0), count - 1);
      if (to === heading) return;
      heading = to;
      latest?.abort(); // a slower render still on its way never hands back a handle
      const mine = (latest = new AbortController());
      try {
        // The click's own cancel plus the frame's: a superseded run in this
        // stage must not paint over the newer one either.
        shown = await render({ slide: deck.slides.get(to), target: stage, zoom: shown.zoom, signal: AbortSignal.any([mine.signal, run]) });
        index = to;
        failure.textContent = "";
        update();
      } catch (error) {
        if (isAbort(error)) return; // a newer click took over
        if (latest === mine) heading = index; // this slide failed: step on from the one on screen
        throw error;
      }
    };
    const zoomBy = (delta: number): void => {
      stepZoom(shown, delta);
      update();
    };

    const previous = button("Previous", () => void go(-1).catch(report));
    const position = label();
    const next = button("Next", () => void go(1).catch(report));
    const zoom = label();
    const update = (): void => {
      position.textContent = `${index + 1} / ${count}`;
      previous.disabled = index <= 0;
      next.disabled = index >= count - 1;
      zoom.textContent = zoomLabel(shown.zoom);
    };
    update();
    const paging = count > 1 ? [previous, position, next] : [];
    const unmountBar = mountBar(stage, ...paging, button("-", () => zoomBy(-0.1), "Zoom out"), zoom, button("+", () => zoomBy(0.1), "Zoom in"), failure);
    return {
      summary: { format: deck.info.format, units: count },
      dispose: () => {
        latest?.abort();
        unmountBar();
        shown.destroy();
        deck.close();
      },
    };
  } catch (error) {
    deck.close();
    throw error;
  }
}
