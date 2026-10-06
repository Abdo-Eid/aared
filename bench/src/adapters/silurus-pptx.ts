import { abortError, type OpenedDoc } from "./types";

/**
 * silurus PPTX via the headless engine. Parse once (`PptxPresentation`),
 * render each slide into our own canvas — same composition the docs advertise
 * for custom UIs, and it gives a reliable `slideCount` where the opinionated
 * `PptxViewer` does not. Renders serialize through one chain: overlapping
 * `renderSlide` calls race on the shared engine.
 *
 * Canvas output, so slides carry no selectable text — the structural contrast
 * with the aared column, same as the pptxviewjs baseline.
 */
export async function openSilurusPptx(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  const { PptxPresentation } = await import("@silurus/ooxml/pptx");
  if (signal.aborted) throw abortError();

  const presentation = await PptxPresentation.load(
    bytes.slice().buffer as ArrayBuffer,
  );
  if (signal.aborted) {
    presentation.destroy();
    throw abortError();
  }
  const slideCount = presentation.slideCount as number;
  if (slideCount <= 0) {
    presentation.destroy();
    throw new Error("This presentation has no slides to show.");
  }

  host.replaceChildren();
  const canvas = document.createElement("canvas");
  canvas.className = "bench-slide-canvas";
  host.append(canvas);

  let chain: Promise<unknown> = Promise.resolve();
  const goToSlide = (index: number): void => {
    chain = chain
      .catch(() => {})
      .then(() => presentation.renderSlide(canvas, index, { width: 960 }));
  };
  goToSlide(0);
  await chain;
  if (signal.aborted) {
    presentation.destroy();
    throw abortError();
  }

  const ms = Math.round(performance.now() - t0);
  return {
    format: "pptx",
    pageCount: slideCount,
    ms,
    goToSlide,
    destroy: () => {
      presentation.destroy();
      host.replaceChildren();
    },
  };
}
