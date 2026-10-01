import { abortError, type OpenedDoc } from "./types";

/**
 * pptx-browser — zero-dependency Canvas 2D renderer (native
 * `DecompressionStream` + `DOMParser`). The weight floor for a PPTX
 * previewer, and a second canvas column beside pptxviewjs: slides carry no
 * DOM text, so RTL correctness rests entirely on the canvas's own shaping.
 * Renders serialize through one chain: overlapping `renderSlide` calls on
 * one canvas would interleave their draws.
 */
export async function openPptxBrowser(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  const { PptxRenderer } = await import("pptx-browser");
  if (signal.aborted) throw abortError();

  const renderer = new PptxRenderer();
  try {
    await renderer.load(bytes);
  } catch (error) {
    renderer.destroy();
    throw error;
  }
  if (signal.aborted) {
    renderer.destroy();
    throw abortError();
  }
  const slideCount = renderer.slideCount;
  if (slideCount <= 0) {
    renderer.destroy();
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
      .then(() => renderer.renderSlide(index, canvas, 1280));
  };
  goToSlide(0);
  await chain;
  if (signal.aborted) {
    renderer.destroy();
    throw abortError();
  }

  const ms = Math.round(performance.now() - t0);
  return {
    format: "pptx",
    pageCount: slideCount,
    ms,
    goToSlide,
    destroy: () => {
      renderer.destroy();
      host.replaceChildren();
    },
  };
}
