import { abortError, type OpenedDoc } from "./types";

/**
 * mdurbar/pptx-viewer — lightweight TS parser→SVG slides, one dependency
 * (`fflate`). The weight floor of the PPTX comparison: parse once with
 * `loadPresentation`, paint with `renderSlideToElement`, release blob URLs
 * with `cleanup()`. Width is explicit (960): without it the renderer measures
 * the container rect, which is 0×0 in our empty host — a valid but invisible
 * zero-size SVG with a working pager.
 */
export async function openMdurbarPptx(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  const { loadPresentation, renderSlideToElement } = await import(
    "pptx-viewer"
  );
  if (signal.aborted) throw abortError();

  const presentation = await loadPresentation(bytes.slice());
  if (signal.aborted) {
    presentation.cleanup();
    throw abortError();
  }
  const slideCount = presentation.slides.length;
  if (slideCount <= 0) {
    presentation.cleanup();
    throw new Error("This presentation has no slides to show.");
  }

  const render = (index: number): void => {
    host.replaceChildren();
    renderSlideToElement(presentation, index, host, { width: 960 });
  };
  render(0);
  if (signal.aborted) {
    presentation.cleanup();
    throw abortError();
  }

  const ms = Math.round(performance.now() - t0);
  return {
    format: "pptx",
    pageCount: slideCount,
    ms,
    goToSlide: render,
    destroy: () => {
      presentation.cleanup();
      host.replaceChildren();
    },
  };
}
