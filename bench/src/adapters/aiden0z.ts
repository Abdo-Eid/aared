import { abortError, type OpenedDoc } from "./types";

/**
 * aiden0z/pptx-renderer — browser-only OOXML→HTML/SVG DOM. The closest PPTX
 * proof that a DOM renderer is viable, so this is the PPTX column where the
 * RTL comparison actually bites. `PptxViewer.open` parses + renders in one
 * call (list mode); the pager steers via `goToSlide`.
 */
export async function openAiden0zPptx(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  const { PptxViewer, RECOMMENDED_ZIP_LIMITS } = await import(
    "@aiden0z/pptx-renderer"
  );
  if (signal.aborted) throw abortError();

  host.replaceChildren();
  const viewer = await PptxViewer.open(bytes.slice().buffer as ArrayBuffer, host, {
    zipLimits: RECOMMENDED_ZIP_LIMITS,
    listOptions: { windowed: false },
  });
  if (signal.aborted) {
    viewer.destroy();
    throw abortError();
  }
  const slideCount = viewer.slideCount;
  if (slideCount <= 0) {
    viewer.destroy();
    throw new Error("This presentation has no slides to show.");
  }

  const ms = Math.round(performance.now() - t0);
  return {
    format: "pptx",
    pageCount: slideCount,
    ms,
    goToSlide: (index: number) => {
      void viewer.goToSlide(index).catch(() => {});
    },
    destroy: () => {
      viewer.destroy();
      host.replaceChildren();
    },
  };
}
