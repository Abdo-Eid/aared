import type { PPTXViewer } from "pptxviewjs";
import { abortError, type OpenedDoc } from "./types";

/**
 * Globals pptxviewjs assigns onto window at module scope. `FileReader` is the
 * damaging one: its replacement class has no `readAsArrayBuffer`, so after a
 * deck opens, uploads and docx-preview's JSZip break with an error that points
 * nowhere near a slide deck. Snapshot and restore around the import — the
 * assignments happen once, so restoring after is enough.
 */
const CLOBBERED_GLOBALS = [
  "FileReader",
  "Logger",
  "EnhancedLogger",
  "ProgressTracker",
  "ErrorHandler",
  "CoordinateTransform",
  "StringUtils",
] as const;

async function importViewer(): Promise<typeof import("pptxviewjs")> {
  const saved = new Map<string, unknown>();
  for (const name of CLOBBERED_GLOBALS) {
    if (name in globalThis) {
      saved.set(name, (globalThis as Record<string, unknown>)[name]);
    }
  }
  try {
    return await import("pptxviewjs");
  } finally {
    for (const [name, value] of saved) {
      (globalThis as Record<string, unknown>)[name] = value;
    }
  }
}

/**
 * Baseline PPTX renderer (pptxviewjs). Canvas output, so slides carry no
 * selectable text — structural, and the reason the aared column exists.
 * Renders serialize through one chain per viewer: overlapping renderSlide
 * calls race on the viewer's current-slide state.
 */
export async function openPptx(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  const { PPTXViewer: Viewer } = await importViewer();
  if (signal.aborted) throw abortError();

  const viewer: PPTXViewer = new Viewer({});
  try {
    await viewer.loadFile(bytes);
  } catch (error) {
    viewer.destroy();
    throw error;
  }
  if (signal.aborted) {
    viewer.destroy();
    throw abortError();
  }
  const slideCount = viewer.getSlideCount();
  if (slideCount <= 0) {
    viewer.destroy();
    throw new Error("This presentation has no slides to show.");
  }

  host.replaceChildren();
  const canvas = document.createElement("canvas");
  canvas.width = 1280;
  canvas.height = 720;
  canvas.className = "bench-slide-canvas";
  host.append(canvas);

  let chain: Promise<unknown> = Promise.resolve();
  const goToSlide = (index: number): void => {
    chain = chain
      .catch(() => {})
      .then(() => viewer.renderSlide(index, canvas));
  };
  goToSlide(0);
  await chain;
  if (signal.aborted) {
    viewer.destroy();
    throw abortError();
  }

  const ms = Math.round(performance.now() - t0);
  return {
    format: "pptx",
    pageCount: slideCount,
    ms,
    goToSlide,
    destroy: () => {
      viewer.destroy();
      host.replaceChildren();
    },
  };
}
