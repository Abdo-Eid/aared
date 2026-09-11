import type { PDFDocumentProxy } from "unpdf/pdfjs";
import { abortError, isAbort, type OpenedDoc } from "./types";

const MAX_THUMBS = 24;
const THUMB_WIDTH = 96;

async function safeDestroy(doc: PDFDocumentProxy): Promise<void> {
  try {
    await doc.loadingTask.destroy();
  } catch {
    // Already destroyed, or never fully opened — nothing to do.
  }
}

/**
 * Baseline PDF renderer (unpdf, pdf.js inside). Pages render into stacked
 * canvases at stage width, so the document fits the way the product's viewer
 * fits it; the rail gets low-res data URLs rendered from the same canvases.
 */
export async function openPdf(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  const { getDocument } = await import("unpdf/pdfjs");
  if (signal.aborted) throw abortError();

  const doc: PDFDocumentProxy = await getDocument({
    data: bytes,
    useSystemFonts: true,
  }).promise;
  if (signal.aborted) {
    await safeDestroy(doc);
    throw abortError();
  }

  host.replaceChildren();
  const stageW = Math.max(host.clientWidth - 32, 320);
  const pageEls: HTMLElement[] = [];
  const thumbs: string[] = [];
  let firstPaint = 0;

  try {
    for (let n = 1; n <= doc.numPages; n += 1) {
      if (signal.aborted) throw abortError();
      const page = await doc.getPage(n);
      const unit = page.getViewport({ scale: 1 });
      const scale = Math.min(Math.max(stageW / unit.width, 0.4), 2.5);
      const viewport = page.getViewport({ scale });
      const canvas = document.createElement("canvas");
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("2d canvas context unavailable");
      await page.render({ canvas, viewport }).promise;

      const wrap = document.createElement("div");
      wrap.className = "bench-page";
      wrap.dataset.page = String(n);
      wrap.append(canvas);
      host.append(wrap);
      pageEls.push(wrap);
      if (n === 1) firstPaint = performance.now();

      if (n <= MAX_THUMBS) {
        const thumb = document.createElement("canvas");
        const ratio = canvas.height / canvas.width;
        thumb.width = THUMB_WIDTH;
        thumb.height = Math.max(1, Math.floor(THUMB_WIDTH * ratio));
        const tctx = thumb.getContext("2d");
        if (tctx) {
          tctx.drawImage(canvas, 0, 0, thumb.width, thumb.height);
          thumbs.push(thumb.toDataURL());
        }
      }
    }
  } catch (error) {
    if (!isAbort(error)) {
      await safeDestroy(doc);
    }
    throw error;
  }

  const ms = Math.round(firstPaint - t0);
  return {
    format: "pdf",
    pageCount: doc.numPages,
    ms,
    thumbs,
    scrollToPage: (page: number) => {
      const el = pageEls[page - 1];
      if (el) el.scrollIntoView({ block: "start", behavior: "auto" });
    },
    destroy: () => {
      void safeDestroy(doc);
      host.replaceChildren();
    },
  };
}
