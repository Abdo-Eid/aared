import { abortError, type OpenedDoc, type OutlineEntry } from "./types";
import { maxSectionWidthPx, scrapeDocxOutline } from "./docx";

/**
 * omni-doc-viewer — one API over proven engines (docx-preview / SheetJS /
 * pptx-preview + emf-converter), per-format lazy load. The bench wires it
 * once per format so the "stitch vs shared core" comparison is visual, not
 * hypothetical. `renderDocument` returns `pages` for paged formats (slides /
 * sections); XLSX navigates by its own sheet tabs, so the bench chrome gets
 * a single "Workbook" tab and a no-op `renderSheet`.
 */
async function openOmni(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
  format: "pptx" | "xlsx" | "docx",
): Promise<OpenedDoc> {
  const t0 = performance.now();
  const { renderDocument } = await import("omni-doc-viewer");
  if (signal.aborted) throw abortError();

  host.replaceChildren();
  // Section widths parse concurrently with the render so the fit width
  // costs no extra wall time (only awaited on the docx path below).
  const sectionWidth =
    format === "docx" ? maxSectionWidthPx(bytes) : Promise.resolve(0);
  const handle = await renderDocument({
    container: host,
    source: bytes.slice(),
  });
  if (signal.aborted) {
    handle.destroy();
    throw abortError();
  }

  const pages = handle.pages ?? [];
  const ms = Math.round(performance.now() - t0);
  const destroy = (): void => {
    handle.destroy();
    host.replaceChildren();
  };

  if (format === "pptx") {
    const pageCount = Math.max(pages.length, 1);
    return {
      format: "pptx",
      pageCount,
      ms,
      goToSlide: (index: number) => {
        pages[index]?.scrollIntoView({ block: "start" });
      },
      destroy,
    };
  }
  if (format === "xlsx") {
    return {
      format: "xlsx",
      pageCount: 1,
      ms,
      sheets: ["Workbook"],
      renderSheet: () => {},
      destroy,
    };
  }
  const outline: OutlineEntry[] = [];
  scrapeDocxOutline(host, outline);
  const first = pages[0];
  // Fit the document's true widest page (file sections), not just the
  // widest rendered one: a renderer that flows a landscape section onto
  // portrait-width pages must still open at the aared column's zoom.
  const paperWidth = Math.max(
    first instanceof HTMLElement ? first.offsetWidth : 0,
    await sectionWidth,
  );
  return {
    format: "docx",
    pageCount: Math.max(pages.length, 1),
    pages: pages.map((element) => ({
      width: element.offsetWidth,
      height: element.offsetHeight,
    })),
    ms,
    outline,
    paperWidth,
    destroy,
  };
}

export function openOmniPptx(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  return openOmni(bytes, host, signal, "pptx");
}

export function openOmniXlsx(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  return openOmni(bytes, host, signal, "xlsx");
}

export function openOmniDocx(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  return openOmni(bytes, host, signal, "docx");
}
