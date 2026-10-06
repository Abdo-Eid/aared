import type { DocxOpened, OpenAdapter, PptxOpened, XlsxOpened } from "./types";

/**
 * aared through its per-format entrypoints (`aared/pptx`, `aared/xlsx`,
 * `aared/docx`), shaped like the baseline adapters so the measurement page
 * can time every renderer through one contract: bytes in, first paint out,
 * then move between units, then destroy.
 *
 * Used by the measurement page only — the bench's own aared column is the
 * `usePreview()` hook (`AaredViewerCard`). `ms` covers the dynamic import,
 * the Worker parse, and the first paint, the same span the baselines report.
 */

export const openAaredPptx: OpenAdapter = async (bytes, host, signal): Promise<PptxOpened> => {
  const t0 = performance.now();
  const { preview } = await import("aared/pptx");
  const viewer = await preview({ source: bytes, target: host, signal });
  return {
    format: "pptx",
    pageCount: viewer.info.count,
    ms: Math.round(performance.now() - t0),
    goToSlide: (index) => viewer.goTo(index),
    destroy: () => viewer.destroy(),
  };
};

export const openAaredXlsx: OpenAdapter = async (bytes, host, signal): Promise<XlsxOpened> => {
  const t0 = performance.now();
  const { preview } = await import("aared/xlsx");
  const viewer = await preview({
    source: bytes, target: host, signal,
    appearance: { spreadsheet: { sheetTabs: false } },
  });
  return {
    format: "xlsx",
    pageCount: viewer.info.count,
    ms: Math.round(performance.now() - t0),
    sheets: [...viewer.info.titles],
    renderSheet: (index) => viewer.goTo(index),
    destroy: () => viewer.destroy(),
  };
};

export const openAaredDocx: OpenAdapter = async (bytes, host, signal): Promise<DocxOpened> => {
  const t0 = performance.now();
  const { preview } = await import("aared/docx");
  const viewer = await preview({ source: bytes, target: host, signal, pagination: { mode: "layout" } });
  const docx = viewer.docx;
  // aared reports inches; the bench's pager uses CSS pixels.
  const pages = docx.pages.map((page) => ({ width: page.width * 96, height: page.height * 96 }));
  return {
    format: "docx",
    pageCount: docx.pageCount,
    ms: Math.round(performance.now() - t0),
    outline: docx.headings.map((heading) => ({ level: heading.level, text: heading.title })),
    paperWidth: pages.length > 0 ? Math.max(...pages.map((page) => page.width)) : 0,
    pages,
    destroy: () => viewer.destroy(),
  };
};
