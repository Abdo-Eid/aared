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

export const openAaredPptx: OpenAdapter = async (bytes, host): Promise<PptxOpened> => {
  const t0 = performance.now();
  const { open, render } = await import("aared/pptx");
  const deck = await open(bytes.slice().buffer);
  host.replaceChildren();
  let handle = await render({ slide: deck.slides.get(0), target: host });
  return {
    format: "pptx",
    pageCount: deck.info.count,
    ms: Math.round(performance.now() - t0),
    goToSlide: async (index) => {
      handle.destroy();
      handle = await render({ slide: deck.slides.get(index), target: host });
    },
    destroy: () => {
      handle.destroy();
      deck.close();
      host.replaceChildren();
    },
  };
};

export const openAaredXlsx: OpenAdapter = async (bytes, host): Promise<XlsxOpened> => {
  const t0 = performance.now();
  const { open, render } = await import("aared/xlsx");
  const workbook = await open(bytes.slice().buffer);
  host.replaceChildren();
  const handle = await render({ sheet: workbook.sheets.get(0), target: host });
  return {
    format: "xlsx",
    pageCount: workbook.info.count,
    ms: Math.round(performance.now() - t0),
    sheets: [...workbook.info.titles],
    renderSheet: (index) => handle.goToSheet(index),
    destroy: () => {
      handle.destroy();
      workbook.close();
      host.replaceChildren();
    },
  };
};

export const openAaredDocx: OpenAdapter = async (bytes, host): Promise<DocxOpened> => {
  const t0 = performance.now();
  const { open, render } = await import("aared/docx");
  const document = await open(bytes.slice().buffer);
  host.replaceChildren();
  const handle = await render({ document, target: host, pagination: { mode: "layout" } });
  const pages = handle.pages.map((page) => ({ width: page.width, height: page.height }));
  return {
    format: "docx",
    pageCount: handle.pageCount,
    ms: Math.round(performance.now() - t0),
    outline: document.headings.list().map((heading) => ({ level: heading.level, text: heading.title })),
    paperWidth: pages.length > 0 ? Math.max(...pages.map((page) => page.width)) : 0,
    pages,
    destroy: () => {
      handle.destroy();
      document.close();
      host.replaceChildren();
    },
  };
};
