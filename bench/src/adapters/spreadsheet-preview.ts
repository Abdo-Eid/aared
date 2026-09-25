import { abortError, type OpenedDoc } from "./types";

/**
 * DataGridXL spreadsheet-preview ("PDF.js for spreadsheets"): virtual
 * scrolling, multi-sheet tabs, frozen panes, cached formula values. The
 * viewer owns its own chrome (tabs, zoom) — the bench pager therefore only
 * steers sheets when the workbook exposes names, and `renderSheet` is a
 * no-op the viewer already handles internally. Community build: the badge
 * stays visible, no license key in this repo.
 */
export async function openSpreadsheetPreview(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  const { default: SpreadsheetPreview } = await import(
    "@datagridxl/spreadsheet-preview"
  );
  if (signal.aborted) throw abortError();

  host.replaceChildren();
  // The grid is virtualized and needs an explicit height — without it only
  // the chrome (tabs/zoom) paints and the cell area collapses to zero.
  const height = host.parentElement?.clientHeight || 640;
  const preview = new SpreadsheetPreview(host, {
    buffer: bytes.slice().buffer as ArrayBuffer,
    height,
  });
  await preview.ready;
  if (signal.aborted) {
    preview.destroy();
    throw abortError();
  }

  const workbook = preview.workbook as {
    sheets?: Array<{ name?: string }>;
    sheetNames?: string[];
  } | null;
  const sheets =
    workbook?.sheetNames ??
    workbook?.sheets?.map((sheet, index) => sheet.name ?? `Sheet ${index + 1}`) ??
    ["Workbook"];

  const ms = Math.round(performance.now() - t0);
  return {
    format: "xlsx",
    pageCount: sheets.length,
    ms,
    sheets,
    renderSheet: () => {},
    destroy: () => {
      preview.destroy();
      host.replaceChildren();
    },
  };
}
