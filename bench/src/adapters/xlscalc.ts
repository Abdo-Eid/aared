import { abortError, type OpenedDoc } from "./types";

/**
 * xlscalc XLSX via its headless HTML renderer. `loadXlsx` parses +
 * evaluates in the root entry (React-free); `renderToHtml` shares style
 * composition with the `ExcelView` component, so this HTML is the renderer's
 * own output, not a re-implementation. The grid is drawn in full (no
 * virtualization) up to the package's own render limits.
 */
export async function openXlscalcXlsx(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  const { loadXlsx, renderToHtml } = await import("@xlscalc/xlsx-preview");
  if (signal.aborted) throw abortError();

  const doc = await loadXlsx(bytes.slice().buffer as ArrayBuffer);
  if (signal.aborted) throw abortError();

  const sheets = (doc.sheets as Array<{ name?: string }>).map(
    (sheet, index) => sheet.name ?? `Sheet ${index + 1}`,
  );
  if (sheets.length === 0) {
    throw new Error("This workbook contains no worksheets.");
  }

  const renderSheet = (index: number): void => {
    const html = renderToHtml(doc, index) as string;
    host.replaceChildren();
    const table = document.createElement("div");
    table.className = "bench-sheet-xlscalc";
    table.innerHTML = html;
    host.append(table);
  };

  renderSheet(0);
  if (signal.aborted) throw abortError();
  const ms = Math.round(performance.now() - t0);
  return {
    format: "xlsx",
    pageCount: sheets.length,
    ms,
    sheets,
    renderSheet,
    destroy: () => {
      host.replaceChildren();
    },
  };
}
