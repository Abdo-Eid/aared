import { abortError, type OpenedDoc } from "./types";

const MAX_ROWS = 200;
const MAX_COLS = 52;

function columnName(index: number): string {
  let name = "";
  let i = index;
  do {
    name = String.fromCharCode(65 + (i % 26)) + name;
    i = Math.floor(i / 26) - 1;
  } while (i >= 0);
  return name;
}

/**
 * Baseline XLSX renderer (SheetJS community build). Formatted text, not raw
 * values, so dates and currencies read as the sheet shows them. Parsed with
 * `cellStyles: true` so author column widths (`!cols[].wpx`) and hidden
 * rows/columns (`!cols[].hidden`, `!rows[].hidden`) survive: hidden
 * rows/columns collapse and merge spans count visible cells only, exactly
 * like Excel. Per-sheet RTL comes from the sheet's own `<sheetView
 * rightToLeft>` (SheetJS exposes only the workbook-level view, so the
 * baseline reads that one XML flag itself). Cell alignment follows Excel's
 * `general` rule from the SheetJS cell type — numbers at the end edge,
 * booleans/errors centered, strings at the start edge. Fills/fonts/borders
 * do not survive (SheetJS CE exposes only fill pattern/colors in `cell.s`)
 * — the grid reads plainer than Excel here.
 */

interface SheetDirCache {
  partForId: Map<string, string>;
  zip: import("jszip");
}

/** Resolve `xl/_rels/workbook.xml.rels` once so sheets map rId → part path. */
async function sheetParts(bytes: Uint8Array): Promise<SheetDirCache | null> {
  try {
    const { default: JSZip } = await import("jszip");
    const zip = await JSZip.loadAsync(bytes);
    const relsFile = zip.file("xl/_rels/workbook.xml.rels");
    if (!relsFile) return null;
    const rels = await relsFile.async("string");
    const partForId = new Map<string, string>();
    for (const match of rels.matchAll(/<Relationship[^>]*>/g)) {
      const tag = match[0];
      const id = tag.match(/\bId="([^"]+)"/)?.[1];
      const target = tag.match(/\bTarget="([^"]+)"/)?.[1];
      if (id !== undefined && target !== undefined && target.includes("worksheets/")) {
        partForId.set(id, `xl/${target.replace(/^\//, "")}`);
      }
    }
    return { partForId, zip };
  } catch {
    return null;
  }
}

/**
 * Per-sheet RTL from the sheet's own `<sheetView rightToLeft>`. Never throws:
 * on any probe failure the sheet stays LTR rather than breaking the baseline.
 */
async function sheetIsRtl(cache: SheetDirCache | null, relId: string | undefined): Promise<boolean> {
  try {
    if (cache === null || relId === undefined) return false;
    const part = cache.partForId.get(relId);
    if (part === undefined) return false;
    const file = cache.zip.file(part);
    if (!file) return false;
    const xml = await file.async("string");
    return /<sheetView\b[^>]*\brightToLeft=(["']?)1\1/.test(xml);
  } catch {
    return false;
  }
}
export async function openXlsx(
  bytes: Uint8Array,
  host: HTMLElement,
  signal: AbortSignal,
): Promise<OpenedDoc> {
  const t0 = performance.now();
  const XLSX = await import("xlsx");
  if (signal.aborted) throw abortError();

  const workbook = XLSX.read(bytes, { type: "array", cellStyles: true });
  const names: string[] = [...workbook.SheetNames];
  if (names.length === 0) {
    throw new Error("This workbook contains no worksheets.");
  }
  // SheetJS CE parses column widths/hidden flags with `cellStyles: true` but
  // not per-sheet RTL — probe the sheet XML for that one flag instead.
  const parts = await sheetParts(bytes);
  if (signal.aborted) throw abortError();
  const relIds = ((workbook.Workbook?.Sheets ?? []) as Array<{ id?: string }>).map((entry) => entry.id);
  const rtlCache = new Map<number, boolean>();

  const renderSheet = async (index: number): Promise<void> => {
    const name = names[index];
    if (name === undefined) return;
    const sheet = workbook.Sheets[name];
    let rtl = rtlCache.get(index);
    if (rtl === undefined) {
      rtl = await sheetIsRtl(parts, relIds[index]);
      rtlCache.set(index, rtl);
    }
    if (signal.aborted) throw abortError();
    host.replaceChildren();

    const table = document.createElement("table");
    table.className = "bench-sheet";
    table.dir = rtl ? "rtl" : "ltr";
    table.dataset.benchSheetDirection = rtl ? "rtl" : "ltr";
    if (!sheet || sheet["!ref"] === undefined) {
      const empty = document.createElement("tbody");
      const row = document.createElement("tr");
      const cell = document.createElement("td");
      cell.textContent = "This sheet is empty.";
      row.append(cell);
      empty.append(row);
      table.append(empty);
      host.append(table);
      return;
    }

    const used = XLSX.utils.decode_range(sheet["!ref"] as string);
    const r0 = used.s.r;
    const c0 = used.s.c;
    const r1 = Math.min(used.e.r, r0 + MAX_ROWS - 1);
    const c1 = Math.min(used.e.c, c0 + MAX_COLS - 1);

    const cols = (sheet["!cols"] ?? []) as Array<{ wpx?: number; hidden?: boolean } | undefined>;
    const rowsMeta = (sheet["!rows"] ?? []) as Array<{ hidden?: boolean } | undefined>;
    const isHiddenCol = (c: number): boolean => cols[c]?.hidden === true;
    const isHiddenRow = (r: number): boolean => rowsMeta[r]?.hidden === true;
    // Hidden rows/columns collapse exactly like Excel: headings jump C→E and
    // …6, 8… when D / row 7 are hidden.
    const visibleCols: number[] = [];
    for (let c = c0; c <= c1; c += 1) {
      if (!isHiddenCol(c)) visibleCols.push(c);
    }
    const visibleRows: number[] = [];
    for (let r = r0; r <= r1; r += 1) {
      if (!isHiddenRow(r)) visibleRows.push(r);
    }

    const colgroup = document.createElement("colgroup");
    const corner = document.createElement("col");
    colgroup.append(corner);
    for (const c of visibleCols) {
      const col = document.createElement("col");
      const wpx = cols[c]?.wpx;
      if (typeof wpx === "number" && wpx > 0) {
        col.style.width = `${Math.round(wpx)}px`;
      }
      colgroup.append(col);
    }
    table.append(colgroup);

    const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      range: { s: { r: r0, c: c0 }, e: { r: r1, c: c1 } },
      defval: "",
      raw: false,
    });

    const head = document.createElement("thead");
    const headRow = document.createElement("tr");
    headRow.append(document.createElement("th"));
    for (const c of visibleCols) {
      const th = document.createElement("th");
      th.textContent = columnName(c);
      th.scope = "col";
      headRow.append(th);
    }
    head.append(headRow);
    table.append(head);

    const merges = (sheet["!merges"] ?? []) as Array<{
      s: { r: number; c: number };
      e: { r: number; c: number };
    }>;
    // Merges with a hidden origin drop — their covered cells render normally.
    const liveMerges = merges.filter((m) => !isHiddenRow(m.s.r) && !isHiddenCol(m.s.c));
    const covered = new Set<string>();
    for (const merge of liveMerges) {
      for (let r = merge.s.r; r <= merge.e.r; r += 1) {
        for (let c = merge.s.c; c <= merge.e.c; c += 1) {
          if (r === merge.s.r && c === merge.s.c) continue;
          if (r < r0 || r > r1 || c < c0 || c > c1) continue;
          if (isHiddenRow(r) || isHiddenCol(c)) continue;
          covered.add(`${r}:${c}`);
        }
      }
    }

    /** Excel `general`: strings at the start edge, numbers at the end edge,
     * booleans/errors centered. SheetJS CE exposes the cell type but not the
     * authored alignment, so `general` is all the baseline can be faithful to. */
    const generalAlign = (r: number, c: number): string => {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })] as { t?: string } | undefined;
      if (cell?.t === "n") return rtl ? "left" : "right";
      if (cell?.t === "b" || cell?.t === "e") return "center";
      return rtl ? "right" : "left";
    };

    const body = document.createElement("tbody");
    for (const r of visibleRows) {
      const row = document.createElement("tr");
      const rowHead = document.createElement("th");
      rowHead.className = "rowhead";
      rowHead.textContent = String(r + 1);
      rowHead.scope = "row";
      row.append(rowHead);
      const values = matrix[r - r0] ?? [];
      for (const c of visibleCols) {
        if (covered.has(`${r}:${c}`)) continue;
        const td = document.createElement("td");
        const value = values[c - c0];
        td.textContent =
          typeof value === "string" ||
          typeof value === "number" ||
          typeof value === "boolean"
            ? String(value)
            : "";
        td.style.textAlign = generalAlign(r, c);
        const merge = liveMerges.find((m) => m.s.r === r && m.s.c === c);
        if (merge) {
          // Spans count visible cells only, so hidden rows/columns inside a
          // merge collapse exactly as Excel collapses them.
          const endRow = Math.min(merge.e.r, r1);
          const endCol = Math.min(merge.e.c, c1);
          let rowSpan = 0;
          for (let probe = r; probe <= endRow; probe++) {
            if (!isHiddenRow(probe)) rowSpan++;
          }
          let colSpan = 0;
          for (let probe = c; probe <= endCol; probe++) {
            if (!isHiddenCol(probe)) colSpan++;
          }
          if (rowSpan > 1) td.rowSpan = rowSpan;
          if (colSpan > 1) td.colSpan = colSpan;
        }
        row.append(td);
      }
      body.append(row);
    }
    table.append(body);
    host.append(table);

    if (used.e.r > r1 || used.e.c > c1) {
      const note = document.createElement("div");
      note.className = "bench-note";
      note.textContent = `Showing rows ${r0 + 1}–${r1 + 1} of ${used.e.r + 1}, columns ${columnName(c0)}–${columnName(c1)} of ${columnName(used.e.c)}.`;
      host.append(note);
    }
  };

  await renderSheet(0);
  if (signal.aborted) throw abortError();
  const ms = Math.round(performance.now() - t0);
  return {
    format: "xlsx",
    pageCount: names.length,
    ms,
    sheets: names,
    renderSheet,
    destroy: () => {
      host.replaceChildren();
    },
  };
}
