/**
 * XLSX entrypoint: open workbook documents in a Worker, read sheet windows,
 * and paint resolved `SheetModel`s to a virtualized DOM grid — merges as
 * spanning cells, frozen panes as sticky regions, sheet-level RTL flips the
 * whole grid. Never canvas. See docs/ARCHITECTURE.md.
 *
 * Structure mirrors `aared/pptx`: `open()` with a format check, a document
 * holding a `WorkerSession`, collections implementing core's
 * `Collection<TItem, TKey>`, lightweight handles, a capabilities constant,
 * `search()`, and an input-object `render()` returning an `EventTarget`
 * controller. The grid painter from the previous revision is kept intact
 * below as the internal `paintSheet()`; only its options plumbing changed
 * (headers default on, sheet tabs default off, the benchmark-only `frame`
 * chrome is gone, gridlines follow `appearance.spreadsheet.gridlines`).
 *
 * Adaptive content (`contentMode` defaults to `"adaptive"`): the content
 * follows the resolved theme — dark theme adapts default surfaces, light
 * theme renders authored. `source` is available as an explicit fidelity
 * escape hatch. Provenance survives to the model: `core/xlsx/styles.ts`
 * `resolveFill` returns `undefined` for `patternType="none"`, missing, or
 * non-solid fills (and `cellStyleFor` only assigns `fill`/`color` when a
 * resolved value exists), while `core/xlsx/sheet.ts` omits `cell.style`
 * entirely when the resolved style object is empty — so an absent
 * `style.fill`/`style.color` reliably means "automatic/default". Adaptive
 * dark therefore re-themes only those defaults (host surface shows through
 * transparent fills; default text follows the theme) while every authored
 * fill, font color, border, and conditional overlay paints verbatim. Never
 * a global inversion or filter.
 *
 * Virtualization stays DOM-only: rows outside the scroll window collapse
 * into spacers. Renderer-driven Worker re-reads on scroll (true
 * range-windowed transfer for huge sheets) are OUT OF SCOPE here — the
 * grid paints whatever model it was handed, and `SheetHandle.read(range)`
 * exists for callers that want windows explicitly.
 *
 * Keyboard/selection notes: the grid host holds a fallback tab stop
 * (`tabindex="0"`) so keyboard focus is never lost when the selected cell
 * is virtualized away. The selected cell receives `tabindex="0"` when it
 * is in the DOM (roving tabindex pattern); all other cells carry
 * `tabindex="-1"`. A move that scrolls re-renders the row window and
 * destroys the focused cell, so selection re-application restores the tab
 * stop (selected cell, else the host) when focus fell through to `<body>`;
 * without that, the second arrow key lands outside the grid and navigation
 * dies after one move. Selection is single-cell (Shift never
 * extends; documented, not deferred). Sheet tabs
 * list VISIBLE sheets only; `goToSheet()` still resolves hidden sheets by
 * name or index. Hyperlink clicks never navigate: they `preventDefault()`
 * and emit `linkactivate`.
 */

import type {
  CellModel,
  CellRange,
  ChartGraphicPrimitive,
  SheetChartModel,
  SheetDrawingModel,
  SheetModel,
  SheetPictureModel,
  SheetRangeModel,
} from "../core/model.js";
import type {
  Appearance,
  Collection,
  ContentMode,
  DocumentInfo,
  GridlineMode,
  MediaCollection as MediaCollectionShape,
  MediaOptions,
  MediaResolver,
  OpenOptions,
  ReadOptions,
  Source,
  ViewerTheme,
  ZoomValue,
} from "../core/index.js";
import { AaredError } from "../core/index.js";
import { WorkerSession } from "../core/session.js";
import type { UnitRange } from "../core/session.js";
import { cellAddress, parseA1Range } from "../core/xlsx/range.js";

/** Document metadata for a workbook; `format` is narrowed to `"xlsx"`. */
export interface XlsxInfo extends DocumentInfo {
  format: "xlsx";
}

/** Lightweight sheet identity from `open()` indexing — no sheet is parsed yet. */
export interface SheetInfo {
  index: number;
  name: string;
  state: "visible" | "hidden" | "veryHidden";
}

/** One matching cell from `workbook.search()`. */
export interface XlsxSearchMatch {
  sheet: string;
  cell: string;
  text: string;
}

/** What a readonly spreadsheet preview can and cannot do. */
export interface XlsxCapabilities {
  textSelection: boolean;
  media: boolean;
  spreadsheet: {
    rangeRead: boolean;
    cellNavigation: boolean;
    selection: boolean;
    formulas: "cached-only";
  };
  editing: false;
}

/**
 * Validate a caller-supplied numeric range (already-numeric form of an A1
 * reference). Bounds mirror `parseA1Range` and the worker's `checkRange`:
 * zero-based, rows 0..1048575, columns 0..16383, start at or before end.
 */
function validateCellRange(range: CellRange): CellRange {
  const invalid = (detail: string): AaredError =>
    new AaredError(
      "invalid-range",
      `aared/xlsx: invalid cell range (${detail}); use "A1:M200" or a single cell like "D14"`,
    );
  const raw = range as unknown as { start?: unknown; end?: unknown };
  if (typeof raw !== "object" || raw === null) throw invalid("not an object");
  const point = (value: unknown, label: string): { row: number; col: number } => {
    const candidate = value as unknown as { row?: unknown; col?: unknown } | null;
    if (typeof candidate !== "object" || candidate === null) throw invalid(`${label} is not a cell`);
    if (!Number.isInteger(candidate.row) || !Number.isInteger(candidate.col)) {
      throw invalid(`${label} coordinates must be integers`);
    }
    const row = candidate.row as number;
    const col = candidate.col as number;
    if (row < 0 || row > 1048575 || col < 0 || col > 16383) {
      throw invalid(`${label} is outside the worksheet grid`);
    }
    return { row, col };
  };
  const start = point(raw.start, "start");
  const end = point(raw.end, "end");
  if (start.row > end.row || start.col > end.col) throw invalid("start must be at or before end");
  return { start, end };
}

/**
 * A single sheet: a lightweight handle, not a parsed model. `read()`
 * parses the sheet in the Worker on demand (index first, content second).
 */
export class SheetHandle {
  constructor(
    readonly info: SheetInfo,
    private readonly session: WorkerSession,
    private readonly owner: XlsxDocument,
  ) {}

  /**
   * Read the sheet. No range (or `null`) resolves the whole sheet and wraps
   * it as a `SheetRangeModel` covering the full extent; a string range is
   * parsed as A1 first; a numeric `CellRange` is validated here before
   * anything is sent to the Worker.
   */
  async read(range?: string | CellRange | null, options?: ReadOptions): Promise<SheetRangeModel> {
    const signal = options?.signal;
    if (range === undefined || range === null) {
      const sheet = await this.session.unit<SheetModel>(
        this.info.index,
        signal === undefined ? undefined : { signal },
      );
      return {
        ...sheet,
        range: {
          start: { row: 0, col: 0 },
          end: {
            row: Math.max(sheet.extent.rows - 1, 0),
            col: Math.max(sheet.extent.cols - 1, 0),
          },
        },
      };
    }
    const parsed = typeof range === "string" ? parseA1Range(range) : validateCellRange(range);
    const requestRange: UnitRange = {
      start: { ...parsed.start },
      end: { ...parsed.end },
    };
    const windowed = await this.session.unit<SheetRangeModel>(
      this.info.index,
      signal === undefined ? { range: requestRange } : { range: requestRange, signal },
    );
    windowed.range = windowed.range ?? {
      start: { ...requestRange.start },
      end: { ...requestRange.end },
    };
    return windowed;
  }

  /** Drawing image bytes for this sheet's `mediaId`s, via the owning session. */
  media(mediaId: string, options?: MediaOptions): Promise<ArrayBuffer> {
    return this.session.media(mediaId, options);
  }

  /** @internal Used by the viewer for tab switching. Not part of the documented public API. */
  sheetDocument(): XlsxDocument {
    return this.owner;
  }
}

/** `workbook.sheets`: addressable sheets in tab order, hidden included. */
export class SheetCollection implements Collection<SheetHandle, number | string> {
  private readonly handles: SheetHandle[];

  constructor(sheets: readonly SheetInfo[], session: WorkerSession, owner: XlsxDocument) {
    this.handles = sheets.map((info) => new SheetHandle(info, session, owner));
  }

  list(): readonly SheetHandle[] {
    return this.handles;
  }

  get(key: number | string): SheetHandle {
    if (typeof key === "number") {
      const handle = this.handles[key];
      if (!Number.isInteger(key) || handle === undefined) {
        throw new AaredError(
          "not-found",
          `aared/xlsx: sheet index ${String(key)} is outside this ${this.handles.length}-sheet workbook`,
        );
      }
      return handle;
    }
    const exact = this.handles.find((handle) => handle.info.name === key);
    if (exact !== undefined) return exact;
    const folded = this.handles.find((handle) => handle.info.name.toLowerCase() === key.toLowerCase());
    if (folded !== undefined) return folded;
    throw new AaredError(
      "not-found",
      `aared/xlsx: no sheet named ${JSON.stringify(key)} in this ${this.handles.length}-sheet workbook`,
    );
  }
}

class XlsxMediaCollection implements MediaCollectionShape {
  constructor(private readonly session: WorkerSession) {}

  get(mediaId: string, options?: MediaOptions): Promise<ArrayBuffer> {
    const signal = options?.signal;
    return this.session.media(mediaId, signal === undefined ? undefined : { signal });
  }
}

/** An opened workbook: collections, search, media, and lifecycle. */
export class XlsxDocument {
  readonly info: XlsxInfo;
  readonly sheets: SheetCollection;
  readonly media: MediaCollectionShape;
  readonly capabilities: XlsxCapabilities = {
    textSelection: true,
    media: true,
    spreadsheet: { rangeRead: true, cellNavigation: true, selection: true, formulas: "cached-only" },
    editing: false,
  };

  constructor(private readonly session: WorkerSession) {
    const info = session.info;
    this.info = { ...info, format: "xlsx" };
    const states = info.sheetStates;
    this.sheets = new SheetCollection(
      info.titles.map((name, index) => ({
        index,
        name,
        state: states?.[index] ?? "visible",
      })),
      session,
      this,
    );
    this.media = new XlsxMediaCollection(session);
  }

  /**
   * Case-insensitive substring search across every sheet, hidden sheets
   * included. Each sheet is fully read in tab order (aborting stops at the
   * next sheet); every matching cell yields one match. A cell matches on its
   * formatted text, falling back to a string cached value when the formatted
   * text is empty.
   */
  async search(query: string, options?: ReadOptions): Promise<XlsxSearchMatch[]> {
    if (query === "") return [];
    const needle = query.toLowerCase();
    const matches: XlsxSearchMatch[] = [];
    for (const handle of this.sheets.list()) {
      const signal = options?.signal;
      const model = await handle.read(null, signal === undefined ? undefined : { signal });
      for (const row of model.rows) {
        for (const { col, cell } of row.cells) {
          const haystack = cell.formatted !== ""
            ? cell.formatted
            : typeof cell.value === "string"
              ? cell.value
              : "";
          if (haystack !== "" && haystack.toLowerCase().includes(needle)) {
            matches.push({ sheet: handle.info.name, cell: cellAddress(row.index, col), text: haystack.slice(0, 120) });
          }
        }
      }
    }
    return matches;
  }

  /** Terminate the Worker session and reject pending reads. */
  close(): void {
    this.session.close();
  }
}

/**
 * Open an `.xlsx` file. Fails with an `unsupported-format` `AaredError`
 * when the file is not actually a spreadsheet workbook.
 */
export async function open(source: Source, options?: OpenOptions): Promise<XlsxDocument> {
  const session = await WorkerSession.open("aared/xlsx", source, options);
  if (session.info.format !== "xlsx") {
    const detected = session.info.format;
    session.close();
    throw new AaredError(
      "unsupported-format",
      `aared/xlsx: expected a spreadsheet workbook; detected "${detected}" — open it with aared/${detected} instead`,
    );
  }
  return new XlsxDocument(session);
}

/** Viewer-only appearance for XLSX. Headers/tabs/selection/gridlines are chrome; content follows `contentMode`. */
export interface XlsxAppearance {
  theme?: ViewerTheme;
  contentMode?: ContentMode;
  spreadsheet?: Appearance["spreadsheet"];
}

/** `render()` input: a sheet handle (read on demand) or an already-resolved model. */
export interface XlsxRenderInput {
  sheet: SheetHandle | SheetModel | SheetRangeModel;
  target: HTMLElement;
  appearance?: XlsxAppearance | Appearance;
  mediaResolver?: MediaResolver;
}

/** Controller returned by `render()`: zoom, navigation, selection, and lifecycle. */
export interface XlsxRenderHandle extends EventTarget {
  /** Disconnect scroll virtualization, abort pending sheet switches, release image object URLs. */
  destroy(): void;
  /** Change zoom; emits `zoomchange`. Numbers scale via the CSS `zoom` property; fit modes clear it. */
  setZoom(zoom: ZoomValue): void;
  /** The active zoom. */
  readonly zoom: ZoomValue;
  /** Select a single cell (no ranges), scroll it into view, and focus it. */
  goTo(cell: string): Promise<void>;
  /** Scroll a cell or range into view without changing the selection. */
  scrollTo(target: { cell?: string; range?: string; align?: "start" | "center" | "end" }): void;
  /** Select a single cell and DOM-focus it with `preventScroll` (no scrolling). */
  focus(cell: string): void;
  /** Switch the visible sheet by name (hidden sheets included) or index; emits `sheetchange`. */
  goToSheet(name: string | number): Promise<void>;
  /** The currently visible sheet. */
  readonly sheet: SheetInfo;
  /** The single selected cell, or `null`. */
  readonly selection: { address: string; row: number; col: number } | null;
}

/** CSS pixels per inch: the shared layout baseline with the PPTX renderer. */
const PIXELS_PER_INCH = 96;
const GRIDLINE_COLOR = "#E7E9EE";
const GRIDLINE_DARK = "#3B3B45";
const ROW_OVERSCAN = 8;
/** Heading chrome dimensions. Viewer chrome, not authored geometry. */
const HEADER_COL_WIDTH_PX = 40;
const HEADER_ROW_HEIGHT_PX = 22;
const HEADER_FILL = "#F4F4F7";
const HEADER_COLOR = "#5C5C64";
const HEADER_FILL_DARK = "#2A2B33";
const HEADER_COLOR_DARK = "#C9CAD1";
const NOTICE_COLOR = "#5B6470";
const NOTICE_COLOR_DARK = "#A7AAB3";
const SURFACE_DARK = "#1E1F26";
const TEXT_DARK = "#E9E9EC";
const SELECTION_LIGHT = "#0B57D0";
const SELECTION_DARK = "#8AB4F8";

/** Zero-based column index to Excel letters: 0 → A … 25 → Z, 26 → AA. */
function columnName(index: number): string {
  let name = "";
  let i = index;
  do {
    name = String.fromCharCode(65 + (i % 26)) + name;
    i = Math.floor(i / 26) - 1;
  } while (i >= 0);
  return name;
}

function pixels(inches: number): string {
  return `${inches * PIXELS_PER_INCH}px`;
}

function borderWidth(style: string | undefined): string {
  switch (style) {
    case "medium": return "2px";
    case "thick": return "3px";
    default: return "1px";
  }
}

function borderStyle(style: string | undefined): string {
  switch (style) {
    case "dashed": case "mediumDashed": return "dashed";
    case "dotted": case "mediumDashDot": case "mediumDashDotDot": case "slantDashDot": return "dotted";
    case "double": return "double";
    default: return "solid";
  }
}

function textAlign(cell: CellModel, sheetRtl: boolean): string {
  const align = cell.style?.align ?? "general";
  if (align !== "general") return align;
  // Excel's `general`: strings at the start edge, numbers at the end edge,
  // booleans and errors centered. Logical values keep DOM order intact.
  if (typeof cell.value === "number") return sheetRtl ? "left" : "right";
  if (typeof cell.value === "boolean" || cell.type === "error") return "center";
  return sheetRtl ? "right" : "left";
}

/** Resolved paint parameters: chrome theme plus content defaults for the active content mode. */
interface GridPaint {
  showHeaders: boolean;
  showGridlines: boolean;
  gridlineColor: string;
  /** Default cell text when neither the style nor a conditional overlay names a color. */
  defaultColor: string;
  /** Opaque fill stamped behind sticky frozen cells whose own fill is transparent. */
  frozenFill: string;
  headerFill: string;
  headerColor: string;
  /** Chrome line/text colors for notices and drawing placeholders. */
  chromeLine: string;
  noticeColor: string;
  hostBackground: string;
  dark: boolean;
  contentMode: "source" | "adaptive";
}

function appendCellContent(cellElement: HTMLElement, cell: CellModel): void {
  if (cell.conditional?.dataBar !== undefined) {
    const bar = document.createElement("div");
    bar.dataset.aaredElement = "data-bar";
    bar.setAttribute("aria-hidden", "true");
    bar.style.position = "absolute";
    bar.style.insetBlock = "0";
    bar.style.insetInlineStart = "0";
    bar.style.width = `${Math.min(Math.max(cell.conditional.dataBar.ratio, 0), 1) * 100}%`;
    bar.style.backgroundColor = cell.conditional.dataBar.color;
    bar.style.opacity = "0.45";
    bar.style.pointerEvents = "none";
    cellElement.append(bar);
  }
  const text = document.createElement("span");
  text.dataset.aaredElement = "cell-text";
  text.style.position = "relative";
  text.style.zIndex = "1";
  if (cell.hyperlink === undefined) {
    text.textContent = cell.formatted;
  } else {
    const link = document.createElement("a");
    link.dataset.aaredElement = "cell-link";
    link.href = cell.hyperlink.target;
    link.textContent = cell.formatted === "" ? (cell.hyperlink.display ?? cell.hyperlink.target) : cell.formatted;
    // Native print shows no underline: link text keeps the cell's own color
    // (blue only when the cell carries no explicit color, as in Excel).
    // Clicks never navigate: the viewer intercepts them and emits
    // `linkactivate` instead (see XlsxViewer).
    link.style.color = cell.style?.color ?? "#1155CC";
    link.style.textDecoration = "none";
    text.append(link);
  }
  cellElement.append(text);
}

function paintCell(cellElement: HTMLElement, cell: CellModel, sheetRtl: boolean, paint: GridPaint): void {
  const style = cell.style;
  cellElement.style.padding = "1px 4px";
  // Default gridlines are content chrome resolved per gridline mode; an
  // authored border always wins below. Hidden gridlines stay transparent
  // (never `none`) so authored and default borders share one box metric.
  cellElement.style.borderTop = "1px solid transparent";
  cellElement.style.borderLeft = "1px solid transparent";
  cellElement.style.borderRight = paint.showGridlines ? `1px solid ${paint.gridlineColor}` : "1px solid transparent";
  cellElement.style.borderBottom = paint.showGridlines ? `1px solid ${paint.gridlineColor}` : "1px solid transparent";
  cellElement.style.fontSize = `${style?.fontSizePt ?? 10}pt`;
  if (style?.fontFamily !== undefined) cellElement.style.fontFamily = style.fontFamily;
  if (style?.bold === true || cell.conditional?.bold === true) cellElement.style.fontWeight = "700";
  if (style?.italic === true) cellElement.style.fontStyle = "italic";
  if (style?.underline === true) cellElement.style.textDecoration = "underline";
  // Adaptive dark re-themes only the default: an authored or conditional
  // color paints verbatim in every mode.
  cellElement.style.color = cell.conditional?.color ?? style?.color ?? paint.defaultColor;
  // Fills are never defaulted: transparent lets the host surface (white in
  // source, themed in adaptive dark) show through, preserving provenance.
  cellElement.style.backgroundColor = cell.conditional?.fill ?? style?.fill ?? "transparent";
  cellElement.style.textAlign = textAlign(cell, sheetRtl);
  cellElement.style.verticalAlign = style?.verticalAlign ?? "bottom";
  if (style?.wrapText === true) {
    cellElement.style.whiteSpace = "pre-wrap";
    cellElement.style.overflowWrap = "break-word";
  } else {
    cellElement.style.whiteSpace = "nowrap";
    cellElement.style.overflow = "hidden";
  }
  if (style?.indent !== undefined) cellElement.style.paddingInlineStart = `calc(4px + ${style.indent * 8}px)`;
  // Sheet RTL flips column order via the table's `dir`, but cell text follows
  // Excel's per-cell reading order: 1 = LTR, 2 = RTL, otherwise context. An
  // unspecified order must not inherit the sheet direction, or LTR values
  // such as "180,000 EGP" and "13 Sep 2026" visually reorder in RTL sheets.
  if (style?.readingOrder === 2) cellElement.dir = "rtl";
  else if (style?.readingOrder === 1) cellElement.dir = "ltr";
  else cellElement.dir = "auto";
  const border = style?.border;
  if (border?.top !== undefined) cellElement.style.borderTop = `${borderWidth(border.top.style)} ${borderStyle(border.top.style)} ${border.top.color ?? "currentColor"}`;
  if (border?.left !== undefined) cellElement.style.borderLeft = `${borderWidth(border.left.style)} ${borderStyle(border.left.style)} ${border.left.color ?? "currentColor"}`;
  if (border?.right !== undefined) cellElement.style.borderRight = `${borderWidth(border.right.style)} ${borderStyle(border.right.style)} ${border.right.color ?? "currentColor"}`;
  if (border?.bottom !== undefined) cellElement.style.borderBottom = `${borderWidth(border.bottom.style)} ${borderStyle(border.bottom.style)} ${border.bottom.color ?? "currentColor"}`;
  if (cell.formula !== undefined) cellElement.dataset.aaredFormula = cell.formula;
  if (cell.type === "error") cellElement.dataset.aaredCellError = cell.formatted;
}

function appendChartPrimitive(svg: SVGSVGElement, primitive: ChartGraphicPrimitive): void {
  const namespace = "http://www.w3.org/2000/svg";
  const stroke = (element: SVGElement, stroke: { color: string; width: number } | undefined): void => {
    if (stroke === undefined) return;
    element.setAttribute("stroke", stroke.color);
    element.setAttribute("stroke-width", String(stroke.width));
  };
  switch (primitive.kind) {
    case "rect": {
      const rect = document.createElementNS(namespace, "rect");
      rect.setAttribute("x", String(primitive.x));
      rect.setAttribute("y", String(primitive.y));
      rect.setAttribute("width", String(primitive.width));
      rect.setAttribute("height", String(primitive.height));
      if (primitive.fill !== undefined) rect.setAttribute("fill", primitive.fill);
      stroke(rect, primitive.stroke);
      if (primitive.opacity !== undefined) rect.setAttribute("opacity", String(primitive.opacity));
      svg.append(rect);
      break;
    }
    case "line": {
      const line = document.createElementNS(namespace, "line");
      line.setAttribute("x1", String(primitive.x1));
      line.setAttribute("y1", String(primitive.y1));
      line.setAttribute("x2", String(primitive.x2));
      line.setAttribute("y2", String(primitive.y2));
      stroke(line, primitive.stroke);
      if (primitive.fill !== undefined) stroke(line, { color: primitive.fill, width: 1 });
      if (primitive.opacity !== undefined) line.setAttribute("opacity", String(primitive.opacity));
      svg.append(line);
      break;
    }
    case "path": {
      const path = document.createElementNS(namespace, "path");
      path.setAttribute("d", primitive.path);
      if (primitive.fill !== undefined) path.setAttribute("fill", primitive.fill);
      stroke(path, primitive.stroke);
      if (primitive.opacity !== undefined) path.setAttribute("opacity", String(primitive.opacity));
      svg.append(path);
      break;
    }
    case "text": {
      const text = document.createElementNS(namespace, "text");
      text.setAttribute("x", String(primitive.x));
      text.setAttribute("y", String(primitive.y));
      text.textContent = primitive.text;
      if (primitive.fontSize !== undefined) text.setAttribute("font-size", String(primitive.fontSize));
      if (primitive.fontFamily !== undefined) text.setAttribute("font-family", primitive.fontFamily);
      if (primitive.fontWeight !== undefined) text.setAttribute("font-weight", primitive.fontWeight);
      if (primitive.textAnchor !== undefined) text.setAttribute("text-anchor", primitive.textAnchor);
      if (primitive.fill !== undefined) text.setAttribute("fill", primitive.fill);
      svg.append(text);
      break;
    }
  }
}

interface GridGeometry {
  visibleColumns: number[];
  visibleRows: number[];
  columnOffsets: number[];
  rowOffsets: number[];
  totalWidth: number;
  totalHeight: number;
  frozenColumns: number[];
  frozenRows: number[];
  frozenColumnOffset: Map<number, number>;
  frozenRowOffset: Map<number, number>;
}

function computeGeometry(
  model: SheetModel,
  heightOf: (row: number) => number,
  headerWidthPx = 0,
  headerHeightPx = 0,
): GridGeometry {
  const hiddenColumns = new Set(model.hiddenColumns);
  const hiddenRows = new Set(model.hiddenRows);
  const visibleColumns: number[] = [];
  for (let col = 0; col < model.extent.cols; col++) {
    if (!hiddenColumns.has(col)) visibleColumns.push(col);
  }
  const visibleRows: number[] = [];
  for (let row = 0; row < model.extent.rows; row++) {
    if (!hiddenRows.has(row)) visibleRows.push(row);
  }
  const columnWidths = visibleColumns.map((col) => (model.columnWidths[col] ?? 0) * PIXELS_PER_INCH);
  const rowHeights = visibleRows.map((row) => heightOf(row));
  const columnOffsets: number[] = [];
  let x = 0;
  for (const width of columnWidths) {
    columnOffsets.push(x);
    x += width;
  }
  const rowOffsets: number[] = [];
  let y = 0;
  for (const height of rowHeights) {
    rowOffsets.push(y);
    y += height;
  }
  // Frozen panes freeze the leading rows/columns; hidden ones inside the
  // frozen band simply never render. Offsets accumulate over visible frozen
  // cells only, so sticky positions stay exact.
  const frozenColumns = visibleColumns.filter((col) => col < model.frozen.cols);
  const frozenRows = visibleRows.filter((row) => row < model.frozen.rows);
  const frozenColumnOffset = new Map<number, number>();
  // Sticky frozen cells sit past the header column/row when headings render.
  let frozenX = headerWidthPx;
  for (const col of frozenColumns) {
    frozenColumnOffset.set(col, frozenX);
    frozenX += columnWidths[visibleColumns.indexOf(col)]!;
  }
  const frozenRowOffset = new Map<number, number>();
  let frozenY = headerHeightPx;
  for (const row of frozenRows) {
    frozenRowOffset.set(row, frozenY);
    frozenY += rowHeights[visibleRows.indexOf(row)]!;
  }
  return {
    visibleColumns,
    visibleRows,
    columnOffsets,
    rowOffsets,
    totalWidth: x,
    totalHeight: y,
    frozenColumns,
    frozenRows,
    frozenColumnOffset,
    frozenRowOffset,
  };
}

interface MergeLayout {
  originOf: Map<string, { row: number; col: number }>;
  rangeOf: Map<string, { row: number; col: number; endRow: number; endCol: number }>;
}

/** Index merges by origin and by covered cell; hidden origins drop the merge. */
function computeMerges(model: SheetModel): MergeLayout {
  const hiddenColumns = new Set(model.hiddenColumns);
  const hiddenRows = new Set(model.hiddenRows);
  const originOf = new Map<string, { row: number; col: number }>();
  const rangeOf = new Map<string, { row: number; col: number; endRow: number; endCol: number }>();
  for (const merge of model.merges) {
    if (hiddenRows.has(merge.start.row) || hiddenColumns.has(merge.start.col)) continue;
    const key = `${merge.start.row}:${merge.start.col}`;
    rangeOf.set(key, { row: merge.start.row, col: merge.start.col, endRow: merge.end.row, endCol: merge.end.col });
    for (let row = merge.start.row; row <= merge.end.row; row++) {
      for (let col = merge.start.col; col <= merge.end.col; col++) {
        if (row === merge.start.row && col === merge.start.col) continue;
        if (hiddenRows.has(row) || hiddenColumns.has(col)) continue;
        originOf.set(`${row}:${col}`, { row: merge.start.row, col: merge.start.col });
      }
    }
  }
  return { originOf, rangeOf };
}

/** One painted grid: DOM handles plus the query/scroll helpers the controller needs. */
interface SheetPaintSession {
  destroy(): void;
  host: HTMLElement;
  scroller: HTMLElement;
  geometry: GridGeometry;
  headerWidthPx: number;
  headerHeightPx: number;
  rowHeightOf: (row: number) => number;
  model: SheetModel;
  /** The rendered cell, falling back to a merge origin when the cell is covered. */
  cellElement(row: number, col: number): HTMLElement | null;
  /** Synchronously re-render the virtualization window (used after programmatic scrolls). */
  renderWindow(): void;
}

/** Paint one already-resolved sheet. The caller owns which sheet is visible. */
function paintSheet(
  model: SheetModel,
  container: HTMLElement,
  mediaResolver: MediaResolver | undefined,
  paint: GridPaint,
  onWindow?: () => void,
): SheetPaintSession {
  const objectUrls: string[] = [];
  let destroyed = false;
  let raf = 0;

  const heightByRow = new Map<number, number>();
  for (const row of model.rows) {
    if (row.height !== undefined) heightByRow.set(row.index, row.height);
  }
  const rowHeightOf = (row: number): number => (heightByRow.get(row) ?? model.defaultRowHeight) * PIXELS_PER_INCH;
  const showHeaders = paint.showHeaders;
  const headerWidthPx = showHeaders ? HEADER_COL_WIDTH_PX : 0;
  const headerHeightPx = showHeaders ? HEADER_ROW_HEIGHT_PX : 0;
  const geometry = computeGeometry(model, rowHeightOf, headerWidthPx, headerHeightPx);
  const merges = computeMerges(model);
  const cellsByRow = new Map<number, Map<number, CellModel>>();
  for (const row of model.rows) {
    const entries = new Map<number, CellModel>();
    for (const { col, cell } of row.cells) entries.set(col, cell);
    cellsByRow.set(row.index, entries);
  }
  const hiddenColumns = new Set(model.hiddenColumns);
  const hiddenRowSet = new Set(model.hiddenRows);
  const frozenColumns: number[] = [...geometry.frozenColumns];
  const frozenRows: number[] = [...geometry.frozenRows];
  const frozenRowSet = new Set(frozenRows);
  const frozenColumnLookup = new Set(frozenColumns);

  const host = document.createElement("div");
  host.dataset.aaredElement = "sheet";
  host.style.position = "relative";
  host.style.width = "100%";
  // Fill the embedder's height when it gives one so the grid below can
  // scroll; without a bounded ancestor the grid grows to its content and
  // virtualization idles (every row renders, still correct).
  host.style.height = "100%";
  host.style.display = "flex";
  host.style.flexDirection = "column";
  host.style.background = paint.hostBackground;
  // Theme is viewer chrome state: chrome reads these variables, and adaptive
  // content defaults resolve through them. Never a filter or inversion.
  host.dataset.aaredTheme = paint.dark ? "dark" : "light";
  host.dataset.aaredContentMode = paint.contentMode;
  host.style.setProperty("--aared-surface", paint.dark ? SURFACE_DARK : "#FFFFFF");
  host.style.setProperty("--aared-text", paint.dark ? TEXT_DARK : "#000000");
  host.style.setProperty("--aared-gridline", paint.gridlineColor);
  host.style.setProperty("--aared-header-fill", paint.headerFill);
  host.style.setProperty("--aared-header-color", paint.headerColor);
  host.style.setProperty("--aared-selection", paint.dark ? SELECTION_DARK : SELECTION_LIGHT);
  host.style.setProperty("--aared-border", paint.gridlineColor);
  host.style.setProperty("--aared-muted-text", paint.dark ? NOTICE_COLOR_DARK : NOTICE_COLOR);
  host.style.setProperty("--aared-accent", paint.dark ? SELECTION_DARK : SELECTION_LIGHT);
  host.style.setProperty("--aared-focus-ring", paint.dark ? SELECTION_DARK : SELECTION_LIGHT);
  // The host holds the single tab stop: cells are programmatically
  // focusable (`tabindex="-1"`) but virtualization unmounts off-window
  // rows, so a roving tabindex on cells would drop the tab stop whenever
  // the selection scrolls out of the window.
  host.tabIndex = 0;
  host.setAttribute("aria-label", model.name);

  const hasFormulas = model.rows.some((row) => row.cells.some(({ cell }) => cell.formula !== undefined));
  if (model.unsupported.length > 0 || hasFormulas) {
    const notices = document.createElement("div");
    notices.dataset.aaredElement = "sheet-notices";
    notices.style.font = "12px system-ui, sans-serif";
    notices.style.color = paint.noticeColor;
    notices.style.padding = "4px 8px";
    notices.style.borderBottom = `1px solid ${paint.chromeLine}`;
    notices.style.flex = "none";
    const list = document.createElement("ul");
    list.style.margin = "0";
    list.style.paddingInlineStart = "1.25em";
    if (hasFormulas) {
      const item = document.createElement("li");
      item.dataset.aaredNotice = "formulas";
      item.textContent = "Formulas show the values saved by Excel and are never recalculated.";
      list.append(item);
    }
    for (const entry of model.unsupported) {
      const item = document.createElement("li");
      item.dataset.aaredNotice = entry.kind;
      item.textContent = entry.detail;
      list.append(item);
    }
    notices.append(list);
    host.append(notices);
  }

  const scroller = document.createElement("div");
  scroller.dataset.aaredElement = "sheet-scroll";
  scroller.style.overflow = "auto";
  scroller.style.flex = "1 1 auto";
  scroller.style.minHeight = "0";
  scroller.style.position = "relative";
  host.append(scroller);

  const canvas = document.createElement("div");
  canvas.dataset.aaredElement = "sheet-canvas";
  canvas.style.position = "relative";
  canvas.style.width = `${geometry.totalWidth + headerWidthPx}px`;
  canvas.style.minHeight = "1px";
  scroller.append(canvas);

  const table = document.createElement("table");
  table.dataset.aaredElement = "grid";
  table.dataset.aaredSheetDirection = model.rightToLeft ? "rtl" : "ltr";
  table.dir = model.rightToLeft ? "rtl" : "ltr";
  table.dataset.aaredHiddenColumns = model.hiddenColumns.join(",");
  table.dataset.aaredHiddenRows = model.hiddenRows.join(",");
  table.dataset.aaredHeaders = showHeaders ? "true" : "false";
  table.dataset.aaredGridlines = paint.showGridlines ? "true" : "false";
  table.setAttribute("role", "grid");
  table.setAttribute("aria-rowcount", String(model.extent.rows));
  table.setAttribute("aria-colcount", String(model.extent.cols));
  table.setAttribute("aria-label", model.name);
  table.style.borderCollapse = "separate";
  table.style.borderSpacing = "0";
  table.style.tableLayout = "fixed";
  table.style.width = `${geometry.totalWidth + headerWidthPx}px`;
  canvas.append(table);

  const colgroup = document.createElement("colgroup");
  if (showHeaders) {
    const gutter = document.createElement("col");
    gutter.style.width = `${HEADER_COL_WIDTH_PX}px`;
    colgroup.append(gutter);
  }
  for (const col of geometry.visibleColumns) {
    const element = document.createElement("col");
    element.style.width = pixels(model.columnWidths[col] ?? 0);
    colgroup.append(element);
  }
  table.append(colgroup);

  /** Paint one heading cell: sticky on its axis, above frozen cells. */
  const paintHeading = (element: HTMLElement, kind: "corner" | "row" | "column"): void => {
    element.style.position = "sticky";
    element.style.zIndex = kind === "corner" ? "5" : "4";
    if (kind === "corner" || kind === "column") element.style.top = "0px";
    if (kind === "corner" || kind === "row") {
      // Headings stick to the inline-start edge: physical `right` in RTL.
      if (model.rightToLeft) element.style.right = "0px";
      else element.style.left = "0px";
    }
    element.style.backgroundColor = paint.headerFill;
    element.style.color = paint.headerColor;
    // Headings share the grid's own hairlines — bottom/right only, never a
    // full box — so the reference headers blend with the grid instead of
    // buttoning it. Borders follow content mode (gridlineColor), not chrome.
    element.style.borderBottom = `1px solid ${paint.gridlineColor}`;
    element.style.borderRight = `1px solid ${paint.gridlineColor}`;
    element.style.fontSize = "9pt";
    element.style.fontWeight = "600";
    element.style.textAlign = "center";
    element.style.verticalAlign = "middle";
    element.style.whiteSpace = "nowrap";
    element.style.userSelect = "none";
  };

  if (showHeaders) {
    // Letters skip hidden columns exactly like Excel: …C, E… with D hidden.
    const head = document.createElement("thead");
    head.dataset.aaredElement = "column-headers";
    const headRow = document.createElement("tr");
    headRow.style.height = `${HEADER_ROW_HEIGHT_PX}px`;
    const corner = document.createElement("th");
    corner.dataset.aaredElement = "corner";
    corner.scope = "col";
    corner.setAttribute("role", "columnheader");
    paintHeading(corner, "corner");
    headRow.append(corner);
    for (const col of geometry.visibleColumns) {
      const letter = document.createElement("th");
      letter.dataset.aaredElement = "column-header";
      letter.dataset.aaredCol = String(col);
      letter.scope = "col";
      letter.setAttribute("role", "columnheader");
      letter.textContent = columnName(col);
      paintHeading(letter, "column");
      headRow.append(letter);
    }
    head.append(headRow);
    table.append(head);
  }
  const body = document.createElement("tbody");
  table.append(body);

  const paintFrozen = (element: HTMLElement, row: number, col: number): void => {
    const frozenRow = frozenRowSet.has(row);
    const frozenColumn = frozenColumnLookup.has(col);
    if (!frozenRow && !frozenColumn) return;
    element.style.position = "sticky";
    element.style.zIndex = frozenRow && frozenColumn ? "3" : "2";
    if (frozenRow) element.style.top = `${geometry.frozenRowOffset.get(row) ?? 0}px`;
    if (frozenColumn) {
      // Sticky follows the inline-start edge: physical `right` in RTL sheets.
      const offset = `${geometry.frozenColumnOffset.get(col) ?? 0}px`;
      if (model.rightToLeft) element.style.right = offset;
      else element.style.left = offset;
    }
    element.dataset.aaredFrozen = frozenRow && frozenColumn ? "corner" : frozenRow ? "row" : "column";
    if (element.style.backgroundColor === "" || element.style.backgroundColor === "transparent") {
      element.style.backgroundColor = paint.frozenFill;
    }
  };

  const paintRow = (row: number): HTMLTableRowElement => {
    const line = document.createElement("tr");
    line.dataset.aaredRow = String(row);
    line.setAttribute("role", "row");
    line.setAttribute("aria-rowindex", String(row + 1));
    line.style.height = `${rowHeightOf(row)}px`;
    if (showHeaders) {
      // Numbers skip hidden rows exactly like Excel: …6, 8… with 7 hidden.
      const number = document.createElement("th");
      number.dataset.aaredElement = "row-header";
      number.dataset.aaredRow = String(row);
      number.scope = "row";
      number.setAttribute("role", "rowheader");
      number.textContent = String(row + 1);
      paintHeading(number, "row");
      line.append(number);
    }
    const entries = cellsByRow.get(row);
    for (const col of geometry.visibleColumns) {
      const covered = merges.originOf.get(`${row}:${col}`);
      if (covered !== undefined) continue;
      const cellElement = document.createElement("td");
      cellElement.dataset.aaredElement = "cell";
      cellElement.dataset.aaredRow = String(row);
      cellElement.dataset.aaredCol = String(col);
      cellElement.setAttribute("role", "gridcell");
      cellElement.setAttribute("aria-rowindex", String(row + 1));
      cellElement.setAttribute("aria-colindex", String(col + 1));
      // Programmatic focus only; the host keeps the tab stop (see above).
      cellElement.tabIndex = -1;
      cellElement.style.position = "relative";
      const merge = merges.rangeOf.get(`${row}:${col}`);
      const cell = entries?.get(col);
      if (merge !== undefined) {
        // Spans count visible cells only, so hidden rows/columns inside a
        // merge collapse exactly as Excel collapses them.
        const endRow = Math.min(merge.endRow, model.extent.rows - 1);
        const endCol = Math.min(merge.endCol, model.extent.cols - 1);
        let rowSpan = 0;
        for (let probe = row; probe <= endRow; probe++) {
          if (!hiddenRowSet.has(probe)) rowSpan++;
        }
        let colSpan = 0;
        for (let probe = col; probe <= endCol; probe++) {
          if (!hiddenColumns.has(probe)) colSpan++;
        }
        if (rowSpan > 1) cellElement.rowSpan = rowSpan;
        if (colSpan > 1) cellElement.colSpan = colSpan;
        cellElement.dataset.aaredMerge = `${merge.row}:${merge.col}:${endRow}:${endCol}`;
      }
      const effective: CellModel = cell ?? { value: null, formatted: "", type: "empty" };
      paintCell(cellElement, effective, model.rightToLeft, paint);
      appendCellContent(cellElement, effective);
      paintFrozen(cellElement, row, col);
      line.append(cellElement);
    }
    return line;
  };

  const expandForMerges = (start: number, end: number): { start: number; end: number } => {
    // A merge intersecting the window renders whole: partial spans cannot be
    // expressed across spacer rows, so the window grows to its boundaries.
    let next = { start, end };
    for (const merge of model.merges) {
      if (merge.end.row >= next.start && merge.start.row <= next.end) {
        next = { start: Math.min(next.start, merge.start.row), end: Math.max(next.end, merge.end.row) };
      }
    }
    return next;
  };

  const renderWindow = (): void => {
    if (destroyed) return;
    body.replaceChildren();
    const dataRows = geometry.visibleRows.filter((row) => !frozenRowSet.has(row));
    const scrollTop = scroller.scrollTop;
    const viewport = scroller.clientHeight === 0 ? 600 : scroller.clientHeight;
    // Offsets over visible rows only; frozen rows always render at the top.
    const offsets: number[] = [];
    let total = 0;
    for (const row of dataRows) {
      offsets.push(total);
      total += rowHeightOf(row);
    }
    let startIndex = dataRows.length;
    let endIndex = -1;
    for (let i = 0; i < dataRows.length; i++) {
      const top = offsets[i]!;
      const bottom = top + rowHeightOf(dataRows[i]!);
      if (bottom > scrollTop - ROW_OVERSCAN * 20 && top < scrollTop + viewport + ROW_OVERSCAN * 20) {
        startIndex = Math.min(startIndex, i);
        endIndex = Math.max(endIndex, i);
      }
    }
    for (const row of frozenRows) body.append(paintRow(row));
    if (dataRows.length !== 0 && endIndex !== -1) {
      startIndex = Math.max(0, startIndex - ROW_OVERSCAN);
      endIndex = Math.min(dataRows.length - 1, endIndex + ROW_OVERSCAN);
      // Expand across merge boundaries using absolute row numbers, then map
      // back into the visible data-row list.
      const absolute = expandForMerges(dataRows[startIndex]!, dataRows[endIndex]!);
      while (startIndex > 0 && dataRows[startIndex]! > absolute.start) startIndex--;
      while (endIndex < dataRows.length - 1 && dataRows[endIndex]! < absolute.end) endIndex++;
      const visibleCount = geometry.visibleColumns.length + (showHeaders ? 1 : 0) || 1;
      if (startIndex > 0) {
        const spacer = document.createElement("tr");
        spacer.dataset.aaredElement = "spacer";
        spacer.setAttribute("aria-hidden", "true");
        const cell = document.createElement("td");
        cell.colSpan = visibleCount;
        cell.style.height = `${offsets[startIndex]!}px`;
        cell.style.padding = "0";
        cell.style.border = "0";
        spacer.append(cell);
        body.append(spacer);
      }
      for (let i = startIndex; i <= endIndex; i++) body.append(paintRow(dataRows[i]!));
      const tail = total - (offsets[endIndex]! + rowHeightOf(dataRows[endIndex]!));
      if (tail > 0) {
        const spacer = document.createElement("tr");
        spacer.dataset.aaredElement = "spacer";
        spacer.setAttribute("aria-hidden", "true");
        const cell = document.createElement("td");
        cell.colSpan = visibleCount;
        cell.style.height = `${tail}px`;
        cell.style.padding = "0";
        cell.style.border = "0";
        spacer.append(cell);
        body.append(spacer);
      }
    }
    // The controller re-applies ephemeral state (selection) that a fresh
    // window would otherwise drop. Windowed Worker re-reads on scroll are
    // deliberately out of scope: the grid only ever paints the model it was
    // handed (see the module header).
    onWindow?.();
  };

  const schedule = (): void => {
    if (destroyed || raf !== 0) return;
    raf = requestAnimationFrame(() => {
      raf = 0;
      renderWindow();
    });
  };
  // Virtualization windows on the scroller's live size: without this, a
  // container resize (or a first paint before layout) leaves a stale window
  // rendered — rows beyond it never appear until the next scroll event.
  const onScroll = (): void => schedule();
  scroller.addEventListener("scroll", onScroll, { passive: true });
  const resizer = new ResizeObserver(() => schedule());
  renderWindow();

  const drawingsLayer = document.createElement("div");
  drawingsLayer.dataset.aaredElement = "drawings";
  drawingsLayer.style.position = "absolute";
  drawingsLayer.style.inset = "0";
  drawingsLayer.style.pointerEvents = "none";
  canvas.append(drawingsLayer);

  const paintDrawingBox = (drawing: SheetDrawingModel): HTMLElement => {
    const box = document.createElement("div");
    box.style.position = "absolute";
    box.style.left = pixels(drawing.bounds.x);
    box.style.top = pixels(drawing.bounds.y);
    box.style.width = pixels(Math.max(drawing.bounds.width, 0.01));
    box.style.height = pixels(Math.max(drawing.bounds.height, 0.01));
    box.style.pointerEvents = "auto";
    box.style.overflow = "hidden";
    return box;
  };

  const paintChart = (drawing: SheetChartModel): void => {
    const box = paintDrawingBox(drawing);
    box.dataset.aaredElement = "chart";
    box.dataset.aaredChartName = drawing.name;
    if (drawing.sourceChartId !== undefined) box.dataset.aaredChartSource = drawing.sourceChartId;
    if (drawing.graphic === undefined) {
      box.dataset.aaredChartFallback = drawing.unsupportedReason ?? "unsupported chart";
      box.style.display = "flex";
      box.style.alignItems = "center";
      box.style.justifyContent = "center";
      box.style.border = `1px dashed ${paint.chromeLine}`;
      box.style.font = "12px system-ui, sans-serif";
      box.style.color = paint.noticeColor;
      box.textContent = drawing.unsupportedReason ?? "This chart type is not previewed.";
      drawingsLayer.append(box);
      return;
    }
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.dataset.aaredElement = "chart-graphic";
    svg.setAttribute("viewBox", `0 0 ${drawing.graphic.viewBox.width} ${drawing.graphic.viewBox.height}`);
    svg.style.width = "100%";
    svg.style.height = "100%";
    svg.setAttribute("role", "img");
    svg.setAttribute("aria-label", drawing.name);
    for (const primitive of drawing.graphic.primitives) appendChartPrimitive(svg, primitive);
    box.append(svg);
    drawingsLayer.append(box);
  };

  const paintPicture = (drawing: SheetPictureModel, resolver: MediaResolver | undefined): void => {
    const box = paintDrawingBox(drawing);
    box.dataset.aaredElement = "picture";
    box.dataset.aaredMedia = drawing.mediaId;
    const placeholder = document.createElement("div");
    placeholder.dataset.aaredElement = "picture-placeholder";
    placeholder.style.width = "100%";
    placeholder.style.height = "100%";
    placeholder.style.display = "flex";
    placeholder.style.alignItems = "center";
    placeholder.style.justifyContent = "center";
    placeholder.style.border = `1px dashed ${paint.chromeLine}`;
    placeholder.style.font = "12px system-ui, sans-serif";
    placeholder.style.color = paint.noticeColor;
    placeholder.textContent = drawing.name;
    box.append(placeholder);
    drawingsLayer.append(box);
    if (resolver === undefined) return;
    void resolver(drawing.mediaId).then(
      (bytes) => {
        if (destroyed || !box.isConnected) return;
        const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: drawing.contentType }));
        objectUrls.push(url);
        const image = document.createElement("img");
        image.src = url;
        image.alt = drawing.name;
        image.style.width = "100%";
        image.style.height = "100%";
        image.style.objectFit = "contain";
        box.replaceChildren(image);
      },
      () => {
        if (destroyed || !box.isConnected) return;
        placeholder.textContent = `${drawing.name} (unavailable)`;
      },
    );
  };

  for (const drawing of model.drawings) {
    if (drawing.kind === "chart") paintChart(drawing);
    else paintPicture(drawing, mediaResolver);
  }

  container.replaceChildren(host);
  // Observe attached: an observer added while the tree is detached never
  // sees layout, and this engine does not reliably backfill the initial
  // notification on attach. Observing here self-corrects the pre-layout
  // first paint (rendered above with a fallback viewport) once real
  // metrics exist.
  resizer.observe(scroller);
  const destroy = (): void => {
    if (destroyed) return;
    destroyed = true;
    if (raf !== 0) cancelAnimationFrame(raf);
    resizer.disconnect();
    scroller.removeEventListener("scroll", onScroll);
    for (const url of objectUrls) URL.revokeObjectURL(url);
  };
  const cellElement = (row: number, col: number): HTMLElement | null => {
    const direct = body.querySelector<HTMLElement>(
      `td[data-aared-element="cell"][data-aared-row="${row}"][data-aared-col="${col}"]`,
    );
    if (direct !== null) return direct;
    // Covered merge cells never render: focus their origin instead.
    const origin = merges.originOf.get(`${row}:${col}`);
    if (origin === undefined) return null;
    return body.querySelector<HTMLElement>(
      `td[data-aared-element="cell"][data-aared-row="${origin.row}"][data-aared-col="${origin.col}"]`,
    );
  };
  return {
    destroy,
    host,
    scroller,
    geometry,
    headerWidthPx,
    headerHeightPx,
    rowHeightOf,
    model,
    cellElement,
    renderWindow,
  };
}

/** Parse a single A1 cell; ranges are rejected for cell-targeted navigation. */
function parseSingleCell(cell: string, what: string): { row: number; col: number } {
  const range = parseA1Range(cell);
  if (range.start.row !== range.end.row || range.start.col !== range.end.col) {
    throw new AaredError(
      "invalid-range",
      `aared/xlsx: ${what} expects a single cell like "D14", not a range ${JSON.stringify(cell)}`,
    );
  }
  return { row: range.start.row, col: range.start.col };
}

interface NormalizedAppearance {
  theme: ViewerTheme;
  dark: boolean;
  contentMode: "source" | "adaptive";
  headers: boolean;
  sheetTabs: boolean;
  selection: boolean;
  gridlines: GridlineMode;
}

function normalizeAppearance(appearance: XlsxAppearance | Appearance | undefined): NormalizedAppearance {
  const theme = appearance?.theme ?? "system";
  let dark = theme === "dark";
  if (theme === "system" && typeof matchMedia === "function") {
    dark = matchMedia("(prefers-color-scheme: dark)").matches;
  }
  return {
    theme,
    dark,
    contentMode: appearance?.contentMode ?? "adaptive",
    headers: appearance?.spreadsheet?.headers ?? true,
    sheetTabs: appearance?.spreadsheet?.sheetTabs ?? false,
    selection: appearance?.spreadsheet?.selection ?? true,
    gridlines: appearance?.spreadsheet?.gridlines ?? "source",
  };
}

function resolveGridPaint(model: SheetModel, appearance: NormalizedAppearance): GridPaint {
  const adaptiveDark = appearance.contentMode === "adaptive" && appearance.dark;
  return {
    showHeaders: appearance.headers,
    showGridlines: appearance.gridlines === "show"
      ? true
      : appearance.gridlines === "hide"
        ? false
        : (model.showGridlines ?? true),
    gridlineColor: adaptiveDark ? GRIDLINE_DARK : GRIDLINE_COLOR,
    defaultColor: adaptiveDark ? TEXT_DARK : "#000000",
    frozenFill: adaptiveDark ? SURFACE_DARK : "white",
    headerFill: adaptiveDark ? HEADER_FILL_DARK : HEADER_FILL,
    headerColor: adaptiveDark ? HEADER_COLOR_DARK : HEADER_COLOR,
    chromeLine: appearance.dark ? GRIDLINE_DARK : GRIDLINE_COLOR,
    noticeColor: appearance.dark ? NOTICE_COLOR_DARK : NOTICE_COLOR,
    hostBackground: adaptiveDark ? SURFACE_DARK : "white",
    dark: appearance.dark,
    contentMode: appearance.contentMode,
  };
}

/**
 * The `render()` controller: zoom, cell navigation, single-cell selection,
 * sheet tabs, and lifecycle over one target element. Owns a stable tab bar
 * plus grid container; each painted sheet swaps the grid host in place.
 */
export class XlsxViewer extends EventTarget implements XlsxRenderHandle {
  readonly #target: HTMLElement;
  readonly #gridBox: HTMLDivElement;
  #tabsBar: HTMLDivElement | null = null;
  #session: SheetPaintSession | null = null;
  #model: SheetModel;
  readonly #handle: SheetHandle | undefined;
  readonly #document: XlsxDocument | undefined;
  readonly #mediaResolver: MediaResolver | undefined;
  readonly #appearance: NormalizedAppearance;
  #zoom: ZoomValue = 1;
  #selection: { address: string; row: number; col: number } | null = null;
  #sheet: SheetInfo;
  #sheetGeneration = 0;
  #sheetController: AbortController | null = null;
  readonly #listeners = new AbortController();
  #destroyed = false;
  readonly #previousTargetDisplay: string;
  readonly #previousTargetFlexDirection: string;

  /** Built only by `render()` after the model (and optional handle) resolve. */
  constructor(
    target: HTMLElement,
    model: SheetModel | SheetRangeModel,
    handle: SheetHandle | undefined,
    mediaResolver: MediaResolver | undefined,
    appearance: XlsxAppearance | Appearance | undefined,
  ) {
    super();
    this.#target = target;
    this.#model = model;
    this.#handle = handle;
    this.#document = handle?.sheetDocument();
    this.#mediaResolver = mediaResolver;
    this.#appearance = normalizeAppearance(appearance);
    this.#sheet = handle === undefined
      ? { index: 0, name: model.name, state: "visible" }
      : { ...handle.info };
    this.#previousTargetDisplay = target.style.display;
    this.#previousTargetFlexDirection = target.style.flexDirection;
    target.style.display = "flex";
    target.style.flexDirection = "column";
    this.#gridBox = document.createElement("div");
    this.#gridBox.dataset.aaredElement = "sheet-body";
    this.#gridBox.style.display = "flex";
    this.#gridBox.style.flexDirection = "column";
    this.#gridBox.style.flex = "1 1 auto";
    this.#gridBox.style.minHeight = "0";
    this.#gridBox.style.minWidth = "0";
    this.#buildTabs();
    if (this.#tabsBar !== null) target.replaceChildren(this.#tabsBar, this.#gridBox);
    else target.replaceChildren(this.#gridBox);
    const signal = this.#listeners.signal;
    this.#gridBox.addEventListener("click", (event) => this.#handleClick(event), { signal });
    // Middle-clicks bypass `click`: the same interception applies so a link
    // can never navigate automatically.
    this.#gridBox.addEventListener("auxclick", (event) => this.#handleClick(event), { signal });
    this.#gridBox.addEventListener("dblclick", (event) => this.#handleDoubleClick(event), { signal });
    this.#gridBox.addEventListener("keydown", (event) => this.#handleKeyDown(event), { signal });
    this.#repaintCurrent();
  }

  get zoom(): ZoomValue {
    return this.#zoom;
  }

  setZoom(zoom: ZoomValue): void {
    this.#assertLive();
    this.#zoom = zoom;
    this.#applyZoom();
    this.dispatchEvent(new CustomEvent("zoomchange", { detail: { zoom } }));
  }

  get sheet(): SheetInfo {
    return { ...this.#sheet };
  }

  get selection(): { address: string; row: number; col: number } | null {
    return this.#selection === null ? null : { ...this.#selection };
  }

  async goTo(cell: string): Promise<void> {
    this.#assertLive();
    const { row, col } = parseSingleCell(cell, "goTo");
    if (this.#appearance.selection) {
      this.#select({ address: cellAddress(row, col), row, col }, true);
    }
    // Center when the cell is out of view; a visible cell keeps its scroll.
    this.#ensureVisible(row, col, "center", true);
    this.#session?.cellElement(row, col)?.focus({ preventScroll: true });
  }

  scrollTo(target: { cell?: string; range?: string; align?: "start" | "center" | "end" }): void {
    this.#assertLive();
    const align = target.align ?? "start";
    let row: number;
    let col: number;
    if (target.cell !== undefined) {
      ({ row, col } = parseA1Range(target.cell).start);
    } else if (target.range !== undefined) {
      const range = parseA1Range(target.range);
      if (align === "center") {
        row = Math.floor((range.start.row + range.end.row) / 2);
        col = Math.floor((range.start.col + range.end.col) / 2);
      } else if (align === "end") {
        ({ row, col } = range.end);
      } else {
        ({ row, col } = range.start);
      }
    } else {
      return;
    }
    this.#ensureVisible(row, col, align, false);
  }

  focus(cell: string): void {
    this.#assertLive();
    const { row, col } = parseSingleCell(cell, "focus");
    if (this.#appearance.selection) {
      this.#select({ address: cellAddress(row, col), row, col }, true);
    }
    // No scrolling by construction: a cell outside the current window keeps
    // its logical selection without DOM focus until it is scrolled into view.
    this.#session?.cellElement(row, col)?.focus({ preventScroll: true });
  }

  async goToSheet(name: string | number): Promise<void> {
    this.#assertLive();
    const document = this.#document;
    if (document === undefined) {
      throw new AaredError(
        "unsupported-feature",
        "aared/xlsx: sheet switching needs a SheetHandle input; this view was rendered from a bare model",
      );
    }
    const handle = document.sheets.get(name);
    const generation = ++this.#sheetGeneration;
    this.#sheetController?.abort();
    const controller = new AbortController();
    this.#sheetController = controller;
    const signal = controller.signal;
    let model: SheetRangeModel;
    try {
      model = await handle.read(null, { signal });
    } catch (error) {
      // A stale switch (aborted by a newer one or by destroy) stays silent.
      if (generation !== this.#sheetGeneration) return;
      throw error;
    }
    if (generation !== this.#sheetGeneration || signal.aborted) return;
    this.#sheetController = null;
    this.#model = model;
    this.#sheet = { ...handle.info };
    const hadSelection = this.#selection !== null;
    this.#selection = null;
    this.#repaintCurrent();
    // Reset the grid scroll: the new sheet paints a fresh scroller at origin.
    if (hadSelection) {
      this.dispatchEvent(new CustomEvent("selectionchange", { detail: { sheet: this.#sheet.name, selection: null } }));
    }
    this.dispatchEvent(new CustomEvent("sheetchange", { detail: { sheet: this.#sheet.name } }));
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    this.#sheetGeneration += 1;
    this.#sheetController?.abort();
    this.#sheetController = null;
    this.#listeners.abort();
    this.#session?.destroy();
    this.#session = null;
    this.#target.style.display = this.#previousTargetDisplay;
    this.#target.style.flexDirection = this.#previousTargetFlexDirection;
  }

  #assertLive(): void {
    if (this.#destroyed) throw new AaredError("aborted", "aared/xlsx: viewer is destroyed");
  }

  #buildTabs(): void {
    // Tabs need the owning document (handle inputs only) and list visible
    // sheets: hidden sheets stay reachable via goToSheet(), never via tabs.
    if (!this.#appearance.sheetTabs || this.#document === undefined) return;
    const bar = document.createElement("div");
    bar.dataset.aaredElement = "sheet-tabs";
    bar.setAttribute("role", "tablist");
    bar.setAttribute("aria-label", "Sheets");
    bar.style.display = "flex";
    bar.style.flex = "none";
    bar.style.gap = "2px";
    bar.style.overflowX = "auto";
    bar.style.padding = "4px 8px 0";
    bar.style.backgroundColor = "var(--aared-header-fill)";
    bar.style.borderBottom = "1px solid var(--aared-gridline)";
    bar.style.setProperty("--aared-header-fill", this.#appearance.dark ? HEADER_FILL_DARK : HEADER_FILL);
    bar.style.setProperty("--aared-gridline", this.#appearance.dark ? GRIDLINE_DARK : GRIDLINE_COLOR);
    bar.style.setProperty("--aared-selection", this.#appearance.dark ? SELECTION_DARK : SELECTION_LIGHT);
    bar.style.setProperty("--aared-text", this.#appearance.dark ? TEXT_DARK : "#000000");
    bar.style.setProperty("--aared-border", this.#appearance.dark ? GRIDLINE_DARK : GRIDLINE_COLOR);
    bar.style.setProperty("--aared-muted-text", this.#appearance.dark ? NOTICE_COLOR_DARK : NOTICE_COLOR);
    bar.style.setProperty("--aared-accent", this.#appearance.dark ? SELECTION_DARK : SELECTION_LIGHT);
    bar.style.setProperty("--aared-focus-ring", this.#appearance.dark ? SELECTION_DARK : SELECTION_LIGHT);
    for (const handle of this.#document.sheets.list()) {
      if (handle.info.state !== "visible") continue;
      const tab = document.createElement("button");
      tab.type = "button";
      tab.dataset.aaredElement = "sheet-tab";
      tab.dataset.aaredSheet = handle.info.name;
      tab.setAttribute("role", "tab");
      const selected = handle.info.name === this.#sheet.name;
      tab.setAttribute("aria-selected", selected ? "true" : "false");
      tab.tabIndex = selected ? 0 : -1;
      tab.textContent = handle.info.name;
      tab.style.font = "12px system-ui, sans-serif";
      tab.style.color = "var(--aared-text)";
      tab.style.background = "transparent";
      tab.style.border = "none";
      tab.style.borderBottom = selected ? "2px solid var(--aared-selection)" : "2px solid transparent";
      tab.style.fontWeight = selected ? "700" : "400";
      tab.style.padding = "4px 10px 6px";
      tab.style.cursor = "pointer";
      tab.style.whiteSpace = "nowrap";
      const sheetName = handle.info.name;
      tab.addEventListener("click", () => {
        // Tabs derive from the live collection, so not-found cannot happen;
        // abort rejections from a superseded switch stay silent in goToSheet.
        void this.goToSheet(sheetName).catch(() => {});
      });
      bar.append(tab);
    }
    // A workbook whose sheets are all hidden still gets the (empty) bar so
    // the marker stays queryable; tabs appear only for visible sheets.
    this.#tabsBar = bar;
  }

  #updateTabs(): void {
    if (this.#tabsBar === null) return;
    for (const tab of this.#tabsBar.querySelectorAll<HTMLElement>('[data-aared-element="sheet-tab"]')) {
      const selected = tab.dataset.aaredSheet === this.#sheet.name;
      tab.setAttribute("aria-selected", selected ? "true" : "false");
      tab.tabIndex = selected ? 0 : -1;
      tab.style.borderBottom = selected ? "2px solid var(--aared-selection)" : "2px solid transparent";
      tab.style.fontWeight = selected ? "700" : "400";
    }
  }

  #repaintCurrent(): void {
    this.#session?.destroy();
    this.#session = null;
    this.#session = paintSheet(
      this.#model,
      this.#gridBox,
      this.#mediaResolver,
      resolveGridPaint(this.#model, this.#appearance),
      () => this.#restoreSelection(),
    );
    this.#applyZoom();
    this.#updateTabs();
  }

  #applyZoom(): void {
    const host = this.#session?.host;
    if (host === undefined) return;
    const zoom = this.#zoom;
    if (typeof zoom === "number") {
      const scale = zoom > 0 && Number.isFinite(zoom) ? zoom : 1;
      host.style.setProperty("zoom", String(scale));
    } else {
      host.style.removeProperty("zoom");
    }
  }

  #select(selection: { address: string; row: number; col: number } | null, emit: boolean): void {
    this.#selection = selection === null ? null : { ...selection };
    this.#restoreSelection();
    if (emit && this.#appearance.selection) {
      this.dispatchEvent(new CustomEvent("selectionchange", {
        detail: { sheet: this.#sheet.name, selection: this.selection },
      }));
    }
  }

  #restoreSelection(): void {
    const session = this.#session;
    if (session === null) return;
    for (const element of session.host.querySelectorAll('[data-aared-selected="true"]')) {
      const stale = element as HTMLElement;
      delete stale.dataset.aaredSelected;
      stale.removeAttribute("aria-selected");
      stale.style.outline = "";
      stale.style.outlineOffset = "";
      stale.tabIndex = -1;
    }
    const selection = this.#selection;
    if (selection === null || !this.#appearance.selection) return;
    const element = session.cellElement(selection.row, selection.col);
    // A window re-render destroys the focused cell, dropping focus to
    // <body> and killing every later keydown (the grid listener is below
    // the destroyed element). When the grid held focus (a selection exists
    // and nothing else took it), restore the roving tab stop: the selected
    // cell when it is in the window, the host fallback when it is not.
    const lostGridFocus = document.activeElement === document.body || document.activeElement === null;
    if (element === null) {
      if (lostGridFocus) session.host.focus({ preventScroll: true });
      return;
    }
    element.dataset.aaredSelected = "true";
    element.setAttribute("aria-selected", "true");
    element.style.outline = "2px solid var(--aared-selection)";
    element.style.outlineOffset = "-2px";
    element.tabIndex = 0;
    if (lostGridFocus) element.focus({ preventScroll: true });
  }

  #cellModel(row: number, col: number): CellModel {
    const found = this.#model.rows.find((entry) => entry.index === row)?.cells.find((entry) => entry.col === col)?.cell;
    return found ?? { value: null, formatted: "", type: "empty" };
  }

  #handleClick(event: MouseEvent): void {
    if (this.#destroyed) return;
    const target = event.target as HTMLElement | null;
    const link = target?.closest?.('a[data-aared-element="cell-link"]') ?? null;
    if (link !== null && link instanceof HTMLElement) {
      // Hyperlinks never navigate: prevent the default and report the
      // activation to the embedder instead.
      event.preventDefault();
      const cell = link.closest('td[data-aared-element="cell"]');
      if (cell === null || !(cell instanceof HTMLElement)) return;
      const row = Number(cell.dataset.aaredRow);
      const col = Number(cell.dataset.aaredCol);
      if (!Number.isInteger(row) || !Number.isInteger(col)) return;
      if (this.#appearance.selection) {
        this.#select({ address: cellAddress(row, col), row, col }, true);
      }
      const href = this.#cellModel(row, col).hyperlink?.target ?? link.getAttribute("href") ?? "";
      this.dispatchEvent(new CustomEvent("linkactivate", {
        detail: { sheet: this.#sheet.name, address: cellAddress(row, col), target: href },
      }));
      return;
    }
    if (!this.#appearance.selection) return;
    const cell = target?.closest?.('td[data-aared-element="cell"]') ?? null;
    if (cell === null || !(cell instanceof HTMLElement)) return;
    const row = Number(cell.dataset.aaredRow);
    const col = Number(cell.dataset.aaredCol);
    if (!Number.isInteger(row) || !Number.isInteger(col)) return;
    this.#select({ address: cellAddress(row, col), row, col }, true);
  }

  #handleDoubleClick(event: MouseEvent): void {
    if (this.#destroyed || !this.#appearance.selection) return;
    const target = event.target as HTMLElement | null;
    const cell = target?.closest?.('td[data-aared-element="cell"]') ?? null;
    if (cell === null || !(cell instanceof HTMLElement)) return;
    const row = Number(cell.dataset.aaredRow);
    const col = Number(cell.dataset.aaredCol);
    if (!Number.isInteger(row) || !Number.isInteger(col)) return;
    this.#select({ address: cellAddress(row, col), row, col }, true);
    this.#activateCell(row, col);
  }

  #activateCell(row: number, col: number): void {
    this.dispatchEvent(new CustomEvent("cellactivate", {
      detail: {
        sheet: this.#sheet.name,
        address: cellAddress(row, col),
        row,
        column: col,
        cell: this.#cellModel(row, col),
      },
    }));
  }

  #handleKeyDown(event: KeyboardEvent): void {
    if (this.#destroyed || !this.#appearance.selection) return;
    const target = event.target as HTMLElement | null;
    // A focused hyperlink keeps native activation: Enter/Space clicks the
    // link (emitting `linkactivate`); the grid must not also `cellactivate`.
    if (target?.closest?.('a[data-aared-element="cell-link"]') !== null && (event.key === "Enter" || event.key === " ")) return;
    const session = this.#session;
    if (session === null) return;
    const { geometry } = session;
    const columns = geometry.visibleColumns;
    const rows = geometry.visibleRows;
    if (columns.length === 0 || rows.length === 0) return;
    const clampIndex = (values: number[], value: number): number => {
      const exact = values.indexOf(value);
      if (exact !== -1) return exact;
      let position = 0;
      while (position < values.length && (values[position] ?? 0) < value) position++;
      return Math.min(position, values.length - 1);
    };
    const current = this.#selection ?? {
      address: cellAddress(rows[0]!, columns[0]!),
      row: rows[0]!,
      col: columns[0]!,
    };
    const rtl = this.#model.rightToLeft;
    let next: { row: number; col: number } | null = null;
    let align: "start" | "end" = "start";
    switch (event.key) {
      case "ArrowUp":
        next = { row: rows[Math.max(0, clampIndex(rows, current.row) - 1)]!, col: current.col };
        align = "start";
        break;
      case "ArrowDown":
        next = { row: rows[Math.min(rows.length - 1, clampIndex(rows, current.row) + 1)]!, col: current.col };
        align = "end";
        break;
      case "ArrowLeft":
        // Visual movement: in RTL sheets the inline-start edge is physical
        // right, so Left advances the column and Right retreats it.
        next = { row: current.row, col: columns[Math.max(0, Math.min(columns.length - 1, clampIndex(columns, current.col) + (rtl ? 1 : -1)))]! };
        align = rtl ? "end" : "start";
        break;
      case "ArrowRight":
        next = { row: current.row, col: columns[Math.max(0, Math.min(columns.length - 1, clampIndex(columns, current.col) + (rtl ? -1 : 1)))]! };
        align = rtl ? "start" : "end";
        break;
      case "Home":
        if (event.ctrlKey || event.metaKey) next = { row: rows[0]!, col: columns[0]! };
        else next = { row: current.row, col: columns[0]! };
        align = "start";
        break;
      case "End":
        if (event.ctrlKey || event.metaKey) next = { row: rows[rows.length - 1]!, col: columns[columns.length - 1]! };
        else next = { row: current.row, col: columns[columns.length - 1]! };
        align = "end";
        break;
      case "PageUp":
      case "PageDown": {
        const viewport = session.scroller.clientHeight || 600;
        const page = Math.max(1, Math.round(viewport / Math.max(((session.model.defaultRowHeight) * PIXELS_PER_INCH), 1)));
        const delta = event.key === "PageUp" ? -page : page;
        next = { row: rows[Math.max(0, Math.min(rows.length - 1, clampIndex(rows, current.row) + delta))]!, col: current.col };
        align = event.key === "PageUp" ? "start" : "end";
        break;
      }
      case "Enter":
        event.preventDefault();
        this.#select({ ...current }, true);
        this.#activateCell(current.row, current.col);
        return;
      case "Escape":
        // Escape keeps the selection: view-only grids have no edit session
        // to cancel, so there is nothing to clear back to.
        return;
      default:
        return;
    }
    event.preventDefault();
    // Clamp the column into the visible list: the current cell may sit in a
    // hidden column after a model change, and rows move within one column.
    const column = columns.includes(next.col) ? next.col : columns[clampIndex(columns, next.col)]!;
    const selection = { address: cellAddress(next.row, column), row: next.row, col: column };
    this.#select(selection, true);
    this.#ensureVisible(selection.row, selection.col, align, true);
    // The move may have scrolled and re-rendered the window, destroying the
    // focused cell. Re-focus the selected cell so the next keydown lands
    // inside the grid instead of falling through to <body>.
    this.#session?.cellElement(selection.row, selection.col)?.focus({ preventScroll: true });
  }

  /**
   * Scroll `(row, col)` into view without touching the selection. Frozen
   * cells are pinned and need no scroll on their frozen axis; hidden cells
   * cannot be revealed and are skipped. `minimal` keeps the current scroll
   * when the cell is already fully visible.
   */
  #ensureVisible(row: number, col: number, align: "start" | "center" | "end", minimal: boolean): void {
    const session = this.#session;
    if (session === null) return;
    const { geometry, scroller } = session;
    const model = session.model;
    const frozenRowSet = new Set(geometry.frozenRows);
    const frozenColumnSet = new Set(geometry.frozenColumns);

    const dataRows = geometry.visibleRows.filter((candidate) => !frozenRowSet.has(candidate));
    const rowIndex = dataRows.indexOf(row);
    if (rowIndex !== -1) {
      const offsets: number[] = [];
      let total = 0;
      for (const candidate of dataRows) {
        offsets.push(total);
        total += session.rowHeightOf(candidate);
      }
      const top = offsets[rowIndex]!;
      const height = session.rowHeightOf(row);
      const viewport = scroller.clientHeight || 600;
      const header = session.headerHeightPx;
      // Sticky headings overlay the scroller top: the visible band starts
      // past them, and centering accounts for the offset band.
      const desired = align === "center"
        ? top + height / 2 - (header + viewport) / 2
        : align === "end"
          ? top + height - viewport
          : top - header;
      const visible = top - scroller.scrollTop >= header && top + height - scroller.scrollTop <= viewport;
      if (!minimal || !visible) scroller.scrollTop = Math.max(0, desired);
    } else if (!frozenRowSet.has(row)) {
      // Hidden (or out-of-extent) rows have no rendered box to reveal.
      return;
    }

    const columnIndex = geometry.visibleColumns.indexOf(col);
    if (columnIndex === -1) {
      if (!frozenColumnSet.has(col)) return;
    } else {
      const left = geometry.columnOffsets[columnIndex]!;
      const width = (model.columnWidths[col] ?? 0) * PIXELS_PER_INCH;
      const viewportWidth = scroller.clientWidth || 800;
      const header = session.headerWidthPx;
      if (!model.rightToLeft) {
        const desired = align === "center"
          ? left + width / 2 - (header + viewportWidth) / 2
          : align === "end"
            ? left + width - viewportWidth
            : left - header;
        const visible = left - scroller.scrollLeft >= header && left + width - scroller.scrollLeft <= viewportWidth;
        if (!minimal || !visible) scroller.scrollLeft = Math.max(0, desired);
      } else {
        // RTL scroll offsets use the Chromium/Firefox negative model
        // (`scrollLeft` in [-(scrollWidth - clientWidth), 0]); Safari's
        // legacy positive-descending model may land elsewhere. The sticky
        // header column sits at the physical right (inline-start edge).
        const scrollWidth = scroller.scrollWidth;
        const min = viewportWidth >= scrollWidth ? 0 : viewportWidth - scrollWidth;
        const desired = align === "center"
          ? left + width / 2 + viewportWidth - scrollWidth - (viewportWidth - header) / 2
          : align === "end"
            ? left + viewportWidth - scrollWidth
            : left + width - scrollWidth + header;
        const point = (p: number): number => p - scroller.scrollLeft + (viewportWidth - scrollWidth);
        const visible = point(left) >= 0 && point(left + width) <= viewportWidth - header;
        if (!minimal || !visible) scroller.scrollLeft = Math.min(0, Math.max(min, desired));
      }
    }
    // Materialize the window synchronously so callers can focus the cell
    // immediately after scrolling (no rAF round-trip).
    session.renderWindow();
  }
}

function isSheetHandle(sheet: SheetHandle | SheetModel | SheetRangeModel): sheet is SheetHandle {
  return sheet instanceof SheetHandle;
}

/**
 * Paint one sheet. Accepts a `SheetHandle` (read from the Worker on demand)
 * or an already-resolved model; when the input is a handle, drawing media
 * resolves through its session unless an explicit `mediaResolver` is given.
 * Painting itself is synchronous — only the handle read awaits.
 */
export async function render(input: XlsxRenderInput): Promise<XlsxRenderHandle> {
  const requested = input.sheet;
  let mediaResolver = input.mediaResolver;
  let model: SheetModel | SheetRangeModel;
  let handle: SheetHandle | undefined;
  if (isSheetHandle(requested)) {
    handle = requested;
    model = await requested.read();
    mediaResolver ??= (mediaId: string) => requested.media(mediaId);
  } else {
    model = requested;
  }
  return new XlsxViewer(input.target, model, handle, mediaResolver, input.appearance);
}
