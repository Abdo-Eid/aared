/**
 * XLSX entrypoint: open workbook documents in a Worker, read sheet windows,
 * and paint resolved `SheetModel`s to a virtualized DOM grid — merges as
 * spanning cells, frozen panes as sticky regions, sheet-level RTL flips the
 * whole grid. Never canvas. See docs/ARCHITECTURE.md.
 *
 * Structure mirrors `aared/pptx`: `open()` with a format check, a document
 * holding a `WorkerSession`, collections implementing core's
 * `Collection<TItem, TKey>`, lightweight handles, `search()`, and an
 * input-object `render()` returning an `EventTarget`
 * controller. The grid painter from the previous revision is kept intact
 * below as the internal `paintSheet()`; only its options plumbing changed
 * (headers default on, sheet tabs default off, the benchmark-only `frame`
 * chrome is gone, gridlines follow `appearance.spreadsheet.gridlines`).
 *
 * Source-faithful content (like PPTX): the grid always paints as authored —
 * white host surface, Excel black automatic text, automatic text over an
 * authored fill resolved to the ink that reads on it. Provenance survives
 * to the model: `core/xlsx/styles.ts` `resolveFill` returns `undefined`
 * for `patternType="none"`, missing, or non-solid fills (and
 * `cellStyleFor` only assigns `fill`/`color` when a resolved value exists),
 * while `core/xlsx/sheet.ts` omits `cell.style` entirely when the resolved
 * style object is empty — so an absent `style.fill`/`style.color`
 * reliably means "automatic/default". Transparent fills let the white host
 * surface show through while every authored fill, font color, border, and
 * conditional overlay paints verbatim. Never a global inversion or filter.
 * The resolved `theme` affects viewer chrome only (tab bar, selection
 * accent); the sheet itself stays light in dark mode, like a slide.
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
  SheetChartModel,
  SheetDrawingModel,
  SheetModel,
  SheetPictureModel,
  SheetRangeModel,
} from "../core/model.js";
import type {
  Appearance,
  CellRef,
  Collection,
  DocumentInfo,
  GridlineMode,
  MediaCollection as MediaCollectionShape,
  MediaOptions,
  MediaResolver,
  OpenOptions,
  ReadOptions,
  Source,
  UnitRef,
  ViewerTheme,
  ZoomValue,
} from "../core/index.js";
import { AaredError } from "../core/index.js";
import { adoptClassSheet, classes, cssClass, type ClassSheetScope } from "../class-sheet.js";
import { clipSnippet, findByTitle, foldQuery, foldText } from "../core/search.js";
import { SessionMedia, WorkerSession } from "../core/session.js";
import type { UnitRange } from "../core/session.js";
import { itemAt, throwIfAborted } from "../core/errors.js";
import { PALETTE, automaticInk, paintTokens, resolveTheme } from "../appearance.js";
import { appendChartPrimitive } from "../chart-svg.js";
import { MediaUrls } from "../media-url.js";
import { RenderHandleBase, fitScale, measureFitRoom, safeScale } from "../render-handle.js";
import { MARK_FALLBACKS } from "../diagnostics.js";
import { PX_PER_INCH, inchesToCss } from "../core/units.js";
import { cssFontFamily } from "../core/typography.js";
import { isInsideGrid } from "../core/xlsx/limits.js";
import { cellAddress, columnName, parseA1Range } from "../core/xlsx/range.js";

/** Document metadata for a workbook; `format` is narrowed to `"xlsx"`. */
export interface XlsxInfo extends DocumentInfo {
  format: "xlsx";
}

/** Lightweight sheet identity from `open()` indexing — no sheet is parsed yet. `title` is the sheet's name. */
export interface SheetInfo extends UnitRef {
  state: "visible" | "hidden" | "veryHidden";
}

/** One matching cell from `workbook.search()`: which sheet, and which cell in it. */
export interface XlsxSearchMatch extends CellRef {
  /** The sheet holding the match; `sheet.index` is its `goTo()` key. */
  sheet: UnitRef;
  text: string;
}

/** Options for `sheet.read()`: a window of the sheet, or all of it when `range` is absent. */
export interface SheetReadOptions extends ReadOptions {
  /** An A1 range (`"A1:M200"`, or one cell) or a numeric `CellRange`. Absent reads the whole sheet. */
  range?: string | CellRange;
}

/**
 * Validate a caller-supplied numeric range (already-numeric form of an A1
 * reference). Bounds are the worksheet grid, as in `parseA1Range` and the
 * worker's window check: zero-based, start at or before end.
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
    if (!isInsideGrid(row, col)) throw invalid(`${label} is outside the worksheet grid`);
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
   * Read the sheet. No `range` resolves the whole sheet and wraps it as a
   * `SheetRangeModel` covering the full extent; a string range is parsed as
   * A1 first; a numeric `CellRange` is validated here before anything is
   * sent to the Worker. One options object, like `slide.read(options)`.
   */
  async read(options?: SheetReadOptions): Promise<SheetRangeModel> {
    const signal = options?.signal;
    const range = options?.range;
    if (range === undefined) {
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

  /** By index, or by title (exact, then ignoring case — never a substring). Hidden sheets included. */
  get(key: number | string): SheetHandle {
    if (typeof key === "number") {
      return itemAt(
        this.handles,
        key,
        () => `aared/xlsx: sheet index ${String(key)} is outside this ${this.handles.length}-sheet workbook`,
      );
    }
    const found = findByTitle(this.handles, key, (handle) => handle.info.title);
    if (found !== undefined) return found;
    throw new AaredError(
      "not-found",
      `aared/xlsx: no sheet titled ${JSON.stringify(key)} in this ${this.handles.length}-sheet workbook`,
    );
  }
}

/** Non-empty cells of one sheet as `search()` scans them, in row-major order (parallel arrays). */
interface SheetSearchCells {
  rows: number[];
  cols: number[];
  text: string[];
  lower: string[];
}

/** An opened workbook: collections, search, media, and lifecycle. */
export class XlsxDocument {
  readonly info: XlsxInfo;
  readonly sheets: SheetCollection;
  readonly media: MediaCollectionShape;
  /** Per-sheet cell text, filled by the first `search()` that reaches each sheet. */
  private readonly searchIndex = new Map<number, SheetSearchCells>();

  constructor(private readonly session: WorkerSession) {
    const info = session.info;
    this.info = { ...info, format: "xlsx" };
    const states = info.sheetStates;
    this.sheets = new SheetCollection(
      info.titles.map((title, index) => ({
        index,
        title,
        state: states?.[index] ?? "visible",
      })),
      session,
      this,
    );
    this.media = new SessionMedia(session);
  }

  /**
   * Case-insensitive substring search across every sheet, hidden sheets
   * included. The first search reads each sheet in tab order (aborting stops
   * at the next sheet) and keeps its cell text; later searches scan that text
   * without touching the Worker. Every matching cell yields one match. A cell
   * matches on its formatted text, falling back to a string cached value when
   * the formatted text is empty.
   */
  async search(query: string, options?: ReadOptions): Promise<XlsxSearchMatch[]> {
    const needle = foldQuery(query);
    if (needle === null) return [];
    const signal = options?.signal;
    const matches: XlsxSearchMatch[] = [];
    for (const handle of this.sheets.list()) {
      throwIfAborted(signal);
      const cells = await this.cellsOf(handle.info.index, signal);
      for (let at = 0; at < cells.lower.length; at++) {
        if (cells.lower[at]!.includes(needle)) {
          const row = cells.rows[at]!;
          const col = cells.cols[at]!;
          matches.push({
            sheet: { index: handle.info.index, title: handle.info.title },
            address: cellAddress(row, col),
            row,
            col,
            text: clipSnippet(cells.text[at]!),
          });
        }
      }
    }
    return matches;
  }

  /** Searchable cell text of one sheet, read once (search never paints, so no font work). */
  private async cellsOf(index: number, signal: AbortSignal | undefined): Promise<SheetSearchCells> {
    const cached = this.searchIndex.get(index);
    if (cached !== undefined) return cached;
    const sheet = await this.session.unit<SheetModel>(index, signal === undefined ? { fonts: false, retain: false } : { signal, fonts: false, retain: false });
    const cells: SheetSearchCells = { rows: [], cols: [], text: [], lower: [] };
    for (const row of sheet.rows) {
      for (const { col, cell } of row.cells) {
        const haystack = cell.formatted !== ""
          ? cell.formatted
          : typeof cell.value === "string"
            ? cell.value
            : "";
        if (haystack === "") continue;
        cells.rows.push(row.index);
        cells.cols.push(col);
        cells.text.push(haystack);
        cells.lower.push(foldText(haystack));
      }
    }
    this.searchIndex.set(index, cells);
    return cells;
  }

  /**
   * End the Worker session and reject pending reads. The Worker itself is
   * kept for the next `open()` and stops about 15 s after the last document
   * closes.
   */
  close(): void {
    this.searchIndex.clear();
    this.session.close();
  }
}

/**
 * Open an `.xlsx` file. Fails with an `unsupported-format` `AaredError`
 * when the file is not actually a spreadsheet workbook.
 */
export async function open(source: Source, options?: OpenOptions): Promise<XlsxDocument> {
  return new XlsxDocument(await WorkerSession.openAs("xlsx", source, options));
}

/** `render()` input: a sheet handle (read on demand) or an already-resolved model. */
export interface XlsxRenderInput {
  sheet: SheetHandle | SheetModel | SheetRangeModel;
  target: HTMLElement;
  appearance?: Appearance;
  /** Starting zoom, applied without a `zoomchange`. Default `1` — 100%, as Excel opens a sheet. */
  zoom?: ZoomValue;
  /** Override the default session-backed media resolution. */
  mediaResolver?: MediaResolver;
}

/** Controller returned by `render()`: zoom, navigation, selection, and lifecycle. */
export interface XlsxRenderHandle extends EventTarget {
  /** Disconnect scroll virtualization, abort pending sheet switches, release image object URLs. Idempotent. */
  destroy(): void;
  /** `true` once destroyed; every method is then a no-op that fires nothing. */
  readonly destroyed: boolean;
  /**
   * Change zoom: a number scales the grid (default `1`); `fit-width` scales
   * the sheet's used columns to the container's width and `fit-page` also
   * keeps its used rows in view — never below 25%, so a huge sheet still
   * scrolls instead of shrinking to nothing. Fits stay current on resize.
   * Fires `zoomchange` only when the value changes.
   */
  setZoom(zoom: ZoomValue): void;
  /** The requested zoom. */
  readonly zoom: ZoomValue;
  /** The scale actually painted — the number, or what the fit mode computed. */
  readonly scale: number;
  /**
   * Replace the viewer appearance in place: repaints the current sheet with
   * the new theme/chrome knobs while keeping the sheet, selection, zoom and
   * scroll position. Replaces, never merges: pass the whole object, or
   * `spreadsheet` knobs you leave out return to their defaults.
   */
  setAppearance(appearance: Appearance): void;
  /** Select one cell (`"D14"`; ranges reject `invalid-range`), scroll it into view if needed, and focus it. */
  goToCell(cell: string): Promise<void>;
  /**
   * Scroll a cell or range into view without changing the selection. Needs
   * exactly one of `cell` (a single cell) or `range`; anything else throws
   * `invalid-range`.
   */
  scrollTo(target: { cell?: string; range?: string; align?: "start" | "center" | "end" }): void;
  /** Select a single cell and DOM-focus it with `preventScroll` (no scrolling). */
  focus(cell: string): void;
  /**
   * Switch the visible sheet in place by index or title (hidden sheets
   * included); fires `sheetchange` with `{ index, title }`. Rejects
   * `not-found` for an unknown key.
   */
  goToSheet(key: string | number): Promise<void>;
  /** The currently visible sheet. */
  readonly sheet: SheetInfo;
  /** The single selected cell, or `null`. */
  readonly selection: CellRef | null;
}

const GRIDLINE_COLOR = "#E7E9EE";
const GRIDLINE_DARK = "#3B3B45";
const ROW_OVERSCAN = 8;
/** Heading chrome dimensions. Viewer chrome, not authored geometry. */
const HEADER_COL_WIDTH_PX = 40;
/** Excel's default column width (8.43 characters), used for viewport-fill columns. */
const FILL_COLUMN_WIDTH_PX = 64;
/** A fill sliver thinner than this is skipped; the cap bounds a huge viewport. */
const FILL_MIN_TRACK_PX = 6;
/** Below this height a fill row is too short to show its row number. */
const FILL_ROW_TEXT_PX = 16;
const FILL_MAX_TRACKS = 500;
const HEADER_ROW_HEIGHT_PX = 22;
const HEADER_FILL = "#F4F4F7";
const HEADER_COLOR = "#5C5C64";
const HEADER_FILL_DARK = "#2A2B33";
const HEADER_COLOR_DARK = "#C9CAD1";

// The grid's default look, as class rules in the render root's one adopted
// sheet (see class-sheet.ts): a data or fill cell carries a class name, not
// ~20 inline longhands, and inline style keeps only what a cell authors.
// The lines read the tokens the host publishes (`paintSheetTokens`). Bases
// are declared before their modifiers, so a modifier wins.
const GRIDLINE = `var(--aared-gridline,${GRIDLINE_COLOR})`;
const CELL_BOX = `padding:1px 4px;border:1px solid transparent;border-right-color:${GRIDLINE};border-bottom-color:${GRIDLINE};font-size:10pt;line-height:1`;
/** A data cell. Single-line cells never outgrow their authored row: see `paintCell`. */
const CELL = cssClass(`${CELL_BOX};color:#000000;vertical-align:bottom;white-space:nowrap;overflow:hidden`);
const CELL_WRAP = cssClass("line-height:1.1;white-space:pre-wrap;overflow-wrap:break-word;overflow:visible");
/** A viewport-fill cell: inert, the same box as a data cell. */
const FILL_CELL = cssClass(`${CELL_BOX};overflow:hidden`);
/** One fill cell standing for a row's run of fill columns: their lines drawn at the fill-column pitch. */
const FILL_TRACKS = cssClass(
  `background-origin:border-box;background-image:repeating-linear-gradient(to right,transparent 0 calc(${FILL_COLUMN_WIDTH_PX}px - var(--aared-hairline,1px)),${GRIDLINE} calc(${FILL_COLUMN_WIDTH_PX}px - var(--aared-hairline,1px)) ${FILL_COLUMN_WIDTH_PX}px)`,
);
const FILL_TRACKS_RTL = cssClass(
  `background-origin:border-box;background-image:repeating-linear-gradient(to left,${GRIDLINE} 0 var(--aared-hairline,1px),transparent var(--aared-hairline,1px) ${FILL_COLUMN_WIDTH_PX}px)`,
);
/** One fill cell standing for a whole empty row: every column's line, from per-sheet custom properties (`syncFillLines`). */
const FILL_ROW = cssClass(
  "background-origin:border-box;background-repeat:no-repeat;background-image:var(--aared-fill-lines,none);background-size:var(--aared-fill-lines-size,auto);background-position:var(--aared-fill-lines-position,0 0)",
);
const NO_GRIDLINES = cssClass("border-right-color:transparent;border-bottom-color:transparent;background-image:none");
/** Text that must sit above a data bar. */
const CELL_TEXT_OVER = cssClass("position:relative;z-index:1");
const CELL_LINK = cssClass("text-decoration:none");
const DATA_BAR = cssClass("position:absolute;inset-block:0;inset-inline-start:0;opacity:0.45;pointer-events:none");
const SPILL = cssClass("position:absolute;top:0;bottom:0;box-sizing:border-box;padding:1px 4px;display:flex;overflow:hidden;pointer-events:none");
const SPILL_TEXT = cssClass("white-space:nowrap;pointer-events:auto");
/**
 * Headings share the grid's own hairlines — bottom/right only, never a full
 * box — so the reference headers blend with the grid instead of buttoning
 * it; they stick to their axis above frozen cells.
 */
const HEADING = cssClass(
  `position:sticky;z-index:4;background-color:var(--aared-header-fill,${HEADER_FILL});color:var(--aared-header-color,${HEADER_COLOR});border-bottom:1px solid ${GRIDLINE};border-right:1px solid ${GRIDLINE};font-size:9pt;line-height:1;font-weight:600;text-align:center;vertical-align:middle;white-space:nowrap;user-select:none`,
);
const HEADING_TOP = cssClass("top:0px");
const HEADING_CORNER = cssClass("z-index:5");
/** Headings stick to the inline-start edge: physical `right` in RTL. */
const HEADING_LEFT = cssClass("left:0px");
const HEADING_RIGHT = cssClass("right:0px");
const HEADING_CLIP = cssClass("overflow:hidden");
const SPACER = cssClass("padding:0;border:0");
/** HTML clamps `colSpan` at 1000; a wider fill row keeps one cell per column. */
const MAX_COL_SPAN = 1000;

/**
 * The tokens only a grid publishes, on top of the shared set: its lines, its
 * header band and its selection (the shared accent). The tab bar sits outside
 * the grid host, so it carries its own copy.
 */
function paintSheetTokens(
  element: HTMLElement,
  tokens: { dark: boolean; gridline: string; headerFill: string; headerColor?: string },
): void {
  element.style.setProperty("--aared-gridline", tokens.gridline);
  element.style.setProperty("--aared-header-fill", tokens.headerFill);
  if (tokens.headerColor !== undefined) element.style.setProperty("--aared-header-color", tokens.headerColor);
  element.style.setProperty("--aared-selection", PALETTE[tokens.dark ? "dark" : "light"].accent);
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

/** Resolved paint parameters: source-faithful content plus theme-driven chrome. */
interface GridPaint {
  showHeaders: boolean;
  /** Continue the grid with inert empty cells until the viewport is full. */
  fillViewport: boolean;
  showGridlines: boolean;
  gridlineColor: string;
  /**
   * Automatic cell text, when neither the style nor a conditional overlay names
   * a color: over an authored (or conditional) `fill` it is whichever ink reads
   * on that fill, over the host surface (`fill` undefined) Excel black.
   */
  automaticColor: (fill: string | undefined) => string;
  /** Opaque fill stamped behind sticky frozen cells whose own fill is transparent. */
  frozenFill: string;
  headerFill: string;
  headerColor: string;
  /** Chrome line/text colors for notices and drawing placeholders. */
  chromeLine: string;
  noticeColor: string;
  hostBackground: string;
  dark: boolean;
}

/**
 * Where an overflowing text cell paints: the pixel box of its run of blank
 * neighbours, and how far that box starts left of the cell's own edge. Excel
 * lets unwrapped text spill over empty cells until it meets content.
 */
interface TextSpill {
  /** Width of the whole run (cell plus blank neighbours), in CSS pixels. */
  width: number;
  /** Physical distance from the run's left edge to the cell's left edge. */
  offset: number;
}

/** Where a cell's text sits, as `paintCell` resolved it (read by the spill overlay). */
interface CellPlacement {
  align: string;
  vertical: string;
}

/** Nothing to show: no text, no link, no data bar. Such a cell paints no content element. */
function isEmptyCell(cell: CellModel): boolean {
  return cell.formatted === "" && cell.hyperlink === undefined && cell.conditional?.dataBar === undefined;
}

function appendCellContent(cellElement: HTMLElement, cell: CellModel, placement: CellPlacement, spill?: TextSpill): void {
  const dataBar = cell.conditional?.dataBar;
  if (dataBar !== undefined) {
    const bar = document.createElement("div");
    bar.setAttribute("data-aared-element", "data-bar");
    bar.setAttribute("aria-hidden", "true");
    bar.setAttribute("class", DATA_BAR);
    bar.style.width = `${Math.min(Math.max(dataBar.ratio, 0), 1) * 100}%`;
    bar.style.backgroundColor = dataBar.color;
    cellElement.append(bar);
  }
  if (cell.formatted === "" && cell.hyperlink === undefined) return;
  const text = document.createElement("span");
  text.setAttribute("data-aared-element", "cell-text");
  if (dataBar !== undefined) text.setAttribute("class", CELL_TEXT_OVER);
  if (cell.hyperlink === undefined) {
    text.textContent = cell.formatted;
  } else {
    const link = document.createElement("a");
    link.setAttribute("data-aared-element", "cell-link");
    link.setAttribute("class", CELL_LINK);
    link.href = cell.hyperlink.target;
    link.textContent = cell.formatted === "" ? (cell.hyperlink.display ?? cell.hyperlink.target) : cell.formatted;
    // Native print shows no underline: link text keeps the cell's own color
    // (blue only when the cell carries no explicit color, as in Excel).
    // Clicks never navigate: the viewer intercepts them and emits
    // `linkactivate` instead (see XlsxViewer).
    link.style.color = cell.style?.color ?? "#1155CC";
    text.append(link);
  }
  if (spill === undefined) {
    cellElement.append(text);
    return;
  }
  // The cell's own box would clip the text, so it is painted in an overlay
  // that covers the whole run and is clipped at the run's far edge instead.
  const overlay = document.createElement("div");
  overlay.setAttribute("data-aared-element", "cell-spill");
  overlay.setAttribute("class", SPILL);
  overlay.dir = "ltr";
  overlay.style.left = `${-spill.offset - 1}px`;
  overlay.style.width = `${spill.width}px`;
  overlay.style.justifyContent = placement.align === "right" ? "flex-end" : placement.align === "center" ? "center" : "flex-start";
  overlay.style.alignItems = placement.vertical === "top" ? "flex-start" : placement.vertical === "middle" ? "center" : "flex-end";
  text.dir = cellElement.dir;
  text.setAttribute("class", classes(text.getAttribute("class"), SPILL_TEXT));
  overlay.append(text);
  cellElement.append(overlay);
}

/**
 * Paint a data cell: the default look is the `CELL` class, and inline style
 * carries only what the cell authors (font, fill, colour, alignment,
 * borders). A single-line cell must never be taller than its authored row:
 * with `normal` leading the embedder's font decides the line box and rows
 * drift taller than the model, which misplaces everything positioned from row
 * offsets (charts, virtualization spacers), so `CELL` sets `line-height: 1`;
 * wrapped text gets just enough leading (`CELL_WRAP`) for two lines to fit a
 * row Excel sized for them, cell padding and borders included.
 */
function paintCell(cellElement: HTMLElement, cell: CellModel, sheetRtl: boolean, paint: GridPaint): CellPlacement {
  const style = cell.style;
  const wrap = style?.wrapText === true;
  // Default gridlines are content chrome resolved per gridline mode; an
  // authored border always wins below. Hidden gridlines stay transparent
  // (never `none`) so authored and default borders share one box metric.
  cellElement.setAttribute("class", classes(CELL, wrap && CELL_WRAP, !paint.showGridlines && NO_GRIDLINES));
  if (style?.fontSizePt !== undefined && style.fontSizePt !== 10) cellElement.style.fontSize = `${style.fontSizePt}pt`;
  if (style?.fontFamily !== undefined) cellElement.style.fontFamily = cssFontFamily(style.fontFamily);
  if (style?.bold === true || cell.conditional?.bold === true) cellElement.style.fontWeight = "700";
  if (style?.italic === true) cellElement.style.fontStyle = "italic";
  if (style?.underline === true) cellElement.style.textDecoration = "underline";
  // Fills are never defaulted: no fill lets the white host surface show
  // through, preserving provenance.
  const fill = cell.conditional?.fill ?? style?.fill;
  if (fill !== undefined) cellElement.style.backgroundColor = fill;
  // An authored or conditional color paints verbatim; automatic text is
  // chosen for the fill it sits on (Excel black, `CELL`'s own, without one).
  const color = cell.conditional?.color ?? style?.color ?? paint.automaticColor(fill);
  if (color !== paint.automaticColor(undefined)) cellElement.style.color = color;
  const placement: CellPlacement = { align: textAlign(cell, sheetRtl), vertical: style?.verticalAlign ?? "bottom" };
  const hasContent = !isEmptyCell(cell) || style?.indent !== undefined;
  if (hasContent) cellElement.style.textAlign = placement.align;
  if (placement.vertical !== "bottom") cellElement.style.verticalAlign = placement.vertical;
  if (style?.indent !== undefined) cellElement.style.paddingInlineStart = `calc(4px + ${style.indent * 8}px)`;
  // Sheet RTL flips column order via the table's `dir`, but cell text follows
  // Excel's per-cell reading order: 1 = LTR, 2 = RTL, otherwise context. An
  // unspecified order must not inherit the sheet direction, or LTR values
  // such as "180,000 EGP" and "13 Sep 2026" visually reorder in RTL sheets
  // (and a data bar or indent would flip side). An empty cell has nothing to
  // order.
  if (style?.readingOrder === 2) cellElement.dir = "rtl";
  else if (style?.readingOrder === 1) cellElement.dir = "ltr";
  else if (hasContent) cellElement.dir = "auto";
  const border = style?.border;
  if (border?.top !== undefined) cellElement.style.borderTop = `${borderWidth(border.top.style)} ${borderStyle(border.top.style)} ${border.top.color ?? "currentColor"}`;
  if (border?.left !== undefined) cellElement.style.borderLeft = `${borderWidth(border.left.style)} ${borderStyle(border.left.style)} ${border.left.color ?? "currentColor"}`;
  if (border?.right !== undefined) cellElement.style.borderRight = `${borderWidth(border.right.style)} ${borderStyle(border.right.style)} ${border.right.color ?? "currentColor"}`;
  if (border?.bottom !== undefined) cellElement.style.borderBottom = `${borderWidth(border.bottom.style)} ${borderStyle(border.bottom.style)} ${border.bottom.color ?? "currentColor"}`;
  if (cell.formula !== undefined) cellElement.setAttribute("data-aared-formula", cell.formula);
  if (cell.type === "error") cellElement.setAttribute("data-aared-cell-error", cell.formatted);
  return placement;
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
  const columnWidths = visibleColumns.map((col) => (model.columnWidths[col] ?? 0) * PX_PER_INCH);
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
  /**
   * The sheet's used size at 100%, in the host's own pixels: the columns and
   * rows the sheet has plus the headings — never the viewport-fill cells,
   * which exist only to fill whatever room is left.
   */
  naturalSize(): { width: number; height: number };
  /**
   * How far the painted sheet overflows the scroller vertically, in host
   * pixels, leaving out the viewport-fill rows. Painted rows can outgrow the
   * model below 100%: every border snaps to at least a device pixel.
   */
  paintedOverflowY(): number;
  /**
   * Tell the session the CSS zoom the controller set on `host`. Everything
   * inside the host mixes only with itself, except the room borrowed from an
   * outer clipping ancestor, which is converted into the host's space with it.
   */
  setScale(scale: number): void;
}

/** Paint one already-resolved sheet. The caller owns which sheet is visible. */
function paintSheet(
  model: SheetModel,
  container: HTMLElement,
  mediaResolver: MediaResolver | undefined,
  paint: GridPaint,
  onWindow?: () => void,
): SheetPaintSession {
  const urls = mediaResolver === undefined ? undefined : new MediaUrls(mediaResolver);
  let destroyed = false;
  let raf = 0;

  const heightByRow = new Map<number, number>();
  for (const row of model.rows) {
    if (row.height !== undefined) heightByRow.set(row.index, row.height);
  }
  const rowHeightOf = (row: number): number => (heightByRow.get(row) ?? model.defaultRowHeight) * PX_PER_INCH;
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
  // Fill a bounded embedder's height so the grid below can scroll; without
  // a bounded ancestor the grid adopts the nearest clipping ancestor's
  // visible room as its own minimum height (see syncViewport), and only
  // then can virtualization window — otherwise every row renders whenever
  // the whole sheet fits, still correct.
  host.style.height = "100%";
  host.style.display = "flex";
  host.style.flexDirection = "column";
  host.style.background = paint.hostBackground;
  // Theme is viewer chrome state (like PPTX): the sheet itself always paints
  // source-faithful, so the host carries light content tokens. The tab bar
  // outside the host carries its own theme-following copy. Never a filter.
  host.dataset.aaredTheme = paint.dark ? "dark" : "light";
  paintTokens(host, "light");
  paintSheetTokens(host, {
    dark: false,
    gridline: paint.gridlineColor,
    headerFill: paint.headerFill,
    headerColor: paint.headerColor,
  });
  // The host holds the single tab stop: cells are programmatically
  // focusable (`tabindex="-1"`) but virtualization unmounts off-window
  // rows, so a roving tabindex on cells would drop the tab stop whenever
  // the selection scrolls out of the window.
  host.tabIndex = 0;
  host.setAttribute("aria-label", model.name);

  // The notice bar (formulas never recalculated, unsupported features) is a
  // development diagnostic: shipped bundles never paint it, so consumers
  // see only the sheet. The model keeps `unsupported` for hosts that want
  // to surface it themselves.
  const hasFormulas = MARK_FALLBACKS && model.rows.some((row) => row.cells.some(({ cell }) => cell.formula !== undefined));
  if (MARK_FALLBACKS && (model.unsupported.length > 0 || hasFormulas)) {
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
  // A right-to-left sheet starts at the physical right: the grid hugs that
  // edge (any spare room is on the left) and scrolling begins at column A.
  if (model.rightToLeft) scroller.dir = "rtl";
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
    element.style.width = inchesToCss(model.columnWidths[col] ?? 0);
    colgroup.append(element);
  }
  table.append(colgroup);

  /** A heading cell's classes: sticky on its axis, above frozen cells (`HEADING`). */
  const headingClass = (kind: "corner" | "row" | "column", clip = false): string =>
    classes(
      HEADING,
      (kind === "corner" || kind === "column") && HEADING_TOP,
      kind === "corner" && HEADING_CORNER,
      (kind === "corner" || kind === "row") && (model.rightToLeft ? HEADING_RIGHT : HEADING_LEFT),
      clip && HEADING_CLIP,
    );
  const paintHeading = (element: HTMLElement, kind: "corner" | "row" | "column", clip = false): void => {
    element.setAttribute("class", headingClass(kind, clip));
  };

  let headRow: HTMLTableRowElement | undefined;
  if (showHeaders) {
    // Letters skip hidden columns exactly like Excel: …C, E… with D hidden.
    const head = document.createElement("thead");
    head.dataset.aaredElement = "column-headers";
    headRow = document.createElement("tr");
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

  // ---- Viewport fill -------------------------------------------------------
  // Excel's grid never ends at the data. When the viewer is bigger than the
  // sheet, keep drawing empty default-size cells (with continuing headings)
  // until the viewport is full. Viewer chrome only: the cells are inert,
  // aria-hidden, and never enter the model, extent, selection or navigation.
  const realWidthPx = geometry.totalWidth + headerWidthPx;
  const realHeightPx = geometry.totalHeight + headerHeightPx;
  const fillRowPx = Math.max(model.defaultRowHeight * PX_PER_INCH, 1);
  let fillColumnWidths: number[] = [];
  let fillRowHeights: number[] = [];

  /** Split a pixel deficit into default-size tracks; the last one is trimmed to fit exactly. */
  const fillTracks = (deficit: number, size: number): number[] => {
    const tracks: number[] = [];
    let remaining = deficit;
    // Under a few pixels the sliver is invisible and would only risk a scrollbar.
    while (remaining >= FILL_MIN_TRACK_PX && tracks.length < FILL_MAX_TRACKS) {
      const track = Math.min(size, remaining);
      tracks.push(track);
      remaining -= track;
    }
    return tracks;
  };
  const sameTracks = (a: number[], b: number[]): boolean => a.length === b.length && a.every((value, i) => value === b[i]);

  // Viewport selection cache: the content probe forces layout and the
  // external room walks ancestors, so both are reused while the scroller's
  // metrics hold. Cleared by the resize observer and the window resize
  // listener, which are the only ways the room can change unnoticed.
  let viewportCache: { unbounded: boolean; roomH: number; clientH: number; scrollH: number } | null = null;
  let observedClipper: Element | null = null;
  /** The CSS zoom on `host` (see `setScale`): outer room divides by it to land in host pixels. */
  let scale = 1;

  /**
   * True when nothing bounds the scroller's height, so it sizes itself to
   * its content. Detected by perturbing the content and watching whether
   * `clientHeight` follows: a bounded viewport ignores the perturbation,
   * a content-sized one absorbs it. The height axis is the dangerous one —
   * `table-layout: fixed` honors column widths exactly, so fill columns
   * have no equivalent feedback loop.
   */
  const scrollerIsUnbounded = (clientH: number): boolean => {
    const probe = document.createElement("div");
    probe.style.height = "10px";
    probe.style.width = "0";
    probe.style.padding = "0";
    probe.style.border = "0";
    probe.style.margin = "0";
    probe.style.pointerEvents = "none";
    canvas.append(probe);
    const unbounded = scroller.clientHeight >= clientH + 9;
    probe.remove();
    return unbounded;
  };

  /** Nearest ancestor that clips vertically (`overflow-y` resolved). */
  const findHeightClipper = (): HTMLElement | null => {
    let node: Element | null = scroller.parentElement;
    while (node !== null && node !== document.documentElement) {
      if (node instanceof HTMLElement) {
        const overflowY = getComputedStyle(node).overflowY;
        if (overflowY === "auto" || overflowY === "scroll" || overflowY === "hidden" || overflowY === "clip") {
          return node;
        }
      }
      node = node.parentElement;
    }
    return null;
  };

  /**
   * Visible room below the grid's top edge in the nearest clipping
   * ancestor's layout space (or the document viewport when there is none).
   * Scroll-invariant: ancestor and page scrolls move both rects together,
   * so painting fill can never change this measurement — the fill
   * feedback loop is structurally impossible. CSS `zoom` on an ancestor
   * mixes visual rects with layout pixels, so the room is approximate
   * there (stable, never growing).
   */
  const externalRoomHeight = (): { roomH: number; clipper: HTMLElement | null } => {
    const scrollerTop = scroller.getBoundingClientRect().top;
    const clipper = findHeightClipper();
    if (clipper === null) {
      const visibleH = document.documentElement.clientHeight;
      const scroll = window.scrollY ?? document.documentElement.scrollTop;
      return { roomH: visibleH - (scrollerTop + scroll), clipper: null };
    }
    const clipRect = clipper.getBoundingClientRect();
    const clipStyle = getComputedStyle(clipper);
    const px = (value: string): number => Number.parseFloat(value || "0") || 0;
    const contentTop = clipRect.top + px(clipStyle.borderTopWidth) + px(clipStyle.paddingTop);
    const offset = scrollerTop - contentTop + clipper.scrollTop;
    // A clipper sized by its content up to a `max-height` (a common scroll
    // box) is as tall as the grid inside it right now: measuring that as the
    // room grew the fill a row per frame until the limit stopped it. Its
    // room is the limit.
    let clientH = clipper.clientHeight;
    const maxH = px(clipStyle.maxHeight);
    if (clipStyle.maxHeight.endsWith("px") && maxH > 0) {
      const limit = clipStyle.boxSizing === "border-box"
        ? maxH - px(clipStyle.borderTopWidth) - px(clipStyle.borderBottomWidth)
        : maxH + px(clipStyle.paddingTop) + px(clipStyle.paddingBottom);
      clientH = Math.max(clientH, limit);
    }
    return { roomH: clientH - offset, clipper };
  };

  /**
   * Pick the fill viewport for this pass. A bounded scroller's own box is
   * the viewport; a content-sized scroller adopts the nearest clipping
   * ancestor's visible room as its minimum height, so the grid fills the
   * space it is rendered in. Returns the fill deficit's baseline height
   * plus the windowing viewport. The baseline never includes painted
   * content: a bounded box is flex/fixed-determined, the external room is
   * scroll-invariant, so fill cannot feed back into itself on any axis.
   */
  const syncViewport = (): { fillH: number; windowH: number; windowW: number } => {
    scroller.style.minHeight = "";
    if (!paint.fillViewport) {
      return {
        fillH: 0,
        windowH: scroller.clientHeight === 0 ? 600 : scroller.clientHeight,
        windowW: scroller.clientWidth,
      };
    }
    const clientH = scroller.clientHeight;
    const scrollH = scroller.scrollHeight;
    let cached = viewportCache;
    if (
      cached === null ||
      Math.abs(cached.clientH - clientH) > 2 ||
      Math.abs(cached.scrollH - scrollH) > 2
    ) {
      const unbounded = scrollerIsUnbounded(clientH);
      const { roomH, clipper } = unbounded ? externalRoomHeight() : { roomH: 0, clipper: null };
      if (clipper !== observedClipper) {
        if (observedClipper !== null) resizer.unobserve(observedClipper);
        observedClipper = clipper;
        if (clipper !== null) resizer.observe(clipper);
      }
      cached = { unbounded, roomH, clientH, scrollH };
      viewportCache = cached;
    }
    // The external room is outer (unzoomed) space; everything it is compared
    // with lives inside the zoomed host, so it is converted into host pixels.
    const roomH = cached.roomH / scale;
    if (cached.unbounded && roomH > clientH) {
      const floor = `${Math.floor(roomH)}px`;
      if (scroller.style.minHeight !== floor) scroller.style.minHeight = floor;
    }
    const windowW = scroller.clientWidth;
    const boxH = scroller.clientHeight;
    return {
      fillH: cached.unbounded ? roomH : clientH,
      windowH: boxH === 0 ? 600 : boxH,
      windowW,
    };
  };

  /**
   * An inert fill cell. `span` > 1 stands for several fill tracks at once:
   * one cell per data row for the row's run of fill columns (`tracks`), one
   * per fill row for the whole empty row (`row`), their inner lines drawn as
   * backgrounds so the fill costs one element per row, not one per column.
   */
  const paintFillCell = (row: number | undefined, span = 1, kind?: "tracks" | "row"): HTMLTableCellElement => {
    const cell = document.createElement("td");
    cell.setAttribute("data-aared-element", "fill-cell");
    cell.setAttribute("aria-hidden", "true");
    const lines = kind === "tracks" ? (model.rightToLeft ? FILL_TRACKS_RTL : FILL_TRACKS) : kind === "row" ? FILL_ROW : undefined;
    cell.setAttribute("class", classes(FILL_CELL, lines, !paint.showGridlines && NO_GRIDLINES));
    if (kind !== undefined) cell.setAttribute("data-aared-fill", kind);
    if (span > 1) cell.colSpan = span;
    if (row !== undefined) paintFrozen(cell, row, -1);
    return cell;
  };

  /**
   * The lines a whole-row fill cell draws, as custom properties on the table:
   * each real column's right edge (the columns are not uniform, so they are
   * explicit stops), then the fill tracks at their pitch. A line is one
   * device pixel thick at the current zoom (`--aared-hairline`), like the
   * cell borders it continues. RTL mirrors: lines on each track's physical
   * right edge, laid out from the right.
   */
  const syncFillLines = (): void => {
    const dpr = typeof window === "undefined" ? 1 : window.devicePixelRatio || 1;
    const device = scale * dpr;
    const hair = Math.max(1, Math.floor(device)) / device;
    table.style.setProperty("--aared-hairline", `${hair}px`);
    const rtl = model.rightToLeft;
    const stops: string[] = [];
    geometry.visibleColumns.forEach((col, i) => {
      const x = geometry.columnOffsets[i]!;
      const w = columnWidthPx(col);
      // LTR from the left edge: the line is the column's last pixel. RTL from
      // the right edge: its first.
      const [a, b] = rtl ? [x, x + hair] : [x + w - hair, x + w];
      stops.push(`transparent ${a}px`, `${GRIDLINE} ${a}px ${b}px`, `transparent ${b}px`);
    });
    const direction = rtl ? "to left" : "to right";
    const layers: string[] = [];
    const sizes: string[] = [];
    const positions: string[] = [];
    if (stops.length > 0) {
      layers.push(`linear-gradient(${direction},${stops.join(",")})`);
      sizes.push(`${geometry.totalWidth}px 100%`);
      positions.push(rtl ? "right 0 top 0" : "left 0 top 0");
    }
    const fillWidth = fillColumnWidths.reduce((sum, width) => sum + width, 0);
    if (fillWidth > 0) {
      layers.push(
        rtl
          ? `repeating-linear-gradient(to left,${GRIDLINE} 0 ${hair}px,transparent ${hair}px ${FILL_COLUMN_WIDTH_PX}px)`
          : `repeating-linear-gradient(to right,transparent 0 ${FILL_COLUMN_WIDTH_PX - hair}px,${GRIDLINE} ${FILL_COLUMN_WIDTH_PX - hair}px ${FILL_COLUMN_WIDTH_PX}px)`,
      );
      sizes.push(`${fillWidth}px 100%`);
      positions.push(rtl ? `right ${geometry.totalWidth}px top 0` : `left ${geometry.totalWidth}px top 0`);
    }
    table.style.setProperty("--aared-fill-lines", layers.length === 0 ? "none" : layers.join(","));
    table.style.setProperty("--aared-fill-lines-size", sizes.length === 0 ? "auto" : sizes.join(","));
    table.style.setProperty("--aared-fill-lines-position", positions.length === 0 ? "0 0" : positions.join(","));
  };

  /** Rebuild the trailing filler `<col>`s and column headings after the fill width changed. */
  const syncFillColumns = (): void => {
    for (const stale of table.querySelectorAll('[data-aared-fill="column"]')) stale.remove();
    let extra = 0;
    fillColumnWidths.forEach((width, i) => {
      extra += width;
      const col = document.createElement("col");
      col.dataset.aaredFill = "column";
      col.style.width = `${width}px`;
      colgroup.append(col);
      if (headRow === undefined) return;
      const letter = document.createElement("th");
      letter.setAttribute("data-aared-element", "fill-column-header");
      letter.setAttribute("data-aared-fill", "column");
      letter.setAttribute("aria-hidden", "true");
      letter.textContent = columnName(model.extent.cols + i);
      paintHeading(letter, "column", true);
      headRow.append(letter);
    });
    const width = `${realWidthPx + extra}px`;
    table.style.width = width;
    canvas.style.width = width;
    syncFillLines();
  };

  /** Recompute how much fill the current viewport needs; true when it changed. */
  const measureFill = (viewportWidth: number, viewportHeight: number): boolean => {
    if (!paint.fillViewport) return false;
    const columns = viewportWidth > 0 ? fillTracks(viewportWidth - realWidthPx, FILL_COLUMN_WIDTH_PX) : [];
    const rows = viewportHeight > 0 ? fillTracks(viewportHeight - realHeightPx, fillRowPx) : [];
    const columnsChanged = !sameTracks(columns, fillColumnWidths);
    const changed = columnsChanged || !sameTracks(rows, fillRowHeights);
    fillColumnWidths = columns;
    fillRowHeights = rows;
    if (columnsChanged) syncFillColumns();
    return changed;
  };

  const paintFrozen =(element: HTMLElement, row: number, col: number): void => {
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
    element.setAttribute("data-aared-frozen", frozenRow && frozenColumn ? "corner" : frozenRow ? "row" : "column");
    if (element.style.backgroundColor === "" || element.style.backgroundColor === "transparent") {
      element.style.backgroundColor = paint.frozenFill;
    }
  };

  const columnPosition = new Map<number, number>();
  geometry.visibleColumns.forEach((col, position) => columnPosition.set(col, position));
  const columnWidthPx = (col: number): number => (model.columnWidths[col] ?? 0) * PX_PER_INCH;

  /** Nothing to paint and not part of a merge: text may run across it. */
  const isBlank = (row: number, col: number, entries: Map<number, CellModel> | undefined): boolean => {
    const key = `${row}:${col}`;
    if (merges.originOf.has(key) || merges.rangeOf.has(key)) return false;
    const cell = entries?.get(col);
    return cell === undefined || (cell.value === null && cell.formatted === "" && cell.conditional?.dataBar === undefined);
  };

  /**
   * Excel lets unwrapped text overflow into empty neighbours (toward the end
   * edge for left-aligned text, the start edge for right-aligned, both for
   * centered) and stops at the first cell with content. Numbers never spill.
   */
  const spillsForRow = (
    row: number,
    entries: Map<number, CellModel> | undefined,
  ): { spills: Map<number, TextSpill>; crossed: Set<number> } => {
    const spills = new Map<number, TextSpill>();
    const crossed = new Set<number>();
    if (entries === undefined) return { spills, crossed };
    const rtl = model.rightToLeft;
    for (const [col, cell] of entries) {
      const position = columnPosition.get(col);
      if (position === undefined || typeof cell.value !== "string" || cell.formatted === "") continue;
      if (cell.style?.wrapText === true || merges.rangeOf.has(`${row}:${col}`) || merges.originOf.has(`${row}:${col}`)) continue;
      const align = textAlign(cell, rtl);
      if (align !== "left" && align !== "right" && align !== "center") continue;
      // Neighbours on each physical side, in visible-column order.
      const run = (step: 1 | -1): number[] => {
        const cols: number[] = [];
        for (let next = position + step; next >= 0 && next < geometry.visibleColumns.length; next += step) {
          const candidate = geometry.visibleColumns[next]!;
          if (!isBlank(row, candidate, entries)) break;
          cols.push(candidate);
        }
        return cols;
      };
      const higher = run(1);
      const lower = run(-1);
      // Physical right is the higher column in an LTR sheet, the lower in RTL.
      const physicalRight = rtl ? lower : higher;
      const physicalLeft = rtl ? higher : lower;
      const right = align === "left" || align === "center" ? physicalRight : [];
      const left = align === "right" || align === "center" ? physicalLeft : [];
      if (left.length === 0 && right.length === 0) continue;
      const offset = left.reduce((sum, c) => sum + columnWidthPx(c), 0);
      const width = offset + columnWidthPx(col) + right.reduce((sum, c) => sum + columnWidthPx(c), 0);
      spills.set(col, { width, offset });
      // Everything in the run except its right-most cell loses its right gridline.
      for (const c of left) crossed.add(c);
      if (right.length > 0) crossed.add(col);
      for (const c of right.slice(0, -1)) crossed.add(c);
    }
    return { spills, crossed };
  };

  const paintRow = (row: number): HTMLTableRowElement => {
    const line = document.createElement("tr");
    line.setAttribute("data-aared-row", String(row));
    line.setAttribute("role", "row");
    line.setAttribute("aria-rowindex", String(row + 1));
    line.style.height = `${rowHeightOf(row)}px`;
    if (showHeaders) {
      // Numbers skip hidden rows exactly like Excel: …6, 8… with 7 hidden.
      const number = document.createElement("th");
      number.setAttribute("data-aared-element", "row-header");
      number.setAttribute("data-aared-row", String(row));
      number.scope = "row";
      number.setAttribute("role", "rowheader");
      number.textContent = String(row + 1);
      paintHeading(number, "row");
      line.append(number);
    }
    const entries = cellsByRow.get(row);
    const { spills, crossed } = spillsForRow(row, entries);
    for (const col of geometry.visibleColumns) {
      const covered = merges.originOf.get(`${row}:${col}`);
      if (covered !== undefined) continue;
      const cellElement = document.createElement("td");
      cellElement.setAttribute("data-aared-element", "cell");
      cellElement.setAttribute("data-aared-row", String(row));
      cellElement.setAttribute("data-aared-col", String(col));
      cellElement.setAttribute("role", "gridcell");
      cellElement.setAttribute("aria-rowindex", String(row + 1));
      cellElement.setAttribute("aria-colindex", String(col + 1));
      // Programmatic focus only; the host keeps the tab stop (see above).
      cellElement.tabIndex = -1;
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
        cellElement.setAttribute("data-aared-merge", `${merge.row}:${merge.col}:${endRow}:${endCol}`);
      }
      const effective: CellModel = cell ?? { value: null, formatted: "", type: "empty" };
      const placement = paintCell(cellElement, effective, model.rightToLeft, paint);
      const spill = spills.get(col);
      // Only a data bar or a spill overlay needs the cell as a containing
      // block (a frozen cell's sticky, set below, is one too).
      if (spill !== undefined || effective.conditional?.dataBar !== undefined) cellElement.style.position = "relative";
      if (spill !== undefined) {
        cellElement.style.overflow = "visible";
        cellElement.style.zIndex = "1";
      }
      // Excel draws no gridline through text that runs across cells.
      if (crossed.has(col) && effective.style?.border?.right === undefined) {
        cellElement.style.borderRightColor = "transparent";
      }
      appendCellContent(cellElement, effective, placement, spill);
      paintFrozen(cellElement, row, col);
      line.append(cellElement);
    }
    if (fillColumnWidths.length > 0) line.append(paintFillCell(row, fillColumnWidths.length, "tracks"));
    return line;
  };

  /** One empty row below the data, numbered on from the last real row. */
  const paintFillRow = (position: number, height: number): HTMLTableRowElement => {
    const line = document.createElement("tr");
    line.setAttribute("data-aared-element", "fill-row");
    line.setAttribute("aria-hidden", "true");
    line.style.height = `${height}px`;
    if (showHeaders) {
      const number = document.createElement("th");
      number.setAttribute("data-aared-element", "fill-row-header");
      number.scope = "row";
      number.textContent = String(model.extent.rows + position + 1);
      paintHeading(number, "row", true);
      // A partial last row shows no number: its text would force the row
      // taller than the room left and reintroduce a scrollbar.
      if (height < FILL_ROW_TEXT_PX) {
        number.style.fontSize = "0";
        number.style.padding = "0";
      }
      line.append(number);
    }
    const columns = geometry.visibleColumns.length + fillColumnWidths.length;
    if (columns <= MAX_COL_SPAN) {
      line.append(paintFillCell(undefined, columns, "row"));
    } else {
      for (let i = 0; i < columns; i++) line.append(paintFillCell(undefined));
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
    // Measure before clearing: a content-sized scroller (no fixed height, as
    // in <aared-viewer>) collapses to its header row once the body is empty,
    // which would window the grid to a couple of rows and leave the rest as
    // a blank spacer. syncViewport also adopts the external room as our own
    // minimum height when content-sized, so the box holds still for the
    // measurement below.
    const scrollTop = scroller.scrollTop;
    const { fillH, windowH, windowW } = syncViewport();
    const viewport = windowH;
    measureFill(windowW, fillH);
    body.replaceChildren();
    const dataRows = geometry.visibleRows.filter((row) => !frozenRowSet.has(row));
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
    // Empty rows below the data, only when the whole sheet fits the viewport.
    fillRowHeights.forEach((height, position) => body.append(paintFillRow(position, height)));
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
  // Any box change can also move the external room, so the viewport cache
  // (probe verdict + room) is dropped alongside: it is recomputed on the
  // next pass, never trusted across a resize.
  const onScroll = (): void => schedule();
  scroller.addEventListener("scroll", onScroll, { passive: true });
  const resizer = new ResizeObserver(() => {
    viewportCache = null;
    schedule();
  });
  // The document-viewport room has no element to observe; the window resize
  // covers it (harmlessly duplicated when a clipping ancestor is observed).
  const onWindowResize = (): void => {
    viewportCache = null;
    // A device-pixel-ratio change (browser zoom, another screen) fires resize: re-snap the fill lines.
    syncFillLines();
    schedule();
  };
  window.addEventListener("resize", onWindowResize);
  // Fill rows draw their column lines from these properties even when no fill
  // column ever appears, so they are set before the first window paints.
  syncFillLines();
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
    // Anchors are measured from cell A1, but the canvas also holds the
    // heading gutter, so the layer origin is shifted past it.
    // In a right-to-left sheet column A is at the right, so anchors run leftward.
    box.style[model.rightToLeft ? "right" : "left"] = `${headerWidthPx + drawing.bounds.x * PX_PER_INCH}px`;
    box.style.top = `${headerHeightPx + drawing.bounds.y * PX_PER_INCH}px`;
    box.style.width = inchesToCss(Math.max(drawing.bounds.width, 0.01));
    box.style.height = inchesToCss(Math.max(drawing.bounds.height, 0.01));
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
    // The chart area's own fill sits under everything, gridlines included,
    // exactly like the frame Office draws.
    const { background } = drawing.graphic;
    if (background?.fill !== undefined) box.style.backgroundColor = background.fill;
    if (background?.stroke !== undefined) {
      box.style.boxSizing = "border-box";
      box.style.border = `${background.stroke.width}px solid ${background.stroke.color}`;
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

  const paintPicture = (drawing: SheetPictureModel): void => {
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
    if (urls === undefined) return;
    void urls.url(drawing.mediaId, drawing.contentType).then(
      (url) => {
        if (url === undefined || !box.isConnected) return;
        const image = document.createElement("img");
        image.src = url;
        image.alt = drawing.name;
        image.style.width = "100%";
        image.style.height = "100%";
        image.style.objectFit = "contain";
        box.replaceChildren(image);
      },
      () => {
        if (urls.closed || !box.isConnected) return;
        placeholder.textContent = `${drawing.name} (unavailable)`;
      },
    );
  };

  for (const drawing of model.drawings) {
    if (drawing.kind === "chart") paintChart(drawing);
    else paintPicture(drawing);
  }

  container.replaceChildren(host);
  // Observe attached: an observer added while the tree is detached never
  // sees layout, and this engine does not reliably backfill the initial
  // notification on attach. Observing here self-corrects the pre-layout
  // first paint (rendered above with a fallback viewport) once real
  // metrics exist.
  resizer.observe(scroller);
  // Measure now that layout exists, in this same task: the viewport fill (and
  // the first window) needs real sizes, and waiting a frame showed a short
  // sheet without its fill first — the grid shrank, then grew back to fill
  // the viewport. A target that isn't in a document yet has no layout to
  // measure: the next frame (or the observer) paints it then.
  if (host.isConnected) renderWindow();
  else schedule();
  const destroy = (): void => {
    if (destroyed) return;
    destroyed = true;
    if (raf !== 0) cancelAnimationFrame(raf);
    resizer.disconnect();
    scroller.removeEventListener("scroll", onScroll);
    window.removeEventListener("resize", onWindowResize);
    urls?.close();
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
  const notices = host.querySelector<HTMLElement>('[data-aared-element="sheet-notices"]');
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
    naturalSize: () => ({ width: realWidthPx, height: realHeightPx + (notices?.offsetHeight ?? 0) }),
    paintedOverflowY: () => {
      let fill = 0;
      for (const row of body.querySelectorAll<HTMLElement>('[data-aared-element="fill-row"]')) fill += row.offsetHeight;
      return Math.max(0, scroller.scrollHeight - fill - scroller.clientHeight);
    },
    setScale: (next: number) => {
      if (next === scale) return;
      scale = next;
      viewportCache = null;
      // The fill lines' thickness snaps to a device pixel at this zoom.
      syncFillLines();
      schedule();
    },
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
  headers: boolean;
  sheetTabs: boolean;
  fill: boolean;
  selection: boolean;
  gridlines: GridlineMode;
}

function normalizeAppearance(appearance: Appearance | undefined): NormalizedAppearance {
  const theme = appearance?.theme ?? "light";
  return {
    theme,
    dark: resolveTheme(theme) === "dark",
    headers: appearance?.spreadsheet?.headers ?? true,
    sheetTabs: appearance?.spreadsheet?.sheetTabs ?? false,
    fill: appearance?.spreadsheet?.fill ?? true,
    selection: appearance?.spreadsheet?.selection ?? true,
    gridlines: appearance?.spreadsheet?.gridlines ?? "source",
  };
}

function resolveGridPaint(model: SheetModel, appearance: NormalizedAppearance): GridPaint {
  return {
    showHeaders: appearance.headers,
    fillViewport: appearance.fill,
    showGridlines: appearance.gridlines === "show"
      ? true
      : appearance.gridlines === "hide"
        ? false
        : (model.showGridlines ?? true),
    gridlineColor: GRIDLINE_COLOR,
    // Automatic text follows its fill (Word/Excel Automatic); unfilled cells are Excel black.
    automaticColor: (fill) => (fill === undefined ? "#000000" : automaticInk(fill)),
    frozenFill: "white",
    headerFill: HEADER_FILL,
    headerColor: HEADER_COLOR,
    chromeLine: GRIDLINE_COLOR,
    noticeColor: PALETTE.light.muted,
    hostBackground: "white",
    dark: appearance.dark,
  };
}

/** A fit never scales a sheet below this: a huge sheet scrolls rather than shrinking to nothing (and the row window stays bounded). */
const MIN_FIT_SCALE = 0.25;

/**
 * The `render()` controller: zoom, cell navigation, single-cell selection,
 * sheet tabs, and lifecycle over one target element. Owns a stable tab bar
 * plus grid container; each painted sheet swaps the grid host in place.
 */
class XlsxViewer extends RenderHandleBase<Appearance> implements XlsxRenderHandle {
  readonly #target: HTMLElement;
  readonly #gridBox: HTMLDivElement;
  #tabsBar: HTMLDivElement | null = null;
  #session: SheetPaintSession | null = null;
  #model: SheetModel;
  readonly #handle: SheetHandle | undefined;
  readonly #document: XlsxDocument | undefined;
  readonly #mediaResolver: MediaResolver | undefined;
  #appearance: NormalizedAppearance;
  #selection: CellRef | null = null;
  #sheet: SheetInfo;
  #sheetGeneration = 0;
  #sheetController: AbortController | null = null;
  readonly #listeners = new AbortController();
  /** Keeps a fit mode current when the room around the grid changes. */
  readonly #fitObserver: ResizeObserver;
  /** The room the last fit measured; a resize that leaves it unchanged is ignored. */
  #fitRoom = { width: -1, height: -1 };
  readonly #previousTargetDisplay: string;
  readonly #previousTargetFlexDirection: string;
  /** This render's hold on the class sheet in the target's root. */
  readonly #classes: ClassSheetScope;

  /** Built only by `render()` after the model (and optional handle) resolve. */
  constructor(
    target: HTMLElement,
    model: SheetModel | SheetRangeModel,
    handle: SheetHandle | undefined,
    mediaResolver: MediaResolver | undefined,
    appearance: Appearance | undefined,
    zoom: ZoomValue,
  ) {
    super(zoom);
    this.#target = target;
    this.#model = model;
    this.#handle = handle;
    this.#document = handle?.sheetDocument();
    this.#mediaResolver = mediaResolver;
    this.#appearance = normalizeAppearance(appearance);
    this.#sheet = handle === undefined
      ? { index: 0, title: model.name, state: "visible" }
      : { ...handle.info };
    this.#previousTargetDisplay = target.style.display;
    this.#previousTargetFlexDirection = target.style.flexDirection;
    // The grid's default look must be in the target's root before any of it is attached or measured.
    this.#classes = adoptClassSheet(target);
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
    this.#fitObserver = new ResizeObserver(() => {
      // The target may have been attached (or moved into a shadow root) since render().
      this.#classes.follow();
      if (typeof this.zoom === "number" || this.#session === null) return;
      const room = measureFitRoom(this.#gridBox, this.#session.host, this.zoom === "fit-page");
      if (room.width === this.#fitRoom.width && room.height === this.#fitRoom.height) return;
      this.refit();
    });
    this.attach({ applyZoom: () => this.#applyZoom(), destroy: () => this.#teardown() });
    this.#repaintCurrent();
    this.#fitObserver.observe(this.#gridBox);
  }

  setAppearance(appearance: Appearance): void {
    if (this.destroyed) return;
    this.#appearance = normalizeAppearance(appearance);
    this.#syncTabsChrome();
    // Preserve scroll across the repaint: the grid paints a fresh scroller.
    const scroller = this.#session?.scroller;
    const scrollTop = scroller?.scrollTop ?? 0;
    const scrollLeft = scroller?.scrollLeft ?? 0;
    this.#repaintCurrent();
    const next = this.#session?.scroller;
    if (next !== undefined) {
      next.scrollTop = scrollTop;
      next.scrollLeft = scrollLeft;
      this.#session?.renderWindow();
    }
  }

  get sheet(): SheetInfo {
    return { ...this.#sheet };
  }

  get selection(): CellRef | null {
    return this.#selection === null ? null : { ...this.#selection };
  }

  async goToCell(cell: string): Promise<void> {
    if (this.destroyed) return;
    const { row, col } = parseSingleCell(cell, "goToCell");
    if (this.#appearance.selection) {
      this.#select({ address: cellAddress(row, col), row, col }, true);
    }
    // Center when the cell is out of view; a visible cell keeps its scroll.
    this.#ensureVisible(row, col, "center", true);
    this.#session?.cellElement(row, col)?.focus({ preventScroll: true });
  }

  scrollTo(target: { cell?: string; range?: string; align?: "start" | "center" | "end" }): void {
    if (this.destroyed) return;
    const align = target.align ?? "start";
    let row: number;
    let col: number;
    if (target.cell !== undefined && target.range === undefined) {
      ({ row, col } = parseSingleCell(target.cell, "scrollTo"));
    } else if (target.range !== undefined && target.cell === undefined) {
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
      throw new AaredError(
        "invalid-range",
        'aared/xlsx: scrollTo needs exactly one of { cell: "D14" } or { range: "A1:M200" }',
      );
    }
    this.#ensureVisible(row, col, align, false);
  }

  focus(cell: string): void {
    if (this.destroyed) return;
    const { row, col } = parseSingleCell(cell, "focus");
    if (this.#appearance.selection) {
      this.#select({ address: cellAddress(row, col), row, col }, true);
    }
    // No scrolling by construction: a cell outside the current window keeps
    // its logical selection without DOM focus until it is scrolled into view.
    this.#session?.cellElement(row, col)?.focus({ preventScroll: true });
  }

  async goToSheet(key: string | number): Promise<void> {
    if (this.destroyed) return;
    const document = this.#document;
    if (document === undefined) {
      throw new AaredError(
        "unsupported-feature",
        "aared/xlsx: sheet switching needs a SheetHandle input; this view was rendered from a bare model",
      );
    }
    const handle = document.sheets.get(key);
    const generation = ++this.#sheetGeneration;
    this.#sheetController?.abort();
    const controller = new AbortController();
    this.#sheetController = controller;
    const signal = controller.signal;
    let model: SheetRangeModel;
    try {
      model = await handle.read({ signal });
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
      this.dispatchEvent(new CustomEvent("selectionchange", { detail: { sheet: this.#sheetRef(), selection: null } }));
    }
    this.dispatchEvent(new CustomEvent("sheetchange", { detail: this.#sheetRef() }));
  }

  /** The visible sheet as a unit reference — how every event names it. */
  #sheetRef(): UnitRef {
    return { index: this.#sheet.index, title: this.#sheet.title };
  }

  #teardown(): void {
    this.#sheetGeneration += 1;
    this.#sheetController?.abort();
    this.#sheetController = null;
    this.#listeners.abort();
    this.#fitObserver.disconnect();
    this.#session?.destroy();
    this.#session = null;
    this.#target.style.display = this.#previousTargetDisplay;
    this.#target.style.flexDirection = this.#previousTargetFlexDirection;
    this.#classes.release();
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
    this.#paintTabsTokens(bar);
    for (const handle of this.#document.sheets.list()) {
      if (handle.info.state !== "visible") continue;
      const tab = document.createElement("button");
      tab.type = "button";
      tab.dataset.aaredElement = "sheet-tab";
      tab.dataset.aaredSheet = handle.info.title;
      tab.dataset.aaredSheetIndex = String(handle.info.index);
      tab.setAttribute("role", "tab");
      const selected = handle.info.index === this.#sheet.index;
      tab.setAttribute("aria-selected", selected ? "true" : "false");
      tab.tabIndex = selected ? 0 : -1;
      tab.textContent = handle.info.title;
      tab.style.font = "12px system-ui, sans-serif";
      tab.style.color = "var(--aared-ink)";
      tab.style.background = "transparent";
      tab.style.border = "none";
      tab.style.borderBottom = selected ? "2px solid var(--aared-selection)" : "2px solid transparent";
      tab.style.fontWeight = selected ? "700" : "400";
      tab.style.padding = "4px 10px 6px";
      tab.style.cursor = "pointer";
      tab.style.whiteSpace = "nowrap";
      const sheetIndex = handle.info.index;
      tab.addEventListener("click", () => {
        // Tabs derive from the live collection, so not-found cannot happen;
        // abort rejections from a superseded switch stay silent in goToSheet.
        void this.goToSheet(sheetIndex).catch(() => {});
      });
      bar.append(tab);
    }
    // A workbook whose sheets are all hidden still gets the (empty) bar so
    // the marker stays queryable; tabs appear only for visible sheets.
    this.#tabsBar = bar;
  }

  #syncTabsChrome(): void {
    if (!this.#appearance.sheetTabs || this.#document === undefined) {
      this.#tabsBar?.remove();
      this.#tabsBar = null;
      return;
    }
    if (this.#tabsBar === null) {
      this.#buildTabs();
      if (this.#tabsBar !== null) this.#target.prepend(this.#tabsBar);
      return;
    }
    this.#paintTabsTokens(this.#tabsBar);
  }

  /** The tab bar's design tokens: the tab bar is viewer chrome, so it follows the resolved theme; the grid always paints light. */
  #paintTabsTokens(bar: HTMLElement): void {
    const { dark } = this.#appearance;
    paintTokens(bar, dark ? "dark" : "light");
    paintSheetTokens(bar, {
      dark,
      gridline: dark ? GRIDLINE_DARK : GRIDLINE_COLOR,
      headerFill: dark ? HEADER_FILL_DARK : HEADER_FILL,
    });
  }

  #updateTabs(): void {
    if (this.#tabsBar === null) return;
    for (const tab of this.#tabsBar.querySelectorAll<HTMLElement>('[data-aared-element="sheet-tab"]')) {
      const selected = tab.dataset.aaredSheetIndex === String(this.#sheet.index);
      tab.setAttribute("aria-selected", selected ? "true" : "false");
      tab.tabIndex = selected ? 0 : -1;
      tab.style.borderBottom = selected ? "2px solid var(--aared-selection)" : "2px solid transparent";
      tab.style.fontWeight = selected ? "700" : "400";
    }
  }

  #repaintCurrent(): void {
    this.#session?.destroy();
    this.#session = null;
    // The grid's look lives in the class sheet: make sure it is on the root
    // the target is in now (moved, or an embedder replaced adoptedStyleSheets).
    this.#classes.follow();
    this.#session = paintSheet(
      this.#model,
      this.#gridBox,
      this.#mediaResolver,
      resolveGridPaint(this.#model, this.#appearance),
      () => {
        this.#classes.follow();
        this.#restoreSelection();
      },
    );
    // A new sheet (or a headers toggle) changes the used size: re-fit it.
    this.refit();
    this.#updateTabs();
  }

  /**
   * Apply the zoom as a CSS `zoom` on the grid host and report the scale.
   * A fit reads only the room `#gridBox` offers — outer space, measured with
   * the host set aside — against the sheet's used size, so it never reads
   * back its own result; everything inside the zoomed host (fill, row
   * window, scroll) then works in host pixels via `setScale`.
   */
  #applyZoom(): number {
    this.#classes.follow();
    const session = this.#session;
    if (session === null) return this.scale;
    const zoom = this.zoom;
    let scale: number;
    if (typeof zoom === "number") {
      scale = safeScale(zoom);
    } else {
      const room = measureFitRoom(this.#gridBox, session.host, zoom === "fit-page");
      this.#fitRoom = room;
      const natural = session.naturalSize();
      const fitted = fitScale(zoom, room, natural);
      // Not laid out yet: keep the current scale; the observer re-fits later.
      scale = fitted === undefined ? this.scale : Math.max(MIN_FIT_SCALE, fitted);
      if (fitted !== undefined) {
        // One scrollbar pass: a vertical scrollbar the fitted grid causes
        // narrows the scroller, so shrink by exactly the width it took. Both
        // widths are host pixels, so the ratio holds however the engine
        // reports them; shrinking can't add a vertical scrollbar.
        this.#setHostZoom(session.host, scale);
        const inner = session.scroller.clientWidth;
        if (inner > 0 && inner < natural.width) scale = Math.max(MIN_FIT_SCALE, Math.min(scale, (inner * scale) / natural.width));
        if (zoom === "fit-page") {
          // One pass for the height the model can't predict: below 100% every
          // border snaps to at least a device pixel, so painted rows outgrow
          // their modeled heights. Shrink by exactly the painted overflow.
          this.#setHostZoom(session.host, scale);
          const client = session.scroller.clientHeight;
          const over = session.paintedOverflowY();
          if (client > 0 && over > 0) scale = Math.max(MIN_FIT_SCALE, (scale * client) / (client + over));
        }
      }
    }
    this.#setHostZoom(session.host, scale);
    session.setScale(scale);
    return scale;
  }

  /** Write `zoom` on the host only when it changes, so a repeated fit causes no re-layout. */
  #setHostZoom(host: HTMLElement, scale: number): void {
    const value = scale === 1 ? "" : String(scale);
    if (host.style.getPropertyValue("zoom") === value) return;
    if (value === "") host.style.removeProperty("zoom");
    else host.style.setProperty("zoom", value);
  }

  #select(selection: CellRef | null, emit: boolean): void {
    this.#selection = selection === null ? null : { ...selection };
    this.#restoreSelection();
    if (emit && this.#appearance.selection) {
      this.dispatchEvent(new CustomEvent("selectionchange", {
        detail: { sheet: this.#sheetRef(), selection: this.selection },
      }));
    }
  }

  #restoreSelection(): void {
    const session = this.#session;
    if (session === null) return;
    for (const element of session.host.querySelectorAll('[data-aared-selected="true"]')) {
      const stale = element as HTMLElement;
      stale.removeAttribute("data-aared-selected");
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
    element.setAttribute("data-aared-selected", "true");
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
    if (this.destroyed) return;
    const target = event.target as HTMLElement | null;
    const link = target?.closest?.('a[data-aared-element="cell-link"]') ?? null;
    if (link !== null && link instanceof HTMLElement) {
      // Hyperlinks never navigate: prevent the default and report the
      // activation to the embedder instead.
      event.preventDefault();
      const cell = link.closest('td[data-aared-element="cell"]');
      if (cell === null || !(cell instanceof HTMLElement)) return;
      const row = Number(cell.getAttribute("data-aared-row"));
      const col = Number(cell.getAttribute("data-aared-col"));
      if (!Number.isInteger(row) || !Number.isInteger(col)) return;
      if (this.#appearance.selection) {
        this.#select({ address: cellAddress(row, col), row, col }, true);
      }
      const href = this.#cellModel(row, col).hyperlink?.target ?? link.getAttribute("href") ?? "";
      this.dispatchEvent(new CustomEvent("linkactivate", {
        detail: { target: href, sheet: this.#sheetRef(), address: cellAddress(row, col), row, col },
      }));
      return;
    }
    if (!this.#appearance.selection) return;
    const cell = target?.closest?.('td[data-aared-element="cell"]') ?? null;
    if (cell === null || !(cell instanceof HTMLElement)) return;
    const row = Number(cell.getAttribute("data-aared-row"));
    const col = Number(cell.getAttribute("data-aared-col"));
    if (!Number.isInteger(row) || !Number.isInteger(col)) return;
    this.#select({ address: cellAddress(row, col), row, col }, true);
  }

  #handleDoubleClick(event: MouseEvent): void {
    if (this.destroyed || !this.#appearance.selection) return;
    const target = event.target as HTMLElement | null;
    const cell = target?.closest?.('td[data-aared-element="cell"]') ?? null;
    if (cell === null || !(cell instanceof HTMLElement)) return;
    const row = Number(cell.getAttribute("data-aared-row"));
    const col = Number(cell.getAttribute("data-aared-col"));
    if (!Number.isInteger(row) || !Number.isInteger(col)) return;
    this.#select({ address: cellAddress(row, col), row, col }, true);
    this.#activateCell(row, col);
  }

  #activateCell(row: number, col: number): void {
    this.dispatchEvent(new CustomEvent("cellactivate", {
      detail: {
        sheet: this.#sheetRef(),
        address: cellAddress(row, col),
        row,
        col,
        model: this.#cellModel(row, col),
      },
    }));
  }

  #handleKeyDown(event: KeyboardEvent): void {
    if (this.destroyed || !this.#appearance.selection) return;
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
        const page = Math.max(1, Math.round(viewport / Math.max(((session.model.defaultRowHeight) * PX_PER_INCH), 1)));
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
      const width = (model.columnWidths[col] ?? 0) * PX_PER_INCH;
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
  return new XlsxViewer(input.target, model, handle, mediaResolver, input.appearance, input.zoom ?? 1);
}
