/**
 * XLSX entrypoint: open workbook documents in a Worker, read sheet windows,
 * and paint resolved `SheetModel`s to a virtualized DOM grid — merges as
 * spanning cells, frozen panes as sticky regions, sheet-level RTL flips the
 * whole grid. Never canvas. See docs/ARCHITECTURE.md.
 *
 * Structure mirrors `aared/pptx`: `open()` with a format check, a document
 * holding a `WorkerSession`, collections implementing core's
 * `Collection<TItem, TKey>`, lightweight handles, `search()`, and an
 * input-object `render()` returning an `EventTarget` controller. The internal
 * `paintSheet()` paints one sheet; headers default on, sheet tabs default off,
 * and gridlines follow `appearance.spreadsheet.gridlines`.
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
 * dies after one move. Selection is a cursor cell with an optional extended
 * range: `Shift`+arrows extends from the anchor, dragging extends from the
 * press point, any plain move re-anchors. Sheet tabs
 * list VISIBLE sheets only; `goToSheet()` still resolves hidden sheets by
 * name or index. Hyperlink clicks never navigate: they `preventDefault()`
 * and emit `linkactivate`.
 */

import type {
  CellModel,
  CellRange,
  RenderWarning,
  SheetChartModel,
  SheetDrawingModel,
  SheetModel,
  SheetPictureModel,
  SheetRangeModel,
  ShapeFillModel,
} from "../core/model.js";
import { mixResolvedColorsSrgb } from "../core/color.js";
import { loadFillPainter, loadedFillPainter } from "../fill-loader.js";
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
  ZoomValue,
} from "../core/index.js";
import { AaredError } from "../core/index.js";
import { adoptClassSheet, classes, cssClass, type ClassSheetScope } from "../class-sheet.js";
import { CONTENT_TEXT } from "../render-defaults.js";
import { clipSnippet, findByTitle, foldQuery, searchKey } from "../core/search.js";
import { SessionMedia, WorkerSession } from "../core/session.js";
import type { SheetText, UnitRange } from "../core/rpc.js";
import { itemAt, throwIfAborted } from "../core/errors.js";
import { reportBackgroundError } from "../background-error.js";
import { contentSignal, contentSource } from "../content-view.js";
import { APPEARANCE_DEFAULTS, PALETTE, automaticInk, mergeAppearance, paintTokens, resolveTheme } from "../appearance.js";
import { chartsSettled, paintChart } from "../chart-svg.js";
import { loadEffectPainter, loadedEffectPainter } from "../effect-loader.js";
import { paintTransform } from "../transform-paint.js";
import { MediaUrls, paintImageSource } from "../media-url.js";
import { DEFAULT_ZOOM, RenderHandleBase, fitScale, measureFitRoom, missingIdWarning, onReplaced, safeScale, settleRender } from "../render-handle.js";
import type { EventTargetOf, ZoomChangeEvent } from "../render-handle.js";
import { MARK_FALLBACKS } from "../diagnostics.js";
import { awaitDocumentFonts, fontFidelityWarnings } from "../font-fidelity.js";
import { PX_PER_INCH, inchesToCss } from "../core/units.js";
import { cssFontFamily } from "../core/typography.js";
import { isInsideGrid } from "../core/xlsx/limits.js";
import { cellAddress, columnName, parseA1Range } from "../core/xlsx/range.js";
import { iconSvg } from "./icons.js";

/** Document metadata for a workbook; `format` is narrowed to `"xlsx"`. */
export interface XlsxInfo extends DocumentInfo {
  format: "xlsx";
}

/** Lightweight sheet identity from `open()` indexing — no sheet is parsed yet. `title` is the sheet's name. */
export interface SheetInfo extends UnitRef {
  state: "visible" | "hidden" | "veryHidden";
}

/** One matching cell from `sheet.search()`: which sheet, and which cell in it. */
export interface XlsxSearchMatch extends CellRef {
  /** The sheet holding the match; `sheet.index` is its `goTo()` key. */
  sheet: UnitRef;
  text: string;
  /**
   * Set when the match is a word of a chart (its title, an axis title, a series
   * name or a category label) rather than a cell: `address`, `row` and `col`
   * are then the cell the chart's top-left corner is anchored to, and `name`
   * is the chart's name in the sheet.
   */
  chart?: { name: string };
}

/** Options for `sheet.read()`: a window of the sheet, or all of it when `range` is absent. */
export interface SheetReadOptions extends ReadOptions {
  /** An A1 range (`"A1:M200"`, or one cell) or a numeric `CellRange`. Absent reads the whole sheet. */
  range?: string | CellRange;
}

/** Viewer-owned sheet input: metadata and range reads, without an owner/session. */
export interface SheetContent {
  readonly info: SheetInfo;
  read(options?: SheetReadOptions): Promise<SheetRangeModel>;
}

/** Read/render capabilities borrowed from a preview; search stays on the viewer. */
export interface XlsxContent {
  readonly info: XlsxInfo;
  readonly sheets: Collection<SheetContent, number | string>;
  readonly media: MediaCollectionShape;
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

  /**
   * Case-insensitive, Arabic-spelling-insensitive substring search over this
   * one sheet, as Excel's Find searches the sheet you are on. Every matching
   * cell yields one match, in row-major order, then a chart's words. A cell
   * matches on its formatted text, falling back to a string cached value when
   * the formatted text is empty. The first search reads the sheet and keeps its
   * cell text (aborting rejects, keeping nothing); later searches scan that
   * text without touching the Worker. A workbook-wide search is a loop over
   * `workbook.sheets.list()`.
   */
  async search(query: string, options?: ReadOptions): Promise<XlsxSearchMatch[]> {
    const needle = foldQuery(query);
    if (needle === null) return [];
    const signal = options?.signal;
    throwIfAborted(signal);
    const cells = await this.owner.cellsOf(this.info.index, signal);
    const matches: XlsxSearchMatch[] = [];
    for (let at = 0; at < cells.key.length; at++) {
      if (!cells.key[at]!.includes(needle)) continue;
      const row = cells.rows[at]!;
      const col = cells.cols[at]!;
      const chart = cells.chart[at];
      matches.push({
        sheet: { index: this.info.index, title: this.info.title },
        address: cellAddress(row, col),
        row,
        col,
        text: clipSnippet(cells.text[at]!),
        ...(typeof chart === "string" ? { chart: { name: chart } } : {}),
      });
    }
    return matches;
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
        "aared/xlsx",
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

/** One sheet's text as `search()` scans it: the Worker's `SheetText` plus the search key of each `text`. */
interface SheetSearchCells extends SheetText {
  key: string[];
}

/** An opened workbook: collections, media, and lifecycle. Search is per sheet (`SheetHandle.search`). */
export class XlsxDocument {
  readonly info: XlsxInfo;
  readonly sheets: SheetCollection;
  readonly media: MediaCollectionShape;
  /** Per-sheet cell text, filled by the first `search()` on each sheet. */
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
   * @internal The searchable cell text of one sheet, read once: the Worker's
   * text read (`sheetText`), its cells' text without the sheet's model. Kept
   * here, not on the handle, so `close()` drops every sheet's text together.
   * Called by `SheetHandle.search()`.
   */
  async cellsOf(index: number, signal: AbortSignal | undefined): Promise<SheetSearchCells> {
    const cached = this.searchIndex.get(index);
    if (cached !== undefined) return cached;
    const text = await this.session.text<SheetText>(index, signal === undefined ? undefined : { signal });
    const cells = { ...text, key: text.text.map(searchKey) };
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

/** `render()` input: an opened handle, borrowed sheet content, or a resolved model. */
export interface XlsxRenderInput {
  sheet: SheetHandle | SheetContent | SheetModel | SheetRangeModel;
  target: HTMLElement;
  appearance?: Appearance;
  /** Starting zoom, applied without a `zoomchange`. Default `1` — 100%, as Excel opens a sheet. */
  zoom?: ZoomValue;
  /** Override the default session-backed media resolution. */
  mediaResolver?: MediaResolver;
  /**
   * Cancel the render until it resolves: before it paints, it paints nothing;
   * after (it then waits for the sheet's fonts), it takes the sheet back out.
   * Either way the target keeps its own layout, and the render rejects with
   * `aborted` and never hands back a handle. Once `render()` has resolved the
   * signal is ignored; `destroy()` is the teardown.
   */
  signal?: AbortSignal;
}

/** The events an XLSX render handle fires, by name: `addEventListener` types each `event.detail`. Every event names its sheet as a unit reference and its cell as a cell location. */
export interface XlsxRenderEventMap {
  /** A built-in sheet-tab action failed. Public methods reject instead. */
  error: CustomEvent<{ message: string; cause: AaredError }>;
  /** The zoom changed. */
  zoomchange: ZoomChangeEvent;
  /** The sheet shown changed (`goToSheet()`, a sheet tab). */
  sheetchange: CustomEvent<UnitRef>;
  /** The selected cell changed, the range extended, or the selection cleared (`null`). */
  selectionchange: CustomEvent<{ sheet: UnitRef; selection: CellRef | null; range: { start: CellRef; end: CellRef } | null }>;
  /** A cell was activated: Enter or a double-click. */
  cellactivate: CustomEvent<CellRef & { sheet: UnitRef; model: CellModel }>;
  /** A cell hyperlink was clicked; nothing navigates. */
  linkactivate: CustomEvent<CellRef & { target: string; sheet: UnitRef }>;
}

/** Controller returned by `render()`: zoom, navigation, selection, and lifecycle. */
export interface XlsxRenderHandle extends EventTargetOf<XlsxRenderEventMap> {
  /** Disconnect scroll virtualization, abort pending sheet switches, release image object URLs. Idempotent. */
  destroy(): void;
  /** `true` once destroyed; every method is then a no-op that fires nothing. */
  readonly destroyed: boolean;
  /** Non-fatal things this render noticed about its own output — currently font fidelity only. See `RenderWarning`. */
  readonly warnings: readonly RenderWarning[];
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
   * Update the viewer appearance in place: repaints the current sheet with
   * the new theme/chrome knobs while keeping the sheet, selection, zoom and
   * scroll position. Changes only what it sets: the theme and `spreadsheet`
   * knobs it leaves out keep their current values. To bring a knob back to
   * its default, pass the default (`{ spreadsheet: { sheetTabs: false } }`).
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
  /** The extended range (`Shift`+arrows, drag), or `null` when the cursor stands alone. */
  readonly range: { start: CellRef; end: CellRef } | null;
}

const GRIDLINE_COLOR = "#E7E9EE";
const GRIDLINE_DARK = "#3B3B45";
const ROW_OVERSCAN = 8;
/** Heading chrome dimensions. Viewer chrome, not authored geometry. */
const HEADER_COL_WIDTH_PX = 40;
/** A fill sliver thinner than this is skipped; the cap bounds a huge viewport. */
const FILL_MIN_TRACK_PX = 6;
/**
 * Below this visible room a clipped sheet is not given its own scroll box:
 * a box too short for the column letters and a few rows is worse than
 * letting the container scroll the sheet.
 */
const MIN_SCROLL_ROOM_PX = 120;
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
/**
 * Room between a cell's 1px border and its text on each side. Excel's PDF
 * export puts left- and right-aligned text 2.4-2.5px from the cell box in
 * Arial 10 and Calibri 11 alike, once its print scale is taken out (it
 * prints a 12-character, 89px column 95.8px wide); border plus padding
 * matches that. 4px here cost every cell 5px of room, enough to wrap a header
 * Excel keeps on one line ("Hours logged", 85.9px in a 91px column).
 */
const CELL_PAD_PX = 1.5;
const CELL_BOX = `padding:1px ${CELL_PAD_PX}px;border:1px solid transparent;border-right-color:${GRIDLINE};border-bottom-color:${GRIDLINE};font-size:10pt;line-height:1`;
/** A data cell. Single-line cells never outgrow their authored row: see `paintCell`. */
const CELL = cssClass(`${CELL_BOX};color:#000000;vertical-align:bottom;white-space:nowrap;overflow:hidden`);
const CELL_WRAP = cssClass("line-height:1.1;white-space:pre-wrap;overflow-wrap:break-word;overflow:visible");
/** Excel's plus pointer over cells that select on click, the viewer's data and empty cells alike. */
const SELECTABLE = cssClass("cursor:cell");
/** A viewport-fill cell: inert, the same box as a data cell. */
const FILL_CELL = cssClass(`${CELL_BOX};overflow:hidden`);
/**
 * One fill cell standing for a row's run of fill columns: their lines drawn at
 * the fill-column pitch, which is the sheet's own default column width
 * (`--aared-fill-col`, set on the table by `syncFillLines`).
 */
const FILL_PITCH = "var(--aared-fill-col,64px)";
const FILL_TRACKS = cssClass(
  `background-origin:border-box;background-image:repeating-linear-gradient(to right,transparent 0 calc(${FILL_PITCH} - var(--aared-hairline,1px)),${GRIDLINE} calc(${FILL_PITCH} - var(--aared-hairline,1px)) ${FILL_PITCH})`,
);
const FILL_TRACKS_RTL = cssClass(
  `background-origin:border-box;background-image:repeating-linear-gradient(to left,${GRIDLINE} 0 var(--aared-hairline,1px),transparent var(--aared-hairline,1px) ${FILL_PITCH})`,
);
/** One fill cell standing for a whole empty row: every column's line, from per-sheet custom properties (`syncFillLines`). */
const FILL_ROW = cssClass(
  "background-origin:border-box;background-repeat:no-repeat;background-image:var(--aared-fill-lines,none);background-size:var(--aared-fill-lines-size,auto);background-position:var(--aared-fill-lines-position,0 0)",
);
const NO_GRIDLINES = cssClass("border-right-color:transparent;border-bottom-color:transparent;background-image:none");
/** Text that must sit above a data bar. */
const CELL_TEXT_OVER = cssClass("position:relative;z-index:1");
const CELL_LINK = cssClass("text-decoration:none");
const DATA_BAR = cssClass("position:absolute;inset-block:0;pointer-events:none");
/** A gradient or pattern cell fill: the cell's whole box, under its text, never interactive. */
const CELL_FILL = cssClass("position:absolute;inset:0;width:100%;height:100%;z-index:-1;pointer-events:none;display:block");
/** An icon-set glyph: 16px at the inline start, centred vertically, never interactive. */
const ICON = cssClass("position:absolute;inset-inline-start:2px;top:50%;width:16px;height:16px;transform:translateY(-50%);pointer-events:none");
/** Text beside an icon keeps clear of the 16px glyph plus its 2px offset. */
const CELL_TEXT_WITH_ICON = cssClass("position:relative;z-index:1;padding-inline-start:20px;box-sizing:border-box;display:inline-block;max-width:100%");
const SPILL = cssClass(`position:absolute;top:0;bottom:0;box-sizing:border-box;padding:1px ${CELL_PAD_PX}px;display:flex;overflow:hidden;pointer-events:none`);
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
/**
 * The extended-range wash on every cell copy the window paints. It is a
 * translucent layer (an inset shadow sits above the cell's fill and below its
 * text), so an authored fill keeps its colour and light text on a dark header
 * stays readable. Scoped to this transient state: removing the class restores
 * the cell, nothing is saved.
 */
const RANGE_WASH = "inset 0 0 0 9999px var(--aared-range-fill)";
const inRangeClass = cssClass(`box-shadow:${RANGE_WASH}`);
/**
 * The cursor cell inside an extended range paints paper-white, as in Excel,
 * but only when it has no fill of its own: white over a header's dark fill
 * would hide its text. Such a cell keeps its look and shows its frame.
 */
const activeInRangeClass = cssClass("box-shadow:inset 0 0 0 9999px #ffffff");
/**
 * Each range a formula reads gets its own colour, in order of first appearance
 * (as Excel's formula editor does), cycling after six: a light tint and a dashed frame.
 */
const REFERENCE_COLORS = [
  { solid: "#2F6FDB", tint: "rgba(47, 111, 219, 0.12)" },
  { solid: "#E8772E", tint: "rgba(232, 119, 46, 0.12)" },
  { solid: "#8E44AD", tint: "rgba(142, 68, 173, 0.12)" },
  { solid: "#2E9E5B", tint: "rgba(46, 158, 91, 0.12)" },
  { solid: "#D4A017", tint: "rgba(212, 160, 23, 0.14)" },
  { solid: "#C0392B", tint: "rgba(192, 57, 43, 0.12)" },
];
const referenceClasses = REFERENCE_COLORS.map(({ tint }) => cssClass(`box-shadow:inset 0 0 0 9999px ${tint}`));
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
  // The extended-range wash: the accent at a low alpha, precomputed per
  // theme (no `color-mix`, which the baseline floor cannot assume).
  element.style.setProperty("--aared-range-fill", tokens.dark ? "rgba(138, 180, 248, 0.18)" : "rgba(11, 87, 208, 0.12)");
}

/**
 * A cell border as CSS, by SpreadsheetML line style: weight (thin and hair
 * 1px, medium 2px, thick 3px, a double needs 3px to show two lines) and the
 * nearest CSS line (CSS has no dash-dot, so those draw dashed).
 */
function borderCss(style: string | undefined, color: string | undefined): string {
  let width = "1px";
  let line = "solid";
  switch (style) {
    case "medium": width = "2px"; break;
    case "thick": width = "3px"; break;
    case "double": width = "3px"; line = "double"; break;
    case "dashed": case "dashDot": case "dashDotDot": line = "dashed"; break;
    case "mediumDashed": case "mediumDashDot": case "mediumDashDotDot": case "slantDashDot": width = "2px"; line = "dashed"; break;
    case "dotted": line = "dotted"; break;
  }
  return `${width} ${line} ${color ?? "currentColor"}`;
}

function textAlign(cell: CellModel, sheetRtl: boolean): string {
  const align = cell.style?.align ?? "general";
  if (align !== "general") return align;
  // Excel's `general`: strings at the start edge, numbers at the end edge,
  // booleans and errors centered. Logical values keep DOM order intact.
  // Deliberately the sheet's start edge for every string: Excel places a
  // General string by its own first strong letter, so in a right-to-left
  // sheet "SKU" sits left beside Arabic sitting right (Notes §116).
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
  /** Cells select on click, so they show Excel's plus pointer (`cursor: cell`). */
  selectable: boolean;
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
  /** Physical distance from the cell's right edge to the run's right edge. */
  after: number;
  /** The run's columns, the cell's own among them, physically left to right. */
  run: number[];
}

/** Where a cell's text sits, as `paintCell` resolved it (read by the spill overlay). */
interface CellPlacement {
  align: string;
  vertical: string;
}

/** Nothing to show: no text, no link, no data bar, no icon. Such a cell paints no content element. */
function isEmptyCell(cell: CellModel): boolean {
  return cell.formatted === "" && cell.hyperlink === undefined && cell.conditional?.dataBar === undefined && cell.conditional?.icon === undefined;
}

function appendCellContent(cellElement: HTMLElement, cell: CellModel, placement: CellPlacement, spill?: TextSpill): void {
  const dataBar = cell.conditional?.dataBar;
  const icon = cell.conditional?.icon;
  if (icon !== undefined) {
    const holder = document.createElement("span");
    holder.setAttribute("data-aared-element", "cell-icon");
    holder.setAttribute("aria-hidden", "true");
    holder.setAttribute("class", ICON);
    // One inline SVG per painted icon; the catalogue shares geometry so the
    // per-cell cost is a 16px box, and virtualization bounds how many paint.
    holder.innerHTML = `<svg viewBox="0 0 16 16" width="16" height="16" focusable="false">${iconSvg(icon.family, icon.icon)}</svg>`;
    cellElement.append(holder);
  }
  if (dataBar !== undefined) {
    const bar = document.createElement("div");
    bar.setAttribute("data-aared-element", "data-bar");
    bar.setAttribute("aria-hidden", "true");
    bar.setAttribute("class", DATA_BAR);
    bar.style.width = `${Math.min(Math.max(dataBar.ratio, 0), 1) * 100}%`;
    const rtl = dataBar.direction === "rightToLeft"
      || (dataBar.direction === "context" && getComputedStyle(cellElement).direction === "rtl");
    if (dataBar.direction === "context") {
      bar.style.insetInlineStart = "0";
    } else {
      bar.style.left = rtl ? "auto" : "0";
      bar.style.right = rtl ? "0" : "auto";
    }
    if (dataBar.gradient) {
      bar.style.backgroundImage = `linear-gradient(to ${rtl ? "left" : "right"}, ${dataBar.color} 0%, transparent 100%)`;
    } else {
      bar.style.backgroundColor = dataBar.color;
    }
    if (dataBar.border) {
      bar.style.boxSizing = "border-box";
      bar.style.border = `1px solid ${dataBar.borderColor ?? dataBar.color}`;
    }
    cellElement.append(bar);
  }
  if ((dataBar?.showValue === false) || (icon?.showValue === false) || (cell.formatted === "" && cell.hyperlink === undefined)) return;
  const text = document.createElement("span");
  text.setAttribute("data-aared-element", "cell-text");
  if (dataBar !== undefined) text.setAttribute("class", CELL_TEXT_OVER);
  else if (icon !== undefined) text.setAttribute("class", CELL_TEXT_WITH_ICON);
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
  if (placement.align === "center" && spill.offset !== spill.after) {
    // Excel centres the text on its own cell, never on the run: each side
    // spills into its own blank neighbours and is clipped at the first cell
    // with content. Read off Excel's PDF export: a word that fits stays
    // centred in its cell beside an empty one, and an overflowing line keeps
    // the same position whether its left or its right neighbour holds text,
    // only cut at that neighbour. Padding the shorter side by the difference
    // puts the centring box's middle on the cell's middle; the overlay still
    // clips at the run's ends.
    const shift = Math.abs(spill.offset - spill.after);
    if (spill.offset > spill.after) overlay.style.paddingLeft = `${CELL_PAD_PX + shift}px`;
    else overlay.style.paddingRight = `${CELL_PAD_PX + shift}px`;
  }
  overlay.style.alignItems = placement.vertical === "top" ? "flex-start" : placement.vertical === "middle" ? "center" : "flex-end";
  text.dir = cellElement.dir;
  text.setAttribute("class", classes(text.getAttribute("class"), SPILL_TEXT));
  overlay.append(text);
  cellElement.append(overlay);
}

/** A painted cell's box in CSS pixels at 100%: its visible columns and rows, merges included. */
interface CellSize {
  width: number;
  height: number;
}

/**
 * Whether any cell of `model` has a gradient or pattern fill, so the shared
 * fill painter is fetched before the grid paints (a sheet with solid fills
 * only never loads it). Cells sharing a style share one object, so each style
 * is checked once.
 */
function needsFillPainter(model: SheetModel): boolean {
  const seen = new Set<object>();
  for (const row of model.rows) {
    for (const { cell } of row.cells) {
      const style = cell.style;
      if (style === undefined || seen.has(style)) continue;
      seen.add(style);
      if (style.fill !== undefined && style.fill.kind !== "solid") return true;
    }
  }
  return false;
}

/**
 * The colour automatic text reads against over a non-solid fill: a
 * gradient's middle, blended in sRGB as Excel blends cells, or a pattern's
 * background (most of its pixels in every pattern but the darkest).
 */
function fillInkColor(fill: ShapeFillModel): string | undefined {
  switch (fill.kind) {
    case "solid":
      return fill.color;
    case "linear-gradient":
    case "path-gradient": {
      const first = fill.stops[0];
      const last = fill.stops[fill.stops.length - 1];
      return first === undefined || last === undefined ? undefined : mixResolvedColorsSrgb(first.color, last.color, 0.5);
    }
    case "pattern":
      return fill.background;
    default:
      return undefined;
  }
}

/**
 * Paint a gradient or pattern cell fill behind the cell's text: an SVG layer
 * the cell's size, through the shared fill painter, as PPTX table cells do.
 * Patterns anchor at the cell's own top-left, as Excel draws them (probe
 * xlsx-cell-fills). `isolation` keeps the layer's negative z-index inside the
 * cell, above the grid and under the text. The painter was loaded before the
 * grid painted (`needsFillPainter`); a fill it cannot paint is marked.
 */
function paintCellFill(cellElement: HTMLElement, fill: ShapeFillModel, size: CellSize): void {
  const painter = loadedFillPainter();
  if (painter === undefined || !(size.width > 0 && size.height > 0)) {
    markCellFill(cellElement, painter === undefined ? "painter-unavailable" : "empty-cell");
    return;
  }
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("data-aared-element", "cell-fill");
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("viewBox", `0 0 ${size.width} ${size.height}`);
  svg.setAttribute("preserveAspectRatio", "none");
  svg.setAttribute("class", CELL_FILL);
  const paint = painter.fillPaint(svg, `aared-cell-fill-${++cellFillIds}`, fill, {
    width: size.width,
    height: size.height,
    unitsPerPxX: 1,
    unitsPerPxY: 1,
    slideX: 0,
    slideY: 0,
  }, { urls: undefined });
  if (paint === undefined) {
    markCellFill(cellElement, painter.fillFallbackReason(fill, undefined));
    return;
  }
  const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
  rect.setAttribute("width", String(size.width));
  rect.setAttribute("height", String(size.height));
  rect.setAttribute("fill", paint);
  svg.append(rect);
  cellElement.style.position = "relative";
  cellElement.style.isolation = "isolate";
  cellElement.prepend(svg);
}

/** A cell fill the painter could not draw: its reason, and in development builds a dashed outline. */
function markCellFill(cellElement: HTMLElement, reason: string): void {
  cellElement.setAttribute("data-aared-fill-fallback", reason);
  if (MARK_FALLBACKS) cellElement.style.outline = "1px dashed #d97706";
}

let cellFillIds = 0;

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
function paintCell(cellElement: HTMLElement, cell: CellModel, sheetRtl: boolean, paint: GridPaint, size: CellSize): CellPlacement {
  const style = cell.style;
  const wrap = style?.wrapText === true;
  // Default gridlines are content chrome resolved per gridline mode; an
  // authored border always wins below. Hidden gridlines stay transparent
  // (never `none`) so authored and default borders share one box metric.
  cellElement.setAttribute("class", classes(CELL, wrap && CELL_WRAP, !paint.showGridlines && NO_GRIDLINES, paint.selectable && SELECTABLE));
  if (style?.fontSizePt !== undefined && style.fontSizePt !== 10) cellElement.style.fontSize = `${style.fontSizePt}pt`;
  if (style?.fontFamily !== undefined) cellElement.style.fontFamily = cssFontFamily(style.fontFamily);
  if (style?.bold === true || cell.conditional?.bold === true) cellElement.style.fontWeight = "700";
  if (style?.italic === true) cellElement.style.fontStyle = "italic";
  if (style?.underline === true) cellElement.style.textDecoration = "underline";
  // Fills are never defaulted: no fill lets the white host surface show
  // through, preserving provenance. A conditional colour overlays the cell's
  // own fill; a gradient or pattern paints as a layer behind the text.
  const ownFill = style?.fill;
  const fill = cell.conditional?.fill ?? (ownFill?.kind === "solid" ? ownFill.color : undefined);
  if (fill !== undefined) cellElement.style.backgroundColor = fill;
  else if (ownFill !== undefined) paintCellFill(cellElement, ownFill, size);
  // An authored or conditional color paints verbatim; automatic text is
  // chosen for the fill it sits on (Excel black, `CELL`'s own, without one).
  const ink = fill ?? (ownFill === undefined ? undefined : fillInkColor(ownFill));
  const color = cell.conditional?.color ?? style?.color ?? paint.automaticColor(ink);
  if (color !== paint.automaticColor(undefined)) cellElement.style.color = color;
  const placement: CellPlacement = { align: textAlign(cell, sheetRtl), vertical: style?.verticalAlign ?? "bottom" };
  const hasContent = !isEmptyCell(cell) || style?.indent !== undefined;
  if (hasContent) cellElement.style.textAlign = placement.align;
  if (placement.vertical !== "bottom") cellElement.style.verticalAlign = placement.vertical;
  if (style?.indent !== undefined) cellElement.style.paddingInlineStart = `${CELL_PAD_PX + style.indent * 8}px`;
  // Sheet RTL flips column order via the table's `dir`, but cell text follows
  // Excel's per-cell reading order: 1 = LTR, 2 = RTL, otherwise context. An
  // unspecified order must not inherit the sheet direction, or LTR values
  // such as "180,000 EGP" and "13 Sep 2026" visually reorder in RTL sheets
  // (and a data bar or indent would flip side). An empty cell has nothing to
  // order.
  if (style?.readingOrder === 2) cellElement.dir = "rtl";
  else if (style?.readingOrder === 1) cellElement.dir = "ltr";
  else if (hasContent) cellElement.dir = "auto";
  // Each cell paints its own four edges. Where neighbours both author the
  // edge between them (Excel's All Borders writes it on both), the table's
  // separate borders draw it twice, 2px against Excel's one line; kept by
  // choice (Notes §116).
  const border = style?.border;
  if (border?.top !== undefined) cellElement.style.borderTop = borderCss(border.top.style, border.top.color);
  if (border?.left !== undefined) cellElement.style.borderLeft = borderCss(border.left.style, border.left.color);
  if (border?.right !== undefined) cellElement.style.borderRight = borderCss(border.right.style, border.right.color);
  if (border?.bottom !== undefined) cellElement.style.borderBottom = borderCss(border.bottom.style, border.bottom.color);
  if (cell.formula !== undefined) cellElement.setAttribute("data-aared-formula", cell.formula);
  if (cell.type === "error") cellElement.setAttribute("data-aared-cell-error", cell.formatted);
  return placement;
}

interface GridGeometry {
  visibleColumns: number[];
  visibleRows: number[];
  columnOffsets: number[];
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
  let y = 0;
  for (const height of rowHeights) y += height;
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
  /** Where a scrolling (not frozen, not hidden) row starts in the scroller, in host pixels; `undefined` for any other row. */
  dataRowTop(row: number): number | undefined;
  /** Synchronously re-render the virtualization window (used after programmatic scrolls). */
  renderWindow(): void;
  /** Resolves once this session's picture effects are painted (or marked). */
  effectsReady: Promise<void>;
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
  /** The cell under a pointer on the viewport fill, numbered past the extent; `null` off the fill. */
  fillCellAt(target: HTMLElement, clientX: number): { row: number; col: number } | null;
  /** The fill rows and columns painted now, as the indexes they continue the extent with. */
  fillSpan(): { rows: number[]; columns: number[] };
  /** Paint the selection over the cells past the extent; the cells inside it are painted by their own classes. */
  paintSelection(range: { top: number; bottom: number; left: number; right: number } | null, cursor: { row: number; col: number } | null): void;
  /** Paint one dashed boundary per same-sheet formula reference range. */
  paintReferenceRanges(ranges: Array<{ range: { top: number; bottom: number; left: number; right: number }; color: string }>): void;
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
  // The scrolling rows (frozen ones always render at the top), where each one
  // starts, and their total height: fixed for this model, so a scroll frame
  // bisects them instead of rebuilding them.
  const dataRows = geometry.visibleRows.filter((row) => !frozenRowSet.has(row));
  const dataTops: number[] = [];
  let dataTotal = 0;
  for (const row of dataRows) {
    dataTops.push(dataTotal);
    dataTotal += rowHeightOf(row);
  }
  /** The first index in `dataRows` for which `isPast` holds, or the row count; `isPast` is false, then true, down the rows. */
  const firstIndexWhere = (isPast: (index: number) => boolean): number => {
    let low = 0;
    let high = dataRows.length;
    while (low < high) {
      const mid = (low + high) >>> 1;
      if (isPast(mid)) high = mid;
      else low = mid + 1;
    }
    return low;
  };
  const dataRowTop = (row: number): number | undefined => {
    const index = firstIndexWhere((candidate) => dataRows[candidate]! >= row);
    return dataRows[index] === row ? dataTops[index] : undefined;
  };

  const host = document.createElement("div");
  host.dataset.aaredElement = "sheet";
  host.setAttribute("class", CONTENT_TEXT);
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
  // until the viewport is full. The cells never enter the model or the
  // extent: they are empty, so they carry no content, and a click on one
  // selects it by position (`fillCellAt`), with the row and column numbers
  // continuing past the extent. They are aria-hidden, as the headings are.
  const realWidthPx = geometry.totalWidth + headerWidthPx;
  const realHeightPx = geometry.totalHeight + headerHeightPx;
  const fillRowPx = Math.max(model.defaultRowHeight * PX_PER_INCH, 1);
  /** Columns past the sheet continue at the sheet's own default column width. */
  const fillColumnPx = Math.max(Math.round(model.defaultColumnWidth * PX_PER_INCH), 1);
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
  let viewportCache: { unbounded: boolean; roomH: number; clipped: boolean; clientH: number; scrollH: number } | null = null;
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
    // A cap already in effect would mask a real size change below it from
    // the dirty check — a caller growing the scroller to measure the
    // sheet's natural size, as the print-capture harness does, reads as
    // unchanged otherwise. Lift it before measuring, only when one is
    // actually set (the common bounded-viewer case never sets one, so this
    // never costs it an extra layout); the scroll position survives either way.
    const priorCap = scroller.style.maxHeight;
    const { scrollTop, scrollLeft } = scroller;
    if (priorCap !== "") scroller.style.maxHeight = "";
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
      cached = { unbounded, roomH, clipped: clipper !== null, clientH, scrollH };
      viewportCache = cached;
    }
    // The external room is outer (unzoomed) space; everything it is compared
    // with lives inside the zoomed host, so it is converted into host pixels.
    const roomH = cached.roomH / scale;
    // A content-sized sheet inside a clipping container (a card, a viewer
    // frame) scrolls itself within the visible room, as Excel's grid does:
    // otherwise the container scrolls the whole sheet, frozen rows and
    // column letters scroll away with it, and the horizontal scrollbar sits
    // at the sheet's far bottom. With no clipping container the page scrolls
    // the sheet, as before.
    const capPx = cached.unbounded && cached.clipped && roomH >= MIN_SCROLL_ROOM_PX ? `${Math.floor(roomH)}px` : "";
    if (scroller.style.maxHeight !== capPx) scroller.style.maxHeight = capPx;
    if (priorCap !== "") {
      scroller.scrollTop = scrollTop;
      scroller.scrollLeft = scrollLeft;
    }
    if (paint.fillViewport && cached.unbounded && roomH > clientH) {
      const floor = `${Math.floor(roomH)}px`;
      if (scroller.style.minHeight !== floor) scroller.style.minHeight = floor;
    }
    const windowW = scroller.clientWidth;
    const boxH = scroller.clientHeight;
    return {
      fillH: !paint.fillViewport ? 0 : cached.unbounded ? roomH : clientH,
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
    cell.setAttribute("class", classes(FILL_CELL, lines, !paint.showGridlines && NO_GRIDLINES, paint.selectable && SELECTABLE));
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
    table.style.setProperty("--aared-fill-col", `${fillColumnPx}px`);
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
          ? `repeating-linear-gradient(to left,${GRIDLINE} 0 ${hair}px,transparent ${hair}px ${fillColumnPx}px)`
          : `repeating-linear-gradient(to right,transparent 0 ${fillColumnPx - hair}px,${GRIDLINE} ${fillColumnPx - hair}px ${fillColumnPx}px)`,
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
    const columns = viewportWidth > 0 ? fillTracks(viewportWidth - realWidthPx, fillColumnPx) : [];
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
    return cell === undefined || (cell.value === null && cell.formatted === "" && cell.conditional?.dataBar === undefined && cell.conditional?.icon === undefined);
  };

  /**
   * Excel lets unwrapped text overflow into empty neighbours (toward the end
   * edge for left-aligned text, the start edge for right-aligned, both for
   * centered) and stops at the first cell with content. Numbers never spill.
   * This is the room the text may take; how far it really reaches, and so
   * which gridlines it crosses, is read off the painted text
   * (`settleSpillLines`).
   */
  const spillsForRow = (
    row: number,
    entries: Map<number, CellModel> | undefined,
  ): Map<number, TextSpill> => {
    const spills = new Map<number, TextSpill>();
    if (entries === undefined) return spills;
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
      const after = right.reduce((sum, c) => sum + columnWidthPx(c), 0);
      // `left` and `right` list the nearest neighbour first.
      spills.set(col, { width: offset + columnWidthPx(col) + after, offset, after, run: [...left].reverse().concat(col, right) });
    }
    return spills;
  };

  /** This window's spilled text and the cells of its run (`TextSpill.run`), for `settleSpillLines`. */
  let windowSpills: Array<{ text: HTMLElement; run: HTMLElement[]; keep: boolean[] }> = [];

  /**
   * Excel draws no gridline through text that runs across cells, and every
   * other gridline as usual: a short word beside empty cells keeps all of
   * them. So the lines are settled from the painted text, not from the room
   * it may take: one read of every spilled text and run cell in the window,
   * then one write. A boundary is a cell's right edge; text reaching past it
   * hides it, unless the cell authors its own right border. Run again when a
   * font finishes loading, since that changes how far the text reaches.
   */
  const settleSpillLines = (): void => {
    const hide = new Set<HTMLElement>();
    const show: HTMLElement[] = [];
    for (const { text, run, keep } of windowSpills) {
      if (!text.isConnected) continue;
      const box = text.getBoundingClientRect();
      for (let i = 0; i < run.length - 1; i++) {
        if (keep[i]) continue;
        const cell = run[i]!;
        const edge = cell.getBoundingClientRect().right;
        if (box.width > 0 && box.left < edge - 0.5 && box.right > edge + 0.5) hide.add(cell);
        else show.push(cell);
      }
    }
    for (const cell of show) if (!hide.has(cell)) cell.style.borderRightColor = "";
    for (const cell of hide) cell.style.borderRightColor = "transparent";
  };

  /** A cell's painted box in CSS pixels: its visible columns and rows, through a merge's span. */
  const cellSize = (row: number, col: number, merge: { endRow: number; endCol: number } | undefined): CellSize => {
    const endRow = Math.min(merge?.endRow ?? row, model.extent.rows - 1);
    const endCol = Math.min(merge?.endCol ?? col, model.extent.cols - 1);
    let width = 0;
    for (let probe = col; probe <= endCol; probe++) {
      if (!hiddenColumns.has(probe)) width += (model.columnWidths[probe] ?? 0) * PX_PER_INCH;
    }
    let height = 0;
    for (let probe = row; probe <= endRow; probe++) {
      if (!hiddenRowSet.has(probe)) height += rowHeightOf(probe);
    }
    return { width, height };
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
      if (frozenRowSet.has(row)) {
        // A frozen row's number pins with its row, as in Excel. Above the
        // scrolling row numbers that pass beneath it (same column, later in
        // the DOM); the corner and column letters never overlap it.
        number.style.top = `${geometry.frozenRowOffset.get(row) ?? 0}px`;
        number.style.zIndex = "5";
      }
      line.append(number);
    }
    const entries = cellsByRow.get(row);
    const spills = spillsForRow(row, entries);
    const painted = new Map<number, HTMLElement>();
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
      const placement = paintCell(cellElement, effective, model.rightToLeft, paint, cellSize(row, col, merge));
      const spill = spills.get(col);
      // Only a data bar, an icon or a spill overlay needs the cell as a
      // containing block (a frozen cell's sticky, set below, is one too).
      if (spill !== undefined || effective.conditional?.dataBar !== undefined || effective.conditional?.icon !== undefined) cellElement.style.position = "relative";
      if (spill !== undefined) {
        cellElement.style.overflow = "visible";
        cellElement.style.zIndex = "1";
      }
      appendCellContent(cellElement, effective, placement, spill);
      paintFrozen(cellElement, row, col);
      painted.set(col, cellElement);
      line.append(cellElement);
    }
    for (const [col, spill] of spills) {
      const text = painted.get(col)?.querySelector<HTMLElement>('[data-aared-element="cell-text"]');
      const run = spill.run.map((c) => painted.get(c));
      if (text == null || run.some((cell) => cell === undefined)) continue;
      const keep = spill.run.map((c) => entries?.get(c)?.style?.border?.right !== undefined);
      windowSpills.push({ text, run: run as HTMLElement[], keep });
    }
    if (fillColumnWidths.length > 0) line.append(paintFillCell(row, fillColumnWidths.length, "tracks"));
    return line;
  };

  /** One empty row below the data, numbered on from the last real row. */
  const paintFillRow = (position: number, height: number): HTMLTableRowElement => {
    const line = document.createElement("tr");
    line.setAttribute("data-aared-element", "fill-row");
    line.setAttribute("data-aared-fill-row", String(position));
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
      // One cell per column: each carries its own column, so a click names it.
      const visible = geometry.visibleColumns.length;
      for (let i = 0; i < columns; i++) {
        const cell = paintFillCell(undefined);
        cell.dataset.aaredCol = String(i < visible ? geometry.visibleColumns[i] : model.extent.cols + i - visible);
        line.append(cell);
      }
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
    windowSpills = [];
    // The rows whose box meets the window plus a pixel margin: they run from the
    // first row ending below its top to the last row starting above its bottom.
    const windowTop = scrollTop - ROW_OVERSCAN * 20;
    const windowBottom = scrollTop + viewport + ROW_OVERSCAN * 20;
    let startIndex = firstIndexWhere((i) => dataTops[i]! + rowHeightOf(dataRows[i]!) > windowTop);
    let endIndex = firstIndexWhere((i) => dataTops[i]! >= windowBottom) - 1;
    if (startIndex > endIndex) endIndex = -1;
    const offsets = dataTops;
    const total = dataTotal;
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
    settleSpillLines();
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
    // A picture box still waiting for attach (see `settleEffects`) paints
    // its effects once layout exists; drained lists make this a no-op.
    settleEffects();
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
  // A face that loads after the first paint changes how far spilled text reaches.
  const fonts = typeof document === "undefined" ? undefined : document.fonts;
  fonts?.addEventListener("loadingdone", settleSpillLines);
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

  /** The notice a chart shows in its frame when it cannot be drawn. */
  const paintChartFallback = (box: HTMLElement, reason: string, message: string): void => {
    box.dataset.aaredChartFallback = reason;
    box.style.display = "flex";
    box.style.alignItems = "center";
    box.style.justifyContent = "center";
    box.style.border = `1px dashed ${paint.chromeLine}`;
    box.style.font = "12px system-ui, sans-serif";
    box.style.color = paint.noticeColor;
    box.textContent = message;
  };

  const paintChartDrawing = (drawing: SheetChartModel): void => {
    const box = paintDrawingBox(drawing);
    box.dataset.aaredElement = "chart";
    box.dataset.aaredChartName = drawing.name;
    if (drawing.sourceChartId !== undefined) box.dataset.aaredChartSource = drawing.sourceChartId;
    drawingsLayer.append(box);
    if (drawing.spec === undefined) {
      paintChartFallback(box, drawing.unsupportedReason ?? "unsupported chart", drawing.unsupportedReason ?? "This chart type is not previewed.");
      return;
    }
    // The chart area's own fill and outline are part of the drawing, under
    // everything, gridlines included, exactly like the frame Office draws.
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.dataset.aaredElement = "chart-graphic";
    svg.style.width = "100%";
    svg.style.height = "100%";
    box.append(svg);
    paintChart(svg, drawing.spec, {
      frame: drawing.bounds,
      urls,
      alt: drawing.alt,
      unavailable: () => {
        svg.remove();
        paintChartFallback(box, "load-error", `${drawing.name} (unavailable)`);
      },
    });
  };

  // Picture effects resolve in the Worker to the shared vocabulary; the
  // effect painter loads on demand and paints them here, with cleanups owned
  // by this paint session (sheet switches and appearance repaints destroy
  // and repaint, so observers never outlive their box).
  const effectCleanups: Array<() => void> = [];
  const effectTargets: Array<{ box: HTMLElement; drawing: SheetPictureModel }> = [];
  const paintPicture = (drawing: SheetPictureModel): void => {
    const box = paintDrawingBox(drawing);
    box.dataset.aaredElement = "picture";
    box.dataset.aaredMedia = drawing.mediaId;
    // A rotated picture turns with its shadow's frame: the transform paints
    // here while the shadow math reads it from the model.
    paintTransform(box, drawing.transform);
    if (drawing.effects !== undefined) effectTargets.push({ box, drawing });
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
    const image = document.createElement("img");
    image.alt = drawing.name;
    image.style.width = "100%";
    image.style.height = "100%";
    image.style.objectFit = "contain";
    // The SVG source paints first; Office's raster fallback covers a read or
    // decode failure, and the named placeholder stays for neither.
    const showPlaceholder = (text: string): void => {
      box.replaceChildren(placeholder);
      placeholder.textContent = text;
    };
    // An `<img>` with no `src` yet paints nothing, so the box stays empty
    // until the source arrives; the placeholder returns only when neither
    // source paints.
    box.replaceChildren(image);
    paintImageSource(image, urls, {
      mediaId: drawing.mediaId,
      contentType: drawing.contentType,
      ...(drawing.fallbackMediaId === undefined || drawing.fallbackContentType === undefined ? {} : {
        fallbackMediaId: drawing.fallbackMediaId,
        fallbackContentType: drawing.fallbackContentType,
      }),
    }, () => showPlaceholder(`${drawing.name} (unavailable)`));
  };

  for (const drawing of model.drawings) {
    if (drawing.kind === "chart") paintChartDrawing(drawing);
    else paintPicture(drawing);
  }
  // Effects paint once their painter arrives; a failed fetch marks rather
  // than drops them, and a torn-down session paints nothing. A box not yet
  // attached (the painter was already loaded, so this pass runs before the
  // host is placed) goes back on the list: the resize observer below settles
  // it once layout exists, so it is never skipped for good.
  let effectsUnavailable = false;
  const settleEffects = (): void => {
    if (destroyed || effectTargets.length === 0) return;
    const painter = loadedEffectPainter();
    // renderWindow also calls this while the lazy chunk is still in flight.
    // Keep the targets until it settles; absence during download is not a
    // failed painter, and draining them here loses the effects permanently.
    if (painter === undefined && !effectsUnavailable) return;
    for (const target of effectTargets.splice(0)) {
      if (destroyed) return;
      if (target.drawing.effects === undefined) continue;
      if (!target.box.isConnected) {
        effectTargets.push(target);
        continue;
      }
      if (painter === undefined) target.box.dataset.aaredEffectFallback = "painter-unavailable";
      else painter.paintDrawingEffects(target.box, target.drawing.effects, drawingsLayer, target.drawing.transform, effectCleanups);
    }
  };
  const effectsReady: Promise<void> = (async () => {
    if (effectTargets.length === 0) return;
    if (loadedEffectPainter() === undefined) {
      await loadEffectPainter().then(() => undefined, () => { effectsUnavailable = true; });
    }
    settleEffects();
  })();

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
    fonts?.removeEventListener("loadingdone", settleSpillLines);
    for (const cleanup of effectCleanups.splice(0)) cleanup();
    urls?.close();
  };
  // ---- Selection past the data ---------------------------------------------
  // Fill cells are not in the model, so a click on one is placed by geometry:
  // the same column and row offsets the drawings and the spacers use.
  const selectionLayer = document.createElement("div");
  selectionLayer.dataset.aaredElement = "sheet-selection";
  selectionLayer.style.position = "absolute";
  selectionLayer.style.inset = "0";
  selectionLayer.style.pointerEvents = "none";
  canvas.append(selectionLayer);
  const visibleColumnIndex = new Map<number, number>();
  geometry.visibleColumns.forEach((col, index) => visibleColumnIndex.set(col, index));
  const visibleRowIndex = new Map<number, number>();
  const visibleRowTops: number[] = [];
  let visibleRowsHeight = 0;
  geometry.visibleRows.forEach((row, index) => {
    visibleRowIndex.set(row, index);
    visibleRowTops.push(visibleRowsHeight);
    visibleRowsHeight += rowHeightOf(row);
  });

  const fillSpan = (): { rows: number[]; columns: number[] } => ({
    rows: fillRowHeights.map((_, position) => model.extent.rows + position),
    columns: fillColumnWidths.map((_, index) => model.extent.cols + index),
  });

  const fillCellAt = (target: HTMLElement, clientX: number): { row: number; col: number } | null => {
    const cell = target.closest<HTMLElement>('td[data-aared-element="fill-cell"]');
    const line = target.closest<HTMLElement>("tr");
    if (cell === null || line === null) return null;
    let row = Number.NaN;
    if (line.dataset.aaredRow !== undefined) row = Number(line.dataset.aaredRow);
    else if (line.dataset.aaredFillRow !== undefined) row = model.extent.rows + Number(line.dataset.aaredFillRow);
    if (!Number.isInteger(row)) return null;
    // A cell that is one column in the DOM names its column; the others
    // are runs of columns, walked from the cell's start edge.
    if (cell.dataset.aaredCol !== undefined) return { row, col: Number(cell.dataset.aaredCol) };
    const runs: Array<{ col: number; width: number }> = [];
    if (cell.dataset.aaredFill === "row") {
      for (const col of geometry.visibleColumns) runs.push({ col, width: columnWidthPx(col) });
    }
    fillColumnWidths.forEach((width, index) => runs.push({ col: model.extent.cols + index, width }));
    const total = runs.reduce((sum, run) => sum + run.width, 0);
    const rect = cell.getBoundingClientRect();
    if (total <= 0 || rect.width <= 0) return null;
    // The host may be zoomed: the pointer's visual distance is scaled back to layout pixels.
    const offset = (model.rightToLeft ? rect.right - clientX : clientX - rect.left) * (total / rect.width);
    let start = 0;
    for (const run of runs) {
      if (offset < start + run.width) return { row, col: run.col };
      start += run.width;
    }
    return null;
  };

  const extentRows = model.extent.rows;
  const extentCols = model.extent.cols;
  const fillColumnStarts: number[] = [];
  let fillX = geometry.totalWidth;
  for (const width of fillColumnWidths) {
    fillColumnStarts.push(fillX);
    fillX += width;
  }
  const fillRowStarts: number[] = [];
  let fillY = geometry.totalHeight;
  for (const height of fillRowHeights) {
    fillRowStarts.push(fillY);
    fillY += height;
  }

  const rangeBounds = (range: { top: number; bottom: number; left: number; right: number }) => {
    let left = Number.POSITIVE_INFINITY;
    let right = Number.NEGATIVE_INFINITY;
    for (let index = 0; index < geometry.visibleColumns.length; index++) {
      const col = geometry.visibleColumns[index]!;
      if (col < range.left || col > range.right) continue;
      const x = geometry.columnOffsets[index]!;
      left = Math.min(left, x);
      right = Math.max(right, x + columnWidthPx(col));
    }
    for (let index = 0; index < fillColumnWidths.length; index++) {
      const col = extentCols + index;
      if (col < range.left || col > range.right) continue;
      const x = fillColumnStarts[index]!;
      left = Math.min(left, x);
      right = Math.max(right, x + fillColumnWidths[index]!);
    }

    let top = Number.POSITIVE_INFINITY;
    let bottom = Number.NEGATIVE_INFINITY;
    for (let index = 0; index < geometry.visibleRows.length; index++) {
      const row = geometry.visibleRows[index]!;
      if (row < range.top || row > range.bottom) continue;
      const y = visibleRowTops[index]!;
      top = Math.min(top, y);
      bottom = Math.max(bottom, y + rowHeightOf(row));
    }
    for (let index = 0; index < fillRowHeights.length; index++) {
      const row = extentRows + index;
      if (row < range.top || row > range.bottom) continue;
      const y = fillRowStarts[index]!;
      top = Math.min(top, y);
      bottom = Math.max(bottom, y + fillRowHeights[index]!);
    }

    if (!Number.isFinite(left) || !Number.isFinite(top)) return null;
    return { x: left, y: top, width: right - left, height: bottom - top };
  };

  const placeSelectionOverlay = (element: HTMLElement, box: { x: number; width: number; y: number; height: number }): void => {
    element.style.position = "absolute";
    element.style[model.rightToLeft ? "right" : "left"] = `${headerWidthPx + box.x}px`;
    element.style.width = `${box.width}px`;
    element.style.top = `${headerHeightPx + box.y}px`;
    element.style.height = `${box.height}px`;
    selectionLayer.append(element);
  };

  const appendRangeBoundary = (
    range: { top: number; bottom: number; left: number; right: number },
    elementName: string,
    color: string,
    offsetX = 0,
    offsetY = 0,
  ): void => {
    const bounds = rangeBounds(range);
    if (bounds === null) return;
    const frame = document.createElement("div");
    frame.dataset.aaredElement = elementName;
    placeSelectionOverlay(frame, { ...bounds, x: bounds.x + offsetX, y: bounds.y + offsetY });
    frame.style.boxSizing = "border-box";
    frame.style.border = `2px dashed ${color}`;
    frame.style.zIndex = "2";
    frame.style.pointerEvents = "none";
  };

  const paintReferenceRanges = (
    ranges: Array<{ range: { top: number; bottom: number; left: number; right: number }; color: string }>,
  ): void => {
    selectionLayer.querySelectorAll('[data-aared-element="sheet-reference-boundary"]').forEach((frame) => frame.remove());
    for (const [index, { range, color }] of ranges.entries()) {
      let offsetX = 0;
      let offsetY = 0;
      const origin = host.querySelector<HTMLElement>(`td[data-aared-formula-reference="true"][data-aared-reference-index="${index}"]`);
      const originRow = Number(origin?.dataset.aaredRow);
      const originCol = Number(origin?.dataset.aaredCol);
      const cellBounds = origin === null ? null : rangeBounds({ top: originRow, bottom: originRow, left: originCol, right: originCol });
      if (origin !== null && cellBounds !== null) {
        // Table layout can offset a cell slightly from the model geometry
        // (notably at a frozen-pane boundary); align the full frame to a real
        // cell in the referenced range while preserving its geometry-derived
        // width and height.
        const canvasRect = canvas.getBoundingClientRect();
        const cellRect = origin.getBoundingClientRect();
        offsetX = model.rightToLeft
          ? canvasRect.right - cellRect.right - headerWidthPx - cellBounds.x
          : cellRect.left - canvasRect.left - headerWidthPx - cellBounds.x;
        offsetY = cellRect.top - canvasRect.top - headerHeightPx - cellBounds.y;
      }
      appendRangeBoundary(range, "sheet-reference-boundary", color, offsetX, offsetY);
    }
  };

  const paintSelection = (
    range: { top: number; bottom: number; left: number; right: number } | null,
    cursor: { row: number; col: number } | null,
  ): void => {
    selectionLayer.replaceChildren();
    /** A cell's box in sheet pixels, or `null` when it is inside the extent and painted by its own cell, or not painted at all. */
    const boxOf = (row: number, col: number): { x: number; width: number; y: number; height: number } | null => {
      const inside = row < extentRows && col < extentCols;
      if (inside) return null;
      let x: number;
      let width: number;
      if (col < extentCols) {
        const index = visibleColumnIndex.get(col);
        if (index === undefined) return null;
        x = geometry.columnOffsets[index]!;
        width = columnWidthPx(col);
      } else {
        const index = col - extentCols;
        if (index >= fillColumnWidths.length) return null;
        x = fillColumnStarts[index]!;
        width = fillColumnWidths[index]!;
      }
      let y: number;
      let height: number;
      if (row < extentRows) {
        const index = visibleRowIndex.get(row);
        if (index === undefined) return null;
        y = visibleRowTops[index]!;
        height = rowHeightOf(row);
      } else {
        const index = row - extentRows;
        if (index >= fillRowHeights.length) return null;
        y = fillRowStarts[index]!;
        height = fillRowHeights[index]!;
      }
      return { x, width, y, height };
    };
    // The range wash goes in one box per run of cells in a row: a drag over
    // empty cells is a handful of boxes, not one per cell.
    if (range !== null) {
      for (let row = range.top; row <= range.bottom; row++) {
        let run: { x: number; end: number; y: number; height: number } | undefined;
        let connected = false;
        for (let col = range.left; col <= range.right; col++) {
          const isCursor = cursor !== null && cursor.row === row && cursor.col === col;
          const box = boxOf(row, col);
          if (box === null || isCursor) {
            connected = false;
            continue;
          }
          if (connected && run !== undefined && run.y === box.y) {
            run.end = box.x + box.width;
          } else {
            if (run !== undefined) {
              const element = document.createElement("div");
              element.dataset.aaredElement = "sheet-selection-range";
              element.style.boxShadow = RANGE_WASH;
              placeSelectionOverlay(element, { x: run.x, width: run.end - run.x, y: run.y, height: run.height });
            }
            run = { x: box.x, end: box.x + box.width, y: box.y, height: box.height };
          }
          connected = true;
        }
        if (run !== undefined) {
          const element = document.createElement("div");
          element.dataset.aaredElement = "sheet-selection-range";
          element.style.boxShadow = RANGE_WASH;
          placeSelectionOverlay(element, { x: run.x, width: run.end - run.x, y: run.y, height: run.height });
        }
      }

    }
    // The active empty cell keeps its solid cursor frame, also when it is
    // inside an extended range; formula-reference ranges use dashed frames.
    if (cursor !== null) {
      const box = boxOf(cursor.row, cursor.col);
      if (box !== null) {
        const element = document.createElement("div");
        element.dataset.aaredElement = "sheet-selection-cursor";
        element.style.backgroundColor = "#ffffff";
        element.style.zIndex = "1";
        element.style.outline = "2px solid var(--aared-selection)";
        element.style.outlineOffset = "-2px";
        placeSelectionOverlay(element, box);
      }
    }
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
    fillCellAt,
    fillSpan,
    paintSelection,
    paintReferenceRanges,
    dataRowTop,
    renderWindow,
    effectsReady,
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
  dark: boolean;
  headers: boolean;
  sheetTabs: boolean;
  fill: boolean;
  selection: boolean;
  gridlines: GridlineMode;
  formulaBar: boolean;
}

function normalizeAppearance(appearance: Appearance | undefined): NormalizedAppearance {
  const defaults = APPEARANCE_DEFAULTS.spreadsheet;
  return {
    dark: resolveTheme(appearance?.theme ?? APPEARANCE_DEFAULTS.theme) === "dark",
    headers: appearance?.spreadsheet?.headers ?? defaults.headers,
    sheetTabs: appearance?.spreadsheet?.sheetTabs ?? defaults.sheetTabs,
    fill: appearance?.spreadsheet?.fill ?? defaults.fill,
    selection: appearance?.spreadsheet?.selection ?? defaults.selection,
    gridlines: appearance?.spreadsheet?.gridlines ?? defaults.gridlines,
    formulaBar: appearance?.spreadsheet?.formulaBar ?? defaults.formulaBar,
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
    selectable: appearance.selection,
    headerFill: HEADER_FILL,
    headerColor: HEADER_COLOR,
    chromeLine: GRIDLINE_COLOR,
    noticeColor: PALETTE.light.muted,
    hostBackground: "white",
    dark: appearance.dark,
  };
}

/**
 * A target's own layout from before the first XLSX render on it set
 * `display: flex`, and the render that owns the target now (the newest).
 */
const targetLayouts = new WeakMap<HTMLElement, { display: string; flexDirection: string; owner: object }>();

/** A fit never scales a sheet below this: a huge sheet scrolls rather than shrinking to nothing (and the row window stays bounded). */
const MIN_FIT_SCALE = 0.25;

/**
 * The `render()` controller: zoom, cell navigation, cursor-plus-range selection,
 * sheet tabs, and lifecycle over one target element. Owns a stable tab bar
 * plus grid container; each painted sheet swaps the grid host in place.
 */
class XlsxViewer extends RenderHandleBase<Appearance> implements XlsxRenderHandle {
  readonly #target: HTMLElement;
  readonly #gridBox: HTMLDivElement;
  #tabsBar: HTMLDivElement | null = null;
  /** Read-only Excel-like chrome, built only when the appearance asks for it. */
  #formulaBar: HTMLDivElement | null = null;
  #nameBox: HTMLDivElement | null = null;
  #formulaText: HTMLDivElement | null = null;
  #aggregateButton: HTMLButtonElement | null = null;
  #aggregateSummary: HTMLSpanElement | null = null;
  #aggregateChevron: HTMLSpanElement | null = null;
  #aggregateMenu: HTMLDivElement | null = null;
  #aggregateKey: "sum" | "average" | "min" | "max" | "count" = "sum";
  #aggregateMenuOpen = false;
  #session: SheetPaintSession | null = null;
  #model: SheetModel;
  readonly #handle: SheetHandle | undefined;
  readonly #document: XlsxDocument | undefined;
  readonly #mediaResolver: MediaResolver | undefined;
  #appearance: NormalizedAppearance;
  /** The appearance as given (merged across `setAppearance()` calls), before defaults fill it in. */
  #givenAppearance: Appearance | undefined;
  #selection: CellRef | null = null;
  /** The range anchor: where an extension started. The cursor (`#selection`) is its moving end. */
  #anchor: CellRef | null = null;
  #range: { start: CellRef; end: CellRef } | null = null;
  /** A drag-select in flight (pointer down on a cell, not yet up). */
  #dragging = false;
  /** Set once the pointer moves past the anchor, so the release click does not collapse the range. */
  #dragged = false;
  #sheet: SheetInfo;
  #sheetGeneration = 0;
  #sheetController: AbortController | null = null;
  readonly #listeners = new AbortController();
  /** Keeps a fit mode current when the room around the grid changes. */
  readonly #fitObserver: ResizeObserver;
  /** The room the last fit measured; a resize that leaves it unchanged is ignored. */
  #fitRoom = { width: -1, height: -1 };
  /** This render's hold on the class sheet in the target's root. */
  readonly #classes: ClassSheetScope;
  /** Stops watching for a newer render replacing this one in the target. */
  readonly #stopWatching: () => void;

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
    this.#givenAppearance = appearance;
    this.#appearance = normalizeAppearance(appearance);
    this.#sheet = handle === undefined
      ? { index: 0, title: model.name, state: "visible" }
      : { ...handle.info };
    // The first render on a target keeps the target's own layout; a render
    // that replaces another keeps that original, not the flex it found, and
    // becomes the one that puts it back.
    const layout = targetLayouts.get(target) ?? { display: target.style.display, flexDirection: target.style.flexDirection, owner: this };
    layout.owner = this;
    targetLayouts.set(target, layout);
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
    this.#syncChrome();
    const children: HTMLElement[] = [];
    if (this.#formulaBar !== null) children.push(this.#formulaBar);
    if (this.#tabsBar !== null) children.push(this.#tabsBar);
    children.push(this.#gridBox);
    target.replaceChildren(...children);
    // A newer render into the same target replaced this one: tear it down, as every format does.
    this.#stopWatching = onReplaced(target, this.#gridBox, () => this.destroy());
    const signal = this.#listeners.signal;
    this.#gridBox.addEventListener("click", (event) => this.#handleClick(event), { signal });
    // Middle-clicks bypass `click`: the same interception applies so a link
    // can never navigate automatically.
    this.#gridBox.addEventListener("auxclick", (event) => this.#handleClick(event), { signal });
    this.#gridBox.addEventListener("mousedown", (event) => this.#handlePointerDown(event), { signal });
    this.#gridBox.addEventListener("mousemove", (event) => this.#handlePointerMove(event), { signal });
    this.#gridBox.addEventListener("mouseup", (event) => this.#handlePointerUp(event), { signal });
    this.#gridBox.addEventListener("dblclick", (event) => this.#handleDoubleClick(event), { signal });
    this.#gridBox.addEventListener("keydown", (event) => this.#handleKeyDown(event), { signal });
    document.addEventListener("pointerdown", (event) => {
      const bar = this.#formulaBar;
      if (this.#aggregateMenuOpen && (bar === null || !event.composedPath().includes(bar))) this.#closeAggregateMenu();
    }, { signal });
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape" && this.#aggregateMenuOpen) this.#closeAggregateMenu();
    }, { signal });
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
    this.#givenAppearance = mergeAppearance(this.#givenAppearance, appearance);
    this.#appearance = normalizeAppearance(this.#givenAppearance);
    this.#syncTabsChrome();
    this.#syncChrome();
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

  /** Resolves once the current paint session's picture effects are painted (or marked). */
  get effectsReady(): Promise<void> {
    return this.#session?.effectsReady ?? Promise.resolve();
  }

  get selection(): CellRef | null {
    return this.#selection === null ? null : { ...this.#selection };
  }

  get range(): { start: CellRef; end: CellRef } | null {
    return this.#range === null ? null : { start: { ...this.#range.start }, end: { ...this.#range.end } };
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
      // As `render()` does: a delivered font loads only when text paints with it,
      // and the fill painter only for a sheet with a gradient or pattern cell.
      await awaitDocumentFonts(model);
      if (needsFillPainter(model)) await loadFillPainter().catch(() => undefined);
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
    this.#anchor = null;
    this.#range = null;
    this.#repaintCurrent();
    this.setWarnings(sheetWarnings(model));
    this.#refreshChrome();
    // Reset the grid scroll: the new sheet paints a fresh scroller at origin.
    if (hadSelection) {
      this.dispatchEvent(new CustomEvent("selectionchange", { detail: { sheet: this.#sheetRef(), selection: null, range: null } }));
    }
    this.dispatchEvent(new CustomEvent("sheetchange", { detail: this.#sheetRef() }));
  }

  /** The visible sheet as a unit reference — how every event names it. */
  #sheetRef(): UnitRef {
    return { index: this.#sheet.index, title: this.#sheet.title };
  }

  #teardown(): void {
    this.#stopWatching();
    this.#sheetGeneration += 1;
    this.#sheetController?.abort();
    this.#sheetController = null;
    this.#listeners.abort();
    this.#fitObserver.disconnect();
    this.#session?.destroy();
    this.#session = null;
    // The target's own layout goes back with the newest XLSX render on it:
    // destroyed, emptied by the host, or replaced by another format. One that
    // a newer sheet replaced leaves it to that sheet.
    const layout = targetLayouts.get(this.#target);
    if (layout?.owner === this) {
      this.#target.style.display = layout.display;
      this.#target.style.flexDirection = layout.flexDirection;
      targetLayouts.delete(this.#target);
    }
    this.#classes.release();
  }

  /**
   * Read-only Excel-like chrome: a formula bar (name box, formula/value, and
   * range aggregates) above the grid. Built only when the appearance asks,
   * kept in step on selection, sheet and appearance changes.
   */
  #syncChrome(): void {
    if (this.#appearance.formulaBar && this.#formulaBar === null) {
      const bar = document.createElement("div");
      bar.dataset.aaredElement = "sheet-formula-bar";
      bar.style.display = "flex";
      bar.style.flex = "none";
      bar.style.alignItems = "center";
      bar.style.position = "relative";
      // Keep the open aggregate menu above sticky sheet headings, which use
      // their own z-index to stay visible while the grid scrolls.
      bar.style.zIndex = "10";
      bar.style.direction = "ltr";
      bar.style.height = "44px";
      bar.style.borderBottom = "1px solid var(--aared-gridline)";
      bar.style.backgroundColor = "var(--aared-surface)";
      const name = document.createElement("div");
      name.dataset.aaredElement = "sheet-name-box";
      name.style.minWidth = "92px";
      name.style.boxSizing = "border-box";
      name.style.alignSelf = "stretch";
      name.style.display = "flex";
      name.style.alignItems = "center";
      name.style.textAlign = "center";
      name.style.justifyContent = "center";
      name.style.font = "500 13px ui-monospace, monospace";
      name.style.color = "var(--aared-ink)";
      name.style.borderRight = "1px solid var(--aared-gridline)";
      const fx = document.createElement("div");
      fx.textContent = "fx";
      fx.setAttribute("aria-hidden", "true");
      fx.style.display = "flex";
      fx.style.alignItems = "center";
      fx.style.alignSelf = "stretch";
      fx.style.padding = "0 14px";
      fx.style.borderRight = "1px solid var(--aared-gridline)";
      fx.style.font = "italic 500 16px Georgia, 'Times New Roman', serif";
      fx.style.color = "var(--aared-muted-text)";
      const text = document.createElement("div");
      text.dataset.aaredElement = "sheet-formula-text";
      text.style.flex = "1 1 auto";
      text.style.minWidth = "0";
      text.style.overflow = "hidden";
      text.style.textOverflow = "ellipsis";
      text.style.whiteSpace = "nowrap";
      text.style.padding = "0 16px";
      text.style.fontSize = "14px";
      text.style.color = "var(--aared-ink)";
      const aggregate = document.createElement("button");
      aggregate.type = "button";
      aggregate.dataset.aaredElement = "sheet-aggregate-button";
      aggregate.setAttribute("aria-haspopup", "menu");
      aggregate.setAttribute("aria-expanded", "false");
      aggregate.style.display = "none";
      aggregate.style.alignItems = "center";
      aggregate.style.gap = "7px";
      aggregate.style.alignSelf = "stretch";
      aggregate.style.padding = "0 14px";
      aggregate.style.border = "0";
      aggregate.style.borderLeft = "1px solid var(--aared-gridline)";
      aggregate.style.background = "transparent";
      aggregate.style.color = "var(--aared-muted-text)";
      aggregate.style.font = "13px system-ui, sans-serif";
      aggregate.style.whiteSpace = "nowrap";
      aggregate.style.cursor = "pointer";
      const summary = document.createElement("span");
      const chevron = document.createElement("span");
      chevron.setAttribute("aria-hidden", "true");
      chevron.style.width = "6px";
      chevron.style.height = "6px";
      chevron.style.borderRight = "1.5px solid var(--aared-muted-text)";
      chevron.style.borderBottom = "1.5px solid var(--aared-muted-text)";
      chevron.style.transform = "translateY(-2px) rotate(45deg)";
      aggregate.append(summary, chevron);
      aggregate.addEventListener("click", () => this.#toggleAggregateMenu());
      const menu = document.createElement("div");
      menu.dataset.aaredElement = "sheet-aggregate-menu";
      menu.setAttribute("role", "menu");
      menu.hidden = true;
      menu.style.position = "absolute";
      menu.style.top = "calc(100% + 6px)";
      menu.style.right = "6px";
      menu.style.width = "240px";
      menu.style.maxWidth = "calc(100% - 12px)";
      menu.style.boxSizing = "border-box";
      menu.style.minWidth = "0";
      menu.style.padding = "4px";
      menu.style.zIndex = "20";
      menu.style.background = "var(--aared-surface)";
      menu.style.color = "var(--aared-ink)";
      menu.style.border = "1px solid var(--aared-border)";
      menu.style.borderRadius = "8px";
      menu.style.boxShadow = "0 10px 28px rgba(0,0,0,.12), 0 2px 6px rgba(0,0,0,.06)";
      menu.addEventListener("click", (event) => {
        const item = (event.target as Element).closest<HTMLButtonElement>("[data-aggregate]");
        if (item === null) return;
        this.#aggregateKey = item.dataset.aggregate as "sum" | "average" | "min" | "max" | "count";
        this.#closeAggregateMenu();
        this.#refreshChrome();
      });
      bar.append(name, fx, text, aggregate, menu);
      paintTokens(bar, this.#appearance.dark ? "dark" : "light");
      paintSheetTokens(bar, {
        dark: this.#appearance.dark,
        gridline: this.#appearance.dark ? GRIDLINE_DARK : GRIDLINE_COLOR,
        headerFill: this.#appearance.dark ? HEADER_FILL_DARK : HEADER_FILL,
      });
      this.#formulaBar = bar;
      this.#nameBox = name;
      this.#formulaText = text;
      this.#aggregateButton = aggregate;
      this.#aggregateSummary = summary;
      this.#aggregateChevron = chevron;
      this.#aggregateMenu = menu;
      this.#target.prepend(bar);
    } else if (!this.#appearance.formulaBar && this.#formulaBar !== null) {
      this.#formulaBar.remove();
      this.#formulaBar = null;
      this.#nameBox = null;
      this.#formulaText = null;
      this.#aggregateButton = null;
      this.#aggregateSummary = null;
      this.#aggregateChevron = null;
      this.#aggregateMenu = null;
      this.#aggregateMenuOpen = false;
    }
    this.#refreshChrome();
  }

  /** The cursor cell's label: its address, or the extended range when one stands. */
  #rangeLabel(): string {
    if (this.#selection === null) return "";
    if (this.#range === null) return this.#selection.address;
    const { top, bottom, left, right } = this.#rangeRect();
    const first = cellAddress(top, left);
    const last = cellAddress(bottom, right);
    return first === last ? first : `${first}:${last}`;
  }

  /** The normalized rectangle of the cursor or the extended range. */
  #rangeRect(): { top: number; bottom: number; left: number; right: number } {
    const start = this.#range?.start ?? this.#selection;
    const end = this.#range?.end ?? this.#selection;
    // `#refreshChrome` returns early without a cursor.
    const first = start ?? { row: 0, col: 0 };
    const last = end ?? first;
    return {
      top: Math.min(first.row, last.row),
      bottom: Math.max(first.row, last.row),
      left: Math.min(first.col, last.col),
      right: Math.max(first.col, last.col),
    };
  }

  #refreshChrome(): void {
    if (this.#nameBox !== null) this.#nameBox.textContent = this.#rangeLabel();
    if (this.#formulaText !== null) {
      const cursor = this.#selection;
      if (cursor === null) this.#formulaText.textContent = "";
      else {
        const cell = this.#cellModel(cursor.row, cursor.col);
        this.#formulaText.textContent = cell.formula ?? cell.formatted;
      }
    }
    this.#refreshAggregates();
  }

  #aggregateValues(): { sum: string; average: string; min: string; max: string; count: string } | null {
    if (this.#range === null) return null;
    const { top, bottom, left, right } = this.#rangeRect();
    const numbers: number[] = [];
    let filled = 0;
    let allPercent = true;
    for (let row = top; row <= bottom; row++) {
      for (let col = left; col <= right; col++) {
        const cell = this.#cellModel(row, col);
        if (cell.value !== null && cell.value !== "") filled++;
        if (typeof cell.value === "number" && Number.isFinite(cell.value)) {
          numbers.push(cell.value);
          if (!cell.style?.numberFormat?.includes("%")) allPercent = false;
        }
      }
    }
    const format = (value: number): string => {
      if (numbers.length === 0) return "-";
      if (allPercent) return `${(value * 100).toFixed(1)}%`;
      return new Intl.NumberFormat("en-US", { maximumFractionDigits: 2 }).format(value);
    };
    const sum = numbers.reduce((total, value) => total + value, 0);
    const min = numbers.reduce((smallest, value) => Math.min(smallest, value), Number.POSITIVE_INFINITY);
    const max = numbers.reduce((largest, value) => Math.max(largest, value), Number.NEGATIVE_INFINITY);
    return {
      sum: format(sum),
      average: format(numbers.length === 0 ? 0 : sum / numbers.length),
      min: format(numbers.length === 0 ? 0 : min),
      max: format(numbers.length === 0 ? 0 : max),
      count: String(filled),
    };
  }

  #refreshAggregates(): void {
    const button = this.#aggregateButton;
    const summary = this.#aggregateSummary;
    const menu = this.#aggregateMenu;
    if (button === null || summary === null || menu === null) return;
    const values = this.#aggregateValues();
    if (values === null) {
      button.style.display = "none";
      summary.textContent = "";
      this.#closeAggregateMenu();
      return;
    }
    const labels = { sum: "Sum", average: "Average", min: "Min", max: "Max", count: "Count" };
    button.style.display = "flex";
    summary.textContent = `${labels[this.#aggregateKey]}: ${values[this.#aggregateKey]}`;
    button.setAttribute("aria-label", `Selection ${labels[this.#aggregateKey]} ${values[this.#aggregateKey]}`);
    menu.replaceChildren();
    for (const key of ["sum", "average", "min", "max", "count"] as const) {
      const item = document.createElement("button");
      item.type = "button";
      item.dataset.aggregate = key;
      item.setAttribute("role", "menuitemradio");
      item.setAttribute("aria-checked", String(key === this.#aggregateKey));
      item.style.display = "grid";
      item.style.gridTemplateColumns = "14px minmax(0, 1fr) auto";
      item.style.alignItems = "center";
      item.style.gap = "10px";
      item.style.width = "100%";
      item.style.boxSizing = "border-box";
      item.style.minWidth = "0";
      item.style.padding = "7px 10px";
      item.style.border = "0";
      item.style.borderRadius = "5px";
      item.style.background = "transparent";
      item.style.color = key === this.#aggregateKey ? "var(--aared-ink)" : "var(--aared-muted-text)";
      item.style.font = "13px system-ui, sans-serif";
      item.style.textAlign = "start";
      item.style.cursor = "pointer";
      const check = document.createElement("span");
      check.textContent = key === this.#aggregateKey ? "x" : "";
      check.style.color = "var(--aared-accent)";
      const label = document.createElement("span");
      label.textContent = labels[key];
      const value = document.createElement("span");
      value.textContent = values[key];
      value.style.font = "12px ui-monospace, monospace";
      value.style.color = "var(--aared-ink)";
      value.style.whiteSpace = "nowrap";
      value.dir = "ltr";
      item.append(check, label, value);
      menu.append(item);
    }
  }

  #toggleAggregateMenu(): void {
    if (this.#aggregateMenu === null) return;
    this.#aggregateMenuOpen = !this.#aggregateMenuOpen;
    this.#aggregateMenu.hidden = !this.#aggregateMenuOpen;
    this.#aggregateButton?.setAttribute("aria-expanded", String(this.#aggregateMenuOpen));
  }

  #closeAggregateMenu(): void {
    this.#aggregateMenuOpen = false;
    if (this.#aggregateMenu !== null) this.#aggregateMenu.hidden = true;
    this.#aggregateButton?.setAttribute("aria-expanded", "false");
  }

  /** Repaint the extended-range wash over the window the grid paints. */
  #paintRange(): void {
    const session = this.#session;
    if (session === null) return;
    for (const element of session.host.querySelectorAll(`.${inRangeClass}, .${activeInRangeClass}`)) {
      const cell = element as HTMLElement;
      cell.classList.remove(inRangeClass, activeInRangeClass);
      cell.removeAttribute("data-aared-in-range");
    }
    if (this.#range !== null && this.#appearance.selection) {
      const { top, bottom, left, right } = this.#rangeRect();
      for (let row = top; row <= bottom; row++) {
        for (let col = left; col <= right; col++) {
          const element = session.cellElement(row, col);
          if (element === null) continue;
          const active = this.#selection !== null && row === this.#selection.row && col === this.#selection.col;
          // The cursor cell turns white only when it has no fill of its own (see `activeInRangeClass`).
          if (!active) element.classList.add(inRangeClass);
          else if (element.style.backgroundColor === "") element.classList.add(activeInRangeClass);
          element.setAttribute("data-aared-in-range", "true");
        }
      }
    }
    // Cells past the data have no cell element to carry a class: the session paints them as boxes.
    const selected = this.#appearance.selection && this.#selection !== null;
    session.paintSelection(selected && this.#range !== null ? this.#rangeRect() : null, selected ? this.#selection : null);
    this.#paintFormulaReferences();
  }

  /** Highlight visible same-sheet precedents of the active formula, not text or foreign-sheet references. */
  #paintFormulaReferences(): void {
    const session = this.#session;
    if (session === null) return;
    for (const element of session.host.querySelectorAll('[data-aared-formula-reference="true"]')) {
      element.classList.remove(...referenceClasses);
      element.removeAttribute("data-aared-formula-reference");
      element.removeAttribute("data-aared-reference-index");
    }
    // A range is its own highlight: its cells would otherwise carry the reference outline on top of the range shading.
    if (this.#selection === null || this.#range !== null) return;
    const formula = this.#cellModel(this.#selection.row, this.#selection.col).formula;
    if (formula === undefined) return;
    // Quoted string literals are not references (for example, "A1").
    const source = formula.replace(/"(?:[^"]|"")*"/g, (literal) => " ".repeat(literal.length));
    const pattern = /(^|[^A-Za-z0-9_.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?(?![A-Za-z0-9_]|\s*\()/gi;
    const ranges: Array<{ top: number; bottom: number; left: number; right: number }> = [];
    // The same range written twice is one range, with one colour.
    const seen = new Set<string>();
    for (const match of source.matchAll(pattern)) {
      const sheetName = match[2]?.replace(/''/g, "'") ?? match[3];
      if (match[1] === "!" || (sheetName !== undefined && sheetName.toLocaleLowerCase() !== this.#sheet.title.toLocaleLowerCase())) continue;
      const start = match[4]?.replace(/\$/g, "");
      const end = (match[5] ?? match[4])?.replace(/\$/g, "");
      if (start === undefined || end === undefined) continue;
      try {
        const reference = parseA1Range(`${start}:${end}`);
        const range = {
          top: Math.min(reference.start.row, reference.end.row),
          bottom: Math.max(reference.start.row, reference.end.row),
          left: Math.min(reference.start.col, reference.end.col),
          right: Math.max(reference.start.col, reference.end.col),
        };
        const key = `${range.top}:${range.bottom}:${range.left}:${range.right}`;
        if (seen.has(key)) continue;
        seen.add(key);
        ranges.push(range);
      } catch {
        // A token that resembles a reference but is outside Aared's address space is ignored.
      }
    }
    if (ranges.length === 0) return;
    for (const element of session.host.querySelectorAll<HTMLElement>('td[data-aared-element="cell"]')) {
      const row = Number(element.getAttribute("data-aared-row"));
      const col = Number(element.getAttribute("data-aared-col"));
      const index = ranges.findIndex((range) => row >= range.top && row <= range.bottom && col >= range.left && col <= range.right);
      if (index === -1) continue;
      element.classList.add(referenceClasses[index % referenceClasses.length]!);
      element.setAttribute("data-aared-formula-reference", "true");
      element.setAttribute("data-aared-reference-index", String(index));
    }
    session.paintReferenceRanges(ranges.map((range, index) => ({
      range,
      color: REFERENCE_COLORS[index % REFERENCE_COLORS.length]!.solid,
    })));
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
      tab.style.font = "13px system-ui, sans-serif";
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
        // The tab owns this action: unlike a public call, its failure has no
        // caller to reject to. Superseded switches and destroyed views stay quiet.
        void this.goToSheet(sheetIndex).catch((error: unknown) => {
          if (!this.destroyed) reportBackgroundError(this, "aared/xlsx", error);
        });
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

  #select(selection: CellRef | null, emit: boolean, range: { start: CellRef; end: CellRef } | null = null): void {
    this.#selection = selection === null ? null : { ...selection };
    this.#range = range === null ? null : { start: { ...range.start }, end: { ...range.end } };
    this.#anchor = range === null
      ? (selection === null ? null : { ...selection })
      : { ...range.start };
    this.#restoreSelection();
    this.#paintRange();
    this.#refreshChrome();
    if (emit && this.#appearance.selection) {
      this.dispatchEvent(new CustomEvent("selectionchange", {
        detail: { sheet: this.#sheetRef(), selection: this.selection, range: this.range },
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
      // A cell past the data has no element; its box is still painted.
      this.#paintRange();
      return;
    }
    element.setAttribute("data-aared-selected", "true");
    element.setAttribute("aria-selected", "true");
    element.style.outline = "2px solid var(--aared-selection)";
    element.style.outlineOffset = "-2px";
    element.tabIndex = 0;
    if (lostGridFocus) element.focus({ preventScroll: true });
    // The range wash rides the same funnel: repaints and scroll windows
    // rebuild row DOM, so both re-apply here or neither survives scrolling.
    this.#paintRange();
  }

  #cellModel(row: number, col: number): CellModel {
    const found = this.#model.rows.find((entry) => entry.index === row)?.cells.find((entry) => entry.col === col)?.cell;
    return found ?? { value: null, formatted: "", type: "empty" };
  }

  #cellFromEvent(event: MouseEvent): CellRef | null {
    const target = event.target as HTMLElement | null;
    // Links keep the click path (which reports `linkactivate`); a drag never
    // starts on one.
    if (target?.closest?.('a[data-aared-element="cell-link"]') !== null) return null;
    // The empty area past the data selects by position, numbered past the extent.
    if (target !== null && target.closest?.('td[data-aared-element="fill-cell"]') != null) {
      const at = this.#session?.fillCellAt(target, event.clientX) ?? null;
      return at === null ? null : { address: cellAddress(at.row, at.col), row: at.row, col: at.col };
    }
    const cell = target?.closest?.('td[data-aared-element="cell"]') ?? null;
    if (cell === null || !(cell instanceof HTMLElement)) return null;
    const row = Number(cell.getAttribute("data-aared-row"));
    const col = Number(cell.getAttribute("data-aared-col"));
    if (!Number.isInteger(row) || !Number.isInteger(col)) return null;
    return { address: cellAddress(row, col), row, col };
  }

  #handlePointerDown(event: MouseEvent): void {
    if (this.destroyed || !this.#appearance.selection || event.button !== 0) return;
    const cell = this.#cellFromEvent(event);
    if (cell === null) return;
    this.#dragging = true;
    this.#dragged = false;
    // Native text selection would paint over the range wash while dragging.
    this.#gridBox.style.userSelect = "none";
    this.#select(cell, true);
  }

  #handlePointerMove(event: MouseEvent): void {
    if (!this.#dragging || this.destroyed) return;
    const cell = this.#cellFromEvent(event);
    if (cell === null) return;
    const cursor = this.#selection;
    if (cursor !== null && cell.row === cursor.row && cell.col === cursor.col) return;
    // No auto-scroll past the window edge: cells outside the painted window
    // have no element to extend to, so the range ends at the window.
    this.#dragged = true;
    const anchor = this.#anchor ?? cell;
    this.#select({ ...cell }, true, { start: anchor, end: { ...cell } });
    this.#ensureVisible(cell.row, cell.col, "start", true);
  }

  #handlePointerUp(_event: MouseEvent): void {
    if (!this.#dragging) return;
    this.#dragging = false;
    this.#gridBox.style.userSelect = "";
  }

  #handleClick(event: MouseEvent): void {
    // The release click after a drag-select must not collapse the range it made.
    if (this.#dragged) {
      this.#dragged = false;
      return;
    }
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
    const cell = this.#cellFromEvent(event);
    if (cell === null) return;
    this.#select(cell, true);
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
    // The arrows walk on into the empty area past the data, as far as it is
    // painted; Home, End and the page keys stay on the data, as Ctrl+End does.
    const fill = session.fillSpan();
    const reachRows = [...rows, ...fill.rows];
    const reachColumns = [...columns, ...fill.columns];
    let next: { row: number; col: number } | null = null;
    let align: "start" | "end" = "start";
    switch (event.key) {
      case "ArrowUp":
        next = { row: reachRows[Math.max(0, clampIndex(reachRows, current.row) - 1)]!, col: current.col };
        align = "start";
        break;
      case "ArrowDown":
        next = { row: reachRows[Math.min(reachRows.length - 1, clampIndex(reachRows, current.row) + 1)]!, col: current.col };
        align = "end";
        break;
      case "ArrowLeft":
        // Visual movement: in RTL sheets the inline-start edge is physical
        // right, so Left advances the column and Right retreats it.
        next = { row: current.row, col: reachColumns[Math.max(0, Math.min(reachColumns.length - 1, clampIndex(reachColumns, current.col) + (rtl ? 1 : -1)))]! };
        align = rtl ? "end" : "start";
        break;
      case "ArrowRight":
        next = { row: current.row, col: reachColumns[Math.max(0, Math.min(reachColumns.length - 1, clampIndex(reachColumns, current.col) + (rtl ? -1 : 1)))]! };
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
    const column = reachColumns.includes(next.col) ? next.col : reachColumns[clampIndex(reachColumns, next.col)]!;
    const selection = { address: cellAddress(next.row, column), row: next.row, col: column };
    // `Shift` extends the range from its anchor instead of moving the cursor
    // alone; any plain move re-anchors. The anchor is the cursor as it stood
    // when `Shift` went down, kept across repeated extends.
    const extendKeys = new Set(["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight", "Home", "End", "PageUp", "PageDown"]);
    if (event.shiftKey && extendKeys.has(event.key)) {
      const anchor = this.#anchor ?? { ...current };
      this.#select(selection, true, { start: anchor, end: { ...selection } });
    } else {
      this.#select(selection, true);
    }
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

    const top = session.dataRowTop(row);
    if (top !== undefined) {
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
 * Paint one sheet. Accepts a `SheetHandle` or borrowed `SheetContent` (read from the Worker on demand)
 * or an already-resolved model; when the input is a handle, drawing media
 * resolves through its session unless an explicit `mediaResolver` is given.
 * Painting itself is synchronous — only the handle read awaits.
 */
export async function render(input: XlsxRenderInput): Promise<XlsxRenderHandle> {
  throwIfAborted(input.signal);
  const requested = contentSource<SheetHandle>(input.sheet) ?? (input.sheet as SheetHandle | SheetModel | SheetRangeModel);
  if (!(requested instanceof SheetHandle) && "read" in requested) {
    throw new AaredError("invalid-argument", "aared/xlsx: render() takes a sheet from open() or viewer.content.sheets, or a resolved sheet model");
  }
  const signal = contentSignal(input.sheet, input.signal);
  let mediaResolver = input.mediaResolver;
  let model: SheetModel | SheetRangeModel;
  let handle: SheetHandle | undefined;
  if (isSheetHandle(requested)) {
    handle = requested;
    model = await requested.read(signal === undefined ? undefined : { signal });
    mediaResolver ??= (mediaId: string) => requested.media(mediaId);
  } else {
    model = requested;
  }
  // A gradient or pattern cell paints through the shared fill painter, which
  // loads only for a sheet that has one; a painter that fails to load leaves
  // those cells marked, never blank.
  if (needsFillPainter(model)) await loadFillPainter().catch(() => undefined);
  // The last moment a cancel can still leave the target untouched: the viewer paints as it is built.
  throwIfAborted(signal);
  const viewer = new XlsxViewer(input.target, model, handle, mediaResolver, input.appearance, input.zoom ?? DEFAULT_ZOOM.xlsx);
  // As in PPTX and DOCX: which fonts painted can only be read once they have
  // arrived (a delivered face loads only when text needs it), so the warnings wait, bounded.
  // A first chart waits for the layout code too (Notes §71), and a
  // first picture effect for its painter: a resolved render has both drawn.
  // Cancelled during that wait, the render takes the sheet back out (its
  // destroy gives the target its own layout back) and rejects (`settleRender`).
  const settled = Promise.all([awaitDocumentFonts(model), chartsSettled(), viewer.effectsReady]).then(() => undefined);
  await settleRender(viewer, input.target, [...input.target.childNodes], settled, signal);
  if (!viewer.destroyed) viewer.setWarnings(sheetWarnings(model));
  return viewer;
}

/**
 * What a render noticed about the sheet: drawings the file gave no id (the file is damaged, but
 * aared never uses a drawing's id, so the sheet is drawn in full), then fonts that painted with a fallback.
 */
function sheetWarnings(model: { name: string; missingIds?: number }): RenderWarning[] {
  const noId = missingIdWarning(
    "xlsx-missing-id", model.missingIds ?? 0, "drawing", `on sheet "${model.name}"`,
    "Excel offers to repair such a file; aared does not use the id, and the sheet is rendered in full.",
  );
  return [...(noId === undefined ? [] : [noId]), ...fontFidelityWarnings(model)];
}
/** Load the embed controller only when the convenience preview API is used. */
export async function preview(input: import("./viewer.js").XlsxPreviewInput): Promise<import("../preview/index.js").XlsxPreview> {
  return (await import("./viewer.js")).preview(input);
}
export type { XlsxPreviewInput } from "./viewer.js";
export type { XlsxPreview, PreviewSearchMatch, PreviewSearchOptions, PreviewEventMap, SearchNavigation, SearchNavigationState } from "../preview/index.js";
