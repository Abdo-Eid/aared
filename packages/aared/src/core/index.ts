/**
 * Shared core: the types, errors, and appearance contracts every format
 * entrypoint (`aared/pptx`, `aared/xlsx`, `aared/docx`) is built from.
 *
 * `aared/core` opens nothing itself — each format exposes its own typed
 * `open()` (see the per-format modules). What lives here is the shared
 * API grammar: sources, options, collections, appearance, media, and the
 * single typed error model. Worker machinery is internal (`./session.js`,
 * `./worker.ts`) and never re-exported.
 */

export type DocumentFormat = "pptx" | "xlsx" | "docx";

export type {
  BoundsModel,
  ChartGraphicModel,
  ChartGraphicPrimitive,
  ChartModel,
  GroupModel,
  LinearGradientFillModel,
  LinearGradientStopModel,
  LocalViewBoxModel,
  PatternFillModel,
  PictureFillModel,
  PictureModel,
  PictureTileModel,
  PathGradientFillModel,
  RelativeRectModel,
  ShapeGeometryModel,
  ShapeFillModel,
  ShapeModel,
  SizeModel,
  SolidFillModel,
  SlideElementModel,
  SlideLayoutModel,
  SlideModel,
  StrokeModel,
  TableCellModel,
  TableModel,
  TableRowModel,
  TextBodyModel,
  InsetsModel,
  TextBodyVerticalAlign,
  TextAlign,
  TextDirection,
  TextAutofitMode,
  TextParagraphModel,
  TextRunModel,
  TextWrappingMode,
  TransformModel,
  UnsupportedShapeFillModel,
} from "./model.js";

/**
 * What `open()` knows before anything renders. Every per-unit array is
 * exactly `count` long, so `titles[i]` and `sizes[i]` always describe unit
 * `i` — the unit `goTo(i)` shows.
 */
export interface DocumentInfo {
  format: DocumentFormat;
  /**
   * Navigable units — what one `render()` draws and what `goTo(i)` switches
   * between: slides (PPTX), sheets (XLSX), or the whole document (DOCX,
   * always `1`, in flow and layout mode alike — pages are a render result).
   */
  count: number;
  /**
   * One label per unit. PPTX: `"Slide N"` (`open()` reads no slide part, so
   * real titles come from `slide.read()`); XLSX: sheet names; DOCX: the
   * document title from its properties, else `"Document"`. The DOCX heading
   * outline lives on `document.headings`.
   */
  titles: string[];
  /**
   * Per-unit size in inches, only where a unit has one fixed size: one entry
   * per slide (PPTX). XLSX omits it (a sheet has no page), and so does DOCX (a
   * flow has no size; per-section page setup is on `document.sections`).
   */
  sizes?: { width: number; height: number }[];
  /**
   * XLSX-only: per-sheet visibility in `titles` order. Other formats omit
   * it; sheet tabs and navigation use it to exclude hidden sheets.
   */
  sheetStates?: Array<"visible" | "hidden" | "veryHidden">;
  /** The file's own Office theme (colours and major/minor fonts) — not the viewer theme. */
  theme: {
    colors: Record<string, string>;
    fonts: { major: string; minor: string };
  };
}

/** A pointer at one unit: its `goTo()` index plus its label. The same shape everywhere a unit is named. */
export interface UnitRef {
  index: number;
  title: string;
}

/** A cell location: the A1 `address` plus its zero-based `row`/`col`. The same shape everywhere a cell is named. */
export interface CellRef {
  address: string;
  row: number;
  col: number;
}

/** Canonical spreadsheet/document models live in `./model.js`; this module re-exports them. */
export type {
  CellRange,
  CellModel,
  RowModel,
  SheetModel,
  SheetRangeModel,
  XlsxCellType,
  XlsxCellStyleModel,
  SheetDrawingModel,
  SheetPictureModel,
  SheetChartModel,
  SheetUnsupportedModel,
  DocxPageBreakNode,
  DocxParagraphNode,
  DocxHeadingNode,
  DocxTableNode,
  DocxTableCellNode,
  DocxTableRowNode,
  DocxImageNode,
  DocxListNode,
  DocxListItemNode,
  DocxFootnoteNode,
  DocxUnsupportedNode,
  DocxParagraphLayout,
  DocxParagraphBorder,
  DocxParagraphBorders,
  DocxTabStop,
  DocxInlineImage,
  DocxTableEdge,
  DocxTableBorders,
  DocumentSection,
  FlowNode,
  SectionBreakType,
  SectionPageGeometry,
  SectionPageMargins,
  SectionProperties,
} from "./model.js";

export type { AaredErrorCode } from "./errors.js";
export { AaredError, isAaredError, isAbort, messageFor } from "./errors.js";

/** Anything `open()` accepts: a file handle, in-memory bytes, or a URL to fetch. */
export type Source = File | Blob | ArrayBuffer | string;

export type FontResolver = (
  requestedFamily: string,
) => Promise<FontFace> | undefined;

/** Main-thread callback used by renderers to retrieve model-referenced bytes. */
export type MediaResolver = (mediaId: string) => Promise<ArrayBuffer>;

/**
 * Override for Worker construction. Use when the default
 * `new Worker(new URL("./worker.mjs", import.meta.url))` does not work for
 * your environment (single-file bundles, strict CSP, SSR, or bundler-specific
 * Worker constructors).
 *
 * - **`string` / `URL`**: treated as the worker script URL (`blob:` URLs
 *   from inlined code work).
 * - **`() => Worker`**: factory that returns a live `Worker` instance
 *   (for Vite `?worker`, webpack `worker-loader`, etc.).
 */
export type WorkerInit = string | URL | (() => Worker);

export interface OpenOptions {
  fontResolver?: FontResolver;
  signal?: AbortSignal;
  /** Override worker construction (see `WorkerInit` for details). */
  worker?: WorkerInit;
}

/** Options for lazy content reads (`SlideHandle.read`, `SheetHandle.read`). */
export interface ReadOptions {
  signal?: AbortSignal;
}

export interface MediaOptions {
  signal?: AbortSignal;
}

/**
 * The shared collection grammar behind `deck.slides`, `workbook.sheets`
 * and `document.sections`. Public collections expose domain-specific
 * names; the shape — synchronous `list()`/`get()` over data resolved at
 * `open()` time — is identical everywhere. `get()` throws a `not-found`
 * `AaredError` for unknown keys.
 */
export interface Collection<TItem, TKey> {
  list(): readonly TItem[];
  get(key: TKey): TItem;
}

/** Collection-style media access shared by every format document. */
export interface MediaCollection {
  get(mediaId: string, options?: MediaOptions): Promise<ArrayBuffer>;
}

/** Viewer chrome theme. Controls `aared`-owned UI only — never authored content. */
export type ViewerTheme = "light" | "dark" | "system";

/** Small appearance knob; visual detail lives in CSS variables, not here. */
export interface Appearance {
  theme?: ViewerTheme;
  spreadsheet?: {
    headers?: boolean;
    sheetTabs?: boolean;
    /**
     * Continue the grid past the data with empty cells (and headings) until
     * the viewport is full, like Excel. Viewer chrome only: the cells are
     * inert and never part of the sheet's extent. Default `true`.
     */
    fill?: boolean;
    selection?: boolean;
    gridlines?: GridlineMode;
  };
}

/** How the XLSX grid draws its lines: from the file, forced on, or hidden. */
export type GridlineMode = "source" | "show" | "hide";

/** Zoom shared by every render controller: fit modes or an explicit scale. */
export type ZoomValue = "fit-width" | "fit-page" | number;

/** Fetch/parse progress for large files (element `progress` events). */
export interface LoadProgress {
  loaded: number;
  total?: number;
}
