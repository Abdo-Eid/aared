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

/**
 * The name a `FontResolver` must register a delivered substitute's
 * `FontFace` under, for the requested family it was asked to resolve. See
 * `cssFontFamily` in `core/typography.ts` (not re-exported: renderer-only)
 * for why — this is the one exported hook a resolver author needs.
 */
export { fontAlias } from "./typography.js";

export type { ChartSpecModel } from "./chart/spec.js";

export type {
  BoundsModel,
  ChartModel,
  GroupModel,
  DrawingEffectsModel,
  LinearGradientFillModel,
  LinearGradientStopModel,
  LocalViewBoxModel,
  PatternFillModel,
  PictureFillModel,
  PictureModel,
  PictureTileModel,
  PathGradientFillModel,
  RelativeRectModel,
  RenderWarning,
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
  DocxChartNode,
  DocxDrawingNode,
  DocxTextBoxNode,
  DocxImageAnchor,
  DocxAnchorPosition,
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
  HeaderFooterParts,
  SectionProperties,
} from "./model.js";

export type { AaredErrorCode } from "./errors.js";
export { AaredError, isAaredError, isAbort, messageFor } from "./errors.js";

/**
 * Anything `open()` accepts: a file handle, in-memory bytes (an `ArrayBuffer`,
 * or a view of one such as a `Uint8Array` or a Node `Buffer`, read over its own
 * range), or a URL to fetch. Bytes are copied once before they go to the Worker,
 * so the caller's stay usable.
 */
export type Source = File | Blob | ArrayBuffer | ArrayBufferView | string;

/**
 * Main-thread hook for a family the file names: register a substitute under
 * `fontAlias(requestedFamily)` and return it, or return `undefined` (directly,
 * or as the promise's value, so an `async` resolver works) to leave the family
 * to the browser's own fonts.
 */
export type FontResolver = (
  requestedFamily: string,
) => Promise<FontFace | undefined> | undefined;

/** Main-thread callback used by renderers to retrieve model-referenced bytes. */
export type MediaResolver = (mediaId: string) => Promise<ArrayBuffer>;

/**
 * Explicit Worker input. Leave out for the bundler-managed default.
 * `url` loads a same-origin module; `crossOrigin: "blob"` opts into a Blob
 * bootstrap importing that URL (requires CORS and CSP permitting it).
 * `source` is a self-contained module's text for single-file builds.
 * Aared owns generated Blob URLs. `create` returns a fresh native Worker
 * for advanced integrations; it is terminated on close and never pooled.
 * Exactly one of `url`, `source`, or `create` is accepted.
 */
export type WorkerInit =
  | { url: string | URL; crossOrigin?: "blob"; source?: never; create?: never }
  | { source: string; url?: never; crossOrigin?: never; create?: never }
  | { create: () => Worker; url?: never; source?: never; crossOrigin?: never };

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
    /**
     * Excel-like viewer chrome: a read-only formula bar (name box plus the
     * active cell's formula or formatted value) above the grid. Default
     * `false`: the grid paints alone, as before.
     */
    formulaBar?: boolean;
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
