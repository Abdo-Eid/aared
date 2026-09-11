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
  LinearGradientFillModel,
  LinearGradientStopModel,
  LocalViewBoxModel,
  PictureModel,
  ShapeGeometryModel,
  ShapeFillModel,
  ShapeModel,
  SizeModel,
  SolidFillModel,
  SlideElementModel,
  SlideModel,
  StrokeModel,
  TableCellModel,
  TableModel,
  TableRowModel,
  TextBodyModel,
  TextBodyInsetsModel,
  TextBodyVerticalAlign,
  TextDirection,
  TextAutofitMode,
  TextParagraphModel,
  TextRunModel,
  TextWrappingMode,
  TransformModel,
  UnsupportedShapeFillModel,
} from "./model.js";

export interface DocumentInfo {
  format: DocumentFormat;
  /** slides / sheets / sections */
  count: number;
  /** slide titles, sheet names, or heading outline */
  titles: string[];
  sizes?: { width: number; height: number }[];
  /**
   * XLSX-only: per-sheet visibility in `titles` order. Other formats omit
   * it; sheet tabs and navigation use it to exclude hidden sheets.
   */
  sheetStates?: Array<"visible" | "hidden" | "veryHidden">;
  theme: {
    colors: Record<string, string>;
    fonts: { major: string; minor: string };
  };
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
  DocxTextAlign,
  DocxParagraphLayout,
  DocxParagraphBorder,
  DocxParagraphBorders,
  DocumentSection,
  FlowNode,
  SectionBreakType,
  SectionPageGeometry,
  SectionPageMargins,
  SectionProperties,
} from "./model.js";

export type { AaredErrorCode } from "./errors.js";
export { AaredError, isAaredError } from "./errors.js";

/** Anything `open()` accepts: a file handle, in-memory bytes, or a URL to fetch. */
export type Source = File | Blob | ArrayBuffer | string;

export type FontResolver = (
  requestedFamily: string,
) => Promise<FontFace> | undefined;

/** Main-thread callback used by renderers to retrieve model-referenced bytes. */
export type MediaResolver = (mediaId: string) => Promise<ArrayBuffer>;

export interface OpenOptions {
  fontResolver?: FontResolver;
  signal?: AbortSignal;
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

/** Viewer chrome theme. Controls `aared`-owned UI only — never authored content by itself. */
export type ViewerTheme = "light" | "dark" | "system";

/**
 * Whether document content may adapt for dark reading. `source` always
 * preserves the file as authored; `adaptive` lets format-default surfaces
 * (never explicit authored styles) follow the viewer theme. PPTX supports
 * `source` only — a slide is a designed canvas — while XLSX and DOCX
 * support both.
 */
export type ContentMode = "source" | "adaptive";

/** Small appearance knob; visual detail lives in CSS variables, not here. */
export interface Appearance {
  theme?: ViewerTheme;
  contentMode?: ContentMode;
  spreadsheet?: {
    headers?: boolean;
    sheetTabs?: boolean;
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
