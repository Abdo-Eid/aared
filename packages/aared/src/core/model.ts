/**
 * Structured-cloneable resolved models shared by parsers, Workers, and
 * renderers. All geometry is expressed in inches unless documented otherwise.
 */

export interface SizeModel {
  width: number;
  height: number;
}

export interface BoundsModel {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** Unitless coordinates local to an SVG graphic. */
export interface LocalViewBoxModel {
  width: number;
  height: number;
}

export type TextDirection = "ltr" | "rtl";

export type TextBodyVerticalAlign = "top" | "middle" | "bottom" | "justify" | "distributed";

export type TextWrappingMode = "none" | "square";

export type TextAutofitMode = "none" | "shrinkText" | "resizeShape";

export interface TextBodyInsetsModel {
  /** All inset values are in inches. */
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
}

export interface TextRunModel {
  text: string;
  bold?: boolean;
  italic?: boolean;
  /** Points, matching DrawingML text sizing. */
  fontSizePt?: number;
  language?: string;
  latinTypeface?: string;
  complexScriptTypeface?: string;
  /** Resolved CSS color, not a scheme token. */
  color?: string;
  /** Resolved Office/theme font family; font files resolve separately. */
  fontFamily?: string;
  /** Word `w:u` single (or heavier) underline; absent means none. */
  underline?: boolean;
  /** Word `w:strike` / `w:dstrike`. */
  strikethrough?: boolean;
  /** Word `w:vertAlign` superscript/subscript (footnote markers, endnotes). */
  verticalAlign?: "subscript" | "superscript";
  /** Resolved CSS highlight color behind the run (`w:highlight`). */
  highlight?: string;
  /** Resolved CSS run shading fill (`w:shd`). */
  shading?: string;
  /** Hyperlink target for runs inside `w:hyperlink`; externals resolved, internals kept as `#anchor`. */
  hyperlink?: string;
  /** Footnote/endnote entry id when this run is its in-text marker. */
  footnoteRef?: string;
  /**
   * Live page field this run evaluates (`w:instrText PAGE`/`NUMPAGES` or
   * `w:fldSimple`). `text` carries the static flow fiction (`"1"`); layout
   * pagination substitutes the real value per page, flow rendering keeps it.
   */
  field?: "PAGE" | "NUMPAGES";
}

export interface TextParagraphModel {
  runs: TextRunModel[];
  /** Effective default run size for paragraph-level layout, in points. */
  fontSizePt?: number;
  /** Absent when the source did not explicitly set paragraph direction. */
  direction?: TextDirection;
  align?: "left" | "right" | "center" | "justify";
  bullet?: { char?: string; typeface?: string };
  level?: number;
  /** Paragraph margins and first-line indent in inches. */
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndent?: number;
  /** Line height as a multiple of the paragraph font size. */
  lineSpacing?: number;
  /** Explicit line height in points when the source does not use a percentage. */
  lineSpacingPt?: number;
  /** Space before and after the paragraph, in points. */
  spaceBeforePt?: number;
  spaceAfterPt?: number;
}

export interface TextBodyModel {
  paragraphs: TextParagraphModel[];
  insets?: TextBodyInsetsModel;
  verticalAlign?: TextBodyVerticalAlign;
  wrapping?: TextWrappingMode;
  autofit?: TextAutofitMode;
}

export interface TransformModel {
  /** Clockwise degrees. */
  rotation?: number;
  flipHorizontal?: boolean;
  flipVertical?: boolean;
}

export interface ShapeGeometryModel {
  /** DrawingML preset name when the source uses `a:prstGeom`. */
  preset?: string;
  /** SVG path data in the local coordinate system defined by `viewBox`. */
  path: string;
  viewBox: LocalViewBoxModel;
  /** Source geometry could not be rendered faithfully; `path` is a visible bounds fallback. */
  fallback?: "unsupported" | "custom" | "missing";
}

export interface StrokeModel {
  color: string;
  /** Stroke width in inches. */
  width: number;
}

export interface SolidFillModel {
  kind: "solid";
  /** Resolved CSS color. */
  color: string;
}

export interface LinearGradientStopModel {
  /** Stop position from 0 through 1, in authored DrawingML order. */
  position: number;
  /** Resolved CSS color. */
  color: string;
}

export interface LinearGradientFillModel {
  kind: "linear-gradient";
  /** Clockwise degrees, matching DrawingML gradient angles. */
  angle: number;
  /** Whether DrawingML scales the gradient with the shape bounds. */
  scaled: boolean;
  /** Whether the gradient follows the shape's transform. */
  rotatesWithShape: boolean;
  stops: LinearGradientStopModel[];
}

/** A deliberately unpainted source fill that cannot be represented faithfully. */
export interface UnsupportedShapeFillModel {
  kind: "unsupported";
  source: "gradient-path" | "gradient-linear" | "image" | "pattern";
}

/** Resolved DrawingML shape paint, never raw OOXML or SVG markup. */
export type ShapeFillModel = SolidFillModel | LinearGradientFillModel | UnsupportedShapeFillModel;

export interface ShapeModel {
  kind: "shape";
  id: string;
  name: string;
  placeholder?: { type: string; index: string };
  bounds?: BoundsModel;
  text?: TextBodyModel;
  /** Present when the parser resolves DrawingML geometry to an SVG path. */
  geometry?: ShapeGeometryModel;
  /** Resolved shape paint when the parser supports the source fill. */
  fill?: ShapeFillModel;
  stroke?: StrokeModel;
  transform?: TransformModel;
}

export interface TableCellModel {
  text: TextBodyModel;
  verticalAlign?: "top" | "middle" | "bottom";
  /** Logical-grid spans; omitted values each mean one cell. */
  rowSpan?: number;
  columnSpan?: number;
  /** Resolved cell insets in inches, from `a:tcPr/@marL|marR|marT|marB`. */
  margins?: { left?: number; right?: number; top?: number; bottom?: number };
  /** Resolved cell edge strokes, from `a:tcPr/a:lnL|lnR|lnT|lnB`. Widths are inches. */
  borders?: {
    left?: { color: string; width: number };
    right?: { color: string; width: number };
    top?: { color: string; width: number };
    bottom?: { color: string; width: number };
  };
}

export interface TableRowModel {
  /** Row height in inches. */
  height?: number;
  /** Logical XML cell order, regardless of table display direction. */
  cells: TableCellModel[];
}

export interface TableModel {
  kind: "table";
  id: string;
  name: string;
  bounds?: BoundsModel;
  /** Table display direction; independent from every paragraph direction. */
  direction: TextDirection;
  /** Column widths in inches, in logical XML order. */
  columns: number[];
  rows: TableRowModel[];
  /** Resolved table-style backgrounds per logical XML cell, if present. */
  cellFills: Array<Array<string | undefined>>;
  transform?: TransformModel;
}

export interface PictureModel {
  kind: "picture";
  id: string;
  name: string;
  bounds?: BoundsModel;
  /** Identifier accepted by `MediaCollection.get()`. */
  mediaId: string;
  /** Resolved MIME content type used to create the media Blob. */
  contentType: string;
  /** Fractions (0..1) cropped from each edge via DrawingML `a:srcRect`. */
  crop?: { left: number; top: number; right: number; bottom: number };
  transform?: TransformModel;
}

interface ChartGraphicStyleModel {
  fill?: string;
  stroke?: { color: string; width: number };
  opacity?: number;
}

interface ChartGraphicRectModel extends ChartGraphicStyleModel {
  kind: "rect";
  x: number;
  y: number;
  width: number;
  height: number;
}

interface ChartGraphicLineModel extends ChartGraphicStyleModel {
  kind: "line";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

interface ChartGraphicPathModel extends ChartGraphicStyleModel {
  kind: "path";
  /** SVG path commands only, never SVG markup. */
  path: string;
}

interface ChartGraphicTextModel extends ChartGraphicStyleModel {
  kind: "text";
  x: number;
  y: number;
  text: string;
  fontSize?: number;
  fontFamily?: string;
  fontWeight?: "normal" | "bold";
  textAnchor?: "start" | "middle" | "end";
}

export type ChartGraphicPrimitive =
  | ChartGraphicRectModel
  | ChartGraphicLineModel
  | ChartGraphicPathModel
  | ChartGraphicTextModel;

/** Safe chart drawing data; all primitive coordinates are local and unitless. */
export interface ChartGraphicModel {
  viewBox: LocalViewBoxModel;
  primitives: ChartGraphicPrimitive[];
}

export interface ChartModel {
  kind: "chart";
  id: string;
  name: string;
  bounds?: BoundsModel;
  /** Typed, markup-free chart drawing data when chart parsing is supported. */
  graphic?: ChartGraphicModel;
  /** Optional source identifier for diagnostics and source correlation. */
  sourceChartId?: string;
  transform?: TransformModel;
}

export type SlideElementModel = ShapeModel | TableModel | PictureModel | ChartModel;

export interface SlideModel {
  /** OPC part name for the source slide. */
  part: string;
  /** Slide dimensions in inches. */
  size: SizeModel;
  /** Back-to-front z-order. */
  elements: SlideElementModel[];
  notes?: string;
}

/**
 * XLSX sheet-model family. Worker-resolved from SpreadsheetML, styles, and
 * shared strings; structured-cloneable plain data. Geometry (column widths,
 * row heights, drawing bounds) is inches, matching the slide model's
 * convention. Rows and cells are sparse: only authored rows appear, and each
 * row carries only its authored cells in ascending column order, so a large
 * mostly-empty used range stays proportional to its content, not its extent.
 * True range-windowed transfer for huge sheets needs a `UnitOptions` range
 * (owned by the worker orchestrator) and is proposed separately.
 */

/** Zero-based, inclusive cell range in logical (A1) order, never display order. */
export interface CellRange {
  start: { row: number; col: number };
  end: { row: number; col: number };
}

/** How a cell's cached value was stored. `formula-string` is a cached string result, never a re-evaluation. */
export type XlsxCellType = "shared" | "number" | "boolean" | "error" | "inline" | "formula-string" | "empty";

export interface XlsxCellBorderSideModel {
  /** Source border style name (`thin`, `medium`, `dashed`, ...). */
  style?: string;
  /** Resolved CSS color. */
  color?: string;
}

/** Resolved direct cell formatting: the SheetJS Community gap (fills, fonts, borders). */
export interface XlsxCellStyleModel {
  /** Resolved CSS fill color. */
  fill?: string;
  /** Resolved CSS font color. */
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** Points, matching SpreadsheetML font sizing. */
  fontSizePt?: number;
  /** Literal typeface requested by the style; font files resolve separately. */
  fontFamily?: string;
  /** `general` keeps Excel's type-dependent default (text at the start edge, numbers at the end edge). */
  align?: "left" | "right" | "center" | "justify" | "general";
  verticalAlign?: "top" | "middle" | "bottom";
  wrapText?: boolean;
  /** Excel indent units. */
  indent?: number;
  /** SpreadsheetML reading order: 1 = LTR context, 2 = RTL context. */
  readingOrder?: 0 | 1 | 2;
  /** Source number-format code that produced `CellModel.formatted`. */
  numberFormat?: string;
  border?: {
    left?: XlsxCellBorderSideModel;
    right?: XlsxCellBorderSideModel;
    top?: XlsxCellBorderSideModel;
    bottom?: XlsxCellBorderSideModel;
  };
}

export interface CellModel {
  /** Cached value. Formulas are never evaluated: this is the file's stored result. */
  value: string | number | boolean | null;
  /** Display string with the cell's number format already applied. */
  formatted: string;
  type: XlsxCellType;
  /** Cached formula source when the cell carries one; present for diagnostics, never evaluated. */
  formula?: string;
  hyperlink?: { target: string; display?: string };
  style?: XlsxCellStyleModel;
  /**
   * Conditional-formatting overlay resolved once in the Worker from cached
   * values. A view-only snapshot, not a live recalculation.
   */
  conditional?: {
    fill?: string;
    color?: string;
    bold?: boolean;
    dataBar?: { ratio: number; color: string };
  };
}

export interface RowModel {
  /** Zero-based row index. */
  index: number;
  /** Sparse authored cells in ascending column order; absent columns are empty. */
  cells: Array<{ col: number; cell: CellModel }>;
  /** Row height in inches. */
  height?: number;
  hidden?: boolean;
}

export interface SheetPictureModel {
  kind: "picture";
  name: string;
  /** Identifier accepted by `MediaCollection.get()`. */
  mediaId: string;
  /** Resolved MIME content type used to create the media Blob. */
  contentType: string;
  /** Absolute bounds in inches from the grid origin. */
  bounds: BoundsModel;
  /** Anchor cells the bounds were derived from. */
  anchor: CellRange;
}

export interface SheetChartModel {
  kind: "chart";
  name: string;
  /**
   * Typed, markup-free chart drawing from the narrow spreadsheet chart
   * adapter. Absent when the source chart type is not covered; the reason is
   * then in `unsupportedReason` and the bounds still reserve its space.
   */
  graphic?: ChartGraphicModel;
  /** Absolute bounds in inches from the grid origin. */
  bounds: BoundsModel;
  /** Anchor cells the bounds were derived from. */
  anchor: CellRange;
  /** Optional source identifier for diagnostics and source correlation. */
  sourceChartId?: string;
  unsupportedReason?: string;
}

export type SheetDrawingModel = SheetPictureModel | SheetChartModel;

/** Visible, non-silent record of content the previewer deliberately does not model. */
export interface SheetUnsupportedModel {
  kind: "pivot" | "macro" | "sparkline" | "data-validation" | "comment" | "conditional-format" | "drawing" | "other";
  detail: string;
  range?: CellRange;
}

export interface SheetModel {
  /** Sheet (tab) name. */
  name: string;
  /** OPC part name for the source worksheet. */
  part: string;
  /** Sheet-level RTL: column A displays at the right. DOM order stays logical. */
  rightToLeft: boolean;
  /** Authored gridline visibility from `sheetView@showGridLines` (default true). */
  showGridlines?: boolean;
  /** Used-range size as counts (columns × rows), from `dimension` or the parsed cells. */
  extent: { cols: number; rows: number };
  /** Sparse authored rows in ascending order. */
  rows: RowModel[];
  merges: CellRange[];
  /** Frozen leading rows/columns derived from the sheet's split pane. */
  frozen: { rows: number; cols: number };
  /** Column widths in inches per column index, dense to `extent.cols`. */
  columnWidths: number[];
  hiddenColumns: number[];
  /** Hidden row indexes, including empty hidden rows that carry no cells. */
  hiddenRows: number[];
  /** Default row height in inches. */
  defaultRowHeight: number;
  drawings: SheetDrawingModel[];
  unsupported: SheetUnsupportedModel[];
}

/**
 * The product of `SheetHandle.read(range?)`: full sheet metadata (name,
 * extent, frozen panes, column widths, merges, drawings) with `rows`
 * windowed to the requested range. `range` echoes the window that was
 * read; absent rows inside `extent` are empty, exactly like sparse
 * authored rows. Renderers paint whatever they are handed, so a range
 * model renders as-is.
 */
export interface SheetRangeModel extends SheetModel {
  /** The zero-based inclusive window this model was read with, in logical order. */
  range: CellRange;
}

/**
 * Worker unit payload: one slide, one sheet (or sheet window), or the
 * whole DOCX section list. Format entrypoints wrap these in native
 * collections; this union stays internal to the worker boundary.
 */
export type UnitModel = SlideModel | SheetModel | SheetRangeModel | DocumentSection[];

/**
 * DOCX flow blocks. These definitions are canonical: `core/index.ts`
 * re-exports them for the public `UnitModel`, and `core/docx/*` plus
 * `aared/docx` read and write them directly.
 *
 * Every block is plain data (structured-cloneable). Geometry is inches,
 * type is points — matching the PPTX model conventions above. `root` order
 * is document order; each entry is an independent virtualization block, so a
 * future block window (`blockStart`/`blockEnd` on `UnitOptions`) can slice
 * `root` without re-parsing.
 */

export type DocxTextAlign = "left" | "right" | "center" | "justify";

/** Shared flow-paragraph layout; Word `w:spacing`/`w:ind` resolved once in the worker. */
export interface DocxParagraphBorder {
  /** Resolved CSS color; `currentColor` when the source names none. */
  color: string;
  /** Border width in points (`w:sz` eighths). */
  widthPt: number;
  /** Gap between border and text in points (`w:space`). */
  spacePt: number;
  style: "solid" | "double" | "dashed" | "dotted";
}

export interface DocxParagraphBorders {
  top?: DocxParagraphBorder;
  right?: DocxParagraphBorder;
  bottom?: DocxParagraphBorder;
  left?: DocxParagraphBorder;
}

export interface DocxParagraphLayout {
  /** Effective default run size for paragraph-level layout, in points. */
  fontSizePt?: number;
  align?: DocxTextAlign;
  /** Authored relative line spacing; the renderer clamps its CSS box, never this value. */
  lineSpacing?: number;
  /** Explicit line height in points when the source does not use a multiplier. */
  lineSpacingPt?: number;
  /** Space before and after the paragraph, in points. */
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  /** Paragraph margins and first-line indent in inches. */
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndent?: number;
  /** Resolved CSS paragraph shading (`w:shd`). */
  shading?: string;
  /** Resolved paragraph borders (`w:pBdr`), per side. */
  borders?: DocxParagraphBorders;
}

export interface DocxParagraphNode extends DocxParagraphLayout {
  type: "paragraph";
  runs: TextRunModel[];
  /** Always resolved: Word stores `w:bidi` explicitly; otherwise `ltr`. */
  direction: TextDirection;
  /** Source paragraph style id, for diagnostics only. */
  styleId?: string;
}

export interface DocxHeadingNode extends DocxParagraphLayout {
  type: "heading";
  /** Word outline level 1–9 (`Heading 1`–`Heading 9` styles or `w:outlineLvl`). */
  level: number;
  /** Joined run text, for heading outlines and document titles. */
  text: string;
  direction: TextDirection;
  runs: TextRunModel[];
}

export interface DocxTableCellNode {
  nodes: FlowNode[];
  verticalAlign?: "top" | "middle" | "bottom";
  /** Resolved CSS cell shading (`w:shd`). */
  shading?: string;
  /** Logical-grid spans; omitted values each mean one cell. */
  rowSpan?: number;
  columnSpan?: number;
  /** Preferred cell width in inches, when authored. */
  width?: number;
}

export interface DocxTableRowNode {
  /** Row height in inches, when authored. */
  height?: number;
  /** Logical XML cell order, regardless of table display direction. */
  cells: DocxTableCellNode[];
}

export interface DocxTableNode {
  type: "table";
  /** Table display direction (`w:tblBidiVisual`); never reorders logical cells. */
  direction: TextDirection;
  /** Column widths in inches, in logical XML order. */
  columns: number[];
  rows: DocxTableRowNode[];
}

export interface DocxImageNode {
  type: "image";
  /** Identifier accepted by `MediaCollection.get()` (the OPC part path). */
  mediaId: string;
  /** Resolved MIME content type used to create the media Blob. */
  contentType: string;
  /** Extent in inches, from `wp:extent` or VML sizing. */
  width: number;
  height: number;
  alt?: string;
  /** Basic `wp:anchor` float; absent means inline flow. */
  float?: "left" | "right";
  /**
   * Host paragraph alignment for inline images (`w:jc`); centered source
   * paragraphs center their image. Absent when unauthored — and never set
   * for floats, which paragraph alignment does not position.
   */
  align?: DocxTextAlign;
}

export interface DocxListItemNode {
  /** Numbering level (`w:ilvl`), 0-based. */
  level: number;
  /** False for bullet levels, true for decimal/letter/roman levels. */
  ordered: boolean;
  /** Resolved marker text for this item (`1.`, `b)`, `●`). */
  marker: string;
  paragraph: DocxParagraphNode;
}

export interface DocxListNode {
  type: "list";
  /** Consecutive paragraphs sharing one `w:numId`; levels may mix inside. */
  items: DocxListItemNode[];
}

export interface DocxFootnoteNode {
  type: "footnote";
  /** Footnote/endnote entry id (`f`/`e`-prefixed; see parser). */
  id: string;
  /** Display number in order of first reference (`1`, `2`, …). */
  label: string;
  nodes: FlowNode[];
  direction: TextDirection;
}

/** A source drawing with no faithful flow representation; rendered as a marked fallback, never skipped silently. */
export interface DocxUnsupportedNode {
  type: "unsupported";
  source: "chart" | "ole" | "textbox" | "ink" | "drawing";
  label: string;
}

/**
 * Authored pagination signal: an explicit `w:br w:type="page"` or inherited
 * `w:pageBreakBefore` (`mandatory`), or a `w:lastRenderedPageBreak` hint from
 * the last Word layout (`preferred`). Flow rendering paints it as a line
 * break; layout pagination splits pages on it according to its strength.
 */
export interface DocxPageBreakNode {
  type: "pageBreak";
  strength: "mandatory" | "preferred";
}

export type FlowNode =
  | DocxParagraphNode
  | DocxHeadingNode
  | DocxTableNode
  | DocxImageNode
  | DocxListNode
  | DocxFootnoteNode
  | DocxUnsupportedNode
  | DocxPageBreakNode;

/** How a section starts relative to the previous one (`w:sectPr/w:type`; absent means `nextPage`). */
export type SectionBreakType = "continuous" | "nextPage" | "evenPage" | "oddPage";

/** Section page margins in twips (`1440 twips = 1 inch`), from `w:pgMar`. */
export interface SectionPageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
  header: number;
  footer: number;
  gutter: number;
}

/** Physical page geometry in twips, from `w:pgSz` plus orientation. */
export interface SectionPageGeometry {
  widthTwips: number;
  heightTwips: number;
  orientation: "portrait" | "landscape";
  margins: SectionPageMargins;
}

/**
 * Resolved `w:sectPr` for one document section: its break behavior, physical
 * page geometry, and header/footer part names in preference order (default
 * first). A paragraph-embedded `sectPr` closes the section that precedes it;
 * the trailing body `sectPr` closes the last section.
 */
export interface SectionProperties {
  breakType: SectionBreakType;
  page: SectionPageGeometry;
  headerRefs: string[];
  footerRefs: string[];
}

/**
 * One DOCX section: flow blocks in document order plus the header/footer
 * content resolved from its own `sectPr` references. Plain data
 * (structured-cloneable); geometry is twips, matching the OOXML source.
 */
export interface DocumentSection {
  properties: SectionProperties;
  root: FlowNode[];
  header?: FlowNode[];
  footer?: FlowNode[];
  /** First heading text within this section, when present — the section title. */
  title?: string;
}
