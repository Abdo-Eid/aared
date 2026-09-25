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

/** Paragraph alignment in physical terms; a bidi paragraph's start/end resolve to left/right before they reach the model. */
export type TextAlign = "left" | "right" | "center" | "justify";

export type TextBodyVerticalAlign = "top" | "middle" | "bottom" | "justify" | "distributed";

export type TextWrappingMode = "none" | "square";

export type TextAutofitMode = "none" | "shrinkText" | "resizeShape";

/**
 * Box insets in inches, every side optional (absent: not authored). The one
 * shape behind a text body's insets (`a:bodyPr`), a PPTX cell's margins
 * (`a:tcPr`) and a Word cell's padding (`w:tblCellMar` / `w:tcMar`).
 */
export interface InsetsModel {
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
  /** East Asian (`a:ea`) and symbol (`a:sym`) typeface slots; families resolve like latin/cs. */
  eastAsianTypeface?: string;
  symbolTypeface?: string;
  /** Resolved CSS color, not a scheme token. */
  color?: string;
  /** Resolved Office/theme font family; font files resolve separately. */
  fontFamily?: string;
  /** Word `w:u` single (or heavier) underline; absent means none. */
  underline?: boolean;
  /**
   * DrawingML underline style token (`a:rPr/@u`: `sng`, `dbl`, `dotted`,
   * `dashed`, `wavy`, …); `none` means no underline. Word keeps `underline`.
   */
  underlineStyle?: string;
  /** Resolved CSS underline color from `a:uFill`/`a:uLn`; absent follows the run color. */
  underlineColor?: string;
  /** Word `w:strike` / `w:dstrike`. */
  strikethrough?: boolean;
  /** DrawingML caps (`a:rPr/@cap`): `all` caps, `small` small-caps. */
  caps?: "all" | "small";
  /** Character spacing in points (DrawingML `a:rPr/@spc` hundredths, Word `w:spacing` twentieths). */
  letterSpacingPt?: number;
  /** DrawingML baseline shift as a percentage of the font size (`a:rPr/@baseline` thousandths of a percent). */
  baselinePct?: number;
  /** DrawingML `a:noFill`: outline-only text (fill explicitly none). */
  noFill?: boolean;
  /** DrawingML `a:grpFill`: fill inherits the group (marked, not resolved). */
  groupFill?: boolean;
  /** DrawingML text outline (`a:ln` on the run): stroke color + width in points. Painted via text-stroke. */
  outline?: { color: string; widthPt: number };
  /**
   * DrawingML text fill gradient (`a:gradFill` on the run). Linear gradients
   * paint via background-clip; path/unsupported kinds are marked, never
   * approximated — same contract as shape fills.
   */
  gradient?: ShapeFillModel;
  /** Run glow (`a:effectLst/a:glow`): color + radius in points. Approximated as text-shadow. */
  glow?: { color: string; radiusPt: number };
  /** Run shadow (`a:outerShdw`/`a:innerShdw`/`a:prstShdw`): color + blur + offset in points. Approximated as text-shadow. */
  shadow?: { color: string; blurPt: number; dxPt: number; dyPt: number };
  /** Word `w:vertAlign` superscript/subscript (footnote markers, endnotes). */
  verticalAlign?: "subscript" | "superscript";
  /** Resolved CSS highlight color behind the run (`w:highlight`, `a:highlight`). */
  highlight?: string;
  /** Resolved CSS run shading fill (`w:shd`). */
  shading?: string;
  /**
   * Hyperlink target. Word: runs inside `w:hyperlink` (externals resolved,
   * internals kept as `#anchor`). PPTX: `a:hlinkClick` `r:id` resolved
   * through slide relationships; slide jumps kept as `#slide-N`.
   */
  hyperlink?: string;
  /** Footnote/endnote entry id when this run is its in-text marker. */
  footnoteRef?: string;
  /**
   * Live field this run evaluates. DOCX: `PAGE`/`NUMPAGES` (`w:instrText` or
   * `w:fldSimple`); `text` carries the field's cached result from the file
   * (`"1"` when it has none), layout pagination substitutes the real value per
   * page, and flow rendering shows the cached text.
   * PPTX: `slidenum` (the renderer shows `SlideModel.number`) and `datetime`
   * (`text` keeps the value cached in the file).
   */
  field?: "PAGE" | "NUMPAGES" | "slidenum" | "datetime";
  /**
   * Word inline drawing that sits on the text line (`wp:inline` inside a
   * paragraph that also carries text). `text` is empty; the renderer paints
   * the picture in the line, so a header logo stays beside its wordmark.
   */
  image?: DocxInlineImage;
}

/** An inline picture riding on a text line; extent in inches. */
export interface DocxInlineImage {
  mediaId: string;
  contentType: string;
  width: number;
  height: number;
  alt?: string;
}

export interface TextParagraphModel {
  runs: TextRunModel[];
  /** Effective default run size for paragraph-level layout, in points. */
  fontSizePt?: number;
  /** Absent when the source did not explicitly set paragraph direction. */
  direction?: TextDirection;
  align?: TextAlign;
  bullet?: {
    char?: string;
    typeface?: string;
    /** Resolved CSS color from `a:buClr`; absent (or `buClrTx`) follows the paragraph text. */
    color?: string;
    /** Bullet size as a percentage of the run size (`a:buSzPct`, 100 = full). Absent (or `buSzTx`) follows text. */
    sizePct?: number;
    /** Bullet size in absolute points (`a:buSzPts`). Wins over `sizePct` when both resolve. */
    sizePts?: number;
    /**
     * Picture bullet (`a:buBlip`): media part target for session fetch plus
     * content type. The renderer paints the image, falling back to `char`
     * (or `•`) when fetch fails or no resolver is wired.
     */
    pictureMediaId?: string;
    pictureContentType?: string;
  };
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
  insets?: InsetsModel;
  verticalAlign?: TextBodyVerticalAlign;
  wrapping?: TextWrappingMode;
  autofit?: TextAutofitMode;
  /**
   * WordArt preset warp (`a:bodyPr/a:prstTxWarp/@prst`, e.g. `textArchUp`).
   * Painted unwarped and marked — true glyph warping needs per-glyph path
   * morphing and is explicitly out of scope.
   */
  textWarp?: string;
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
  /** SVG path data in the local coordinate system defined by `viewBox`; painted with the shape's fill and outline. */
  path: string;
  /**
   * Subpaths painted with the fill but no outline, beneath `path` — DrawingML
   * paths marked `stroke="false"` (a can's body, a cube's shaded faces).
   */
  fillOnly?: string;
  /**
   * Subpaths painted with the outline but no fill, above `path` — DrawingML
   * paths marked `fill="none"` (a can's rim, a callout's leader line).
   */
  strokeOnly?: string;
  viewBox: LocalViewBoxModel;
  /** Source geometry could not be rendered faithfully; `path` is a visible bounds fallback. */
  fallback?: "unsupported" | "custom" | "missing";
  /**
   * @internal Worker-to-session transport, never present on models a session
   * returns: geometry the worker described but did not evaluate. The session
   * resolves it on the main thread, lazily loading the evaluator and preset
   * catalog, then removes it — replacing `path` and `fallback` on success.
   */
  deferred?: DeferredGeometryModel;
}

/** @internal See `ShapeGeometryModel.deferred`. */
export interface DeferredGeometryModel {
  /** Encoded `a:custGeom`; absent for a preset, which the catalog supplies by `preset` name. */
  program?: string;
  /** Authored `a:avLst` values overriding the program's defaults by name. */
  adjustments?: Record<string, number>;
}

export interface StrokeModel {
  color: string;
  /** Stroke width in inches. */
  width: number;
  /** Dash family from `a:prstDash`; absent means solid. Read for table cell edges only so far. */
  dash?: "dash" | "dot";
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

/**
 * Edges of a rectangle as fractions of a fill box (1 = the box's full width
 * or height), each measured inward from the matching edge of the box. A
 * negative value extends that edge outward, as DrawingML `CT_RelativeRect`
 * allows.
 */
export interface RelativeRectModel {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/**
 * An `a:path` gradient: the first stop paints the `focus` rectangle and
 * colour runs outward to the last stop. How far "outward" reaches depends on
 * `path`, each measured against PowerPoint renders (DECISION_LOG §28):
 *
 * - `circle` — a true circle in slide units (never an ellipse fitted to the
 *   box): the weight is the distance from the focus divided by half the
 *   diagonal of `extent`.
 * - `rect` — per side: each side runs from the focus edge to the matching
 *   `extent` edge, so the sides meet on the focus-to-extent diagonals.
 * - `shape` — the fill's own outline: the weight is how far along the ray
 *   from the focus centre to the outline a point sits.
 *
 * Both rectangles are fractions of the fill box. PowerPoint reads
 * `fillToRect` against the fill box, not against `tileRect` as the spec
 * text suggests: a `tileRect l=t=-100%` + `fillToRect r=b=100%` gradient
 * puts the first stop at the box's top-left corner (`pptx-coverage.pptx`,
 * slides 4 and 7).
 */
export interface PathGradientFillModel {
  kind: "path-gradient";
  path: "circle" | "rect" | "shape";
  /** The first stop's rectangle (`a:fillToRect`), measured from `extent`'s edges. */
  focus: RelativeRectModel;
  /** The gradient box (`a:tileRect`); zeros are the fill box itself. */
  extent: RelativeRectModel;
  stops: LinearGradientStopModel[];
}

/**
 * A two-colour pattern (`a:pattFill`): one of DrawingML's 54 preset 8×8
 * bitmaps. Each bitmap cell is one CSS pixel (the tile is 1/12 inch), the
 * pattern scales with the slide, and it is anchored to the slide's origin
 * rather than to each shape, as PowerPoint draws it.
 */
export interface PatternFillModel {
  kind: "pattern";
  /** `ST_PresetPatternVal`, e.g. `dkDnDiag`, `smGrid`, `pct50`. */
  preset: string;
  /** Resolved CSS colour of the bitmap's set cells (`a:fgClr`). */
  foreground: string;
  /** Resolved CSS colour of the remaining cells (`a:bgClr`). */
  background: string;
}

/** How a tiled picture fill (`a:tile`) repeats. */
export interface PictureTileModel {
  /** Tile width and height as a fraction of the image's natural size at 96 DPI (1 = 100 %). */
  scaleX: number;
  scaleY: number;
  /** Offset of the tile grid in inches, after alignment. */
  offsetX: number;
  offsetY: number;
  /** Which point of the fill box the first tile aligns to. */
  align: "tl" | "t" | "tr" | "l" | "ctr" | "r" | "bl" | "b" | "br";
  /** Mirror every other tile horizontally (`x`), vertically (`y`), or both. */
  flip: "none" | "x" | "y" | "xy";
}

/**
 * A picture fill (`a:blipFill`) on a shape, table cell, or slide background.
 * Without `tile` the picture stretches over `stretch` (the fill box when
 * absent); `crop` removes a fraction of the image from each edge first.
 */
export interface PictureFillModel {
  kind: "picture";
  /** Identifier accepted by `MediaCollection.get()` (the OPC part path). */
  mediaId: string;
  contentType: string;
  /** `a:srcRect`: fractions of the image cut from each edge; negative pads. */
  crop?: RelativeRectModel;
  /** `a:stretch/a:fillRect`: where the stretched image sits inside the fill box. */
  stretch?: RelativeRectModel;
  tile?: PictureTileModel;
  /** True when inherited from the layout or master, so many slides share one image. */
  shared?: true;
}

/** A deliberately unpainted source fill that cannot be represented faithfully. */
export interface UnsupportedShapeFillModel {
  kind: "unsupported";
  source: "gradient-path" | "gradient-linear" | "image" | "pattern";
}

/**
 * Resolved DrawingML paint for shapes, table cells and slide backgrounds,
 * never raw OOXML or SVG markup.
 */
export type ShapeFillModel =
  | SolidFillModel
  | LinearGradientFillModel
  | PathGradientFillModel
  | PatternFillModel
  | PictureFillModel
  | UnsupportedShapeFillModel;

export interface ShapeModel {
  kind: "shape";
  id: string;
  name: string;
  placeholder?: { type: string; index: string };
  bounds?: BoundsModel;
  text?: TextBodyModel;
  /**
   * Whole-shape hyperlink (`p:cNvPr/a:hlinkClick`), resolved through the
   * owning part's rels. The renderer surfaces clicks as `linkactivate`;
   * run-level links (`TextRunModel.hyperlink`) win for their own runs.
   */
  hyperlink?: string;
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
  margins?: InsetsModel;
  /**
   * Resolved cell edge strokes: the direct `a:tcPr/a:lnL|lnR|lnT|lnB` edge,
   * else the table style's region edge. A missing side paints no line.
   * Widths are inches.
   */
  borders?: {
    left?: StrokeModel;
    right?: StrokeModel;
    top?: StrokeModel;
    bottom?: StrokeModel;
    /** Diagonals: `lnTlToBr` (top-left to bottom-right) and `lnBlToTr`. */
    diagonalDown?: StrokeModel;
    diagonalUp?: StrokeModel;
  };
  /** Rotated cell text from `a:tcPr/@vert`: `vert` reads top to bottom, `vert270` bottom to top. */
  textRotation?: 90 | 270;
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
  /**
   * Resolved background per logical XML cell: the cell's own `a:tcPr` fill,
   * else its table-style band colour. `undefined` paints no background.
   */
  cellFills: Array<Array<ShapeFillModel | undefined>>;
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
  /**
   * The chart area's own fill and outline (`c:chartSpace/c:spPr`), painted
   * behind the whole frame. Absent when the chart authors neither, so hosts
   * keep whatever surface is under it.
   */
  background?: {
    fill?: string;
    /** Outline width in CSS pixels. */
    stroke?: { color: string; width: number };
  };
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

/**
 * A grouped shape tree (`p:grpSp`). Children keep their authored coordinates
 * in `childSpace`; the renderer maps that space onto `bounds`. Nested groups
 * recurse through `children`.
 */
export interface GroupModel {
  kind: "group";
  id: string;
  name: string;
  bounds?: BoundsModel;
  /** Group rotation/flips from `p:grpSpPr/a:xfrm`. */
  transform?: TransformModel;
  /** Child coordinate space (`a:chOff` + `a:chExt`), absent when identity. */
  childSpace?: BoundsModel;
  /** Back-to-front children: shapes, tables, pictures, charts, nested groups. */
  children: SlideElementModel[];
}

export type SlideElementModel = ShapeModel | TableModel | PictureModel | ChartModel | GroupModel;

/**
 * The shared, non-placeholder graphics a layout and its master paint under
 * every slide that uses the layout (logos, bars, header text boxes). Stored
 * once per layout: the Worker parses it once, the session sends it to the
 * main thread once, and every slide on that layout carries the SAME object.
 * Treat it as read-only. Placeholders are prototypes, never painted here.
 */
export interface SlideLayoutModel {
  /** OPC part name of the layout. */
  id: string;
  /** Master graphics first, then the layout's, back-to-front. */
  graphics: SlideElementModel[];
}

export interface SlideModel {
  /** OPC part name for the source slide. */
  part: string;
  /** Slide dimensions in inches. */
  size: SizeModel;
  /** Back-to-front z-order. */
  elements: SlideElementModel[];
  /**
   * Resolved slide background (slide, else layout, else master). Absent means
   * the slide paints no authored background (white). `unsupported` marks a
   * background the renderer cannot paint yet (an off-center circle or
   * `shape`-path gradient, a missing picture).
   */
  background?: ShapeFillModel;
  /** 1-based position in the deck, shown by `slidenum` fields. */
  number?: number;
  /** Set by the Worker when the slide's layout has shared graphics to fetch. */
  layoutId?: string;
  /**
   * Shared layout/master graphics, attached by `SlideHandle.read()` from
   * `layoutId`. Painted under `elements` unless `hideBackgroundGraphics`.
   */
  layout?: SlideLayoutModel;
  /** The slide hides background graphics (`showMasterSp="0"`). */
  hideBackgroundGraphics?: boolean;
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

/** One side of a Word paragraph or table border (`w:pBdr`, `w:tblBorders`); widths are points. */
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

/** One Word tab stop (`w:tabs/w:tab`); `pos` is inches from the text margin's start edge. */
export interface DocxTabStop {
  pos: number;
  align: "left" | "center" | "right" | "decimal";
  /** Fill drawn across the tab's width; absent means blank. */
  leader?: "dot" | "hyphen" | "underscore" | "middleDot";
}

/**
 * Shared flow-paragraph layout, `w:spacing`/`w:ind` resolved once in the
 * worker: the paragraph fields a PPTX paragraph has (font size, alignment,
 * line spacing, space before/after, indents) plus Word's tab stops, shading
 * and borders. `lineSpacing` is the authored multiple; the renderer clamps its
 * CSS box, never this value.
 */
export interface DocxParagraphLayout
  extends Pick<
    TextParagraphModel,
    | "fontSizePt"
    | "align"
    | "lineSpacing"
    | "lineSpacingPt"
    | "spaceBeforePt"
    | "spaceAfterPt"
    | "leftIndent"
    | "rightIndent"
    | "firstLineIndent"
  > {
  /** Custom tab stops in ascending position; text tabs past the last stop use the default grid. */
  tabStops?: DocxTabStop[];
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

/** A table or cell edge: `"none"` is an authored `nil` that clears an inherited border. */
export type DocxTableEdge = DocxParagraphBorder | "none";

/** Table-level borders (`w:tblBorders`); the `inside*` edges sit between cells. */
export interface DocxTableBorders {
  top?: DocxTableEdge;
  right?: DocxTableEdge;
  bottom?: DocxTableEdge;
  left?: DocxTableEdge;
  insideH?: DocxTableEdge;
  insideV?: DocxTableEdge;
}

export interface DocxTableCellNode {
  nodes: FlowNode[];
  /** Per-cell padding (`w:tcMar`); absent falls back to the table's `cellMargins`. */
  margins?: InsetsModel;
  /** Per-cell edges (`w:tcBorders`), overriding the table's borders on that side. */
  borders?: Pick<DocxTableBorders, "top" | "right" | "bottom" | "left">;
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
  /** Preferred table width in inches (`w:tblW` in dxa); absent means the columns' sum. */
  width?: number;
  /** Table borders as authored; absent means no borders. */
  borders?: DocxTableBorders;
  /** Default cell padding (`w:tblCellMar`, else the Normal Table default). */
  cellMargins?: InsetsModel;
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
  align?: TextAlign;
}

export interface DocxListItemNode {
  /** Numbering level (`w:ilvl`), 0-based. */
  level: number;
  /** False for bullet levels, true for decimal/letter/roman levels. */
  ordered: boolean;
  /** Resolved marker text for this item (`1.`, `b)`, `●`). */
  marker: string;
  /**
   * The marker's own run formatting from the numbering level (`w:lvl/w:rPr`):
   * colour, weight, size, family. Absent when the level authors none, in which
   * case the marker follows the paragraph text. `text` repeats `marker`.
   */
  markerRun?: TextRunModel;
  /** What follows the marker (`w:suff`); absent is Word's default, a tab to the text indent. */
  markerSuffix?: "space" | "nothing";
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
