/**
 * Structured-cloneable resolved models shared by parsers, Workers, and
 * renderers. All geometry is expressed in inches unless documented otherwise.
 */

import type { ChartSpecModel } from "./chart/spec.js";

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

/**
 * One non-fatal thing a render noticed about its own output: a DOCX
 * paginator block that couldn't fit a page even alone, a footnote that ran
 * past its page, a font this render couldn't actually paint with (see
 * `font-fidelity.ts`'s `detectFontFidelity`). `code` is a stable, format- and
 * source-specific slug an embedder can switch on; `message` is for a human.
 * Shared by every format's render handle (`RenderHandleBase.warnings`,
 * `render-handle.ts`) so an embedder reads one shape regardless of source.
 */
export interface RenderWarning {
  code: string;
  message: string;
}

export type TextDirection = "ltr" | "rtl";

/** Paragraph alignment in physical terms; a bidi paragraph's start/end resolve to left/right before they reach the model. */
export type TextAlign = "left" | "right" | "center" | "justify";

export type TextBodyVerticalAlign = "top" | "middle" | "bottom" | "justify" | "distributed";

export type TextWrappingMode = "none" | "square";

export type TextAutofitMode = "none" | "shrinkText" | "resizeShape";

/**
 * Box insets in inches, every side optional. The one shape behind a text
 * body's insets (`a:bodyPr`), a PPTX cell's margins (`a:tcPr`) and a Word
 * cell's padding (`w:tblCellMar` / `w:tcMar`). A PPTX text body and cell always
 * carry all four sides, DrawingML's defaults filled in for any none states;
 * a Word cell's absent side means Word's own default was not overridden.
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
  /**
   * A left-to-right run (DOCX without `w:rtl`, PPTX in a left-to-right
   * `lang`) with no right-to-left letters, inside a right-to-left paragraph.
   * Word and PowerPoint keep it together left to right: "71%" and
   * "15 Oct 2026" stay as written, where the Unicode bidi algorithm alone
   * would show "%71" and "Oct 2026 15" (probe run-direction).
   */
  ltrRun?: true;
  latinTypeface?: string;
  complexScriptTypeface?: string;
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
  /** DrawingML text `a:grpFill`; PowerPoint paints no fill, and the source remains flagged. */
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
   * page, and flow rendering shows the cached text. `PAGEREF` is a table-of-
   * contents page reference: `hyperlink` names its target fragment
   * (`#aared-heading-N` for an outline heading, else the target heading's
   * `tocAnchor`); layout mode substitutes the target's printed page number
   * and flow rendering shows the cached text (empty for entries added since
   * the TOC was last saved).
   * PPTX: `slidenum` (the renderer shows `SlideModel.number`) and `datetime`
   * (`text` keeps the value cached in the file).
   */
  field?: "PAGE" | "NUMPAGES" | "PAGEREF" | "slidenum" | "datetime";
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
  /**
   * Office's own raster fallback beside an SVG source (`a:blip` next to
   * `asvg:svgBlip`): painted when the SVG cannot be read or decoded.
   */
  fallbackMediaId?: string;
  fallbackContentType?: string;
  width: number;
  height: number;
  alt?: string;
  effects?: DrawingEffectsModel;
  /** Word's authored effect overflow, in inches; separate from picture geometry. */
  effectInsets?: InsetsModel;
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
    /** The raster beside an SVG picture bullet, painted when the SVG fails. */
    pictureFallbackMediaId?: string;
    pictureFallbackContentType?: string;
  };
  level?: number;
  /** Paragraph margins and first-line indent in inches. */
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndent?: number;
  /** Line height as a multiple of the paragraph font size. */
  lineSpacing?: number;
  /**
   * Explicit line height in points when the source does not use a
   * percentage: an exact pitch, whatever the runs on the line are set in.
   */
  lineSpacingPt?: number;
  /** Space before and after the paragraph, in points (`a:spcPts`). */
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  /**
   * Space before and after the paragraph in lines (`a:spcPct`, 1 = 100%),
   * set instead of the points field. A PowerPoint line here is always 1.2 ×
   * the paragraph's largest run size, whatever the font or the paragraph's
   * own line spacing.
   */
  spaceBeforeLines?: number;
  spaceAfterLines?: number;
}

export interface TextBodyModel {
  paragraphs: TextParagraphModel[];
  insets?: InsetsModel;
  verticalAlign?: TextBodyVerticalAlign;
  wrapping?: TextWrappingMode;
  autofit?: TextAutofitMode;
  /**
   * The type scale PowerPoint stored with `a:normAutofit/@fontScale` (1 = 100%,
   * 0.925 = 92.5%), already applied to every resolved run and paragraph size
   * (each rounded to a whole point, as PowerPoint paints it). Absent when the
   * file stores none: PowerPoint does not shrink on open, so neither does aared.
   */
  fontScale?: number;
  /**
   * The line-spacing cut PowerPoint stored with `a:normAutofit/@lnSpcReduction`
   * (0.1 = 10 percentage points off each paragraph's line spacing), already
   * applied to `lineSpacing`.
   */
  lineSpacingReduction?: number;
  /**
   * WordArt preset warp (`a:bodyPr/a:prstTxWarp/@prst`, e.g. `textArchUp`).
   * Painted unwarped and marked. Most presets are free-form envelopes with no
   * DOM form; the arch and circle ones could follow an SVG `textPath`
   * (docs/ARCHITECTURE.md § Gaps, limits and fidelity gaps, PPTX WordArt).
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
  /** Geometry-defined text rectangle edges in local `viewBox` units, before body insets. */
  textRect?: { left: number; top: number; right: number; bottom: number };
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
  /** Dash family from `a:prstDash` on shape, line and table-cell strokes; absent means solid (or a dash style not read). */
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
  /** `srgb`: blend straight in sRGB, as Excel paints cell gradients. Absent: PowerPoint's blend (see `gradientColorAt`). */
  blend?: "srgb";
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
 * `path`, each measured against PowerPoint renders (Notes §28):
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
  /** `srgb`: blend straight in sRGB, as Excel paints cell gradients. Absent: PowerPoint's blend. */
  blend?: "srgb";
}

/**
 * A two-colour pattern: one of DrawingML's 54 preset 8×8 bitmaps
 * (`a:pattFill`) or SpreadsheetML's 17 cell patterns (`patternFill`). Each
 * bitmap cell is one CSS pixel (the tile is 1/12 inch) and the pattern scales
 * with its content. PowerPoint anchors the tiles to the slide's origin, Excel
 * to each cell's top-left; the painter's fill box says which.
 */
export interface PatternFillModel {
  kind: "pattern";
  /** `ST_PresetPatternVal` (`dkDnDiag`, `smGrid`, `pct50`) or `ST_PatternType` (`darkGrid`, `gray125`, `lightTrellis`). */
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
  /**
   * Office's own raster fallback beside an SVG source (`a:blip` next to
   * `asvg:svgBlip`): painted when the SVG cannot be read or decoded.
   */
  fallbackMediaId?: string;
  fallbackContentType?: string;
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
  source: "gradient-path" | "gradient-linear" | "image" | "pattern" | "group";
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

/** Resolved 2D DrawingML effects, applied to one shape or a group's composite. */
export interface DrawingEffectsModel {
  outerShadow?: { color: string; blurPt: number; dxPt: number; dyPt: number; rotatesWithShape: boolean };
  innerShadow?: { color: string; blurPt: number; dxPt: number; dyPt: number; rotatesWithShape: boolean };
  glow?: { color: string; radiusPt: number };
  softEdgePt?: number;
  reflection?: { distancePt: number; startOpacity: number; endOpacity: number; startPosition: number; endPosition: number };
  /** Effects outside the supported 2D vocabulary; never silently discarded. */
  unsupported?: string[];
}

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
  /** Group-inherited paint box in the shape's parent coordinates, in inches. */
  fillBounds?: BoundsModel;
  effects?: DrawingEffectsModel;
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
  /**
   * Office's own raster fallback beside an SVG source (`a:blip` next to
   * `asvg:svgBlip`): painted when the SVG cannot be read or decoded.
   */
  fallbackMediaId?: string;
  fallbackContentType?: string;
  /** Fractions (0..1) cropped from each edge via DrawingML `a:srcRect`. */
  crop?: { left: number; top: number; right: number; bottom: number };
  transform?: TransformModel;
  effects?: DrawingEffectsModel;
}

export interface ChartModel {
  kind: "chart";
  id: string;
  name: string;
  /** The author's alt text (`p:cNvPr/@descr`, else its `title`): the chart's accessible name. */
  alt?: string;
  bounds?: BoundsModel;
  /**
   * The resolved chart: what it plots and how it looks, every Office default
   * applied. Laid out and painted on the page, to the frame's size. Absent
   * when the chart plots a type the preview does not draw; the frame then
   * shows a placeholder saying why (`unsupportedReason`).
   */
  spec?: ChartSpecModel;
  /** Why `spec` is absent, phrased for a reader ("Map charts are not previewed"). */
  unsupportedReason?: string;
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
  /** Effects operate on the combined child appearance, not each child separately. */
  effects?: DrawingEffectsModel;
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
  /**
   * How many of the slide's elements (nested groups included) state no
   * `p:cNvPr/@id`. The file is damaged there; each was kept and drawn with an id
   * of its own (`auto-N`). Absent when none. The renderer turns it into a
   * `pptx-missing-id` warning.
   */
  missingIds?: number;
  /** Set by the Worker when the slide's layout has shared graphics to fetch. */
  layoutId?: string;
  /**
   * Shared layout/master graphics, attached by the private model reader from
   * `layoutId`. Painted under `elements` unless `hideBackgroundGraphics`.
   */
  layout?: SlideLayoutModel;
  /** The slide hides background graphics (`showMasterSp="0"`). */
  hideBackgroundGraphics?: boolean;
  /**
   * The slide's speaker notes as plain text (the notes page's body
   * placeholder). Read only: no renderer paints it and search skips it.
   */
  notes?: string;
}

// XLSX rendering models are Worker-resolved, structured-cloneable data.
// Track geometry is inches; sparse rows/cells stay proportional to content.
// UnitRequest.range slices a cached sheet; it still parses the full sheet once.

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
  /**
   * The cell background, through the same fill model as shapes: `solid`, a
   * `linear-gradient` or `path-gradient` (`gradientFill`, blended in sRGB) or
   * a `pattern` (`patternFill`'s 17 non-solid types).
   */
  fill?: ShapeFillModel;
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
    dataBar?: {
      /** Final bar length as a fraction of the cell width, including min/max length settings. */
      ratio: number;
      color: string;
      gradient: boolean;
      showValue: boolean;
      direction: "context" | "leftToRight" | "rightToLeft";
      border: boolean;
      borderColor?: string;
      axisPosition: "automatic" | "middle" | "none";
      axisColor?: string;
      negativeFillColor?: string;
      negativeBorderColor?: string;
    };
    icon?: {
      /** Family as written (`3Arrows`, `4TrafficLights`, `5Boxes`, ...). */
      family: string;
      /** Zero-based icon within the family in Excel's defined order (0 is the lowest tier). */
      icon: number;
      /** When false, the icon paints alone and the cell's text is hidden. */
      showValue: boolean;
    };
  };
}

export interface RowModel {
  /** Zero-based row index. */
  index: number;
  /** Sparse authored cells in ascending column order; absent columns are empty. */
  cells: Array<{ col: number; cell: CellModel }>;
  /** Row height in inches. */
  height?: number;
}

export interface SheetPictureModel {
  kind: "picture";
  name: string;
  /** Identifier accepted by `MediaCollection.get()`. */
  mediaId: string;
  /** Resolved MIME content type used to create the media Blob. */
  contentType: string;
  /**
   * Office's own raster fallback beside an SVG source (`a:blip` next to
   * `asvg:svgBlip`): painted when the SVG cannot be read or decoded.
   */
  fallbackMediaId?: string;
  fallbackContentType?: string;
  /** Resolved DrawingML 2D effects (`pic/spPr/effectLst`), painted by the shared effect painter. */
  effects?: DrawingEffectsModel;
  /** The picture's own rotation/flips, for effects that do not rotate with it. */
  transform?: TransformModel;
  /** Absolute bounds in inches from the grid origin. */
  bounds: BoundsModel;
  /** Anchor cells the bounds were derived from. */
  anchor: CellRange;
}

export interface SheetChartModel {
  kind: "chart";
  name: string;
  /** The author's alt text (`xdr:cNvPr/@descr`, else its `title`): the chart's accessible name. */
  alt?: string;
  /**
   * The resolved chart, read by the same reader as PPTX charts and painted
   * the same way. Absent when the chart plots a type the preview does not
   * draw; the reason is then in `unsupportedReason` and the bounds still
   * reserve its space.
   */
  spec?: ChartSpecModel;
  /** Absolute bounds in inches from the grid origin. */
  bounds: BoundsModel;
  /** Anchor cells the bounds were derived from. */
  anchor: CellRange;
  /** Optional source identifier for diagnostics and source correlation. */
  sourceChartId?: string;
  unsupportedReason?: string;
}

export type SheetDrawingModel = SheetPictureModel | SheetChartModel;

/**
 * Visible, non-silent record of content the preview does not show. Every kind
 * but `damaged` is a feature the previewer deliberately does not model;
 * `damaged` is the file itself at fault (an unreadable cell address, a
 * drawing pointing at a missing part), which Excel never writes, so a reader
 * can be told the file may be damaged rather than that a feature is missing.
 */
export interface SheetUnsupportedModel {
  kind: "pivot" | "macro" | "sparkline" | "data-validation" | "comment" | "conditional-format" | "drawing" | "damaged" | "other";
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
  /** Width in inches of a column no `<col>` sizes: what the grid continues with past `columnWidths`. */
  defaultColumnWidth: number;
  drawings: SheetDrawingModel[];
  unsupported: SheetUnsupportedModel[];
  /**
   * How many of the sheet's pictures and charts state no `xdr:cNvPr/@id`. The file is
   * damaged there (Excel offers to repair it); a drawing's id is not used, so each is
   * kept and drawn. Absent when none. The renderer turns it into an `xlsx-missing-id` warning.
   */
  missingIds?: number;
}

/**
 * The private sheet model reader's product: full sheet metadata (name,
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
 * Internal DOCX flow blocks: the canonical Worker/renderer representation.
 * Public section metadata is a separate projection, never these flow nodes.
 *
 * Every block is plain data (structured-cloneable). Geometry is inches,
 * type is points — matching the PPTX model conventions above. `root` order
 * is document order; each entry is an independent virtualization block.
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
  /** Authored paragraph spacing in line units, preserved until the section grid is known. */
  spaceBeforeLines?: number;
  /** Authored paragraph spacing in line units, preserved until the section grid is known. */
  spaceAfterLines?: number;
  /** Resolved `w:snapToGrid`; when absent, Word uses a defined section grid. */
  snapToGrid?: boolean;
  /**
   * `w:lineRule="atLeast"`: `lineSpacingPt` is a floor a taller run grows its
   * line past. Absent, `lineSpacingPt` is Word's `exact` pitch, which no run
   * on the line changes — the same meaning PowerPoint's points spacing has.
   */
  lineSpacingAtLeast?: true;
  /**
   * A justified paragraph Word lays out with its 2013 engine
   * (`compatibilityMode` 15), which takes back part of each space to fit one
   * more word on a line rather than breaking before it.
   */
  justifyShrink?: true;
  /** Custom tab stops in ascending position; text tabs past the last stop use the default grid. */
  tabStops?: DocxTabStop[];
  /** Resolved CSS paragraph shading (`w:shd`). */
  shading?: string;
  /** Resolved paragraph borders (`w:pBdr`), per side. */
  borders?: DocxParagraphBorders;
  /**
   * `w:keepNext`: layout mode keeps this paragraph on the page of the block
   * after it; a table row whose paragraphs all keep with next never ends a page.
   */
  keepNext?: true;
}

export interface DocxParagraphNode extends DocxParagraphLayout {
  type: "paragraph";
  runs: TextRunModel[];
  /** Always resolved: Word stores `w:bidi` explicitly; otherwise `ltr`. */
  direction: TextDirection;
  /** Source paragraph style id, for diagnostics only. */
  styleId?: string;
  /** Pictures and charts anchored in this paragraph (`wp:anchor`), placed relative to it; text wraps or runs over them. */
  floats?: DocxDrawingNode[];
}

export interface DocxHeadingNode extends DocxParagraphLayout {
  type: "heading";
  /** Pictures and charts anchored in this heading, as `DocxParagraphNode.floats`. */
  floats?: DocxDrawingNode[];
  /** Word outline level 1–9 (`Heading 1`–`Heading 9` styles or `w:outlineLvl`). */
  level: number;
  /**
   * Joined run text, for heading outlines and document titles. When the
   * heading carries numbering (`marker`), the marker leads it (`1. Project
   * Overview`), matching what Word paints and what a refreshed table of
   * contents lists; `runs` hold the authored text alone.
   */
  text: string;
  direction: TextDirection;
  runs: TextRunModel[];
  /** Source paragraph style id, for diagnostics only (as `DocxParagraphNode.styleId`). */
  styleId?: string;
  /**
   * Resolved numbering marker for a numbered heading (`1.`, `b)`), advanced
   * through the same sequence as list items. Absent when the heading carries
   * no numbering, or its level suppresses the marker (`w:numFmt w:val="none"`).
   */
  marker?: string;
  /**
   * The marker's own run formatting from the numbering level (`w:lvl/w:rPr`):
   * colour, weight, size, family. Absent when the level authors none, in which
   * case the marker follows the heading text (as `DocxListItemNode.markerRun`).
   * `text` repeats the marker without its following gap.
   */
  markerRun?: TextRunModel;
  /** What follows the marker (`w:suff`); absent is Word's default, a tab to the text indent. */
  markerSuffix?: "space" | "nothing";
  /**
   * Fragment id for a refreshed table-of-contents entry to link at when the
   * heading is outside the outline (a table-cell heading Word's TOC lists).
   * Outline headings keep their `aared-heading-N` fragment instead.
   */
  tocAnchor?: string;
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
  /**
   * Row height in inches, when authored.
   *
   * The authored `w:trHeight/@w:val` alone; `w:hRule` is not read. The height
   * is a floor, not a ceiling: a row taller than this grows to fit its cells,
   * so a `hRule="exact"` row paints taller here than Word clips it to. Nothing
   * in aared clips a row, matching its rule that rows never split.
   */
  height?: number;
  /** Logical XML cell order, regardless of table display direction. */
  cells: DocxTableCellNode[];
  /**
   * Word's own saved layout put a page break right before this row
   * (`w:lastRenderedPageBreak` leading the first cell's first paragraph).
   * Tables break between rows only, so this is the one place that hint can
   * still land. Read only, never authored: pagination treats it as a
   * `"preferred"` break the same way a paragraph-level hint is
   * (`SoftBreakPolicy`), not a forced one.
   */
  pageBreakBefore?: true;
}

export interface DocxTableNode {
  type: "table";
  /**
   * Table display direction (`w:bidiVisual`); never reorders logical cells.
   * In a right-to-left table, `left`/`right` borders are the start/end sides
   * (Word paints a `w:left` border on the right), while cell margins keep
   * their physical sides.
   */
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
  /**
   * Indent in inches from the start side (`w:tblInd` in dxa): the left of a
   * left-to-right table, the right of a right-to-left one. May be negative.
   */
  indent?: number;
}

export interface DocxImageNode {
  type: "image";
  /** Identifier accepted by `MediaCollection.get()` (the OPC part path). */
  mediaId: string;
  /** Resolved MIME content type used to create the media Blob. */
  contentType: string;
  /**
   * Office's own raster fallback beside an SVG source (`a:blip` next to
   * `asvg:svgBlip`): painted when the SVG cannot be read or decoded.
   */
  fallbackMediaId?: string;
  fallbackContentType?: string;
  /** Extent in inches, from `wp:extent` or VML sizing. */
  width: number;
  height: number;
  alt?: string;
  effects?: DrawingEffectsModel;
  effectInsets?: InsetsModel;
  /**
   * Where an anchored picture (`wp:anchor`) sits and how text wraps around it.
   * Absent for an inline picture, which flows like a glyph. An anchored picture
   * is carried by its anchor paragraph (`floats`), not placed in the flow.
   */
  anchor?: DocxImageAnchor;
  /**
   * Host paragraph alignment for inline images (`w:jc`); centered source
   * paragraphs center their image. A picture has no text to justify, so an
   * unauthored or justified alignment resolves to the paragraph's start side
   * (right in a right-to-left paragraph). Absent when unauthored - and never
   * set for anchored pictures, which paragraph alignment does not position.
   */
  align?: TextAlign;
}

/**
 * A chart in the flow (`w:drawing` holding a `c:chart`), placed exactly as a
 * picture is: inline in its own block, or anchored with text wrapping around
 * it. The chart reads with the same reader as PPTX and XLSX charts and lays
 * out at its extent.
 */
export interface DocxChartNode {
  type: "chart";
  /** The resolved chart, as `ChartModel.spec`. */
  spec: ChartSpecModel;
  /** Extent in inches, from `wp:extent`: the frame the chart lays out in. */
  width: number;
  height: number;
  /** `wp:docPr` description or title. */
  alt?: string;
  /** As `DocxImageNode.anchor`. */
  anchor?: DocxImageAnchor;
  /** As `DocxImageNode.align`. */
  align?: TextAlign;
}

type CellVerticalAlign = NonNullable<DocxTableCellNode["verticalAlign"]>;

/** A Word text box and the ordinary flow content inside it. Treat resolved data as read-only. */
export interface DocxTextBoxNode {
  type: "textBox";
  /** Content uses the same paragraph, table and list model as the document body. */
  content: FlowNode[];
  width: number;
  height: number;
  direction: TextDirection;
  insets?: InsetsModel;
  fill?: ShapeFillModel;
  border?: DocxParagraphBorder;
  geometry?: ShapeGeometryModel;
  transform?: TransformModel;
  effects?: DrawingEffectsModel;
  effectInsets?: InsetsModel;
  verticalAlign?: CellVerticalAlign;
  /** Internal text wrapping, independent of the anchor's wrapping around the shape. */
  wrapping?: TextWrappingMode;
  /**
   * Word grows the frame for `spAutoFit`; all other modes keep fixed extents and
   * font sizes. `shrinkText` is never resolved for DOCX, because Word ignores a
   * stored `normAutofit` scale (probe `autofit`, Notes §94).
   */
  autofit?: TextAutofitMode;
  /** Set for a legacy `w:pict`/VML text box. */
  legacy?: true;
  /** As `DocxImageNode.anchor`; absent for an inline box. */
  anchor?: DocxImageAnchor;
}

/** What a paragraph carries anchored: a picture, chart, or text box, placed the same way. */
export type DocxDrawingNode = DocxImageNode | DocxChartNode | DocxTextBoxNode;

/**
 * The placement Word saves on a `wp:anchor`: one wrap type, whether the
 * picture sits behind the text, a horizontal and a vertical position, the
 * distance text keeps from each side, and for tight and through wrapping the
 * polygon text follows. All lengths are inches.
 */
export interface DocxImageAnchor {
  /**
   * `square` follows the picture's box, `tight` and `through` its polygon,
   * `topAndBottom` leaves the lines beside it empty, `none` lets text run
   * over it (`wp:wrapNone`: behind or in front of the text, see `behind`).
   */
  wrap: "square" | "tight" | "through" | "topAndBottom" | "none";
  /** `behindDoc`: painted under the text. Independent of `wrap`: Word saves a tight wrap with it set. */
  behind: boolean;
  /** `wrapText`: which side of the picture text may use. */
  wrapText: "bothSides" | "left" | "right" | "largest";
  horizontal: DocxAnchorPosition<"left" | "center" | "right" | "inside" | "outside">;
  vertical: DocxAnchorPosition<"top" | "center" | "bottom" | "inside" | "outside">;
  /** Space text keeps from each side, from `distT`/`distB`/`distL`/`distR`. */
  distance: { top: number; bottom: number; left: number; right: number };
  /** Tight and through only: points as fractions (0-1) of the picture's width and height, from Word's 21600-unit square. */
  polygon?: Array<[x: number, y: number]>;
}

/** One axis of an anchor: the edge it is measured from, and either an alignment or an offset. */
export interface DocxAnchorPosition<Align extends string> {
  /** `relativeFrom`: `column`, `margin`, `page`, `paragraph`, `line`, `character` or a margin/page part. */
  from: string;
  align?: Align;
  /** `wp:posOffset` in inches from the `from` edge; may be negative. */
  offset?: number;
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
  source: "chart" | "ole" | "drawing";
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
  | DocxChartNode
  | DocxTextBoxNode
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
 * The header or footer parts a section's own references name, by kind (`w:type`
 * absent means `default`). A kind the section names none of is absent: it
 * inherits the previous section's, which `DocumentSection` already resolves.
 */
export interface HeaderFooterParts {
  default?: string;
  first?: string;
  even?: string;
}

/**
 * Resolved `w:sectPr` for one document section: its break behavior, physical
 * page geometry, the header/footer parts of its own references, whether its
 * first page has headers of its own (`w:titlePg`), and where its page numbers
 * start (`w:pgNumType`). A paragraph-embedded `sectPr` closes the section that
 * precedes it; the trailing body `sectPr` closes the last section.
 */
export interface SectionProperties {
  breakType: SectionBreakType;
  page: SectionPageGeometry;
  /** Section grid. `linePitch` is in twentieths of a point, as authored in `w:linePitch`. */
  docGrid?: { type: string; linePitch?: number; adjustLineHeightInTable?: true };
  headerParts?: HeaderFooterParts;
  footerParts?: HeaderFooterParts;
  /** `w:titlePg`: the section's first page takes the `first` header and footer. */
  titlePage?: true;
  /** `w:evenAndOddHeaders` of the document's settings: even page numbers take the `even` header and footer. */
  evenAndOddPages?: true;
  /** `w:pgNumType/@w:start`: the number of the section's first page; absent continues the previous section's. */
  pageNumberStart?: number;
}

/**
 * One DOCX section: flow blocks in document order plus the header/footer
 * content resolved from its own `sectPr` references, each kind (`default`, and
 * the `first`- and `even`-page ones) inheriting from the previous section on
 * its own. Plain data (structured-cloneable); geometry is twips, matching the
 * OOXML source. Flow mode paints `header` and `footer`, the default pair.
 */
export interface DocumentSection {
  properties: SectionProperties;
  /**
   * How many drawings in the whole document (body, headers, footers and notes) state no
   * `wp:docPr/@id`. Set on the first section only; absent when none. The file is damaged
   * there (Word offers to recover it); a drawing's id is not used, so each is kept and
   * drawn. The renderer turns it into a `docx-missing-id` warning.
   */
  missingIds?: number;
  root: FlowNode[];
  header?: FlowNode[];
  footer?: FlowNode[];
  firstHeader?: FlowNode[];
  firstFooter?: FlowNode[];
  evenHeader?: FlowNode[];
  evenFooter?: FlowNode[];
  /** First heading text within this section, when present — the section title. */
  title?: string;
}
