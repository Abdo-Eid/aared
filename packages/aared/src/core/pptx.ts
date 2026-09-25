/**
 * Focused PPTX reader for the Arabic/table gate fixture.
 *
 * This is deliberately headless and returns only plain data. It reads
 * PresentationML for slide order and DrawingML for text/table content;
 * DOM/SVG painting belongs to `aared/pptx`, never here.
 */

import { BULLET } from "./chars.js";
import { attr, childElements, firstChild, numberAttr, parseNumber, textContent } from "./xml.js";
import type { XmlElement } from "./xml.js";
import { guessMediaContentType } from "./opc.js";
import type { OpcPackage } from "./opc.js";
import { EMU_PER_PT, emuToInches, emuToPt } from "./units.js";
import { parseHexColor, relativeLuminance } from "./color.js";
import { readDrawingTheme, resolveDrawingColor, resolveThemeTypeface } from "./theme.js";
import { usesArabic } from "./typography.js";
import type { DrawingTheme } from "./theme.js";
import { AaredError, throwIfAborted } from "./errors.js";
import { parseChartGraphic } from "./chart.js";
import type {
  BoundsModel,
  ChartModel,
  GroupModel,
  InsetsModel,
  LinearGradientStopModel,
  PictureFillModel,
  PictureModel,
  PictureTileModel,
  RelativeRectModel,
  ShapeFillModel,
  ShapeModel,
  SlideLayoutModel,
  SlideModel,
  StrokeModel,
  TableCellModel,
  TableModel,
  TextBodyModel,
  TextParagraphModel,
  TextRunModel,
} from "./model.js";
import { geometryForPreset, unresolvedGeometry } from "./drawingml-shape.js";

/** `a:ln` width PowerPoint paints when the line states none: one point. */
const DEFAULT_LINE_WIDTH_EMU = EMU_PER_PT;

/** Cheap deck index: presentation metadata only, no slide XML inflated yet. */
export interface PptxDeck {
  size: { width: number; height: number };
  slideParts: string[];
  theme: DrawingTheme;
  tableStyles: Map<string, PptxTableStyle>;
  /** Embedded-font parts (`ppt/fonts/*.fntdata`) for future resolver delivery. */
  embeddedFonts: string[];
  /**
   * The deck's default run size per outline level (`p:defaultTextStyle`), in
   * points: what text with no placeholder or master style (text boxes,
   * shapes, table cells) falls back to. Colour is deliberately not taken
   * from it — see `parseRun` on automatic contrast.
   */
  defaultTextSizes: Array<number | undefined>;
}

/** A table-style edge: a resolved stroke, or `null` for an authored `a:noFill` that clears what a weaker region drew. */
type TableStyleEdge = StrokeModel | null;

/** One `a:tblStyle` region (`wholeTbl`, `firstRow`, `band1H`, …), theme-resolved. */
interface TableStyleRegion {
  fill?: string;
  text?: string;
  bold?: boolean;
  italic?: boolean;
  /** `tcTxStyle` typeface (`a:font/a:latin` or a `fontRef` theme slot). */
  font?: string;
  borders?: Partial<Record<"left" | "right" | "top" | "bottom" | "insideH" | "insideV" | "tl2br" | "tr2bl", TableStyleEdge>>;
}

/** The regions a table style can author. */
const TABLE_STYLE_REGIONS = [
  "wholeTbl", "band1H", "band2H", "band1V", "band2V",
  "lastCol", "firstCol", "lastRow", "firstRow",
  "seCell", "swCell", "neCell", "nwCell",
] as const;

type TableStyleRegionName = typeof TABLE_STYLE_REGIONS[number];

type PptxTableStyle = Partial<Record<TableStyleRegionName, TableStyleRegion>>;

function onOff(element: XmlElement, name: string): boolean | undefined {
  const value = attr(element, name);
  if (value === undefined) return undefined;
  if (value === "1" || value === "true" || value === "on") return true;
  if (value === "0" || value === "false" || value === "off") return false;
  throw new Error(`aared/core: invalid ${name}="${value}" on <${element.name}>`);
}

function align(value: string | undefined): TextParagraphModel["align"] {
  switch (value) {
    case "l":
      return "left";
    case "r":
      return "right";
    case "ctr":
      return "center";
    case "just":
    case "dist":
      return "justify";
    case undefined:
      return undefined;
    default:
      throw new Error(`aared/core: unsupported paragraph alignment "${value}"`);
  }
}

function verticalAlign(value: string | undefined): TableCellModel["verticalAlign"] {
  switch (value) {
    case "t":
    case undefined:
      return undefined;
    case "ctr":
      return "middle";
    case "b":
      return "bottom";
    default:
      throw new Error(`aared/core: unsupported table-cell vertical alignment "${value}"`);
  }
}

function parseBounds(xfrm: XmlElement | undefined): BoundsModel | undefined {
  if (xfrm === undefined) return undefined;
  const off = firstChild(xfrm, "off");
  const ext = firstChild(xfrm, "ext");
  if (off === undefined || ext === undefined) return undefined;
  const x = numberAttr(off, "x", "a:off");
  const y = numberAttr(off, "y", "a:off");
  const width = numberAttr(ext, "cx", "a:ext");
  const height = numberAttr(ext, "cy", "a:ext");
  if (x === undefined || y === undefined || width === undefined || height === undefined) return undefined;
  return { x: emuToInches(x), y: emuToInches(y), width: emuToInches(width), height: emuToInches(height) };
}

function resolveColor(element: XmlElement | undefined, theme: DrawingTheme): string | undefined {
  return resolveDrawingColor(element, theme);
}

function resolveSolidFill(element: XmlElement | undefined, theme: DrawingTheme): string | undefined {
  return resolveColor(element === undefined ? undefined : firstChild(element, "solidFill"), theme);
}

function resolveGradientStops(gradient: XmlElement, theme: DrawingTheme): LinearGradientStopModel[] | undefined {
  const stopList = firstChild(gradient, "gsLst");
  if (stopList === undefined) return undefined;
  const stops: LinearGradientStopModel[] = [];
  for (const stop of childElements(stopList, "gs")) {
    const position = numberAttr(stop, "pos", "a:gs");
    const color = resolveColor(stop, theme);
    if (position === undefined || position < 0 || position > 100000 || color === undefined) return undefined;
    stops.push({ position: position / 100000, color });
  }
  return stops.length > 0 ? stops : undefined;
}

/** A `CT_RelativeRect` (`l`/`t`/`r`/`b` in 1/1000 %, default 0) as fractions; undefined when the element is absent. */
function relativeRect(element: XmlElement | undefined): RelativeRectModel | undefined {
  if (element === undefined) return undefined;
  const side = (name: string): number => (numberAttr(element, name, `a:${element.local}`) ?? 0) / 100000;
  return { left: side("l"), top: side("t"), right: side("r"), bottom: side("b") };
}

const NO_INSET: RelativeRectModel = { left: 0, top: 0, right: 0, bottom: 0 };

function isZeroRect(rect: RelativeRectModel | undefined): boolean {
  return rect === undefined || (rect.left === 0 && rect.top === 0 && rect.right === 0 && rect.bottom === 0);
}

function resolveGradientFill(gradient: XmlElement, theme: DrawingTheme): ShapeFillModel {
  const path = firstChild(gradient, "path");
  if (path !== undefined) {
    const stops = resolveGradientStops(gradient, theme);
    if (stops === undefined) return { kind: "unsupported", source: "gradient-path" };
    const kind = attr(path, "path");
    if (kind !== "circle" && kind !== "rect" && kind !== "shape") return { kind: "unsupported", source: "gradient-path" };
    // Absent rectangles take the schema default: zero insets.
    return {
      kind: "path-gradient",
      path: kind,
      focus: relativeRect(firstChild(path, "fillToRect")) ?? { ...NO_INSET },
      extent: relativeRect(firstChild(gradient, "tileRect")) ?? { ...NO_INSET },
      stops,
    };
  }
  const linear = firstChild(gradient, "lin");
  if (linear === undefined) return { kind: "unsupported", source: "gradient-linear" };
  const stops = resolveGradientStops(gradient, theme);
  if (stops === undefined) return { kind: "unsupported", source: "gradient-linear" };
  const drawingMlAngle = numberAttr(linear, "ang", "a:lin") ?? 0;
  return {
    kind: "linear-gradient",
    angle: ((drawingMlAngle / 60000) % 360 + 360) % 360,
    // An absent `scaled` is false: PowerPoint draws slide 4's `<a:lin
    // ang="2700000"/>` with bands square to a true 45° across a 1.7:1 box
    // (DECISION_LOG §28), not squashed to the box's diagonal.
    scaled: onOff(linear, "scaled") ?? false,
    rotatesWithShape: onOff(gradient, "rotWithShape") ?? true,
    stops,
  };
}

/**
 * A shape/table style reference (`a:fillRef`, `a:lnRef`) as a color.
 * `idx="0"` (or a missing `idx`) means "no reference": PowerPoint renders no
 * fill/line even though the schema requires a vestigial color child, so the
 * child must not resolve. Ground truth: the coverage deck's unfilled right
 * brace (`fillRef idx="0"` + accent1 child) paints no fill in PowerPoint,
 * while its `lnRef idx="2"` sibling resolves normally.
 */
export function resolveStyleRef(reference: XmlElement | undefined, theme: DrawingTheme): string | undefined {
  if (reference === undefined) return undefined;
  const index = attr(reference, "idx");
  if (index === undefined || index === "0") return undefined;
  return resolveColor(reference, theme);
}

/**
 * The part a fill belongs to, so a picture fill's `r:embed` resolves against
 * that part's relationships: the slide, or the layout/master a background or
 * shared graphic comes from.
 */
interface FillOwner {
  relationships: ReadonlyArray<{ id: string; type: string; target: string; targetMode: string }>;
  mediaContentType: (target: string) => string;
  /** Inherited from the layout or master, so many slides share the picture. */
  inherited: boolean;
}

/**
 * `a:pattFill`: a preset bitmap in two colours. The bitmaps themselves live
 * with the renderer; an unknown preset name is marked there, not here.
 * `fgClr`/`bgClr` default to black on white, as PowerPoint writes them.
 */
function resolvePatternFill(pattern: XmlElement, theme: DrawingTheme): ShapeFillModel {
  const preset = attr(pattern, "prst");
  if (preset === undefined) return { kind: "unsupported", source: "pattern" };
  return {
    kind: "pattern",
    preset,
    foreground: resolveColor(firstChild(pattern, "fgClr"), theme) ?? "#000000",
    background: resolveColor(firstChild(pattern, "bgClr"), theme) ?? "#FFFFFF",
  };
}

const TILE_ALIGNMENTS = new Set(["tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"]);
const TILE_FLIPS = new Set(["none", "x", "y", "xy"]);

/**
 * `a:blipFill` as a fill: stretched (optionally into `a:stretch/a:fillRect`)
 * or tiled (`a:tile`: `sx`/`sy` in 1/1000 %, `tx`/`ty` in EMU). A picture the
 * owner's relationships cannot resolve stays `unsupported`.
 */
function resolvePictureFill(blipFill: XmlElement, owner: FillOwner | undefined): ShapeFillModel {
  const unsupported = { kind: "unsupported", source: "image" } as const;
  if (owner === undefined) return unsupported;
  const blip = firstChild(blipFill, "blip");
  const embed = blip === undefined ? undefined : (attr(blip, "embed") ?? attr(blip, "r:embed"));
  const rel = embed === undefined ? undefined : owner.relationships.find((candidate) => candidate.id === embed);
  if (rel === undefined || rel.targetMode === "External" || !rel.type.endsWith("/image")) return unsupported;
  const out: PictureFillModel = {
    kind: "picture",
    mediaId: rel.target,
    contentType: owner.mediaContentType(rel.target),
    ...(owner.inherited ? { shared: true as const } : {}),
  };
  const crop = relativeRect(firstChild(blipFill, "srcRect"));
  if (!isZeroRect(crop)) out.crop = crop!;
  const tile = firstChild(blipFill, "tile");
  if (tile !== undefined) {
    const align = attr(tile, "algn") ?? "tl";
    const flip = attr(tile, "flip") ?? "none";
    out.tile = {
      scaleX: (numberAttr(tile, "sx", "a:tile") ?? 100000) / 100000,
      scaleY: (numberAttr(tile, "sy", "a:tile") ?? 100000) / 100000,
      offsetX: emuToInches(numberAttr(tile, "tx", "a:tile") ?? 0),
      offsetY: emuToInches(numberAttr(tile, "ty", "a:tile") ?? 0),
      align: (TILE_ALIGNMENTS.has(align) ? align : "tl") as PictureTileModel["align"],
      flip: (TILE_FLIPS.has(flip) ? flip : "none") as PictureTileModel["flip"],
    };
    return out;
  }
  const stretch = firstChild(blipFill, "stretch");
  const fillRect = relativeRect(stretch === undefined ? undefined : firstChild(stretch, "fillRect"));
  if (!isZeroRect(fillRect)) out.stretch = fillRect!;
  return out;
}

/**
 * The fill authored directly on `container` (`p:spPr`, `p:bgPr`, `a:tcPr`):
 * `null` for an explicit `a:noFill`, `undefined` when no fill is authored so
 * the caller's style fallback applies.
 */
function resolveDirectFill(container: XmlElement, theme: DrawingTheme, owner: FillOwner | undefined): ShapeFillModel | null | undefined {
  if (firstChild(container, "noFill") !== undefined) return null;
  const gradient = firstChild(container, "gradFill");
  if (gradient !== undefined) return resolveGradientFill(gradient, theme);
  const blipFill = firstChild(container, "blipFill");
  if (blipFill !== undefined) return resolvePictureFill(blipFill, owner);
  const pattern = firstChild(container, "pattFill");
  if (pattern !== undefined) return resolvePatternFill(pattern, theme);
  const color = resolveSolidFill(container, theme);
  return color === undefined ? undefined : { kind: "solid", color };
}

function resolveShapeFill(
  spPr: XmlElement | undefined,
  style: XmlElement | undefined,
  theme: DrawingTheme,
  owner: FillOwner | undefined,
): ShapeFillModel | undefined {
  const direct = spPr === undefined ? undefined : resolveDirectFill(spPr, theme, owner);
  if (direct === null) return undefined;
  if (direct !== undefined) return direct;
  const color = resolveStyleRef(style === undefined ? undefined : firstChild(style, "fillRef"), theme);
  return color === undefined ? undefined : { kind: "solid", color };
}

/** `a:prstDash` as a dash family; solid and unknown values paint solid. */
function lineDash(prstDash: XmlElement | undefined): StrokeModel["dash"] {
  const value = prstDash === undefined ? undefined : attr(prstDash, "val");
  if (value === undefined || value === "solid") return undefined;
  return value === "dot" || value === "sysDot" ? "dot" : "dash";
}

/**
 * Line width PowerPoint's default themes give `a:lnRef idx="1|2|3"` (the
 * theme's `a:lnStyleLst`, which aared does not read yet): 0.5, 1 and 1.5 pt.
 */
const THEME_LINE_WIDTHS_EMU = [6350, 12700, 19050];

/**
 * One table edge: a resolved stroke, `null` for an authored `a:noFill` (an
 * explicit "no line" that overrides a weaker region's edge), or undefined
 * when nothing paintable is authored. Accepts a direct cell edge (`a:lnL`)
 * or a table-style edge wrapper (`a:left`, holding `a:ln` or `a:lnRef`).
 */
function parseTableEdge(edge: XmlElement | undefined, theme: DrawingTheme): StrokeModel | null | undefined {
  if (edge === undefined) return undefined;
  const ln = edge.local.startsWith("ln") && edge.local !== "lnRef" ? edge : firstChild(edge, "ln");
  if (ln !== undefined) {
    if (firstChild(ln, "noFill") !== undefined) return null;
    const color = resolveSolidFill(ln, theme);
    if (color === undefined) return undefined;
    const out: StrokeModel = { color, width: emuToInches(numberAttr(ln, "w", "a:ln") ?? DEFAULT_LINE_WIDTH_EMU) };
    const dash = lineDash(firstChild(ln, "prstDash"));
    if (dash !== undefined) out.dash = dash;
    return out;
  }
  const lnRef = firstChild(edge, "lnRef");
  if (lnRef === undefined) return undefined;
  const index = numberAttr(lnRef, "idx", "a:lnRef") ?? 0;
  if (index <= 0) return null;
  const color = resolveDrawingColor(lnRef, theme);
  if (color === undefined) return undefined;
  return { color, width: emuToInches(THEME_LINE_WIDTHS_EMU[Math.min(index, 3) - 1]!) };
}

function readTableStyleRegion(region: XmlElement | undefined, theme: DrawingTheme): TableStyleRegion | undefined {
  if (region === undefined) return undefined;
  const out: TableStyleRegion = {};
  const tcStyle = firstChild(region, "tcStyle");
  if (tcStyle !== undefined) {
    const fillElement = firstChild(tcStyle, "fill");
    const fill = resolveSolidFill(fillElement, theme)
      ?? resolveStyleRef(firstChild(tcStyle, "fillRef"), theme);
    if (fill !== undefined) out.fill = fill;
    const bdr = firstChild(tcStyle, "tcBdr");
    if (bdr !== undefined) {
      const borders: NonNullable<TableStyleRegion["borders"]> = {};
      for (const side of ["left", "right", "top", "bottom", "insideH", "insideV", "tl2br", "tr2bl"] as const) {
        const edge = parseTableEdge(firstChild(bdr, side), theme);
        if (edge !== undefined) borders[side] = edge;
      }
      if (Object.keys(borders).length > 0) out.borders = borders;
    }
  }
  // `tcTxStyle` carries its colour as a direct `schemeClr`/`srgbClr` child;
  // the sibling `fontRef` names the typeface slot. Ground truth: Medium
  // Style 2 Accent 1 in pptx-coverage.pptx (wholeTbl dk1, firstRow lt1).
  const txStyle = firstChild(region, "tcTxStyle");
  if (txStyle !== undefined) {
    const text = resolveDrawingColor(txStyle, theme);
    if (text !== undefined) out.text = text;
    const bold = onOff(txStyle, "b");
    if (bold !== undefined) out.bold = bold;
    const italic = onOff(txStyle, "i");
    if (italic !== undefined) out.italic = italic;
    const font = firstChild(txStyle, "font");
    const latin = font === undefined ? undefined : firstChild(font, "latin");
    const fontRef = firstChild(txStyle, "fontRef");
    const slot = fontRef === undefined ? undefined : attr(fontRef, "idx");
    const typeface = latin !== undefined ? attr(latin, "typeface")
      : slot === "major" ? "+mj-lt" : slot === "minor" ? "+mn-lt" : undefined;
    if (typeface !== undefined && typeface !== "") out.font = typeface;
  }
  return Object.keys(out).length === 0 ? undefined : out;
}

/** Which grid rows and columns a table-style region covers, given the table's flags. */
interface TableShape {
  rows: number;
  columns: number;
  firstRow: boolean;
  lastRow: boolean;
  firstCol: boolean;
  lastCol: boolean;
  bandRow: boolean;
  bandCol: boolean;
}

/**
 * The regions that apply to the cell spanning rows `r0..r1` and columns
 * `c0..c1`, weakest first (PowerPoint's precedence: whole table, column
 * bands, row bands, last/first column, last/first row, then the corner
 * cells), each with the extent its outer edges belong to.
 */
function tableCellRegions(
  shape: TableShape,
  r0: number, r1: number, c0: number, c1: number,
): Array<{ name: TableStyleRegionName; rows: [number, number]; columns: [number, number] }> {
  const all: [number, number] = [0, shape.rows - 1];
  const allColumns: [number, number] = [0, shape.columns - 1];
  const regions: Array<{ name: TableStyleRegionName; rows: [number, number]; columns: [number, number] }> = [
    { name: "wholeTbl", rows: all, columns: allColumns },
  ];
  const isHeaderRow = shape.firstRow && r0 === 0;
  const isTotalRow = shape.lastRow && r1 === shape.rows - 1;
  const isFirstColumn = shape.firstCol && c0 === 0;
  const isLastColumn = shape.lastCol && c1 === shape.columns - 1;
  if (shape.bandCol && !isFirstColumn && !isLastColumn) {
    const band = (c0 - (shape.firstCol ? 1 : 0)) % 2 === 0 ? "band1V" : "band2V";
    regions.push({ name: band, rows: all, columns: [c0, c1] });
  }
  if (shape.bandRow && !isHeaderRow && !isTotalRow) {
    const band = (r0 - (shape.firstRow ? 1 : 0)) % 2 === 0 ? "band1H" : "band2H";
    regions.push({ name: band, rows: [r0, r1], columns: allColumns });
  }
  if (isLastColumn) regions.push({ name: "lastCol", rows: all, columns: [shape.columns - 1, shape.columns - 1] });
  if (isFirstColumn) regions.push({ name: "firstCol", rows: all, columns: [0, 0] });
  if (isTotalRow) regions.push({ name: "lastRow", rows: [shape.rows - 1, shape.rows - 1], columns: allColumns });
  if (isHeaderRow) regions.push({ name: "firstRow", rows: [0, 0], columns: allColumns });
  const corner = (name: TableStyleRegionName, r: number, c: number): void => {
    regions.push({ name, rows: [r, r], columns: [c, c] });
  };
  if (isTotalRow && isLastColumn) corner("seCell", shape.rows - 1, shape.columns - 1);
  if (isTotalRow && isFirstColumn) corner("swCell", shape.rows - 1, 0);
  if (isHeaderRow && isLastColumn) corner("neCell", 0, shape.columns - 1);
  if (isHeaderRow && isFirstColumn) corner("nwCell", 0, 0);
  return regions;
}

/** A cell's table-style formatting: every applicable region layered weakest first, outer edges from the region's extent, inner ones from its `insideH`/`insideV`. */
function resolveTableStyleCell(
  style: PptxTableStyle | undefined,
  shape: TableShape,
  r0: number, r1: number, c0: number, c1: number,
  /** An RTL table paints logical column 0 on the right, so the style's physical `left` outer edge belongs to the logical last column. */
  rtl: boolean,
): {
  fill?: string; text?: string; bold?: boolean; italic?: boolean; font?: string;
  edges: Partial<Record<"left" | "right" | "top" | "bottom" | "diagonalDown" | "diagonalUp", TableStyleEdge>>;
} {
  const out: ReturnType<typeof resolveTableStyleCell> = { edges: {} };
  if (style === undefined) return out;
  for (const { name, rows, columns } of tableCellRegions(shape, r0, r1, c0, c1)) {
    const region = style[name];
    if (region === undefined) continue;
    if (region.fill !== undefined) out.fill = region.fill;
    if (region.text !== undefined) out.text = region.text;
    if (region.bold !== undefined) out.bold = region.bold;
    if (region.italic !== undefined) out.italic = region.italic;
    if (region.font !== undefined) out.font = region.font;
    const borders = region.borders;
    if (borders === undefined) continue;
    const pick = (outer: boolean, outerKey: "left" | "right" | "top" | "bottom", innerKey: "insideH" | "insideV"): TableStyleEdge | undefined =>
      outer ? borders[outerKey] : borders[innerKey];
    const edges: Array<["left" | "right" | "top" | "bottom", TableStyleEdge | undefined]> = [
      ["top", pick(r0 === rows[0], "top", "insideH")],
      ["bottom", pick(r1 === rows[1], "bottom", "insideH")],
      ["left", pick(rtl ? c1 === columns[1] : c0 === columns[0], "left", "insideV")],
      ["right", pick(rtl ? c0 === columns[0] : c1 === columns[1], "right", "insideV")],
    ];
    for (const [side, edge] of edges) if (edge !== undefined) out.edges[side] = edge;
    if (borders.tl2br !== undefined) out.edges.diagonalDown = borders.tl2br;
    if (borders.tr2bl !== undefined) out.edges.diagonalUp = borders.tr2bl;
  }
  return out;
}

async function readTableStyles(pkg: OpcPackage, theme: DrawingTheme, signal?: AbortSignal): Promise<Map<string, PptxTableStyle>> {
  throwIfAborted(signal);
  const part = pkg.partNames().find((name) => pkg.contentTypeOf(name)?.endsWith("tableStyles+xml"));
  if (part === undefined) return new Map();
  const root = await pkg.readXml(part, signal);
  const styles = new Map<string, PptxTableStyle>();
  for (const style of childElements(root, "tblStyle")) {
    throwIfAborted(signal);
    const id = attr(style, "styleId");
    if (id === undefined) continue;
    const resolved: PptxTableStyle = {};
    for (const name of TABLE_STYLE_REGIONS) {
      const region = readTableStyleRegion(firstChild(style, name), theme);
      if (region !== undefined) resolved[name] = region;
    }
    styles.set(id, resolved);
  }
  throwIfAborted(signal);
  return styles;
}

function firstAttribute(elements: XmlElement[], name: string): string | undefined {
  for (const element of elements) {
    const value = attr(element, name);
    if (value !== undefined) return value;
  }
  return undefined;
}

function firstElement(elements: XmlElement[], name: string): XmlElement | undefined {
  for (const element of elements) {
    const child = firstChild(element, name);
    if (child !== undefined) return child;
  }
  return undefined;
}

/**
 * Hyperlink targets for one part: external URLs verbatim, slide jumps and
 * other internal targets kept as `#`-prefixed part paths. The renderer
 * surfaces them via `linkactivate`; the host decides.
 */
/** Relationship lists are cached per part by `OpcPackage`, so their link maps can be too. */
const linkMaps = new WeakMap<object, Map<string, string>>();

function hyperlinkTargets(relationships: Array<{ id: string; type: string; target: string; targetMode: string }>): Map<string, string> {
  const known = linkMaps.get(relationships);
  if (known !== undefined) return known;
  const out = new Map<string, string>();
  linkMaps.set(relationships, out);
  for (const rel of relationships) {
    if (rel.type.endsWith("/hyperlink")) {
      out.set(rel.id, rel.targetMode === "External" ? rel.target : `#${rel.target}`);
    } else if (rel.type.endsWith("/slide")) {
      if (!out.has(rel.id)) out.set(rel.id, `#${rel.target}`);
    }
  }
  return out;
}

/**
 * PowerPoint "Automatic" run color: fill-less text contrasts its slide
 * background (white on dark, black otherwise). Only solid backgrounds resolve
 * exactly; gradients/images and missing backgrounds default to black (a
 * documented approximation — sampling those needs its own probe).
 * Ground truth: coverage slides 3–4 titles on `#3A4961` paint white.
 */
export function automaticTextColor(background: SlideModel["background"]): string {
  const channels = background?.kind === "solid" ? parseHexColor(background.color) : undefined;
  return channels !== undefined && relativeLuminance(channels) < 0.35 ? "#FFFFFF" : "#000000";
}

/** Picture-bullet sources for one part: image rels by `r:id`, typed through the package's media rule. */
function bulletEmbeds(
  relationships: Array<{ id: string; type: string; target: string; targetMode: string }> | undefined,
  mediaContentType: (target: string) => string,
): Map<string, { mediaId: string; contentType: string }> | undefined {
  if (relationships === undefined) return undefined;
  const out = new Map<string, { mediaId: string; contentType: string }>();
  for (const rel of relationships) {
    if (rel.type.endsWith("/image") && rel.targetMode !== "External") {
      out.set(rel.id, { mediaId: rel.target, contentType: mediaContentType(rel.target) });
    }
  }
  return out.size === 0 ? undefined : out;
}

/**
 * First embedded-image reference (`r:embed`) anywhere under a picture
 * bullet (`a:buBlip/a:blip`, including `asvg:svgBlip` extension embeds).
 * `attr()` matches the local name, so prefixed `r:embed` resolves.
 */
function findBulletEmbed(buBlip: XmlElement): string | undefined {
  const stack = [...buBlip.children];
  while (stack.length > 0) {
    const element = stack.pop()!;
    const embed = attr(element, "embed");
    if (embed !== undefined) return embed;
    stack.push(...element.children);
  }
  return undefined;
}

function runProperties(run: XmlElement, paragraphProperties: XmlElement[]): XmlElement[] {
  const direct = firstChild(run, "rPr");
  return [
    ...(direct === undefined ? [] : [direct]),
    ...paragraphProperties.flatMap((paragraph) => {
      const defaults = firstChild(paragraph, "defRPr");
      return defaults === undefined ? [] : [defaults];
    }),
  ];
}

function parseRun(
  run: XmlElement,
  theme: DrawingTheme,
  paragraphProperties: XmlElement[],
  paragraphText: string,
  fontRefs?: XmlElement[],
  links?: Map<string, string>,
  autoColor?: string,
): TextRunModel {
  const properties = runProperties(run, paragraphProperties);
  const sizeElement = properties.find((property) => attr(property, "sz") !== undefined);
  const size = sizeElement === undefined ? undefined : numberAttr(sizeElement, "sz", "a:rPr");
  const solidFill = firstElement(properties, "solidFill");
  const latin = firstElement(properties, "latin");
  const cs = firstElement(properties, "cs");
  const ea = firstElement(properties, "ea");
  const sym = firstElement(properties, "sym");
  const t = firstChild(run, "t");
  const out: TextRunModel = { text: t === undefined ? "" : textContent(t) };
  const boldElement = properties.find((property) => attr(property, "b") !== undefined);
  const italicElement = properties.find((property) => attr(property, "i") !== undefined);
  const bold = boldElement === undefined ? undefined : onOff(boldElement, "b");
  const italic = italicElement === undefined ? undefined : onOff(italicElement, "i");
  const language = firstAttribute(properties, "lang");
  if (bold !== undefined) out.bold = bold;
  if (italic !== undefined) out.italic = italic;
  if (language !== undefined) out.language = language;
  if (size !== undefined) out.fontSizePt = size / 100;
  const latinTypeface = latin === undefined ? undefined : attr(latin, "typeface");
  const complexScriptTypeface = cs === undefined ? undefined : attr(cs, "typeface");
  const eastAsianTypeface = ea === undefined ? undefined : attr(ea, "typeface");
  const symbolTypeface = sym === undefined ? undefined : attr(sym, "typeface");
  if (latinTypeface !== undefined) out.latinTypeface = latinTypeface;
  if (complexScriptTypeface !== undefined) out.complexScriptTypeface = complexScriptTypeface;
  if (eastAsianTypeface !== undefined) out.eastAsianTypeface = eastAsianTypeface;
  if (symbolTypeface !== undefined) out.symbolTypeface = symbolTypeface;
  const color = resolveColor(solidFill, theme)
    ?? fontRefs?.map((ref) => resolveDrawingColor(ref, theme)).find((value): value is string => value !== undefined);
  // `p:defaultTextStyle` intentionally does not interpose: the coverage
  // fixture's dark-slide titles carry `tx1` there yet PowerPoint paints them
  // white — "Automatic" contrast wins for fill-less runs.
  if (color !== undefined) out.color = color;
  else if (autoColor !== undefined) out.color = autoColor;
  // Underline: `u="none"` (or absent) means none; any other token underlines.
  // Explicit `uFill`/`uLn` colors resolve, else the line follows the run.
  const underlineToken = firstAttribute(properties, "u");
  if (underlineToken !== undefined && underlineToken !== "none") {
    out.underline = true;
    out.underlineStyle = underlineToken;
    const uFill = firstElement(properties, "uFill");
    const uLn = firstElement(properties, "uLn");
    const uLnFill = uLn === undefined ? undefined : firstChild(uLn, "solidFill");
    const underlineColor = (uFill === undefined ? undefined : resolveDrawingColor(uFill, theme))
      ?? (uLnFill === undefined ? undefined : resolveColor(uLnFill, theme));
    if (underlineColor !== undefined) out.underlineColor = underlineColor;
  }
  const strikeToken = firstAttribute(properties, "strike");
  if (strikeToken === "sngStrike" || strikeToken === "dblStrike") out.strikethrough = true;
  const capToken = firstAttribute(properties, "cap");
  if (capToken === "all") out.caps = "all";
  else if (capToken === "small") out.caps = "small";
  const spcRaw = firstAttribute(properties, "spc");
  if (spcRaw !== undefined) {
    const spc = Number(spcRaw);
    // `a:rPr/@spc` is hundredths of a point (same scale as `@sz`); needs a
    // ground-truth probe against native tracking before trusting sub-point values.
    if (Number.isFinite(spc)) out.letterSpacingPt = spc / 100;
  }
  const baselineRaw = firstAttribute(properties, "baseline");
  if (baselineRaw !== undefined) {
    const baseline = Number(baselineRaw);
    // `a:rPr/@baseline` is thousandths of a percent (100000 = 100%).
    if (Number.isFinite(baseline) && baseline !== 0) out.baselinePct = baseline / 1000;
  }
  const highlight = firstElement(properties, "highlight");
  const highlightColor = highlight === undefined ? undefined : resolveDrawingColor(highlight, theme);
  if (highlightColor !== undefined) out.highlight = highlightColor;
  if (firstElement(properties, "noFill") !== undefined) out.noFill = true;
  if (firstElement(properties, "grpFill") !== undefined) out.groupFill = true;
  const hlink = firstElement(properties, "hlinkClick");
  if (hlink !== undefined && links !== undefined) {
    const rid = attr(hlink, "id") ?? attr(hlink, "r:id");
    if (rid !== undefined) {
      const target = links.get(rid);
      if (target !== undefined) out.hyperlink = target;
    }
  }
  // WordArt text outline (`a:ln` on the run): the stroke width is in EMUs.
  const runLine = firstElement(properties, "ln");
  if (runLine !== undefined) {
    const outlineColor = resolveSolidFill(runLine, theme);
    if (outlineColor !== undefined) {
      out.outline = {
        color: outlineColor,
        widthPt: emuToPt(numberAttr(runLine, "w", "a:ln") ?? DEFAULT_LINE_WIDTH_EMU),
      };
    }
  }
  // WordArt text gradient fill. Linear gradients paint; path gradients (and
  // other failures) resolve to the marked `unsupported` kind — same contract
  // as shape fills, never silently approximated.
  const runGradient = firstElement(properties, "gradFill");
  if (runGradient !== undefined) out.gradient = resolveGradientFill(runGradient, theme);
  const effectList = firstElement(properties, "effectLst");
  if (effectList !== undefined) {
    const glow = firstChild(effectList, "glow");
    if (glow !== undefined) {
      const glowColor = resolveDrawingColor(glow, theme);
      if (glowColor !== undefined) {
        out.glow = { color: glowColor, radiusPt: emuToPt(numberAttr(glow, "rad", "a:glow") ?? 0) };
      }
    }
    const shadow = firstChild(effectList, "outerShdw") ?? firstChild(effectList, "innerShdw") ?? firstChild(effectList, "prstShdw");
    if (shadow !== undefined) {
      const shadowColor = resolveDrawingColor(shadow, theme);
      if (shadowColor !== undefined) {
        // `blurRad`/`dist` are EMUs; `dir` is 60000ths of a degree,
        // clockwise from 3 o'clock in screen coordinates (y down).
        const blurPt = emuToPt(numberAttr(shadow, "blurRad", "a:outerShdw") ?? 0);
        const distPt = emuToPt(numberAttr(shadow, "dist", "a:outerShdw") ?? 0);
        const dirDeg = (numberAttr(shadow, "dir", "a:outerShdw") ?? 0) / 60000;
        const dirRad = (dirDeg * Math.PI) / 180;
        out.shadow = {
          color: shadowColor,
          blurPt,
          dxPt: distPt * Math.cos(dirRad),
          dyPt: distPt * Math.sin(dirRad),
        };
      }
    }
  }
  // Whitespace follows its paragraph's script slot, so inherited runs do not
  // lose their family merely because their own text has no script characters.
  const fontText = out.text.trim() === "" ? paragraphText : out.text;
  const preferredTypeface = usesArabic(fontText)
    ? complexScriptTypeface ?? latinTypeface ?? "+mn-cs"
    : latinTypeface ?? complexScriptTypeface ?? "+mn-lt";
  const fontFamily = resolveThemeTypeface(preferredTypeface, fontText, theme);
  if (fontFamily !== undefined) out.fontFamily = fontFamily;
  return out;
}

/** The live field kinds aared evaluates; other `a:fld` types keep their cached text as plain text. */
function fieldKind(type: string | undefined): TextRunModel["field"] {
  if (type === "slidenum") return "slidenum";
  if (type !== undefined && type.startsWith("datetime")) return "datetime";
  return undefined;
}

interface TextBodyInheritance {
  /** Text bodies in decreasing precedence after the slide's own body. */
  bodies?: XmlElement[];
  /** The applicable p:txStyles entry from the slide master. */
  masterStyle?: XmlElement;
  /**
   * Shape style font colors (`p:style/a:fontRef`) in decreasing precedence:
   * own shape, layout placeholder shape, master placeholder shape.
   * Applied only when no direct/`defRPr` solid fill resolves.
   */
  fontRefs?: XmlElement[];
  /** Hyperlink targets by `r:id`, resolved through the owning part's rels. */
  links?: Map<string, string>;
  /**
   * PowerPoint "Automatic" text color for runs with no resolved fill: white
   * on dark slide backgrounds, black otherwise. Never left undefined into
   * CSS inheritance (which would follow embedding chrome, not the file).
   */
  autoColor?: string;
  /** Picture-bullet bytes by `r:id`: media part target + content type. */
  embeds?: Map<string, { mediaId: string; contentType: string }>;
  /**
   * Deck default run size per level (`PptxDeck.defaultTextSizes`), for text
   * that inherits no placeholder or master style; used only when nothing in
   * the paragraph's own chain states a size.
   */
  defaultSizes?: ReadonlyArray<number | undefined>;
}

function levelProperties(body: XmlElement, level: number): XmlElement | undefined {
  const listStyle = firstChild(body, "lstStyle");
  return listStyle === undefined ? undefined : firstChild(listStyle, `lvl${level + 1}pPr`);
}

function paragraphProperties(paragraph: XmlElement, bodies: XmlElement[], masterStyle: XmlElement | undefined): XmlElement[] {
  const direct = firstChild(paragraph, "pPr");
  const directLevel = direct === undefined ? undefined : numberAttr(direct, "lvl", "a:pPr");
  const level = directLevel ?? 0;
  return [
    ...(direct === undefined ? [] : [direct]),
    ...bodies.flatMap((body) => {
      const defaults = levelProperties(body, level);
      return defaults === undefined ? [] : [defaults];
    }),
    ...(masterStyle === undefined ? [] : [firstChild(masterStyle, `lvl${level + 1}pPr`)].filter((value): value is XmlElement => value !== undefined)),
  ];
}

function spacingValue(properties: XmlElement[], name: "lnSpc" | "spcBef" | "spcAft"): { percent?: number; points?: number } | undefined {
  const spacing = firstElement(properties, name);
  if (spacing === undefined) return undefined;
  const percent = firstChild(spacing, "spcPct");
  const points = firstChild(spacing, "spcPts");
  if (percent !== undefined) {
    const value = numberAttr(percent, "val", "a:spcPct");
    if (value !== undefined) return { percent: value / 100000 };
  }
  if (points !== undefined) {
    const value = numberAttr(points, "val", "a:spcPts");
    if (value !== undefined) return { points: value / 100 };
  }
  return undefined;
}

function toRoman(value: number, upper: boolean): string {
  if (value <= 0) return String(value);
  const table: Array<[number, string]> = [
    [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"],
    [100, "C"], [90, "XC"], [50, "L"], [40, "XL"],
    [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
  ];
  let remaining = Math.min(value, 3999);
  let out = "";
  for (const [numeric, glyph] of table) {
    while (remaining >= numeric) {
      out += glyph;
      remaining -= numeric;
    }
  }
  return upper ? out : out.toLowerCase();
}

function toAlpha(value: number, upper: boolean): string {
  if (value <= 0) return String(value);
  let remaining = value;
  let out = "";
  while (remaining > 0) {
    remaining -= 1;
    out = String.fromCharCode((upper ? 65 : 97) + (remaining % 26)) + out;
    remaining = Math.floor(remaining / 26);
  }
  return out;
}

function formatAutoNum(type: string, value: number): string {
  const normalized = type.toLowerCase();
  const isRomanUc = normalized.includes("roman") && (normalized.includes("uc") || normalized.includes("upper") || type.includes("Uc"));
  const isRomanLc = normalized.includes("roman") && !isRomanUc;
  const isAlphaUc = normalized.includes("alpha") && (normalized.includes("uc") || normalized.includes("upper") || type.includes("Uc"));
  const isAlphaLc = normalized.includes("alpha") && !isAlphaUc;
  let core: string;
  if (isRomanUc) core = toRoman(value, true);
  else if (isRomanLc) core = toRoman(value, false);
  else if (isAlphaUc) core = toAlpha(value, true);
  else if (isAlphaLc) core = toAlpha(value, false);
  else core = String(value);
  if (normalized.includes("parenboth") || normalized.includes("parenboth")) return `(${core})`;
  if (normalized.includes("parenr") || normalized.includes("paren")) return `${core})`;
  if (normalized.includes("plain")) return core;
  // Default: period form ("1.", "I.", "A.", …) covers arabicPeriod and unknown types.
  return `${core}.`;
}

export function parseTextBody(txBody: XmlElement, theme: DrawingTheme, inheritance?: TextBodyInheritance): TextBodyModel {
  const bodies = [txBody, ...(inheritance?.bodies ?? [])];
  // Anchor, wrapping, and autofit retain the existing highest-bodyPr behavior.
  // Insets are resolved independently below, as PowerPoint does for
  // placeholders. Ground truth: tools/probe/title-insets.
  const bodyProperties = bodies.map((body) => firstChild(body, "bodyPr")).find((properties): properties is XmlElement => properties !== undefined
    && (Object.keys(properties.attrs).length > 0 || properties.children.length > 0));
  // Insets inherit independently: a layout can override the anchor while the
  // master still supplies the text-frame margins. Ground truth:
  // tools/probe/title-insets.
  const bodyPropertySources = bodies
    .map((body) => firstChild(body, "bodyPr"))
    .filter((properties): properties is XmlElement => properties !== undefined);
  const paragraphElements = childElements(txBody, "p");
  const autoNumCounters = new Map<string, number>();
  const out: TextBodyModel = {
    paragraphs: paragraphElements.map((paragraph): TextParagraphModel => {
      const properties = paragraphProperties(paragraph, bodies, inheritance?.masterStyle);
      const levelElement = properties.find((property) => attr(property, "lvl") !== undefined);
      const rtlElement = properties.find((property) => attr(property, "rtl") !== undefined);
      const alignment = align(firstAttribute(properties, "algn"));
      const level = levelElement === undefined ? undefined : numberAttr(levelElement, "lvl", "a:pPr");
      const leftElement = properties.find((property) => attr(property, "marL") !== undefined);
      const rightElement = properties.find((property) => attr(property, "marR") !== undefined);
      const firstLineElement = properties.find((property) => attr(property, "indent") !== undefined);
      const leftIndent = leftElement === undefined ? undefined : numberAttr(leftElement, "marL", "a:pPr");
      const rightIndent = rightElement === undefined ? undefined : numberAttr(rightElement, "marR", "a:pPr");
      const firstLineIndent = firstLineElement === undefined ? undefined : numberAttr(firstLineElement, "indent", "a:pPr");
      // `a:fld` (slide number, date) carries the same rPr/t as `a:r`; keep
      // both in document order.
      const runs = paragraph.children.filter((child) => child.local === "r" || child.local === "fld");
      const paragraphText = runs.map((run) => textContent(firstChild(run, "t") ?? run)).join("");
      const out: TextParagraphModel = {
        runs: runs.map((run) => {
          const model = parseRun(run, theme, properties, paragraphText, inheritance?.fontRefs, inheritance?.links, inheritance?.autoColor);
          const field = run.local === "fld" ? fieldKind(attr(run, "type")) : undefined;
          return field === undefined ? model : { ...model, field };
        }),
      };
      const defaultRunProperties = properties
        .map((property) => firstChild(property, "defRPr"))
        .filter((property): property is XmlElement => property !== undefined);
      const defaultSize = defaultRunProperties.find((property) => attr(property, "sz") !== undefined);
      if (defaultSize !== undefined) out.fontSizePt = numberAttr(defaultSize, "sz", "a:defRPr")! / 100;
      else {
        const deckSize = inheritance?.defaultSizes?.[level ?? 0];
        if (deckSize !== undefined) out.fontSizePt = deckSize;
      }
      const rtl = rtlElement === undefined ? undefined : onOff(rtlElement, "rtl");
      if (rtl !== undefined) out.direction = rtl ? "rtl" : "ltr";
      if (alignment !== undefined) out.align = alignment;
      if (level !== undefined) out.level = level;
      if (leftIndent !== undefined) out.leftIndent = emuToInches(leftIndent);
      if (rightIndent !== undefined) out.rightIndent = emuToInches(rightIndent);
      if (firstLineIndent !== undefined) out.firstLineIndent = emuToInches(firstLineIndent);
      const lineSpacing = spacingValue(properties, "lnSpc");
      const spaceBefore = spacingValue(properties, "spcBef");
      const spaceAfter = spacingValue(properties, "spcAft");
      if (lineSpacing?.percent !== undefined) out.lineSpacing = lineSpacing.percent;
      if (lineSpacing?.points !== undefined) out.lineSpacingPt = lineSpacing.points;
      if (spaceBefore?.points !== undefined) out.spaceBeforePt = spaceBefore.points;
      if (spaceAfter?.points !== undefined) out.spaceAfterPt = spaceAfter.points;
      const bulletProperties = properties.find((property) => firstChild(property, "buNone") !== undefined
        || firstChild(property, "buChar") !== undefined || firstChild(property, "buAutoNum") !== undefined
        || firstChild(property, "buBlip") !== undefined);
      const hasBuNone = bulletProperties !== undefined && firstChild(bulletProperties, "buNone") !== undefined;
      const autoNumElement = hasBuNone ? undefined : firstElement(properties, "buAutoNum");
      const buChar = bulletProperties === undefined || hasBuNone || autoNumElement !== undefined
        ? undefined
        : firstElement(properties, "buChar");
      const buBlip = bulletProperties === undefined || hasBuNone || autoNumElement !== undefined
        ? undefined
        : firstElement(properties, "buBlip");
      // Shared bullet paint: explicit `buClr` color, `buSzPct`/`buSzPts`
      // size; absent groups mean follow-text (`buClrTx`/`buSzTx`/default).
      const bulletExtras: NonNullable<TextParagraphModel["bullet"]> = {};
      const buClr = firstElement(properties, "buClr");
      const buColor = buClr === undefined ? undefined : resolveDrawingColor(buClr, theme);
      if (buColor !== undefined) bulletExtras.color = buColor;
      const buSzPct = firstElement(properties, "buSzPct");
      const buSzPctVal = buSzPct === undefined ? undefined : numberAttr(buSzPct, "val", "a:buSzPct");
      // `a:buSzPct/@val` is thousandths of a percent (100000 = 100%).
      if (buSzPctVal !== undefined) bulletExtras.sizePct = buSzPctVal / 1000;
      const buSzPts = firstElement(properties, "buSzPts");
      const buSzPtsVal = buSzPts === undefined ? undefined : numberAttr(buSzPts, "val", "a:buSzPts");
      // `a:buSzPts/@val` is hundredths of a point (same scale as `@sz`);
      // needs a ground-truth probe before trusting sub-point values.
      if (buSzPtsVal !== undefined) bulletExtras.sizePts = buSzPtsVal / 100;
      if (autoNumElement !== undefined) {
        // Numbered lists resolve to their display string in the model so the
        // DOM renderer paints them through the same bullet path as buChar.
        // Counters are per numbering type and indentation level: a new type
        // starts a new list (ground truth: the coverage bullets slide shows
        // arabic 1. 2. 3. then alpha a) b), not d) e)). Deeper levels reset
        // when a shallower numbered paragraph follows.
        const paragraphLevel = level ?? 0;
        const autoNumType = attr(autoNumElement, "type") ?? "arabicPeriod";
        const counterKey = `${autoNumType}@${paragraphLevel}`;
        const startAtRaw = attr(autoNumElement, "startAt");
        const startAt = startAtRaw === undefined ? undefined : Number(startAtRaw);
        let count: number;
        if (startAt !== undefined && Number.isFinite(startAt)) {
          count = Math.trunc(startAt);
          autoNumCounters.set(counterKey, count);
        } else {
          count = (autoNumCounters.get(counterKey) ?? 0) + 1;
          autoNumCounters.set(counterKey, count);
        }
        for (const key of [...autoNumCounters.keys()]) {
          const at = key.lastIndexOf("@");
          const keyLevel = at === -1 ? 0 : Number(key.slice(at + 1));
          if (Number.isFinite(keyLevel) && keyLevel > paragraphLevel) autoNumCounters.delete(key);
        }
        const buFontForNum = firstElement(properties, "buFont");
        const numTypeface = buFontForNum === undefined ? undefined : attr(buFontForNum, "typeface");
        out.bullet = {
          char: formatAutoNum(autoNumType, count),
          ...(numTypeface === undefined ? {} : { typeface: numTypeface }),
          ...bulletExtras,
        };
      } else if (buBlip !== undefined) {
        // Picture bullets resolve their embed through the owning part's
        // rels (threaded as `embeds`); the renderer paints the image with a
        // character fallback when fetch fails or no resolver is wired.
        const rid = findBulletEmbed(buBlip);
        const picture = rid === undefined ? undefined : inheritance?.embeds?.get(rid);
        const buCharText = buChar === undefined ? undefined : attr(buChar, "char");
        out.bullet = {
          ...(buCharText === undefined ? {} : { char: buCharText }),
          ...bulletExtras,
          ...(picture === undefined ? {} : {
            pictureMediaId: picture.mediaId,
            pictureContentType: picture.contentType,
          }),
          ...(picture === undefined && rid !== undefined ? { char: BULLET } : {}),
        };
        if (out.bullet !== undefined && Object.keys(out.bullet).length === 0) out.bullet = { char: BULLET };
      } else {
        const buFont = buChar === undefined ? undefined : firstElement(properties, "buFont");
        const char = buChar === undefined ? undefined : attr(buChar, "char");
        const typeface = buFont === undefined ? undefined : attr(buFont, "typeface");
        if (char !== undefined || typeface !== undefined || Object.keys(bulletExtras).length > 0) {
          out.bullet = {
            ...(char === undefined ? {} : { char }),
            ...(typeface === undefined ? {} : { typeface }),
            ...bulletExtras,
          };
        }
      }
      return out;
    }),
  };
  const insets = {
    left: firstAttribute(bodyPropertySources, "lIns"),
    right: firstAttribute(bodyPropertySources, "rIns"),
    top: firstAttribute(bodyPropertySources, "tIns"),
    bottom: firstAttribute(bodyPropertySources, "bIns"),
  };
  if (Object.values(insets).some((value) => value !== undefined)) {
    out.insets = {
      ...(insets.left === undefined ? {} : { left: emuToInches(parseNumber(insets.left, "lIns", "a:bodyPr")) }),
      ...(insets.right === undefined ? {} : { right: emuToInches(parseNumber(insets.right, "rIns", "a:bodyPr")) }),
      ...(insets.top === undefined ? {} : { top: emuToInches(parseNumber(insets.top, "tIns", "a:bodyPr")) }),
      ...(insets.bottom === undefined ? {} : { bottom: emuToInches(parseNumber(insets.bottom, "bIns", "a:bodyPr")) }),
    };
  }
  switch (firstAttribute(bodyProperties === undefined ? [] : [bodyProperties], "anchor")) {
    case "t":
      out.verticalAlign = "top";
      break;
    case "ctr":
      out.verticalAlign = "middle";
      break;
    case "b":
      out.verticalAlign = "bottom";
      break;
    case "just":
      out.verticalAlign = "justify";
      break;
    case "dist":
      out.verticalAlign = "distributed";
      break;
  }
  switch (firstAttribute(bodyProperties === undefined ? [] : [bodyProperties], "wrap")) {
    case "none":
      out.wrapping = "none";
      break;
    case "square":
      out.wrapping = "square";
      break;
  }
  if (bodyProperties !== undefined && (firstChild(bodyProperties, "noAutofit") !== undefined
    || firstChild(bodyProperties, "normAutofit") !== undefined || firstChild(bodyProperties, "spAutoFit") !== undefined)) {
    if (firstChild(bodyProperties, "noAutofit") !== undefined) out.autofit = "none";
    else if (firstChild(bodyProperties, "normAutofit") !== undefined) out.autofit = "shrinkText";
    else out.autofit = "resizeShape";
  }
  // WordArt preset warp lives on the body's own properties (whole-body
  // transform). Painted unwarped and marked — see `TextBodyModel.textWarp`.
  if (bodyProperties !== undefined) {
    const warp = firstChild(bodyProperties, "prstTxWarp");
    const preset = warp === undefined ? undefined : attr(warp, "prst");
    if (preset !== undefined) out.textWarp = preset;
  }
  return out;
}

function shapeIdentity(shape: XmlElement): { id: string; name: string; placeholder?: NonNullable<ShapeModel["placeholder"]> } {
  const nvSpPr = firstChild(shape, "nvSpPr");
  const cNvPr = nvSpPr === undefined ? undefined : firstChild(nvSpPr, "cNvPr");
  const nvPr = nvSpPr === undefined ? undefined : firstChild(nvSpPr, "nvPr");
  const ph = nvPr === undefined ? undefined : firstChild(nvPr, "ph");
  const id = cNvPr === undefined ? undefined : attr(cNvPr, "id");
  if (id === undefined) throw new Error("aared/core: a shape is missing p:cNvPr/@id");
  const out: { id: string; name: string; placeholder?: NonNullable<ShapeModel["placeholder"]> } = {
    id,
    name: cNvPr === undefined ? "" : attr(cNvPr, "name") ?? "",
  };
  if (ph !== undefined) out.placeholder = { type: attr(ph, "type") ?? "body", index: attr(ph, "idx") ?? "0" };
  return out;
}

function shapeBounds(shape: XmlElement): BoundsModel | undefined {
  const spPr = firstChild(shape, "spPr");
  return parseBounds(spPr === undefined ? undefined : firstChild(spPr, "xfrm"));
}

function localGeometrySize(xfrm: XmlElement | undefined): { width?: number; height?: number } | undefined {
  const ext = xfrm === undefined ? undefined : firstChild(xfrm, "ext");
  if (ext === undefined) return undefined;
  const width = numberAttr(ext, "cx", "a:ext");
  const height = numberAttr(ext, "cy", "a:ext");
  return {
    ...(width === undefined ? {} : { width }),
    ...(height === undefined ? {} : { height }),
  };
}

function parseAdjustments(prstGeom: XmlElement | undefined): Record<string, number> {
  const values: Record<string, number> = {};
  const avLst = prstGeom === undefined ? undefined : firstChild(prstGeom, "avLst");
  if (avLst === undefined) return values;
  for (const guide of childElements(avLst, "gd")) {
    const name = attr(guide, "name");
    const formula = attr(guide, "fmla");
    const match = formula === undefined ? null : /^val\s+(-?(?:\d+(?:\.\d*)?|\.\d+))$/.exec(formula);
    if (name !== undefined && match !== null) values[name] = Number(match[1]);
  }
  return values;
}

function matchingPlaceholder(root: XmlElement, target: NonNullable<ShapeModel["placeholder"]>): XmlElement | undefined {
  const cSld = firstChild(root, "cSld");
  const tree = cSld === undefined ? undefined : firstChild(cSld, "spTree");
  if (tree === undefined) return undefined;
  return childElements(tree, "sp").find((shape) => {
    const placeholder = shapeIdentity(shape).placeholder;
    if (placeholder === undefined || placeholder.index !== target.index) return false;
    // PowerPoint maps a centered-title layout/slide placeholder to the title
    // placeholder on its master when their default indices agree.
    return placeholder.type === target.type || (target.type === "ctrTitle" && placeholder.type === "title");
  });
}

/**
 * The master's placeholder types: a master has one `title`, one `body`, and
 * the `dt`/`ftr`/`sldNum` trio. Every other slide placeholder type (subtitle,
 * content, table, chart, picture, …) inherits from the master body.
 */
function masterPlaceholderType(type: string): string {
  switch (type) {
    case "ctrTitle":
    case "title":
      return "title";
    case "dt":
    case "ftr":
    case "sldNum":
      return type;
    default:
      return "body";
  }
}

/**
 * Layout -> master matching is by placeholder *type*: a master numbers its
 * placeholders on its own (date 2, footer 3, slide number 4), while layouts and
 * slides use 10/11/12, so comparing indices never finds them. Ground truth:
 * PowerPoint places the slide number at the master's position; see
 * `test/pptx-fields.test.ts`.
 */
function matchingMasterPlaceholder(root: XmlElement, target: NonNullable<ShapeModel["placeholder"]>): XmlElement | undefined {
  const cSld = firstChild(root, "cSld");
  const tree = cSld === undefined ? undefined : firstChild(cSld, "spTree");
  if (tree === undefined) return undefined;
  const wanted = masterPlaceholderType(target.type);
  return childElements(tree, "sp").find((shape) => {
    const placeholder = shapeIdentity(shape).placeholder;
    return placeholder !== undefined && masterPlaceholderType(placeholder.type) === wanted;
  });
}

interface PlaceholderInheritance {
  bounds?: BoundsModel;
  text: TextBodyInheritance;
}

function masterTextStyle(master: XmlElement, placeholder: NonNullable<ShapeModel["placeholder"]>): XmlElement | undefined {
  const styles = firstChild(master, "txStyles");
  if (styles === undefined) return undefined;
  switch (placeholder.type) {
    case "title":
    case "ctrTitle":
      return firstChild(styles, "titleStyle");
    case "body":
      return firstChild(styles, "bodyStyle");
    default:
      return firstChild(styles, "otherStyle");
  }
}

async function placeholderInheritance(pkg: OpcPackage, slidePart: string, shape: XmlElement, signal?: AbortSignal): Promise<PlaceholderInheritance | undefined> {
  throwIfAborted(signal);
  const placeholder = shapeIdentity(shape).placeholder;
  if (placeholder === undefined) return undefined;
  const relationships = await pkg.relationships(slidePart, signal);
  throwIfAborted(signal);
  const layoutRel = relationships.find((relationship) => relationship.type.endsWith("/slideLayout"));
  if (layoutRel === undefined) return undefined;
  // Layout and master trees come from the per-package chain cache: parsing
  // them here again would re-inflate both parts for every placeholder.
  const chain = await layoutChain(pkg, layoutRel.target, signal);
  throwIfAborted(signal);
  const layoutShape = matchingPlaceholder(chain.layoutRoot, placeholder);
  const layoutBounds = layoutShape === undefined ? undefined : shapeBounds(layoutShape);
  const shapeFontRef = (source: XmlElement | undefined): XmlElement[] => {
    if (source === undefined) return [];
    const styleElement = firstChild(source, "style");
    if (styleElement === undefined) return [];
    const ref = firstChild(styleElement, "fontRef");
    return ref === undefined ? [] : [ref];
  };
  const masterRoot = chain.masterRoot;
  if (masterRoot === undefined) {
    return {
      ...(layoutBounds === undefined ? {} : { bounds: layoutBounds }),
      text: {
        bodies: layoutShape === undefined ? [] : [firstChild(layoutShape, "txBody")].filter((value): value is XmlElement => value !== undefined),
        fontRefs: shapeFontRef(layoutShape),
      },
    };
  }
  const masterShape = matchingMasterPlaceholder(masterRoot, placeholder);
  const masterBounds = masterShape === undefined ? undefined : shapeBounds(masterShape);
  const bounds = layoutBounds ?? masterBounds;
  const style = masterTextStyle(masterRoot, placeholder);
  return {
    ...(bounds === undefined ? {} : { bounds }),
    text: {
      bodies: [layoutShape, masterShape].flatMap((source) => {
        const text = source === undefined ? undefined : firstChild(source, "txBody");
        return text === undefined ? [] : [text];
      }),
      ...(style === undefined ? {} : { masterStyle: style }),
      fontRefs: [...shapeFontRef(layoutShape), ...shapeFontRef(masterShape)],
    },
  };
}

async function parseShape(
  pkg: OpcPackage,
  slidePart: string,
  shape: XmlElement,
  theme: DrawingTheme,
  signal?: AbortSignal,
  slideRels?: Array<{ id: string; type: string; target: string; targetMode: string }>,
  autoColor?: string,
  defaultSizes?: ReadonlyArray<number | undefined>,
): Promise<ShapeModel> {
  throwIfAborted(signal);
  const text = firstChild(shape, "txBody");
  const spPr = firstChild(shape, "spPr");
  const xfrm = spPr === undefined ? undefined : firstChild(spPr, "xfrm");
  const prstGeom = spPr === undefined ? undefined : firstChild(spPr, "prstGeom");
  const preset = prstGeom === undefined ? undefined : attr(prstGeom, "prst");
  const line = spPr === undefined ? undefined : firstChild(spPr, "ln");
  const style = firstChild(shape, "style");
  const inherited = shapeIdentity(shape).placeholder === undefined ? undefined : await placeholderInheritance(pkg, slidePart, shape, signal);
  throwIfAborted(signal);
  const bounds = shapeBounds(shape) ?? inherited?.bounds;
  // Own shape style first, then layout/master placeholder styles: direct
  // fills always win, these only fill runs with no resolved color.
  // Ground truth: slide 3 title resolves white via lt1 fontRef.
  const ownFontRef = style === undefined ? undefined : firstChild(style, "fontRef");
  const fontRefs = [
    ...(ownFontRef === undefined ? [] : [ownFontRef]),
    ...(inherited?.text.fontRefs ?? []),
  ];
  const links = slideRels === undefined || slideRels.length === 0 ? undefined : hyperlinkTargets(slideRels);
  const embeds = bulletEmbeds(slideRels, (target) => pkg.mediaContentType(target));
  const base = inherited?.text;
  const extra = {
    ...(fontRefs.length === 0 ? {} : { fontRefs }),
    ...(links === undefined ? {} : { links }),
    ...(autoColor === undefined ? {} : { autoColor }),
    ...(embeds === undefined ? {} : { embeds }),
    // Placeholders size from the master's text styles; everything else from
    // the deck default.
    ...(inherited === undefined && defaultSizes !== undefined ? { defaultSizes } : {}),
  };
  const inheritance = base === undefined
    ? (Object.keys(extra).length === 0 ? undefined : extra)
    : (Object.keys(extra).length === 0 ? base : { ...base, ...extra });
  // Whole-shape links (`p:cNvPr/a:hlinkClick`) make the entire shape
  // clickable; run-level links still win for their own runs at paint time.
  const nvSpPr = firstChild(shape, "nvSpPr");
  const cNvPr = nvSpPr === undefined ? undefined : firstChild(nvSpPr, "cNvPr");
  const shapeHlink = cNvPr === undefined ? undefined : firstChild(cNvPr, "hlinkClick");
  const shapeRid = shapeHlink === undefined ? undefined : attr(shapeHlink, "id");
  const shapeHyperlink = shapeRid === undefined ? undefined : links?.get(shapeRid);
  const out: ShapeModel = {
    kind: "shape",
    ...shapeIdentity(shape),
    ...(bounds === undefined ? {} : { bounds }),
    ...(text === undefined ? {} : { text: parseTextBody(text, theme, inheritance) }),
    ...(shapeHyperlink === undefined ? {} : { hyperlink: shapeHyperlink }),
  };
  const size = localGeometrySize(xfrm);
  const adjustments = parseAdjustments(prstGeom);
  // Hand-coded presets resolve here; everything else (the rest of the preset
  // catalog, custGeom) is described now and evaluated by the session.
  out.geometry = geometryForPreset(preset, size, adjustments)
    ?? unresolvedGeometry(size, preset, adjustments, spPr === undefined ? undefined : firstChild(spPr, "custGeom"));
  const fill = resolveShapeFill(spPr, style, theme, {
    relationships: slideRels ?? [],
    mediaContentType: (target) => pkg.mediaContentType(target),
    inherited: false,
  });
  if (fill !== undefined) out.fill = fill;
  const directLineHasNoFill = line !== undefined && firstChild(line, "noFill") !== undefined;
  const styleLine = style === undefined ? undefined : firstChild(style, "lnRef");
  const strokeColor = directLineHasNoFill ? undefined : resolveSolidFill(line, theme) ?? (line === undefined ? resolveStyleRef(styleLine, theme) : undefined);
  if (strokeColor !== undefined) {
    const strokeWidth = line === undefined ? undefined : numberAttr(line, "w", "a:ln");
    out.stroke = { color: strokeColor, width: emuToInches(strokeWidth ?? DEFAULT_LINE_WIDTH_EMU) };
  }
  if (xfrm !== undefined) {
    const rotation = numberAttr(xfrm, "rot", "a:xfrm");
    const flipHorizontal = onOff(xfrm, "flipH");
    const flipVertical = onOff(xfrm, "flipV");
    if (rotation !== undefined || flipHorizontal !== undefined || flipVertical !== undefined) {
      out.transform = {
        ...(rotation === undefined ? {} : { rotation: rotation / 60000 }),
        ...(flipHorizontal === undefined ? {} : { flipHorizontal }),
        ...(flipVertical === undefined ? {} : { flipVertical }),
      };
    }
  }
  throwIfAborted(signal);
  return out;
}

function parseXfrmTransform(xfrm: XmlElement | undefined): ChartModel["transform"] {
  if (xfrm === undefined) return undefined;
  const rotation = numberAttr(xfrm, "rot", "a:xfrm");
  const flipHorizontal = onOff(xfrm, "flipH");
  const flipVertical = onOff(xfrm, "flipV");
  if (rotation === undefined && flipHorizontal === undefined && flipVertical === undefined) return undefined;
  return {
    ...(rotation === undefined ? {} : { rotation: rotation / 60000 }),
    ...(flipHorizontal === undefined ? {} : { flipHorizontal }),
    ...(flipVertical === undefined ? {} : { flipVertical }),
  };
}

function findSrcRect(blipFill: XmlElement | undefined): XmlElement | undefined {
  if (blipFill === undefined) return undefined;
  const visit = (element: XmlElement): XmlElement | undefined => {
    if (element.local === "srcRect") return element;
    for (const child of element.children) {
      const found = visit(child);
      if (found !== undefined) return found;
    }
    return undefined;
  };
  return visit(blipFill);
}

/**
 * Parse a `p:pic` element to a resolved picture model. Returns undefined for
 * external or unresolvable images so one bad reference never breaks its slide.
 * Bounds, rotation/flips, and `a:srcRect` crops resolve here; bytes flow later
 * through the existing `MediaCollection.get()` contract.
 */
export function parsePicture(
  pic: XmlElement,
  relationships: Array<{ id: string; type: string; target: string; targetMode: string }>,
  pkg: OpcPackage,
): PictureModel | undefined {
  const nvPicPr = firstChild(pic, "nvPicPr");
  const cNvPr = nvPicPr === undefined ? undefined : firstChild(nvPicPr, "cNvPr");
  const id = cNvPr === undefined ? undefined : attr(cNvPr, "id");
  if (id === undefined) throw new Error("aared/core: a picture is missing p:cNvPr/@id");
  const name = cNvPr === undefined ? "" : attr(cNvPr, "name") ?? "";
  const blipFill = firstChild(pic, "blipFill");
  const blip = blipFill === undefined ? undefined : firstChild(blipFill, "blip");
  const embed = blip === undefined ? undefined : (attr(blip, "embed") ?? attr(blip, "r:embed"));
  if (embed === undefined) return undefined;
  const rel = relationships.find((candidate) => candidate.id === embed);
  if (rel === undefined || rel.targetMode === "External") return undefined;
  if (!rel.type.endsWith("/image")) return undefined;
  const mediaId = rel.target;
  const contentType = pkg.mediaContentType(mediaId);
  const spPr = firstChild(pic, "spPr");
  const xfrm = spPr === undefined ? undefined : firstChild(spPr, "xfrm");
  const bounds = parseBounds(xfrm);
  const out: PictureModel = {
    kind: "picture",
    id,
    name,
    ...(bounds === undefined ? {} : { bounds }),
    mediaId,
    contentType,
  };
  const srcRect = findSrcRect(blipFill);
  if (srcRect !== undefined) {
    const fraction = (name: string): number => {
      const raw = attr(srcRect, name);
      if (raw === undefined) return 0;
      const numeric = Number(raw);
      if (!Number.isFinite(numeric)) return 0;
      return Math.min(Math.max(numeric / 100000, 0), 1);
    };
    const left = fraction("l");
    const top = fraction("t");
    const right = fraction("r");
    const bottom = fraction("b");
    if (left !== 0 || top !== 0 || right !== 0 || bottom !== 0) {
      out.crop = { left, top, right, bottom };
    }
  }
  const transform = parseXfrmTransform(xfrm);
  if (transform !== undefined) out.transform = transform;
  return out;
}

/** List embedded-font parts for future resolver delivery (no font bytes decoded here). */
function listEmbeddedFonts(pkg: OpcPackage): string[] {
  return pkg.partNames().filter((name) => name.startsWith("ppt/fonts/"));
}

/** Extract slide notes as plain text from its `notesSlide` part, if present. */
async function parseSlideNotes(pkg: OpcPackage, slidePart: string, signal?: AbortSignal): Promise<string | undefined> {
  throwIfAborted(signal);
  const relationships = await pkg.relationships(slidePart, signal);
  throwIfAborted(signal);
  const notesRel = relationships.find((relationship) => relationship.type.endsWith("/notesSlide"));
  if (notesRel === undefined || notesRel.targetMode === "External") return undefined;
  let root: XmlElement;
  try {
    root = await pkg.readXml(notesRel.target, signal);
  } catch {
    return undefined;
  }
  throwIfAborted(signal);
  const cSld = firstChild(root, "cSld");
  const tree = cSld === undefined ? undefined : firstChild(cSld, "spTree");
  if (tree === undefined) return undefined;
  const bodies: string[] = [];
  for (const shape of childElements(tree, "sp")) {
    throwIfAborted(signal);
    const nvSpPr = firstChild(shape, "nvSpPr");
    const nvPr = nvSpPr === undefined ? undefined : firstChild(nvSpPr, "nvPr");
    const ph = nvPr === undefined ? undefined : firstChild(nvPr, "ph");
    // Notes text lives in the body placeholder; slide-image and slide-number
    // placeholders are layout chrome, not notes.
    if (ph === undefined || (attr(ph, "type") ?? "body") !== "body") continue;
    const txBody = firstChild(shape, "txBody");
    if (txBody === undefined) continue;
    const paragraphs = childElements(txBody, "p").map((paragraph) => textContent(paragraph).trim());
    const text = paragraphs.filter((line) => line !== "").join("\n").trim();
    if (text !== "") bodies.push(text);
  }
  throwIfAborted(signal);
  if (bodies.length === 0) return undefined;
  return bodies.join("\n\n");
}

async function parseChartFrame(
  pkg: OpcPackage,
  frame: XmlElement,
  relationships: Array<{ id: string; type: string; target: string; targetMode: string }>,
  theme: DrawingTheme,
  signal?: AbortSignal,
): Promise<ChartModel | undefined> {
  throwIfAborted(signal);
  const nvGraphicFramePr = firstChild(frame, "nvGraphicFramePr");
  const cNvPr = nvGraphicFramePr === undefined ? undefined : firstChild(nvGraphicFramePr, "cNvPr");
  const graphic = firstChild(frame, "graphic");
  const graphicData = graphic === undefined ? undefined : firstChild(graphic, "graphicData");
  const chartRef = graphicData === undefined ? undefined : firstChild(graphicData, "chart");
  if (chartRef === undefined) return undefined;
  const id = cNvPr === undefined ? undefined : attr(cNvPr, "id");
  if (id === undefined) throw new Error("aared/core: a chart is missing p:cNvPr/@id");
  const name = cNvPr === undefined ? "" : attr(cNvPr, "name") ?? "";
  const xfrm = firstChild(frame, "xfrm");
  const bounds = parseBounds(xfrm);
  const base: ChartModel = {
    kind: "chart",
    id,
    name,
    ...(bounds === undefined ? {} : { bounds }),
  };
  const transform = parseXfrmTransform(xfrm);
  if (transform !== undefined) base.transform = transform;
  const chartId = attr(chartRef, "id") ?? attr(chartRef, "r:id");
  if (chartId === undefined) return base;
  const rel = relationships.find((candidate) => candidate.id === chartId);
  if (rel === undefined || rel.targetMode === "External") return base;
  const out: ChartModel = { ...base, sourceChartId: rel.target };
  let chartRoot: XmlElement;
  try {
    chartRoot = await pkg.readXml(rel.target, signal);
  } catch {
    return out;
  }
  throwIfAborted(signal);
  const chartGraphic = parseChartGraphic(chartRoot, theme);
  throwIfAborted(signal);
  if (chartGraphic !== undefined) out.graphic = chartGraphic;
  return out;
}

/**
 * Cell insets in inches: authored `a:tcPr/@marL|marR|marT|marB`, else
 * DrawingML's defaults (0.1in left/right, 0.05in top/bottom) — what
 * PowerPoint paints for a cell that states none.
 */
function parseCellMargins(tcPr: XmlElement | undefined): InsetsModel {
  const side = (name: string, fallback: number): number =>
    emuToInches((tcPr === undefined ? undefined : numberAttr(tcPr, name, "a:tcPr")) ?? fallback);
  return { left: side("marL", 91440), right: side("marR", 91440), top: side("marT", 45720), bottom: side("marB", 45720) };
}

/** A cell's edges: each side the direct `a:tcPr` line when authored (an `a:noFill` one clears it), else the table style's. */
function resolveCellBorders(
  tcPr: XmlElement | undefined,
  styleEdges: ReturnType<typeof resolveTableStyleCell>["edges"],
  theme: DrawingTheme,
): TableCellModel["borders"] {
  const direct: Record<keyof typeof styleEdges, string> = {
    left: "lnL", right: "lnR", top: "lnT", bottom: "lnB", diagonalDown: "lnTlToBr", diagonalUp: "lnBlToTr",
  };
  const out: NonNullable<TableCellModel["borders"]> = {};
  for (const side of Object.keys(direct) as Array<keyof typeof direct>) {
    const own = parseTableEdge(tcPr === undefined ? undefined : firstChild(tcPr, direct[side]), theme);
    const edge = own !== undefined ? own : styleEdges[side];
    if (edge !== undefined && edge !== null) out[side] = edge;
  }
  return Object.keys(out).length === 0 ? undefined : out;
}

export function parseTable(
  frame: XmlElement,
  theme: DrawingTheme,
  styles: Map<string, PptxTableStyle>,
  slideRels?: Array<{ id: string; type: string; target: string; targetMode: string }>,
  autoColor?: string,
  /** Types a picture bullet's image; the slide path passes the package's, headless callers get the extension guess. */
  mediaContentType: (target: string) => string = guessMediaContentType,
  /** The deck's default run sizes (`PptxDeck.defaultTextSizes`) for cell text that states none. */
  defaultSizes?: ReadonlyArray<number | undefined>,
): TableModel | undefined {
  const nvGraphicFramePr = firstChild(frame, "nvGraphicFramePr");
  const cNvPr = nvGraphicFramePr === undefined ? undefined : firstChild(nvGraphicFramePr, "cNvPr");
  const graphic = firstChild(frame, "graphic");
  const graphicData = graphic === undefined ? undefined : firstChild(graphic, "graphicData");
  const table = graphicData === undefined ? undefined : firstChild(graphicData, "tbl");
  if (table === undefined) return undefined;
  const id = cNvPr === undefined ? undefined : attr(cNvPr, "id");
  if (id === undefined) throw new Error("aared/core: a table is missing p:cNvPr/@id");
  const tblPr = firstChild(table, "tblPr");
  const grid = firstChild(table, "tblGrid");
  const styleIdElement = tblPr === undefined ? undefined : firstChild(tblPr, "tableStyleId");
  const styleId = styleIdElement === undefined ? undefined : textContent(styleIdElement);
  const style = styleId === undefined ? undefined : styles.get(styleId);
  const flag = (name: string): boolean => tblPr === undefined ? false : onOff(tblPr, name) ?? false;
  const rtl = flag("rtl");
  const rowElements = childElements(table, "tr");
  const columns = grid === undefined ? [] : childElements(grid, "gridCol").map((column) => emuToInches(numberAttr(column, "w", "a:gridCol") ?? 0));
  const shape: TableShape = {
    rows: rowElements.length,
    columns: columns.length,
    firstRow: flag("firstRow"),
    lastRow: flag("lastRow"),
    firstCol: flag("firstCol"),
    lastCol: flag("lastCol"),
    bandRow: flag("bandRow"),
    bandCol: flag("bandCol"),
  };
  const bounds = parseBounds(firstChild(frame, "xfrm"));
  const xfrm = firstChild(frame, "xfrm");
  const fillOwner: FillOwner = { relationships: slideRels ?? [], mediaContentType, inherited: false };
  const cellFills: Array<Array<ShapeFillModel | undefined>> = [];
  const rows = rowElements.map((row, rowIndex) => {
    const height = numberAttr(row, "h", "a:tr");
    const cells: TableCellModel[] = [];
    const fills: Array<ShapeFillModel | undefined> = [];
    // Every grid column has a `tc`, merge continuations included, so a
    // cell's grid column is its position among its row's `tc` elements.
    for (const [columnIndex, cell] of childElements(row, "tc").entries()) {
      const tcPr = firstChild(cell, "tcPr");
      // Horizontal/vertical merge continuations are covered by their master
      // cell's span and carry no independent content. Skipping them keeps the
      // logical grid aligned with what PowerPoint paints.
      const gridSpan = tcPr === undefined ? undefined : numberAttr(tcPr, "gridSpan", "a:tcPr");
      const rowSpan = tcPr === undefined ? undefined : numberAttr(tcPr, "rowSpan", "a:tcPr");
      const hMerge = tcPr === undefined ? undefined : onOff(tcPr, "hMerge");
      const vMerge = tcPr === undefined ? undefined : onOff(tcPr, "vMerge");
      if ((hMerge === true || vMerge === true) && (gridSpan === undefined || gridSpan <= 1) && (rowSpan === undefined || rowSpan <= 1)) {
        continue;
      }
      const body = firstChild(cell, "txBody");
      if (body === undefined) throw new Error("aared/core: a table cell is missing a:txBody");
      const lastRowIndex = rowIndex + (rowSpan !== undefined && rowSpan > 1 ? rowSpan : 1) - 1;
      const lastColumnIndex = columnIndex + (gridSpan !== undefined && gridSpan > 1 ? gridSpan : 1) - 1;
      const styled = resolveTableStyleCell(style, shape, rowIndex, lastRowIndex, columnIndex, lastColumnIndex, rtl);
      const cellLinks = slideRels === undefined || slideRels.length === 0 ? undefined : hyperlinkTargets(slideRels);
      // Authored style text beats automatic contrast, so automatic only
      // enters the body when no style text applies to this cell.
      const cellAuto = styled.text === undefined ? autoColor : undefined;
      const cellEmbeds = bulletEmbeds(slideRels, mediaContentType);
      const cellInheritance = cellLinks === undefined && cellAuto === undefined && cellEmbeds === undefined && defaultSizes === undefined
        ? undefined
        : {
          ...(cellLinks === undefined ? {} : { links: cellLinks }),
          ...(cellAuto === undefined ? {} : { autoColor: cellAuto }),
          ...(cellEmbeds === undefined ? {} : { embeds: cellEmbeds }),
          ...(defaultSizes === undefined ? {} : { defaultSizes }),
        };
      const cellOut: TableCellModel = {
        text: parseTextBody(body, theme, cellInheritance),
      };
      // Table-style character formatting fills in what a run leaves unset;
      // anything the run authors keeps precedence. The style typeface only
      // reaches runs that name no typeface of their own.
      for (const paragraph of cellOut.text.paragraphs) {
        for (const run of paragraph.runs) {
          if (styled.bold !== undefined) run.bold ??= styled.bold;
          if (styled.italic !== undefined) run.italic ??= styled.italic;
          if (styled.text !== undefined) run.color ??= styled.text;
          if (styled.font !== undefined && run.latinTypeface === undefined && run.complexScriptTypeface === undefined) {
            const family = resolveThemeTypeface(styled.font, run.text, theme);
            if (family !== undefined) run.fontFamily = family;
          }
        }
      }
      const anchor = verticalAlign(tcPr === undefined ? undefined : attr(tcPr, "anchor"));
      if (anchor !== undefined) cellOut.verticalAlign = anchor;
      if (gridSpan !== undefined && gridSpan > 1) cellOut.columnSpan = gridSpan;
      if (rowSpan !== undefined && rowSpan > 1) cellOut.rowSpan = rowSpan;
      cellOut.margins = parseCellMargins(tcPr);
      const borders = resolveCellBorders(tcPr, styled.edges, theme);
      if (borders !== undefined) cellOut.borders = borders;
      const vert = tcPr === undefined ? undefined : attr(tcPr, "vert");
      if (vert === "vert" || vert === "eaVert" || vert === "wordArtVertRtl") cellOut.textRotation = 90;
      else if (vert === "vert270") cellOut.textRotation = 270;
      cells.push(cellOut);
      // Direct cell fills (any kind) win over the table style; an explicit
      // `a:noFill` leaves the cell unpainted.
      const direct = tcPr === undefined ? undefined : resolveDirectFill(tcPr, theme, fillOwner);
      const styleRef = resolveStyleRef(tcPr === undefined ? undefined : firstChild(tcPr, "fillRef"), theme);
      const styleFill: ShapeFillModel | undefined = styled.fill === undefined ? undefined : { kind: "solid", color: styled.fill };
      fills.push(direct === null ? undefined
        : direct !== undefined ? direct
          : styleRef !== undefined ? { kind: "solid", color: styleRef }
            : styleFill);
    }
    cellFills.push(fills);
    return {
      ...(height === undefined ? {} : { height: emuToInches(height) }),
      cells,
    };
  });
  const out: TableModel = {
    kind: "table",
    id,
    name: cNvPr === undefined ? "" : attr(cNvPr, "name") ?? "",
    ...(bounds === undefined ? {} : { bounds }),
    direction: rtl ? "rtl" : "ltr",
    columns,
    rows,
    cellFills,
  };
  if (xfrm !== undefined) {
    const rotation = numberAttr(xfrm, "rot", "a:xfrm");
    const flipHorizontal = onOff(xfrm, "flipH");
    const flipVertical = onOff(xfrm, "flipV");
    if (rotation !== undefined || flipHorizontal !== undefined || flipVertical !== undefined) {
      out.transform = {
        ...(rotation === undefined ? {} : { rotation: rotation / 60000 }),
        ...(flipHorizontal === undefined ? {} : { flipHorizontal }),
        ...(flipVertical === undefined ? {} : { flipVertical }),
      };
    }
  }
  return out;
}

async function slideParts(pkg: OpcPackage, presentationPart: string, signal?: AbortSignal): Promise<string[]> {
  throwIfAborted(signal);
  const root = await pkg.readXml(presentationPart, signal);
  const list = firstChild(root, "sldIdLst");
  if (list === undefined) return [];
  const relationships = await pkg.relationships(presentationPart, signal);
  throwIfAborted(signal);
  return childElements(list, "sldId").map((slide) => {
    throwIfAborted(signal);
    const id = attr(slide, "r:id");
    const relationship = relationships.find((candidate) => candidate.id === id && candidate.type.endsWith("/slide"));
    if (relationship === undefined || relationship.targetMode === "External") {
      throw new Error(`aared/core: presentation slide relationship "${id ?? "missing"}" is invalid`);
    }
    return relationship.target;
  });
}

/**
 * `bg1`/`bg2`/`tx1`/`tx2` are slide-level names that the master's `p:clrMap`
 * points at theme slots (`lt1`, `dk1`, …); dark-master decks swap them. The
 * default map is Office's light one.
 */
function themeForColorMap(theme: DrawingTheme, clrMap: XmlElement | undefined): DrawingTheme {
  const slot = (name: string, fallback: string): string | undefined => theme.colors[(clrMap === undefined ? undefined : attr(clrMap, name)) ?? fallback];
  const mapped: Record<string, string> = {};
  for (const [name, fallback] of [["bg1", "lt1"], ["bg2", "lt2"], ["tx1", "dk1"], ["tx2", "dk2"]] as const) {
    const color = slot(name, fallback);
    if (color !== undefined) mapped[name] = color;
  }
  return { ...theme, colors: { ...theme.colors, ...mapped } };
}

/**
 * A `p:bg` element as a fill. `bgPr` carries any fill kind, so it shares the
 * shape-fill reader; a picture is a relationship of the part that owns the
 * `p:bg` (slide, layout, or master), not necessarily the slide. `bgRef`
 * points at one of the theme's background fill styles (`bgFillStyleLst`),
 * recoloured with the reference's own colour.
 */
function readBackgroundFill(
  bg: XmlElement,
  theme: DrawingTheme,
  owner: { relationships: PartRelationships; pkg: OpcPackage; inherited: boolean },
): SlideModel["background"] {
  const properties = firstChild(bg, "bgPr");
  if (properties !== undefined) {
    const fill = resolveDirectFill(properties, theme, {
      relationships: owner.relationships,
      mediaContentType: (target) => owner.pkg.mediaContentType(target),
      inherited: owner.inherited,
    });
    return fill ?? undefined;
  }
  const reference = firstChild(bg, "bgRef");
  if (reference === undefined) return undefined;
  const color = resolveColor(reference, theme);
  if (color === undefined) return undefined;
  // `idx` 1001+ picks a theme background style; its `phClr` placeholder is the
  // reference's own colour, then the style's transforms (tint, satMod, …) apply.
  const style = theme.backgroundFills?.[(numberAttr(reference, "idx", "p:bgRef") ?? 0) - 1001];
  if (style === undefined || !/^#[0-9A-F]{6}$/.test(color)) return { kind: "solid", color };
  const styled = { ...theme, colors: { ...theme.colors, phClr: color } };
  if (style.local === "solidFill") {
    const solid = resolveColor(style, styled);
    return solid === undefined ? { kind: "solid", color } : { kind: "solid", color: solid };
  }
  if (style.local === "gradFill") return resolveGradientFill(style, styled);
  if (style.local === "blipFill") return { kind: "unsupported", source: "image" };
  return { kind: "solid", color };
}

/**
 * What a slide inherits through its layout and that layout's master. Shared
 * by every slide on the layout, so it is read once per package and reused.
 */
interface LayoutChain {
  layoutPart: string;
  layoutRoot: XmlElement;
  layoutRelationships: PartRelationships;
  masterPart?: string;
  masterRoot?: XmlElement;
  masterRelationships?: PartRelationships;
  /** The master's `p:clrMap` (slide-level colour names -> theme slots). */
  clrMap?: XmlElement;
  /** The layout's `p:bg`, else the master's. */
  bg?: XmlElement;
  /** Relationships of the part that owns `bg`, for a picture background. */
  bgRelationships?: PartRelationships;
}

type PartRelationships = Awaited<ReturnType<OpcPackage["relationships"]>>;

const layoutChains = new WeakMap<OpcPackage, Map<string, LayoutChain>>();

async function layoutChain(pkg: OpcPackage, layoutPart: string, signal?: AbortSignal): Promise<LayoutChain> {
  let byLayout = layoutChains.get(pkg);
  if (byLayout === undefined) {
    byLayout = new Map();
    layoutChains.set(pkg, byLayout);
  }
  const cached = byLayout.get(layoutPart);
  if (cached !== undefined) return cached;
  const backgroundOf = (root: XmlElement): XmlElement | undefined => {
    const cSld = firstChild(root, "cSld");
    return cSld === undefined ? undefined : firstChild(cSld, "bg");
  };
  const layoutRoot = await pkg.readXml(layoutPart, signal);
  const layoutRelationships = await pkg.relationships(layoutPart, signal);
  throwIfAborted(signal);
  const masterRel = layoutRelationships.find((relationship) => relationship.type.endsWith("/slideMaster"));
  const masterRoot = masterRel === undefined ? undefined : await pkg.readXml(masterRel.target, signal);
  const masterRelationships = masterRel === undefined ? undefined : await pkg.relationships(masterRel.target, signal);
  throwIfAborted(signal);
  const layoutBg = backgroundOf(layoutRoot);
  const bg = layoutBg ?? (masterRoot === undefined ? undefined : backgroundOf(masterRoot));
  const bgRelationships = layoutBg === undefined ? masterRelationships : layoutRelationships;
  const clrMap = masterRoot === undefined ? undefined : firstChild(masterRoot, "clrMap");
  const chain: LayoutChain = {
    layoutPart,
    layoutRoot,
    layoutRelationships,
    ...(masterRel === undefined ? {} : { masterPart: masterRel.target }),
    ...(masterRoot === undefined ? {} : { masterRoot }),
    ...(masterRelationships === undefined ? {} : { masterRelationships }),
    ...(bg === undefined ? {} : { bg }),
    ...(bgRelationships === undefined ? {} : { bgRelationships }),
    ...(clrMap === undefined ? {} : { clrMap }),
  };
  byLayout.set(layoutPart, chain);
  return chain;
}

/**
 * Parse one shape tree (`p:spTree`) into elements, back-to-front. Slides use
 * it for their own content; layouts and masters use it for shared graphics,
 * where placeholders are prototypes to skip and charts are out of scope.
 * Groups (`p:grpSp`) recurse: children stay in their authored child-space
 * coordinates and the group carries bounds/transform for the renderer.
 */
async function parseShapeTree(
  pkg: OpcPackage,
  part: string,
  tree: XmlElement,
  relationships: PartRelationships,
  deck: PptxDeck,
  theme: DrawingTheme,
  options: { graphicsOnly: boolean; autoColor?: string },
  signal?: AbortSignal,
): Promise<SlideModel["elements"]> {
  const isPlaceholder = (element: XmlElement): boolean => options.graphicsOnly && hasPlaceholder(element);
  const visit = async (child: XmlElement, target: SlideModel["elements"]): Promise<void> => {
    throwIfAborted(signal);
    const pushFallbackPicture = (container: XmlElement): void => {
      const fallback = firstChild(container, "Fallback");
      const pic = fallback === undefined ? undefined : firstChild(fallback, "pic");
      if (pic === undefined) return;
      const picture = parsePicture(pic, relationships, pkg);
      if (picture !== undefined) target.push(picture);
    };
    if (child.local === "sp") {
      if (isPlaceholder(child)) return;
      target.push(await parseShape(pkg, part, child, theme, signal, relationships, options.autoColor, deck.defaultTextSizes));
      throwIfAborted(signal);
      return;
    }
    if (child.local === "pic") {
      if (isPlaceholder(child)) return;
      const picture = parsePicture(child, relationships, pkg);
      if (picture !== undefined) target.push(picture);
      return;
    }
    if (child.local === "AlternateContent") {
      // Ink and other modern content surface their raster fallback here.
      pushFallbackPicture(child);
      return;
    }
    if (child.local === "grpSp") {
      const group = await parseGroup(child, async (nested, nestedTarget) => visit(nested, nestedTarget));
      if (group !== undefined) target.push(group);
      return;
    }
    if (child.local === "graphicFrame") {
      throwIfAborted(signal);
      const table = parseTable(child, theme, deck.tableStyles, relationships, options.autoColor, (name) => pkg.mediaContentType(name), deck.defaultTextSizes);
      if (table !== undefined) {
        target.push(table);
        return;
      }
      if (options.graphicsOnly) return;
      const chart = await parseChartFrame(pkg, child, relationships, theme, signal);
      throwIfAborted(signal);
      if (chart !== undefined) target.push(chart);
    }
  };
  const elements: SlideModel["elements"] = [];
  for (const child of tree.children) await visit(child, elements);
  return elements;
}

/** Child coordinate space of a group: `a:chOff` + `a:chExt` from `p:grpSpPr/a:xfrm`. */
function parseChildSpace(xfrm: XmlElement | undefined): BoundsModel | undefined {
  if (xfrm === undefined) return undefined;
  const chOff = firstChild(xfrm, "chOff");
  const chExt = firstChild(xfrm, "chExt");
  if (chOff === undefined || chExt === undefined) return undefined;
  const x = numberAttr(chOff, "x", "a:chOff");
  const y = numberAttr(chOff, "y", "a:chOff");
  const width = numberAttr(chExt, "cx", "a:chExt");
  const height = numberAttr(chExt, "cy", "a:chExt");
  if (x === undefined || y === undefined || width === undefined || height === undefined) return undefined;
  if (width <= 0 || height <= 0) return undefined;
  return { x: emuToInches(x), y: emuToInches(y), width: emuToInches(width), height: emuToInches(height) };
}

/**
 * One `p:grpSp` as a `GroupModel`: identity from `p:nvGrpSpPr/p:cNvPr`,
 * bounds/transform/child-space from `p:grpSpPr/a:xfrm`, then every child
 * type visited in document order into the group's own children array.
 */
async function parseGroup(
  group: XmlElement,
  visit: (child: XmlElement, target: SlideModel["elements"]) => Promise<void>,
): Promise<GroupModel | undefined> {
  const nvGrpSpPr = firstChild(group, "nvGrpSpPr");
  const cNvPr = nvGrpSpPr === undefined ? undefined : firstChild(nvGrpSpPr, "cNvPr");
  const id = cNvPr === undefined ? undefined : attr(cNvPr, "id");
  if (id === undefined) return undefined;
  const grpSpPr = firstChild(group, "grpSpPr");
  const xfrm = grpSpPr === undefined ? undefined : firstChild(grpSpPr, "xfrm");
  const bounds = parseBounds(xfrm);
  const transform = parseXfrmTransform(xfrm);
  const childSpace = parseChildSpace(xfrm);
  const children: SlideModel["elements"] = [];
  for (const child of group.children) await visit(child, children);
  return {
    kind: "group",
    id,
    name: cNvPr === undefined ? "" : attr(cNvPr, "name") ?? "",
    ...(bounds === undefined ? {} : { bounds }),
    ...(transform === undefined ? {} : { transform }),
    ...(childSpace === undefined ? {} : { childSpace }),
    children,
  };
}

/** A shape or picture that is a placeholder instance (`p:nvPr/p:ph`). */
function hasPlaceholder(element: XmlElement): boolean {
  const nonVisual = firstChild(element, "nvSpPr") ?? firstChild(element, "nvPicPr");
  const nvPr = nonVisual === undefined ? undefined : firstChild(nonVisual, "nvPr");
  return nvPr !== undefined && firstChild(nvPr, "ph") !== undefined;
}

function shapeTreeOf(root: XmlElement): XmlElement | undefined {
  const cSld = firstChild(root, "cSld");
  return cSld === undefined ? undefined : firstChild(cSld, "spTree");
}

const layoutModels = new WeakMap<OpcPackage, Map<string, SlideLayoutModel>>();

/**
 * The shared graphics of one layout and its master, parsed once per package.
 * A layout's `showMasterSp="0"` leaves the master's shapes out.
 */
export async function parsePptxLayout(pkg: OpcPackage, deck: PptxDeck, layoutPart: string, signal?: AbortSignal): Promise<SlideLayoutModel> {
  throwIfAborted(signal);
  let byLayout = layoutModels.get(pkg);
  if (byLayout === undefined) {
    byLayout = new Map();
    layoutModels.set(pkg, byLayout);
  }
  const cached = byLayout.get(layoutPart);
  if (cached !== undefined) return cached;
  const chain = await layoutChain(pkg, layoutPart, signal);
  const theme = themeForColorMap(deck.theme, chain.clrMap);
  const graphics: SlideLayoutModel["graphics"] = [];
  const showsMaster = onOff(chain.layoutRoot, "showMasterSp") !== false;
  const masterTree = chain.masterRoot === undefined ? undefined : shapeTreeOf(chain.masterRoot);
  if (showsMaster && masterTree !== undefined && chain.masterPart !== undefined && chain.masterRelationships !== undefined) {
    graphics.push(...(await parseShapeTree(pkg, chain.masterPart, masterTree, chain.masterRelationships, deck, theme, { graphicsOnly: true }, signal)));
  }
  const layoutTree = shapeTreeOf(chain.layoutRoot);
  if (layoutTree !== undefined) {
    graphics.push(...(await parseShapeTree(pkg, layoutPart, layoutTree, chain.layoutRelationships, deck, theme, { graphicsOnly: true }, signal)));
  }
  throwIfAborted(signal);
  const model: SlideLayoutModel = { id: layoutPart, graphics };
  byLayout.set(layoutPart, model);
  return model;
}

/** The slide's background: its own `p:bg`, else the layout's, else the master's. */
function resolveSlideBackground(
  pkg: OpcPackage,
  cSld: XmlElement | undefined,
  slideRelationships: PartRelationships,
  chain: LayoutChain | undefined,
  theme: DrawingTheme,
): SlideModel["background"] {
  const own = cSld === undefined ? undefined : firstChild(cSld, "bg");
  if (own !== undefined) return readBackgroundFill(own, theme, { relationships: slideRelationships, pkg, inherited: false });
  if (chain?.bg === undefined) return undefined;
  return readBackgroundFill(chain.bg, theme, { relationships: chain.bgRelationships ?? [], pkg, inherited: true });
}

/**
 * Index a PPTX without opening a single slide part. This is the performance
 * boundary: a 200-slide deck should be as cheap to open as a two-slide deck;
 * `parsePptxSlide()` inflates only the unit currently entering the viewport.
 */
export async function indexPptx(pkg: OpcPackage, signal?: AbortSignal): Promise<PptxDeck> {
  throwIfAborted(signal);
  const presentationPart = await pkg.officeDocument("pptx", signal);
  throwIfAborted(signal);
  const presentation = await pkg.readXml(presentationPart, signal);
  const sldSz = firstChild(presentation, "sldSz");
  const width = sldSz === undefined ? undefined : numberAttr(sldSz, "cx", "p:sldSz");
  const height = sldSz === undefined ? undefined : numberAttr(sldSz, "cy", "p:sldSz");
  if (width === undefined || height === undefined) throw new Error("aared/core: PPTX is missing p:sldSz");
  const theme = await readDrawingTheme(pkg, signal);
  throwIfAborted(signal);
  const parts = await slideParts(pkg, presentationPart, signal);
  throwIfAborted(signal);
  const tableStyles = await readTableStyles(pkg, theme, signal);
  throwIfAborted(signal);
  return {
    size: { width: emuToInches(width), height: emuToInches(height) },
    slideParts: parts,
    theme,
    tableStyles,
    embeddedFonts: listEmbeddedFonts(pkg),
    defaultTextSizes: defaultTextSizes(firstChild(presentation, "defaultTextStyle")),
  };
}

/** `p:defaultTextStyle`'s run size per outline level (`lvl1pPr` … `lvl9pPr`), in points. */
function defaultTextSizes(style: XmlElement | undefined): Array<number | undefined> {
  return Array.from({ length: 9 }, (_, level) => {
    const properties = style === undefined ? undefined : firstChild(style, `lvl${level + 1}pPr`);
    const run = properties === undefined ? undefined : firstChild(properties, "defRPr");
    const size = run === undefined ? undefined : numberAttr(run, "sz", "a:defRPr");
    return size === undefined ? undefined : size / 100;
  });
}

/** Inflate and parse exactly one slide previously listed by `indexPptx()`. */
export async function parsePptxSlide(pkg: OpcPackage, deck: PptxDeck, index: number, signal?: AbortSignal): Promise<SlideModel> {
  throwIfAborted(signal);
  const part = deck.slideParts[index];
  if (part === undefined) throw new AaredError("not-found", `aared/core: slide index ${index} is outside this ${deck.slideParts.length}-slide deck`);
  const root = await pkg.readXml(part, signal);
  const cSld = firstChild(root, "cSld");
  const tree = cSld === undefined ? undefined : firstChild(cSld, "spTree");
  if (tree === undefined) throw new Error(`aared/core: slide "${part}" is missing p:spTree`);
  const relationships = await pkg.relationships(part, signal);
  throwIfAborted(signal);
  const layoutRel = relationships.find((relationship) => relationship.type.endsWith("/slideLayout"));
  const chain = layoutRel === undefined ? undefined : await layoutChain(pkg, layoutRel.target, signal);
  throwIfAborted(signal);
  // Colours resolve through this slide's master colour map, not the bare theme.
  const theme = themeForColorMap(deck.theme, chain?.clrMap);
  // Background first: fill-less runs take PowerPoint "Automatic" contrast
  // against it (white on dark), never CSS inheritance from embedding chrome.
  const background = resolveSlideBackground(pkg, cSld, relationships, chain, theme);
  const autoColor = automaticTextColor(background);
  const elements = await parseShapeTree(pkg, part, tree, relationships, deck, theme, { graphicsOnly: false, autoColor }, signal);
  throwIfAborted(signal);
  const notes = await parseSlideNotes(pkg, part, signal);
  throwIfAborted(signal);
  // Only layouts with shared graphics need a fetch; most decks' empty layouts don't.
  const layout = chain === undefined ? undefined : await parsePptxLayout(pkg, deck, chain.layoutPart, signal);
  const layoutId = layout === undefined || layout.graphics.length === 0 ? undefined : layout.id;
  const hidesGraphics = onOff(root, "showMasterSp") === false;
  return {
    part,
    size: deck.size,
    elements,
    number: index + 1,
    ...(background === undefined ? {} : { background }),
    ...(layoutId === undefined ? {} : { layoutId }),
    ...(hidesGraphics ? { hideBackgroundGraphics: true } : {}),
    ...(notes === undefined ? {} : { notes }),
  };
}
