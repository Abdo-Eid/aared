/**
 * PPTX reader: the Worker-side half of `@aared/viewer/pptx`.
 *
 * Deliberately headless; it returns only plain data. It reads PresentationML
 * (slide order, layouts, masters, placeholders) and DrawingML (text, shapes,
 * pictures, charts, groups, tables, fills, theme and table styles) into the
 * resolved slide model; DOM/SVG painting belongs to `@aared/viewer/pptx`, never here.
 */

import { BULLET } from "./chars.js";
import { attr, childElements, firstChild, numberAttr, onOff, parseNumber, parseXml, textContent } from "./xml.js";
import { directFillProperty, resolveDirectFill, resolveGradientFill } from "./fill.js";
import type { FillOwner } from "./fill.js";
import type { XmlElement } from "./xml.js";
import { guessMediaContentType } from "./opc.js";
import type { OpcPackage } from "./opc.js";
import { EMU_PER_PT, emuToInches, emuToPt } from "./units.js";
import { clampSpan } from "./table-span.js";
import { TABLE_STYLE_REGIONS, tableCellRegions } from "./table-style.js";
import type { TableStyleFlags, TableStyleRegionName } from "./table-style.js";
import { parseHexColor, relativeLuminance } from "./color.js";
import { readDrawingTheme, readThemePart, resolveDrawingColor, resolveThemeTypeface, themeForColorMap } from "./theme.js";
import { arabicPunctuationPieces, isLtrRunInRtl, usesArabic } from "./typography.js";
import type { DrawingTheme } from "./theme.js";
import { AaredError, damagedDocument, throwIfAborted } from "./errors.js";
import { chartFillOwner, readChartModel } from "./chart.js";
import { chartText } from "./chart/text.js";
import type {
  BoundsModel,
  ChartModel,
  GroupModel,
  InsetsModel,
  PictureModel,
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
import { drawingGeometryAdjustments, drawingShapeTransform, geometryForPreset, unresolvedGeometry } from "./drawingml-shape.js";
import { symbolFontStandIn } from "./symbol-fonts.js";
import { readDrawingEffects } from "./drawingml-effects.js";
import { resolveImageSource } from "./image-source.js";

interface GroupPaint {
  fill?: ShapeFillModel | null;
  bounds?: BoundsModel;
}

/** `a:ln` width PowerPoint paints when the line states none: one point. */
const DEFAULT_LINE_WIDTH_EMU = EMU_PER_PT;
/** DrawingML's default text-frame and cell insets, in inches (91440 / 45720 EMU). */
const DEFAULT_INSET_X = 0.1;
const DEFAULT_INSET_Y = 0.05;

/** Right-to-left language tags: Arabic, Hebrew, Persian, Urdu, Yiddish and their variants. */
const RTL_LANGUAGE = /^(ar|he|iw|fa|ur|yi|ps|sd|ug|ku-arab|dv)(-|$)/i;

/** Cheap deck index: presentation metadata only, no slide XML inflated yet. */
export interface PptxDeck {
  size: { width: number; height: number };
  slideParts: string[];
  /** The first master's theme (the deck's, for what no slide names: the document info). */
  theme: DrawingTheme;
  tableStyles: Map<string, PptxTableStyle>;
  /**
   * Each slide master's own theme and the table styles resolved against it, by master part:
   * a slide takes the theme of the master its layout belongs to (see `themeOfMaster`).
   */
  masters: Map<string, { theme: DrawingTheme; tableStyles: Map<string, PptxTableStyle> }>;
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

type PptxTableStyle = Partial<Record<TableStyleRegionName, TableStyleRegion>>;

/** PowerPoint's paragraph alignments; an unknown value paints as the default (left), not a failed slide. */
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
    default:
      return undefined;
  }
}

/** A table cell's vertical anchor; an unknown value paints as the default (top). */
function verticalAlign(value: string | undefined): TableCellModel["verticalAlign"] {
  switch (value) {
    case "ctr":
      return "middle";
    case "b":
      return "bottom";
    default:
      return undefined;
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
 * Line width `a:lnRef idx="1|2|3"` draws at when the theme states none (the
 * theme's `a:lnStyleLst` wins): 0.5, 1 and 1.5 pt, the 2013-era Office theme.
 * The newer Office (Aptos) theme states 1, 1.5 and 2 pt, so widths depend on
 * the theme's version and must be read, not assumed.
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
  const slot = Math.min(index, 3) - 1;
  return { color, width: emuToInches(theme.lineWidths?.[slot] ?? THEME_LINE_WIDTHS_EMU[slot]!) };
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

/** PowerPoint's own spelling of the shared region flags (`a:tbl` calls the
 * column flags `firstCol`/`bandCol`; the shared module says `firstColumn`). */
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

function sharedFlags(shape: TableShape): TableStyleFlags {
  return {
    rows: shape.rows, columns: shape.columns,
    firstRow: shape.firstRow, lastRow: shape.lastRow,
    firstColumn: shape.firstCol, lastColumn: shape.lastCol,
    bandRow: shape.bandRow, bandColumn: shape.bandCol,
  };
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
  for (const { name, rows, columns } of tableCellRegions(sharedFlags(shape), r0, r1, c0, c1)) {
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

/**
 * PowerPoint's built-in Medium Style 2 family by style id: `accent` is the
 * theme slot the style is built on. PowerPoint writes the styles a deck uses
 * into `tableStyles.xml`, but other generators (python-pptx, PptxGenJS) only
 * reference the id, and PowerPoint then draws the style from its own copy.
 */
const BUILTIN_MEDIUM_STYLE_2: Record<string, string> = {
  "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}": "dk1",
  "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}": "accent1",
  "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}": "accent2",
  "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}": "accent3",
  "{00A15C55-8517-42AA-B614-E9B94910E393}": "accent4",
  "{7DF18680-E054-41AD-8BC1-D1AEF772440D}": "accent5",
  "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}": "accent6",
};

/** Medium Style 2 as PowerPoint stores it (the `pptx-coverage.pptx` definition), on one accent slot. */
function mediumStyle2Xml(accent: string): string {
  const white = `<a:solidFill><a:schemeClr val="lt1"/></a:solidFill>`;
  const edge = (side: string, width: number): string => `<a:${side}><a:ln w="${width}" cmpd="sng">${white}</a:ln></a:${side}>`;
  const tint = (amount: number): string => `<a:fill><a:solidFill><a:schemeClr val="${accent}"><a:tint val="${amount}"/></a:schemeClr></a:solidFill></a:fill>`;
  const solid = `<a:fill><a:solidFill><a:schemeClr val="${accent}"/></a:solidFill></a:fill>`;
  const text = (bold: boolean, color: string): string => `<a:tcTxStyle${bold ? ` b="on"` : ""}><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="${color}"/></a:tcTxStyle>`;
  const accented = (border: string): string => `${text(true, "lt1")}<a:tcStyle><a:tcBdr>${border}</a:tcBdr>${solid}</a:tcStyle>`;
  return `<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`
    + `<a:wholeTbl>${text(false, "dk1")}<a:tcStyle><a:tcBdr>${["left", "right", "top", "bottom", "insideH", "insideV"].map((side) => edge(side, 12700)).join("")}</a:tcBdr>${tint(20000)}</a:tcStyle></a:wholeTbl>`
    + `<a:band1H><a:tcStyle><a:tcBdr/>${tint(40000)}</a:tcStyle></a:band1H>`
    + `<a:band1V><a:tcStyle><a:tcBdr/>${tint(40000)}</a:tcStyle></a:band1V>`
    + `<a:lastCol>${accented("")}</a:lastCol><a:firstCol>${accented("")}</a:firstCol>`
    + `<a:lastRow>${accented(edge("top", 38100))}</a:lastRow><a:firstRow>${accented(edge("bottom", 38100))}</a:firstRow>`
    + `</a:tblStyle>`;
}

/**
 * The style a table's `a:tableStyleId` names: the deck's own definition, else
 * a built-in one built on first use and kept in `styles`. Any other missing
 * style paints unstyled (PowerPoint draws those from its own copy).
 */
function tableStyleFor(styleId: string | undefined, styles: Map<string, PptxTableStyle>, theme: DrawingTheme): PptxTableStyle | undefined {
  if (styleId === undefined) return undefined;
  const known = styles.get(styleId);
  if (known !== undefined) return known;
  const accent = BUILTIN_MEDIUM_STYLE_2[styleId];
  if (accent === undefined) return undefined;
  const root = parseXml(mediumStyle2Xml(accent));
  const built: PptxTableStyle = {};
  for (const name of TABLE_STYLE_REGIONS) {
    const region = readTableStyleRegion(firstChild(root, name), theme);
    if (region !== undefined) built[name] = region;
  }
  styles.set(styleId, built);
  return built;
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

/** Relationship lists are cached per part by `OpcPackage`, so their link maps can be too. */
const linkMaps = new WeakMap<object, Map<string, string>>();

/**
 * Hyperlink targets for one part: external URLs verbatim, slide jumps and
 * other internal targets kept as `#`-prefixed part paths. The renderer
 * surfaces them via `linkactivate`; the host decides.
 */
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
 * Image references under a picture bullet (`a:buBlip/a:blip`): the `svgBlip`
 * extension embed and the main `blip` embed, each by `r:id`. The SVG paints
 * first with the raster as its fallback; a lone `svgBlip` (what Office
 * writes for a bullet, with no main `blip` reference) paints alone.
 * `attr()` matches the local name, so prefixed `r:embed` resolves.
 */
function findBulletEmbeds(buBlip: XmlElement): { svg?: string; main?: string } {
  const stack = [...buBlip.children];
  let blip: XmlElement | undefined;
  while (stack.length > 0) {
    const element = stack.pop()!;
    if (element.local === "blip" && blip === undefined) blip = element;
    stack.push(...element.children);
  }
  if (blip === undefined) return {};
  const out: { svg?: string; main?: string } = {};
  const main = attr(blip, "embed");
  if (main !== undefined) out.main = main;
  const list = firstChild(blip, "extLst");
  if (list !== undefined) {
    for (const ext of childElements(list, "ext")) {
      const svg = firstChild(ext, "svgBlip");
      const embed = svg === undefined ? undefined : attr(svg, "embed");
      if (embed !== undefined && out.svg === undefined) out.svg = embed;
    }
  }
  return out;
}

/** A paragraph's runs in document order: `a:r`, and `a:fld` (slide number, date), which carries the same rPr/t. `a:br` is not text. */
function paragraphRuns(paragraph: XmlElement): XmlElement[] {
  return paragraph.children.filter((child) => child.local === "r" || child.local === "fld");
}

/** A run's own text: its `a:t`, or nothing. */
function runText(run: XmlElement): string {
  const t = firstChild(run, "t");
  return t === undefined ? "" : textContent(t);
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
  const textFill = properties.map(directFillProperty).find((fill) => fill !== undefined);
  const solidFill = textFill?.local === "solidFill" ? textFill : undefined;
  const latin = firstElement(properties, "latin");
  const cs = firstElement(properties, "cs");
  const out: TextRunModel = { text: runText(run) };
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
  if (latinTypeface !== undefined) out.latinTypeface = latinTypeface;
  if (complexScriptTypeface !== undefined) out.complexScriptTypeface = complexScriptTypeface;
  const color = resolveColor(solidFill, theme)
    ?? fontRefs?.map((ref) => resolveDrawingColor(ref, theme)).find((value): value is string => value !== undefined);
  // `p:defaultTextStyle` intentionally does not interpose: the coverage
  // fixture's dark-slide titles carry `tx1` there yet PowerPoint paints them
  // white — "Automatic" contrast wins for fill-less runs.
  if (color !== undefined) out.color = color;
  else if (autoColor !== undefined) out.color = autoColor;
  // Native PowerPoint preserves text grpFill but paints no ink, not the
  // parent's blue fill: its PNG has no glyphs and its PDF trace has opacity
  // zero. The PDF's extracted black colour alone misleadingly looks painted.
  if (textFill?.local === "grpFill") out.noFill = true;
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
  if (textFill?.local === "noFill") out.noFill = true;
  if (textFill?.local === "grpFill") out.groupFill = true;
  const hlink = firstElement(properties, "hlinkClick");
  if (hlink !== undefined && links !== undefined) {
    const rid = attr(hlink, "id");
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
  const runGradient = textFill?.local === "gradFill" ? textFill : undefined;
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
  // `ST_TextAutonumberScheme` names: arabicPeriod, romanUcPeriod, alphaLcParenBoth, arabicPlain, …
  const normalized = type.toLowerCase();
  const upper = normalized.includes("uc");
  let core: string;
  if (normalized.includes("roman")) core = toRoman(value, upper);
  else if (normalized.includes("alpha")) core = toAlpha(value, upper);
  else core = String(value);
  if (normalized.includes("parenboth")) return `(${core})`;
  if (normalized.includes("paren")) return `${core})`;
  if (normalized.includes("plain")) return core;
  // Default: period form ("1.", "I.", "A.", …) covers arabicPeriod and unknown types.
  return `${core}.`;
}

export function parseTextBody(txBody: XmlElement, theme: DrawingTheme, inheritance?: TextBodyInheritance): TextBodyModel {
  const bodies = [txBody, ...(inheritance?.bodies ?? [])];
  // Anchor, wrapping, and autofit retain the existing highest-bodyPr behavior.
  // Insets are resolved independently below. In the title-insets Office
  // probe, setting a title's bottom margin on master, layout and slide in
  // turn showed that an authored slide value wins, otherwise layout then
  // master supplies it; an explicit zero must not fall through.
  const bodyProperties = bodies.map((body) => firstChild(body, "bodyPr")).find((properties): properties is XmlElement => properties !== undefined
    && (Object.keys(properties.attrs).length > 0 || properties.children.length > 0));
  // Insets inherit independently: a layout can override the anchor while the
  // master still supplies the text-frame margins (the same Office probe).
  const bodyPropertySources = bodies
    .map((body) => firstChild(body, "bodyPr"))
    .filter((properties): properties is XmlElement => properties !== undefined);
  const paragraphElements = childElements(txBody, "p");
  const autoNumCounters = new Map<string, number>();
  // PowerPoint's editor measures the overflow and stores the outcome on the
  // shape's own `a:normAutofit`; opening a file never re-measures, and a bare
  // `<a:normAutofit/>` paints at 100% (probe `line-spacing-units`). So the
  // stored scale is the only shrink there is: each painted size is the
  // authored size times `fontScale`, rounded to a whole point, and each line
  // step loses `lnSpcReduction` of its line spacing. Space in points is left
  // alone. A layout's or master's `normAutofit` never carries a slide's scale.
  const ownBodyProperties = firstChild(txBody, "bodyPr");
  const storedAutofit = ownBodyProperties === undefined ? undefined : firstChild(ownBodyProperties, "normAutofit");
  const storedScale = storedAutofit === undefined ? undefined : numberAttr(storedAutofit, "fontScale", "a:normAutofit");
  const storedReduction = storedAutofit === undefined ? undefined : numberAttr(storedAutofit, "lnSpcReduction", "a:normAutofit");
  const fontScale = storedScale !== undefined && storedScale > 0 && storedScale < 100000 ? storedScale / 100000 : undefined;
  const lineSpacingReduction = storedReduction !== undefined && storedReduction > 0 ? storedReduction / 100000 : undefined;
  const scaleSize = (size: number): number => (fontScale === undefined ? size : Math.max(1, Math.round(size * fontScale)));
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
      const runs = paragraphRuns(paragraph);
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
      // PowerPoint has no run direction: a run in a left-to-right `lang`
      // keeps its numbers and punctuation together left to right inside a
      // right-to-left paragraph (probe run-direction). Unlabelled runs follow
      // the Unicode bidi algorithm alone.
      if (rtl === true) {
        for (const run of out.runs) {
          if (run.language !== undefined && isLtrRunInRtl(run.text, true, RTL_LANGUAGE.test(run.language))) run.ltrRun = true;
        }
      } else {
        // The mirror case, in a left-to-right paragraph: a run in a
        // left-to-right `lang` holding Arabic words keeps them in written
        // order around ASCII punctuation (probe en-us-arabic-run).
        out.runs = out.runs.flatMap((run) => {
          const pieces = run.language === undefined || RTL_LANGUAGE.test(run.language) || run.field !== undefined
            ? undefined
            : arabicPunctuationPieces(run.text);
          return pieces === undefined
            ? [run]
            : pieces.map((piece) => ({ ...run, text: piece.text, ...(piece.punctuation ? { ltrRun: true } : {}) }));
        });
      }
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
      else if (lineSpacingReduction !== undefined) out.lineSpacing = Math.max(0, (lineSpacing?.percent ?? 1) - lineSpacingReduction);
      if (fontScale !== undefined) {
        if (out.fontSizePt !== undefined) out.fontSizePt = scaleSize(out.fontSizePt);
        for (const run of out.runs) if (run.fontSizePt !== undefined) run.fontSizePt = scaleSize(run.fontSizePt);
      }
      if (spaceBefore?.points !== undefined) out.spaceBeforePt = spaceBefore.points;
      if (spaceAfter?.points !== undefined) out.spaceAfterPt = spaceAfter.points;
      // `a:spcPct` spacing counts in lines of 1.2 x the paragraph's largest run
      // size, whatever the font or the line spacing (probe `line-spacing-units`).
      if (spaceBefore?.percent !== undefined) out.spaceBeforeLines = spaceBefore.percent;
      if (spaceAfter?.percent !== undefined) out.spaceAfterLines = spaceAfter.percent;
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
        const numberText = formatAutoNum(autoNumType, count);
        // A theme token (`+mj-lt`, the heading font) resolves like a run's.
        const numTypeface = buFontForNum === undefined ? undefined : resolveThemeTypeface(attr(buFontForNum, "typeface"), numberText, theme);
        out.bullet = {
          char: numberText,
          ...(numTypeface === undefined ? {} : { typeface: numTypeface }),
          ...bulletExtras,
        };
      } else if (buBlip !== undefined) {
        // Picture bullets resolve their embed through the owning part's
        // rels (threaded as `embeds`); the renderer paints the image with a
        // character fallback when fetch fails or no resolver is wired.
        const rids = findBulletEmbeds(buBlip);
        // A dangling SVG reference falls back to the main `blip` as the
        // source itself, not just as a decode fallback.
        const primary = rids.svg !== undefined && (inheritance?.embeds?.get(rids.svg) !== undefined || rids.main === undefined)
          ? rids.svg
          : rids.main;
        const picture = primary === undefined ? undefined : inheritance?.embeds?.get(primary);
        const fallback = primary === rids.svg && rids.main !== undefined && rids.main !== rids.svg
          ? inheritance?.embeds?.get(rids.main)
          : undefined;
        const buCharText = buChar === undefined ? undefined : attr(buChar, "char");
        out.bullet = {
          ...(buCharText === undefined ? {} : { char: buCharText }),
          ...bulletExtras,
          ...(picture === undefined ? {} : {
            pictureMediaId: picture.mediaId,
            pictureContentType: picture.contentType,
            ...(fallback === undefined ? {} : {
              pictureFallbackMediaId: fallback.mediaId,
              pictureFallbackContentType: fallback.contentType,
            }),
          }),
          ...(picture === undefined && primary !== undefined ? { char: BULLET } : {}),
        };
        if (out.bullet !== undefined && Object.keys(out.bullet).length === 0) out.bullet = { char: BULLET };
      } else {
        const buFont = buChar === undefined ? undefined : firstElement(properties, "buFont");
        const authoredChar = buChar === undefined ? undefined : attr(buChar, "char");
        const authoredTypeface = buFont === undefined ? undefined : resolveThemeTypeface(attr(buFont, "typeface"), authoredChar ?? "", theme);
        // A Wingdings or Symbol bullet becomes its Unicode shape, painted in
        // any font that has it, not a letter on a device without Office.
        const standIn = authoredChar === undefined ? undefined : symbolFontStandIn(authoredTypeface, authoredChar);
        const char = standIn ?? authoredChar;
        const typeface = standIn === undefined ? authoredTypeface : undefined;
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
  // A side no `a:bodyPr` in the chain states takes DrawingML's default (0.1in
  // left/right, 0.05in top/bottom), which is what an `AddTextbox` box, whose
  // bodyPr authors no insets, paints in PowerPoint.
  out.insets = {
    left: insets.left === undefined ? DEFAULT_INSET_X : emuToInches(parseNumber(insets.left, "lIns", "a:bodyPr")),
    right: insets.right === undefined ? DEFAULT_INSET_X : emuToInches(parseNumber(insets.right, "rIns", "a:bodyPr")),
    top: insets.top === undefined ? DEFAULT_INSET_Y : emuToInches(parseNumber(insets.top, "tIns", "a:bodyPr")),
    bottom: insets.bottom === undefined ? DEFAULT_INSET_Y : emuToInches(parseNumber(insets.bottom, "bIns", "a:bodyPr")),
  };
  // Attributes cascade one by one, nearest first: a slide's `bodyPr` that only
  // carries `a:normAutofit` still takes the master's `anchor` (slide 2 of
  // `pptx-coverage.pptx`: a centred title).
  switch (firstAttribute(bodyPropertySources, "anchor")) {
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
  switch (firstAttribute(bodyPropertySources, "wrap")) {
    case "none":
      out.wrapping = "none";
      break;
    case "square":
      out.wrapping = "square";
      break;
  }
  if (bodyProperties !== undefined) {
    if (firstChild(bodyProperties, "noAutofit") !== undefined) out.autofit = "none";
    else if (firstChild(bodyProperties, "normAutofit") !== undefined) out.autofit = "shrinkText";
    else if (firstChild(bodyProperties, "spAutoFit") !== undefined) out.autofit = "resizeShape";
  }
  if (fontScale !== undefined) out.fontScale = fontScale;
  if (lineSpacingReduction !== undefined) out.lineSpacingReduction = lineSpacingReduction;
  // WordArt preset warp lives on the body's own properties (whole-body
  // transform). Painted unwarped and marked — see `TextBodyModel.textWarp`.
  if (bodyProperties !== undefined) {
    const warp = firstChild(bodyProperties, "prstTxWarp");
    const preset = warp === undefined ? undefined : attr(warp, "prst");
    if (preset !== undefined) out.textWarp = preset;
  }
  return out;
}

/** A shape's non-visual properties: `p:nvSpPr`, or `p:nvCxnSpPr` on a connector. */
function nonVisualProps(shape: XmlElement): XmlElement | undefined {
  return firstChild(shape, "nvSpPr") ?? firstChild(shape, "nvCxnSpPr");
}

/** Elements of a shape tree that carry an identity (`p:cNvPr`), and the wrappers that can hold some. */
const IDENTIFIED = new Set(["sp", "cxnSp", "pic", "grpSp", "graphicFrame"]);
const WRAPPERS = new Set(["grpSp", "AlternateContent", "Choice", "Fallback"]);

/** Ids made up for elements that state none, by element; filled by `numberShapeTree`. */
const syntheticIds = new WeakMap<XmlElement, string>();

/**
 * Give every identified element of a shape tree its position-derived id, for
 * the ones that state none. It is the element's place in a depth-first walk,
 * so it is the same on every read, and `auto-N` can never equal a real id
 * (always a number).
 */
function numberShapeTree(tree: XmlElement): void {
  let ordinal = 0;
  const walk = (parent: XmlElement): void => {
    for (const child of parent.children) {
      if (IDENTIFIED.has(child.local)) syntheticIds.set(child, `auto-${++ordinal}`);
      if (WRAPPERS.has(child.local)) walk(child);
    }
  };
  walk(tree);
}

/** How many of `elements`, groups' children included, carry an id `numberShapeTree` made up (a real id is always a number). */
function countSyntheticIds(elements: readonly SlideModel["elements"][number][]): number {
  let count = 0;
  for (const element of elements) {
    if (element.id.startsWith("auto-")) count++;
    if (element.kind === "group") count += countSyntheticIds(element.children);
  }
  return count;
}

/**
 * A shape-tree element's identity: `p:cNvPr` inside its non-visual properties
 * element `nonVisual` (`p:nvPicPr`, `p:nvGrpSpPr`, `p:nvGraphicFramePr`, …).
 *
 * A missing `p:cNvPr` or `@id` is not damage. Probe `group-id` (PowerPoint, Oct
 * 2026): copies of `fixtures/pptx-coverage.pptx` with the `id` attribute removed
 * from slide 3's rotated group (`<p:cNvPr id="4" name="Group 3">`), from its
 * Oval 32 and from its Picture 67. The group opened with no prompt and was drawn
 * normally. The shape and the picture each raised "found a problem, repair?" and
 * then "removed some content", but diffing the repaired and saved file shows the
 * repair only adds `id="6"` to the oval's `<p:cNvPr>` (the lowest id unused on
 * the slide); nothing is removed, and both are drawn. aared cannot prompt, so it
 * does the same: the element is kept and takes the id `numberShapeTree` gave it
 * (`auto-N`, from its position; ids are only internal, so the number does not
 * matter). The render reports it (`SlideModel.missingIds`, the `pptx-missing-id`
 * warning).
 *
 * A chart and a table (both `graphicFrame`) were probed the same way (probe
 * `drawing-id`: slide 2's Table 3, the chart-oracle deck's Chart 1): the same
 * repair prompt, and both still drawn. `KEEP_WITHOUT_ID` is the one place to
 * change the rule: a kind taken out of the set makes its missing id a damaged
 * slide again.
 */
const KEEP_WITHOUT_ID = new Set(["grpSp", "sp", "cxnSp", "pic", "graphicFrame"]);

function identityOf(element: XmlElement, nonVisual: XmlElement | undefined): { cNvPr: XmlElement | undefined; id: string; name: string } {
  const cNvPr = nonVisual === undefined ? undefined : firstChild(nonVisual, "cNvPr");
  const stated = cNvPr === undefined ? undefined : attr(cNvPr, "id");
  if (stated === undefined && !KEEP_WITHOUT_ID.has(element.local)) throw damagedDocument(`p:${element.local} is missing p:cNvPr/@id`);
  return { cNvPr, id: stated ?? syntheticIds.get(element) ?? "auto-0", name: (cNvPr === undefined ? undefined : attr(cNvPr, "name")) ?? "" };
}

function shapeIdentity(shape: XmlElement): { id: string; name: string; placeholder?: NonNullable<ShapeModel["placeholder"]> } {
  const nvSpPr = nonVisualProps(shape);
  const { id, name } = identityOf(shape, nvSpPr);
  const nvPr = nvSpPr === undefined ? undefined : firstChild(nvSpPr, "nvPr");
  const ph = nvPr === undefined ? undefined : firstChild(nvPr, "ph");
  const out: { id: string; name: string; placeholder?: NonNullable<ShapeModel["placeholder"]> } = { id, name };
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
  // Same mapping as the master placeholder lookup: a subtitle or content
  // placeholder takes the master's body style (its 90% line spacing and 10 pt
  // space before), and only date/footer/slide-number take `otherStyle`.
  switch (masterPlaceholderType(placeholder.type)) {
    case "title":
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
  groupPaint?: GroupPaint,
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
  const nvSpPr = nonVisualProps(shape);
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
  const adjustments = drawingGeometryAdjustments(prstGeom);
  // Hand-coded presets resolve here; everything else (the rest of the preset
  // catalog, custGeom) is described now and evaluated by the session.
  out.geometry = geometryForPreset(preset, size, adjustments)
    ?? unresolvedGeometry(size, preset, adjustments, spPr === undefined ? undefined : firstChild(spPr, "custGeom"));
  const fill = resolveShapeFill(spPr, style, theme, {
    relationships: slideRels ?? [],
    mediaContentType: (target) => pkg.mediaContentType(target),
    inherited: false,
    ...(groupPaint?.fill === undefined ? {} : { groupFill: groupPaint.fill }),
  });
  if (fill !== undefined) out.fill = fill;
  if (spPr !== undefined && firstChild(spPr, "grpFill") !== undefined && groupPaint?.bounds !== undefined) {
    out.fillBounds = groupPaint.bounds;
  }
  const effects = readDrawingEffects(spPr, theme);
  if (effects !== undefined) out.effects = effects;
  const directLineHasNoFill = line !== undefined && firstChild(line, "noFill") !== undefined;
  const styleLine = style === undefined ? undefined : firstChild(style, "lnRef");
  const strokeColor = directLineHasNoFill ? undefined : resolveSolidFill(line, theme) ?? (line === undefined ? resolveStyleRef(styleLine, theme) : undefined);
  if (strokeColor !== undefined) {
    const strokeWidth = line === undefined ? undefined : numberAttr(line, "w", "a:ln");
    out.stroke = { color: strokeColor, width: emuToInches(strokeWidth ?? DEFAULT_LINE_WIDTH_EMU) };
    const dash = lineDash(line === undefined ? undefined : firstChild(line, "prstDash"));
    if (dash !== undefined) out.stroke.dash = dash;
  }
  const transform = drawingShapeTransform(xfrm);
  if (transform !== undefined) out.transform = transform;
  throwIfAborted(signal);
  return out;
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
  const { id, name } = identityOf(pic, firstChild(pic, "nvPicPr"));
  const blipFill = firstChild(pic, "blipFill");
  // An embedded SVG (`asvg:svgBlip`) paints first; the main `blip` stays as
  // its raster fallback, so an SVG that cannot be decoded still shows
  // Office's own PNG of it.
  const source = resolveImageSource(blipFill === undefined ? undefined : firstChild(blipFill, "blip"), {
    relationships,
    mediaContentType: (target) => pkg.mediaContentType(target),
    hasPart: (target) => pkg.hasPart(target),
  });
  if (source === undefined) return undefined;
  const spPr = firstChild(pic, "spPr");
  const xfrm = spPr === undefined ? undefined : firstChild(spPr, "xfrm");
  const bounds = parseBounds(xfrm);
  const out: PictureModel = {
    kind: "picture",
    id,
    name,
    ...(bounds === undefined ? {} : { bounds }),
    mediaId: source.mediaId,
    contentType: source.contentType,
    ...(source.fallbackMediaId === undefined || source.fallbackContentType === undefined ? {} : {
      fallbackMediaId: source.fallbackMediaId,
      fallbackContentType: source.fallbackContentType,
    }),
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
  const transform = drawingShapeTransform(xfrm);
  if (transform !== undefined) out.transform = transform;
  return out;
}

/** Extract slide notes as plain text from its `notesSlide` part, if present. */
export async function parseSlideNotes(pkg: OpcPackage, slidePart: string, signal?: AbortSignal): Promise<string | undefined> {
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

/** The `p:graphicFrame` of an Office 2016 chart: inside `mc:AlternateContent/mc:Choice`, holding a `cx:chart`. */
function chartExFrame(alternate: XmlElement): XmlElement | undefined {
  for (const choice of alternate.children) {
    if (choice.local !== "Choice") continue;
    const frame = firstChild(choice, "graphicFrame");
    const graphic = frame === undefined ? undefined : firstChild(frame, "graphic");
    const data = graphic === undefined ? undefined : firstChild(graphic, "graphicData");
    if (frame !== undefined && data !== undefined && firstChild(data, "chart") !== undefined) return frame;
  }
  return undefined;
}

async function parseChartFrame(
  pkg: OpcPackage,
  frame: XmlElement,
  relationships: Array<{ id: string; type: string; target: string; targetMode: string }>,
  theme: DrawingTheme,
  signal?: AbortSignal,
): Promise<ChartModel | undefined> {
  throwIfAborted(signal);
  const graphic = firstChild(frame, "graphic");
  const graphicData = graphic === undefined ? undefined : firstChild(graphic, "graphicData");
  const chartRef = graphicData === undefined ? undefined : firstChild(graphicData, "chart");
  if (chartRef === undefined) return undefined;
  const { cNvPr, id, name } = identityOf(frame, firstChild(frame, "nvGraphicFramePr"));
  const alt = cNvPr === undefined ? undefined : attr(cNvPr, "descr") || attr(cNvPr, "title") || undefined;
  const xfrm = firstChild(frame, "xfrm");
  const bounds = parseBounds(xfrm);
  const base: ChartModel = {
    kind: "chart",
    id,
    name,
    ...(alt === undefined ? {} : { alt }),
    ...(bounds === undefined ? {} : { bounds }),
  };
  const transform = drawingShapeTransform(xfrm);
  if (transform !== undefined) base.transform = transform;
  const missing = "this chart points at a part the file does not hold; the file may be damaged";
  const chartId = attr(chartRef, "id");
  if (chartId === undefined) return { ...base, unsupportedReason: missing };
  const rel = relationships.find((candidate) => candidate.id === chartId);
  if (rel === undefined) return { ...base, unsupportedReason: missing };
  if (rel.targetMode === "External") return { ...base, unsupportedReason: "this chart is linked from another file and is not previewed" };
  const out: ChartModel = { ...base, sourceChartId: rel.target };
  let chartRoot: XmlElement;
  try {
    chartRoot = await pkg.readXml(rel.target, signal);
  } catch {
    return { ...out, unsupportedReason: missing };
  }
  throwIfAborted(signal);
  const read = readChartModel(chartRoot, theme, { owner: await chartFillOwner(pkg, rel.target, signal) });
  if ("spec" in read) out.spec = read.spec;
  else out.unsupportedReason = read.unsupportedReason;
  return out;
}

/**
 * Cell insets in inches: authored `a:tcPr/@marL|marR|marT|marB`, else
 * DrawingML's defaults (0.1in left/right, 0.05in top/bottom) — what
 * PowerPoint paints for a cell that states none.
 */
function parseCellMargins(tcPr: XmlElement | undefined): InsetsModel {
  const side = (name: string, fallback: number): number => {
    const emu = tcPr === undefined ? undefined : numberAttr(tcPr, name, "a:tcPr");
    return emu === undefined ? fallback : emuToInches(emu);
  };
  return { left: side("marL", DEFAULT_INSET_X), right: side("marR", DEFAULT_INSET_X), top: side("marT", DEFAULT_INSET_Y), bottom: side("marB", DEFAULT_INSET_Y) };
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

/** A slide table's `a:tbl`, when the graphic frame holds one. */
function tableOf(frame: XmlElement): XmlElement | undefined {
  const graphic = firstChild(frame, "graphic");
  const graphicData = graphic === undefined ? undefined : firstChild(graphic, "graphicData");
  return graphicData === undefined ? undefined : firstChild(graphicData, "tbl");
}

/**
 * The cells of one table row that hold content, with their text bodies. Every
 * grid column has a `tc`, merge continuations included, so a cell's grid
 * column is its position among its row's `tc` elements. A continuation is
 * covered by its master cell's span and is left out, keeping the logical grid
 * aligned with what PowerPoint paints. A span past the table's last column or
 * row covers nothing more: it is cut here, so per-cell loops stay bounded by
 * the grid however large a file claims it is. `gridSpan`, `rowSpan`, `hMerge`
 * and `vMerge` are attributes of `a:tc` itself (ECMA-376 21.1.3.16).
 */
function contentCells(row: XmlElement, rowIndex: number, rowCount: number, columnCount: number) {
  const rowCells = childElements(row, "tc");
  const gridColumns = Math.max(columnCount, rowCells.length);
  return rowCells.flatMap((cell, columnIndex) => {
    const gridSpan = clampSpan(numberAttr(cell, "gridSpan", "a:tc"), gridColumns - columnIndex);
    const rowSpan = clampSpan(numberAttr(cell, "rowSpan", "a:tc"), rowCount - rowIndex);
    const merged = onOff(cell, "hMerge") === true || onOff(cell, "vMerge") === true;
    if (merged && (gridSpan === undefined || gridSpan <= 1) && (rowSpan === undefined || rowSpan <= 1)) return [];
    const body = firstChild(cell, "txBody");
    if (body === undefined) throw damagedDocument("a table cell is missing a:txBody");
    return [{ cell, columnIndex, gridSpan, rowSpan, body }];
  });
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
  const table = tableOf(frame);
  if (table === undefined) return undefined;
  const { id, name } = identityOf(frame, firstChild(frame, "nvGraphicFramePr"));
  const tblPr = firstChild(table, "tblPr");
  const grid = firstChild(table, "tblGrid");
  const styleIdElement = tblPr === undefined ? undefined : firstChild(tblPr, "tableStyleId");
  const styleId = styleIdElement === undefined ? undefined : textContent(styleIdElement);
  const style = tableStyleFor(styleId, styles, theme);
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
    for (const { cell, columnIndex, gridSpan, rowSpan, body } of contentCells(row, rowIndex, rowElements.length, columns.length)) {
      const tcPr = firstChild(cell, "tcPr");
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
    name,
    ...(bounds === undefined ? {} : { bounds }),
    direction: rtl ? "rtl" : "ltr",
    columns,
    rows,
    cellFills,
  };
  const transform = drawingShapeTransform(xfrm);
  if (transform !== undefined) out.transform = transform;
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
      throw damagedDocument(`presentation slide relationship "${id ?? "missing"}" is invalid`);
    }
    return relationship.target;
  });
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
 * A slide's or layout's own colour map: `p:clrMapOvr/a:overrideClrMapping`
 * carries a full replacement map, while `a:masterClrMapping` (or no element)
 * keeps the one above it. Slide beats layout beats master.
 */
export function overrideColorMap(root: XmlElement): XmlElement | undefined {
  const override = firstChild(root, "clrMapOvr");
  return override === undefined ? undefined : firstChild(override, "overrideClrMapping");
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
  /** The colour map the layout paints with: its own override, else the master's `p:clrMap` (slide-level colour names -> theme slots). */
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
  const clrMap = overrideColorMap(layoutRoot) ?? (masterRoot === undefined ? undefined : firstChild(masterRoot, "clrMap"));
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
  options: { graphicsOnly: boolean; autoColor?: string; tableStyles?: Map<string, PptxTableStyle> },
  signal?: AbortSignal,
): Promise<SlideModel["elements"]> {
  numberShapeTree(tree);
  const isPlaceholder = (element: XmlElement): boolean => options.graphicsOnly && hasPlaceholder(element);
  const visit = async (child: XmlElement, target: SlideModel["elements"], groupPaint?: GroupPaint): Promise<void> => {
    throwIfAborted(signal);
    const pushFallbackPicture = (container: XmlElement): void => {
      const fallback = firstChild(container, "Fallback");
      const pic = fallback === undefined ? undefined : firstChild(fallback, "pic");
      if (pic === undefined) return;
      const picture = parsePicture(pic, relationships, pkg);
      if (picture !== undefined) target.push(picture);
    };
    // A connector (`p:cxnSp`) is a shape without text: same spPr geometry
    // and line, so it paints through the shape path.
    if (child.local === "sp" || child.local === "cxnSp") {
      if (isPlaceholder(child)) return;
      target.push(await parseShape(pkg, part, child, theme, signal, relationships, options.autoColor, deck.defaultTextSizes, groupPaint));
      throwIfAborted(signal);
      return;
    }
    if (child.local === "pic") {
      if (isPlaceholder(child)) return;
      const picture = parsePicture(child, relationships, pkg);
      if (picture !== undefined) {
        const effects = readDrawingEffects(firstChild(child, "spPr"), theme);
        if (effects !== undefined) picture.effects = effects;
        target.push(picture);
      }
      return;
    }
    if (child.local === "AlternateContent") {
      // An Office 2016 chart (waterfall, treemap, ...) is the Choice's graphic frame; its Fallback is a picture of it, never painted beside it.
      const chartFrame = options.graphicsOnly ? undefined : chartExFrame(child);
      if (chartFrame !== undefined) {
        const chart = await parseChartFrame(pkg, chartFrame, relationships, theme, signal);
        throwIfAborted(signal);
        if (chart !== undefined) target.push(chart);
        return;
      }
      // Ink and other modern content surface their raster fallback here.
      pushFallbackPicture(child);
      return;
    }
    if (child.local === "grpSp") {
      target.push(await parseGroup(child, visit, theme, {
        relationships, mediaContentType: (target) => pkg.mediaContentType(target), inherited: options.graphicsOnly,
      }, groupPaint));
      return;
    }
    if (child.local === "graphicFrame") {
      throwIfAborted(signal);
      const table = parseTable(child, theme, options.tableStyles ?? deck.tableStyles, relationships, options.autoColor, (name) => pkg.mediaContentType(name), deck.defaultTextSizes);
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
  visit: (child: XmlElement, target: SlideModel["elements"], groupPaint?: GroupPaint) => Promise<void>,
  theme: DrawingTheme,
  owner: FillOwner,
  parentPaint?: GroupPaint,
): Promise<GroupModel> {
  const { id, name } = identityOf(group, firstChild(group, "nvGrpSpPr"));
  const grpSpPr = firstChild(group, "grpSpPr");
  const xfrm = grpSpPr === undefined ? undefined : firstChild(grpSpPr, "xfrm");
  const bounds = parseBounds(xfrm);
  const transform = drawingShapeTransform(xfrm);
  const childSpace = parseChildSpace(xfrm);
  const effects = readDrawingEffects(grpSpPr, theme);
  const fill = grpSpPr === undefined ? undefined : resolveDirectFill(grpSpPr, theme, {
    ...owner, ...(parentPaint?.fill === undefined ? {} : { groupFill: parentPaint.fill }),
  });
  // A nested grpFill keeps the ancestor's paint box, mapped into its child
  // space; a directly authored group fill starts a new box. PowerPoint's COM
  // group formatting writes grpFill on both nested groups and child shapes.
  let fillBounds = childSpace ?? bounds;
  if (grpSpPr !== undefined && firstChild(grpSpPr, "grpFill") !== undefined && parentPaint?.bounds !== undefined
    && bounds !== undefined && childSpace !== undefined && bounds.width > 0 && bounds.height > 0) {
    const sx = childSpace.width / bounds.width, sy = childSpace.height / bounds.height;
    fillBounds = { x: childSpace.x + (parentPaint.bounds.x - bounds.x) * sx,
      y: childSpace.y + (parentPaint.bounds.y - bounds.y) * sy,
      width: parentPaint.bounds.width * sx, height: parentPaint.bounds.height * sy };
  }
  const paint: GroupPaint = { ...(fill === undefined ? {} : { fill }), ...(fillBounds === undefined ? {} : { bounds: fillBounds }) };
  const children: SlideModel["elements"] = [];
  for (const child of group.children) await visit(child, children, paint);
  return {
    kind: "group",
    id,
    name,
    ...(bounds === undefined ? {} : { bounds }),
    ...(transform === undefined ? {} : { transform }),
    ...(childSpace === undefined ? {} : { childSpace }),
    ...(effects === undefined ? {} : { effects }),
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
 * A layout's `showMasterSp="0"` leaves the master's shapes out; a slide's
 * leaves out the layout's and the master's together, and a master's picture
 * background is inherited. `pptx-coverage.pptx` slides 15-18 carry a master's
 * logo with each flag set, exported from PowerPoint and compared against aared.
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
  const owner = themeOfMaster(deck, chain.masterPart);
  const theme = themeForColorMap(owner.theme, chain.clrMap);
  const graphics: SlideLayoutModel["graphics"] = [];
  const showsMaster = onOff(chain.layoutRoot, "showMasterSp") !== false;
  const masterTree = chain.masterRoot === undefined ? undefined : shapeTreeOf(chain.masterRoot);
  if (showsMaster && masterTree !== undefined && chain.masterPart !== undefined && chain.masterRelationships !== undefined) {
    graphics.push(...(await parseShapeTree(pkg, chain.masterPart, masterTree, chain.masterRelationships, deck, theme, { graphicsOnly: true, tableStyles: owner.tableStyles }, signal)));
  }
  const layoutTree = shapeTreeOf(chain.layoutRoot);
  if (layoutTree !== undefined) {
    graphics.push(...(await parseShapeTree(pkg, layoutPart, layoutTree, chain.layoutRelationships, deck, theme, { graphicsOnly: true, tableStyles: owner.tableStyles }, signal)));
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
  if (width === undefined || height === undefined) throw damagedDocument("the presentation is missing p:sldSz");
  const masters = await readMasters(pkg, presentationPart, presentation, signal);
  const first = masters.values().next().value as { theme: DrawingTheme; tableStyles: Map<string, PptxTableStyle> } | undefined;
  // A deck whose masters name no theme part (a hand-made one) falls back to the first theme in the package.
  const theme = first?.theme ?? await readDrawingTheme(pkg, signal);
  throwIfAborted(signal);
  const parts = await slideParts(pkg, presentationPart, signal);
  throwIfAborted(signal);
  const tableStyles = first?.tableStyles ?? await readTableStyles(pkg, theme, signal);
  throwIfAborted(signal);
  return {
    size: { width: emuToInches(width), height: emuToInches(height) },
    slideParts: parts,
    theme,
    tableStyles,
    masters,
    defaultTextSizes: defaultTextSizes(firstChild(presentation, "defaultTextStyle")),
  };
}

/**
 * Every slide master in the presentation's own order (`p:sldMasterIdLst`), with the theme its
 * relationships name and the table styles resolved against it. Probe: PowerPoint gives each
 * master its own theme part (a second master's accent colours are its own, not the first's).
 */
async function readMasters(
  pkg: OpcPackage,
  presentationPart: string,
  presentation: XmlElement,
  signal?: AbortSignal,
): Promise<PptxDeck["masters"]> {
  const rels = await pkg.relationships(presentationPart, signal);
  throwIfAborted(signal);
  const list = firstChild(presentation, "sldMasterIdLst");
  const listed = list === undefined ? [] : childElements(list, "sldMasterId").map((entry) => attr(entry, "r:id"));
  const masterParts = listed.length > 0
    ? listed.flatMap((id) => rels.filter((rel) => rel.id === id).map((rel) => rel.target))
    : rels.filter((rel) => rel.type.endsWith("/slideMaster")).map((rel) => rel.target);
  const masters: PptxDeck["masters"] = new Map();
  for (const masterPart of masterParts) {
    const themeRel = (await pkg.relationships(masterPart, signal)).find((rel) => rel.type.endsWith("/theme"));
    throwIfAborted(signal);
    if (themeRel === undefined) continue;
    const theme = await readThemePart(pkg, themeRel.target, signal);
    masters.set(masterPart, { theme, tableStyles: await readTableStyles(pkg, theme, signal) });
  }
  return masters;
}

/** The theme and table styles of the master a layout belongs to; the deck's own when it names none. */
function themeOfMaster(deck: PptxDeck, masterPart: string | undefined): { theme: DrawingTheme; tableStyles: Map<string, PptxTableStyle> } {
  return (masterPart === undefined ? undefined : deck.masters.get(masterPart)) ?? { theme: deck.theme, tableStyles: deck.tableStyles };
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

/** What every read of one slide starts from: its part, shape tree and relationships, its layout chain, and the theme its colours resolve through. */
async function openSlide(pkg: OpcPackage, deck: PptxDeck, index: number, signal?: AbortSignal) {
  throwIfAborted(signal);
  const part = deck.slideParts[index];
  if (part === undefined) throw new AaredError("not-found", `@aared/viewer/core: slide index ${index} is outside this ${deck.slideParts.length}-slide deck`);
  const root = await pkg.readXml(part, signal);
  const cSld = firstChild(root, "cSld");
  const tree = cSld === undefined ? undefined : firstChild(cSld, "spTree");
  if (tree === undefined) throw damagedDocument(`slide "${part}" is missing p:spTree`);
  const relationships = await pkg.relationships(part, signal);
  throwIfAborted(signal);
  const layoutRel = relationships.find((relationship) => relationship.type.endsWith("/slideLayout"));
  const chain = layoutRel === undefined ? undefined : await layoutChain(pkg, layoutRel.target, signal);
  throwIfAborted(signal);
  // Colours resolve through this slide's colour map (its own override, else
  // its layout's, else the master's), not the bare theme.
  const owner = themeOfMaster(deck, chain?.masterPart);
  const theme = themeForColorMap(owner.theme, overrideColorMap(root) ?? chain?.clrMap);
  return { part, root, cSld, tree, relationships, chain, owner, theme };
}

/**
 * A slide's words, for search: each paragraph of its shapes and table cells,
 * and its charts' words, in the order `parsePptxSlide` lists their elements.
 * The same walk as `parseShapeTree` over a slide's own tree, without resolving
 * what search never shows (colours, fonts, placeholder inheritance, geometry,
 * the background, notes), and nothing is cloned to the page but these strings.
 * It must equal the paragraph text of the full model: `pptx-text.test.ts`
 * checks every slide of the fixture decks.
 */
export async function readPptxSlideText(pkg: OpcPackage, deck: PptxDeck, index: number, signal?: AbortSignal): Promise<string[]> {
  const { tree, relationships, theme } = await openSlide(pkg, deck, index, signal);
  const words: string[] = [];
  const addBody = (body: XmlElement): void => {
    for (const paragraph of childElements(body, "p")) {
      const text = paragraphRuns(paragraph).map(runText).join("").trim();
      if (text !== "") words.push(text);
    }
  };
  const addChart = async (frame: XmlElement): Promise<void> => {
    const chart = await parseChartFrame(pkg, frame, relationships, theme, signal);
    if (chart?.spec !== undefined) words.push(...chartText(chart.spec));
  };
  const visit = async (child: XmlElement): Promise<void> => {
    throwIfAborted(signal);
    if (child.local === "sp" || child.local === "cxnSp") {
      const body = firstChild(child, "txBody");
      if (body !== undefined) addBody(body);
    } else if (child.local === "grpSp") {
      for (const member of child.children) await visit(member);
    } else if (child.local === "AlternateContent") {
      const frame = chartExFrame(child);
      if (frame !== undefined) await addChart(frame);
    } else if (child.local === "graphicFrame") {
      const table = tableOf(child);
      if (table === undefined) return addChart(child);
      const rows = childElements(table, "tr");
      const grid = firstChild(table, "tblGrid");
      const columnCount = grid === undefined ? 0 : childElements(grid, "gridCol").length;
      for (const [rowIndex, row] of rows.entries()) {
        for (const { body } of contentCells(row, rowIndex, rows.length, columnCount)) addBody(body);
      }
    }
  };
  for (const child of tree.children) await visit(child);
  return words;
}

/** Inflate and parse exactly one slide previously listed by `indexPptx()`. */
export async function parsePptxSlide(pkg: OpcPackage, deck: PptxDeck, index: number, signal?: AbortSignal): Promise<SlideModel> {
  const { part, root, cSld, tree, relationships, chain, owner, theme } = await openSlide(pkg, deck, index, signal);
  // Background first: fill-less runs take PowerPoint "Automatic" contrast
  // against it (white on dark), never CSS inheritance from embedding chrome.
  const background = resolveSlideBackground(pkg, cSld, relationships, chain, theme);
  const autoColor = automaticTextColor(background);
  const elements = await parseShapeTree(pkg, part, tree, relationships, deck, theme, { graphicsOnly: false, autoColor, tableStyles: owner.tableStyles }, signal);
  throwIfAborted(signal);
  const notes = await parseSlideNotes(pkg, part, signal);
  throwIfAborted(signal);
  // Only layouts with shared graphics need a fetch; most decks' empty layouts don't.
  const layout = chain === undefined ? undefined : await parsePptxLayout(pkg, deck, chain.layoutPart, signal);
  const layoutId = layout === undefined || layout.graphics.length === 0 ? undefined : layout.id;
  const hidesGraphics = onOff(root, "showMasterSp") === false;
  const missingIds = countSyntheticIds(elements);
  return {
    part,
    size: deck.size,
    elements,
    number: index + 1,
    ...(missingIds === 0 ? {} : { missingIds }),
    ...(background === undefined ? {} : { background }),
    ...(layoutId === undefined ? {} : { layoutId }),
    ...(hidesGraphics ? { hideBackgroundGraphics: true } : {}),
    ...(notes === undefined ? {} : { notes }),
  };
}
