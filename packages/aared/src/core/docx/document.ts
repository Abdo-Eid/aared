/**
 * Word `document.xml` reader for the DOCX route: body → resolved
 * `DocumentSection` flow blocks. Headless, plain data — DOM painting belongs to
 * `aared/docx`, never here.
 *
 * Coverage: style + numbering inheritance, bidirectional paragraphs
 * (mixed Arabic/Latin runs in one paragraph), heading outline, tables
 * (logical order, `bidiVisual`, spans, shading), lists (bullet + ordered,
 * nested levels), inline images, charts (read by the shared `core/chart.ts`)
 * and `wp:anchor` pictures and charts with full placement, footnotes and
 * endnotes, tracked-changes-as-final (`w:ins` kept, `w:del`/`w:moveFrom`
 * dropped), per-section headers/footers, page geometry (`w:pgSz`/`w:pgMar`),
 * explicit page breaks and soft `w:lastRenderedPageBreak` hints.
 * Pagination itself is a renderer concern (`paginate` consumes the sections);
 * this module only preserves the signals faithfully.
 *
 * Reuses `core` machinery only: OPC package/rels, the headless XML reader,
 * and the shared DrawingML theme (for `w:asciiTheme`-style run typefaces).
 * No copies of OPC/XML/theme/DrawingML/chart code live under `core/docx/`.
 */

import { BULLET, NON_BREAKING_HYPHEN } from "../chars.js";
import { attr, childElements, firstChild, numberAttr, ownString } from "../xml.js";
import type { XmlElement } from "../xml.js";
import type { OpcPackage } from "../opc.js";
import { parseHexColor, parseHexDigits, toHexColor } from "../color.js";
import { runsText } from "../search.js";
import { readDrawingTheme, resolveDrawingColor, resolveWordRunColor, themeFontFamily, themeForColorMap } from "../theme.js";
import { directFillProperty, resolveDirectFill } from "../fill.js";
import { resolveImageSource } from "../image-source.js";
import { readDrawingEffects } from "../drawingml-effects.js";
import type { FillOwner } from "../fill.js";
import { drawingGeometryAdjustments, drawingShapeTransform, geometryForPreset, unresolvedGeometry } from "../drawingml-shape.js";
import { isLtrRunInRtl, usesArabic } from "../typography.js";
import type { DrawingTheme } from "../theme.js";
import { damagedDocument, throwIfAborted } from "../errors.js";
import { TWIPS_PER_PT, emuToInches, ptToInches, twipsToInches } from "../units.js";
import { clampSpan } from "../table-span.js";
import { chartFillOwner, readChartModel } from "../chart.js";
import type {
  DocxAnchorPosition,
  DocxChartNode,
  DocxDrawingNode,
  DocxHeadingNode,
  DocxImageAnchor,
  DocxImageNode,
  DocxListItemNode,
  DocxPageBreakNode,
  DocxParagraphBorder,
  DocxParagraphBorders,
  DocxParagraphLayout,
  DocxParagraphNode,
  InsetsModel,
  DocxTabStop,
  DocxTableBorders,
  DocxTableCellNode,
  DocxTableEdge,
  DocxTableRowNode,
  DocxTextBoxNode,
  DocxUnsupportedNode,
  DocumentSection,
  TextAlign,
  FlowNode,
  HeaderFooterParts,
  SectionBreakType,
  SectionProperties,
  TextDirection,
  TextRunModel,
} from "../model.js";
import { characterAncestry, headingLevelForStyle, paragraphAncestry, parseStyles, tableStyleAncestry,
  tableStyleRegionsFor, tableStyleTextPropsFor } from "./styles.js";
import type { StyleData } from "./styles.js";
import { tableCellRegions, tableLookFromVal } from "../table-style.js";
import type { TableStyleFlags } from "../table-style.js";
import { isOrderedFormat, markerFor, parseNumbering, startFor } from "./numbering.js";
import type { NumberingData, NumberingLevel } from "./numbering.js";
import { parseNotes } from "./footnotes.js";
import { symbolFontStandIn } from "../symbol-fonts.js";

/** `w:spacing/@w:line` value of a single line (240 twentieths of a line). */
const SINGLE_LINE = 240;

/**
 * Word on/off: a bare element (`<w:b/>`) means true; only an explicit false
 * value (`false`, `0`, `off`, `none`) turns it off.
 */
function onOffValue(value: string | undefined): boolean {
  if (value === undefined) return true;
  return value !== "false" && value !== "0" && value !== "off" && value !== "none";
}

function onOffElement(element: XmlElement | undefined): boolean | undefined {
  if (element === undefined) return undefined;
  return onOffValue(attr(element, "val"));
}

/** A Word `w:color`/`w:fill` value as `#RRGGBB`; `auto` and anything else that is not six hex digits resolve to nothing. */
function hexColor(value: string | undefined): string | undefined {
  const channels = value === undefined ? undefined : parseHexDigits(value);
  return channels === undefined ? undefined : toHexColor(channels);
}

/** Resolve one `w:pBdr` edge to CSS terms; `nil`/`none` (or missing) draws nothing. */
function paragraphBorder(edge: XmlElement): DocxParagraphBorder | undefined {
  const val = attr(edge, "val");
  if (val === undefined || val === "nil" || val === "none") return undefined;
  const style: DocxParagraphBorder["style"] =
    val === "dotted" ? "dotted" : val === "double" ? "double" : val.startsWith("dash") ? "dashed" : "solid";
  const size = numberAttr(edge, "sz", "w:pBdr");
  const space = numberAttr(edge, "space", "w:pBdr");
  return {
    color: hexColor(attr(edge, "color")) ?? "currentColor",
    widthPt: (size ?? 4) / 8,
    spacePt: space ?? 0,
    style,
  };
}

const HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: "#FFFF00",
  green: "#00FF00",
  cyan: "#00FFFF",
  magenta: "#FF00FF",
  blue: "#0000FF",
  red: "#FF0000",
  darkBlue: "#000080",
  darkCyan: "#008080",
  darkGreen: "#008000",
  darkMagenta: "#800080",
  darkRed: "#800000",
  darkYellow: "#808000",
  darkGray: "#808080",
  lightGray: "#C0C0C0",
  black: "#000000",
};

interface ResolvedRunProps {
  bold?: boolean;
  italic?: boolean;
  /** `w:bCs` / `w:iCs`: the bold and italic a complex-script run takes (see `finalizeRun`). */
  boldCs?: boolean;
  italicCs?: boolean;
  fontSizePt?: number;
  fontSizeCsPt?: number;
  color?: string;
  ascii?: string;
  hAnsi?: string;
  cs?: string;
  eastAsia?: string;
  asciiTheme?: string;
  hAnsiTheme?: string;
  csTheme?: string;
  eastAsiaTheme?: string;
  hintCs?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  verticalAlign?: "subscript" | "superscript";
  highlight?: string;
  shading?: string;
  language?: string;
  languageBidi?: string;
  rtl?: boolean;
  letterSpacingPt?: number;
  caps?: "all" | "small";
}

function readRunProps(rPr: XmlElement | undefined, out: ResolvedRunProps, theme: DrawingTheme): void {
  if (rPr === undefined) return;
  for (const child of rPr.children) {
    switch (child.local) {
      case "b":
        out.bold = onOffValue(attr(child, "val"));
        break;
      case "bCs":
        out.boldCs = onOffValue(attr(child, "val"));
        break;
      case "i":
        out.italic = onOffValue(attr(child, "val"));
        break;
      case "iCs":
        out.italicCs = onOffValue(attr(child, "val"));
        break;
      case "sz": {
        const value = numberAttr(child, "val", "w:sz");
        if (value !== undefined) out.fontSizePt = value / 2;
        break;
      }
      case "szCs": {
        const value = numberAttr(child, "val", "w:szCs");
        if (value !== undefined) out.fontSizeCsPt = value / 2;
        break;
      }
      case "color": {
        const color = resolveWordRunColor(child, theme);
        if (color !== undefined) out.color = color;
        break;
      }
      case "rFonts": {
        const ascii = attr(child, "ascii");
        const hAnsi = attr(child, "hAnsi");
        const cs = attr(child, "cs");
        const eastAsia = attr(child, "eastAsia");
        const hint = attr(child, "hint");
        if (ascii !== undefined) out.ascii = ascii;
        if (hAnsi !== undefined) out.hAnsi = hAnsi;
        if (cs !== undefined) out.cs = cs;
        if (eastAsia !== undefined) out.eastAsia = eastAsia;
        if (hint === "cs") out.hintCs = true;
        const asciiTheme = attr(child, "asciiTheme");
        const hAnsiTheme = attr(child, "hAnsiTheme");
        const csTheme = attr(child, "cstheme") ?? attr(child, "csTheme");
        const eastAsiaTheme = attr(child, "eastAsiaTheme");
        if (asciiTheme !== undefined) out.asciiTheme = asciiTheme;
        if (hAnsiTheme !== undefined) out.hAnsiTheme = hAnsiTheme;
        if (csTheme !== undefined) out.csTheme = csTheme;
        if (eastAsiaTheme !== undefined) out.eastAsiaTheme = eastAsiaTheme;
        break;
      }
      case "u":
        out.underline = onOffValue(attr(child, "val"));
        break;
      case "strike":
      case "dstrike":
        out.strikethrough = onOffValue(attr(child, "val"));
        break;
      case "vertAlign": {
        const value = attr(child, "val");
        if (value === "superscript" || value === "subscript") out.verticalAlign = value;
        break;
      }
      case "highlight": {
        const value = attr(child, "val");
        if (value === "none" || value === undefined) {
          delete out.highlight;
        } else {
          const mapped = HIGHLIGHT_COLORS[value];
          if (mapped !== undefined) out.highlight = mapped;
        }
        break;
      }
      case "shd": {
        const fill = hexColor(attr(child, "fill"));
        if (fill !== undefined) out.shading = fill;
        break;
      }
      case "lang": {
        const val = attr(child, "val");
        const bidi = attr(child, "bidi");
        if (val !== undefined) out.language = val;
        if (bidi !== undefined) out.languageBidi = bidi;
        break;
      }
      case "rtl":
        out.rtl = onOffValue(attr(child, "val"));
        break;
      case "spacing": {
        const value = numberAttr(child, "val", "w:spacing");
        if (value !== undefined) out.letterSpacingPt = value / TWIPS_PER_PT;
        break;
      }
      case "caps":
      case "smallCaps": {
        // All caps outranks small caps when both are on, as in Word; an explicit
        // off (`w:val="0"`) clears what a style inherited.
        const on = onOffValue(attr(child, "val"));
        const kind = child.local === "caps" ? "all" : "small";
        if (on && (kind === "all" || out.caps !== "all")) out.caps = kind;
        else if (!on && out.caps === kind) delete out.caps;
        break;
      }
      default:
        break;
    }
  }
}

/** Resolve a `majorHAnsi`-style theme token through the shared DrawingML theme (see `themeFontFamily`). */
function themeTokenFamily(token: string | undefined, theme: DrawingTheme, arabic: boolean): string | undefined {
  if (token === undefined) return undefined;
  const lower = token.toLowerCase();
  return themeFontFamily(lower.startsWith("major") ? theme.fonts.major : theme.fonts.minor, lower.endsWith("bidi") ? "cs" : "latin", arabic);
}

function resolveRunFamily(props: ResolvedRunProps, text: string, theme: DrawingTheme): Pick<TextRunModel, "fontFamily" | "latinTypeface" | "complexScriptTypeface"> {
  // Whitespace follows its paragraph's script slot (same rule as the PPTX
  // parser): an inherited run must not lose its family merely because its
  // own text has no script characters. Callers pass paragraph text for
  // blank runs; see `parseRun`.
  const arabic = usesArabic(text) || props.rtl === true || props.hintCs === true;
  const ascii = themeTokenFamily(props.asciiTheme, theme, arabic) ?? props.ascii ?? themeTokenFamily(props.hAnsiTheme, theme, arabic) ?? props.hAnsi;
  const cs = themeTokenFamily(props.csTheme, theme, arabic) ?? props.cs;
  const out: Pick<TextRunModel, "fontFamily" | "latinTypeface" | "complexScriptTypeface"> = {};
  if (ascii !== undefined) out.latinTypeface = ascii;
  if (cs !== undefined) out.complexScriptTypeface = cs;
  const preferred = arabic
    ? cs ?? ascii ?? themeFontFamily(theme.fonts.minor, "cs", true)
    : ascii ?? cs ?? themeFontFamily(theme.fonts.minor, "latin", false);
  if (preferred !== undefined && preferred !== "") out.fontFamily = preferred;
  return out;
}

interface ResolvedParaProps {
  styleId?: string;
  jc?: string;
  bidi?: boolean;
  snapToGrid?: boolean;
  /** Paragraph-level default run size in points, from style `w:sz`. */
  fontSizePt?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  /** `w:beforeLines`/`w:afterLines` in lines, inherited apart from the points (see `paragraphSpacePt`). */
  spaceBeforeLines?: number;
  spaceAfterLines?: number;
  lineSpacing?: number;
  lineSpacingPt?: number;
  lineSpacingAtLeast?: true;
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndent?: number;
  outlineLvl?: number;
  numId?: string;
  ilvl?: number;
  shading?: string;
  pageBreakBefore?: boolean;
  keepNext?: boolean;
  borders?: DocxParagraphBorders;
  /** Tab stops merged through docDefaults → style chain → direct formatting. */
  tabs?: DocxTabStop[];
}

const TAB_LEADERS: Record<string, DocxTabStop["leader"]> = {
  dot: "dot",
  hyphen: "hyphen",
  underscore: "underscore",
  heavy: "underscore",
  middleDot: "middleDot",
};

function readTabStops(tabs: XmlElement, out: ResolvedParaProps): void {
  for (const tab of tabs.children) {
    if (tab.local !== "tab") continue;
    const pos = numberAttr(tab, "pos", "w:tab");
    const val = attr(tab, "val");
    if (pos === undefined) continue;
    const inches = twipsToInches(pos);
    const kept = (out.tabs ?? []).filter((stop) => Math.abs(stop.pos - inches) > 1e-6);
    // `clear` removes an inherited stop at the same position.
    if (val !== "clear" && val !== "bar" && val !== "num") {
      const leader = TAB_LEADERS[attr(tab, "leader") ?? ""];
      kept.push({
        pos: inches,
        align: val === "center" ? "center" : val === "right" || val === "end" ? "right" : val === "decimal" ? "decimal" : "left",
        ...(leader === undefined ? {} : { leader }),
      });
    }
    out.tabs = kept.sort((a, b) => a.pos - b.pos);
  }
}

function readParaProps(pPr: XmlElement | undefined, out: ResolvedParaProps): void {
  if (pPr === undefined) return;
  for (const child of pPr.children) {
    switch (child.local) {
      case "pStyle": {
        const value = attr(child, "val");
        if (value !== undefined) out.styleId = value;
        break;
      }
      case "jc": {
        const value = attr(child, "val");
        if (value !== undefined) out.jc = value;
        break;
      }
      case "bidi":
        out.bidi = onOffValue(attr(child, "val"));
        break;
      case "snapToGrid":
        out.snapToGrid = onOffValue(attr(child, "val"));
        break;
      case "keepNext":
        out.keepNext = onOffValue(attr(child, "val"));
        break;
      case "spacing": {
        const before = numberAttr(child, "before", "w:spacing");
        const after = numberAttr(child, "after", "w:spacing");
        const beforeLines = numberAttr(child, "beforeLines", "w:spacing");
        const afterLines = numberAttr(child, "afterLines", "w:spacing");
        const line = numberAttr(child, "line", "w:spacing");
        if (before !== undefined) out.spaceBeforePt = before / TWIPS_PER_PT;
        if (after !== undefined) out.spaceAfterPt = after / TWIPS_PER_PT;
        if (beforeLines !== undefined) out.spaceBeforeLines = beforeLines / 100;
        if (afterLines !== undefined) out.spaceAfterLines = afterLines / 100;
        if (line !== undefined) {
          // `w:lineRule="auto"` (or absent) is a 240ths multiplier;
          // `exact`/`atLeast` are twips. The authored value is preserved in
          // the model; the renderer clamps its CSS box (ARCHITECTURE § Format
          // rules: line-box clipping is a metrics defect). Each
          // level replaces the inherited rule outright: a style's multiple
          // over an exact docDefault is a multiple, not both.
          const rule = attr(child, "lineRule") ?? "auto";
          if (rule === "auto") {
            out.lineSpacing = line / SINGLE_LINE;
            delete out.lineSpacingPt;
            delete out.lineSpacingAtLeast;
          } else {
            out.lineSpacingPt = line / TWIPS_PER_PT;
            delete out.lineSpacing;
            if (rule === "atLeast") out.lineSpacingAtLeast = true;
            else delete out.lineSpacingAtLeast;
          }
        }
        break;
      }
      case "ind": {
        const left = numberAttr(child, "left", "w:ind");
        const right = numberAttr(child, "right", "w:ind");
        const firstLine = numberAttr(child, "firstLine", "w:ind");
        const hanging = numberAttr(child, "hanging", "w:ind");
        if (left !== undefined) out.leftIndent = twipsToInches(left);
        if (right !== undefined) out.rightIndent = twipsToInches(right);
        if (firstLine !== undefined) out.firstLineIndent = twipsToInches(firstLine);
        else if (hanging !== undefined) out.firstLineIndent = -twipsToInches(hanging);
        break;
      }
      case "outlineLvl": {
        const value = numberAttr(child, "val", "w:outlineLvl");
        if (value !== undefined) out.outlineLvl = value;
        break;
      }
      case "numPr": {
        const numId = firstChild(child, "numId");
        const ilvl = firstChild(child, "ilvl");
        const numValue = numId === undefined ? undefined : attr(numId, "val");
        const ilvlValue = ilvl === undefined ? undefined : attr(ilvl, "val");
        if (numValue !== undefined) out.numId = numValue;
        if (ilvlValue !== undefined) {
          const parsed = Number(ilvlValue);
          if (Number.isInteger(parsed) && parsed >= 0) out.ilvl = parsed;
        }
        break;
      }
      case "shd": {
        const fill = hexColor(attr(child, "fill"));
        if (fill !== undefined) out.shading = fill;
        break;
      }
      case "pBdr": {
        // Sides merge independently through the style chain: a direct left
        // border composes with a styled top border rather than replacing it.
        out.borders = out.borders ?? {};
        for (const edge of child.children) {
          if (edge.local !== "top" && edge.local !== "right" && edge.local !== "bottom" && edge.local !== "left") continue;
          const border = paragraphBorder(edge);
          if (border !== undefined) out.borders[edge.local] = border;
        }
        break;
      }
      case "pageBreakBefore":
        out.pageBreakBefore = onOffValue(attr(child, "val"));
        break;
      case "tabs":
        readTabStops(child, out);
        break;
      default:
        break;
    }
  }
}

/**
 * Resolve `w:jc` to a physical alignment. Word reads `left`/`right` as
 * start/end in a bidi paragraph (`<w:bidi/><w:jc w:val="right"/>` paints on
 * the left), so the two swap when the paragraph is RTL — verified against a
 * native Word export of the bidi-coverage fixture.
 */
function alignFor(jc: string | undefined, direction: TextDirection): DocxParagraphLayout["align"] {
  const rtl = direction === "rtl";
  switch (jc) {
    case undefined:
      return undefined;
    case "left":
    case "start":
      return rtl ? "right" : undefined;
    case "right":
    case "end":
      return rtl ? "left" : "right";
    case "center":
      return "center";
    case "both":
    case "distribute":
      return "justify";
    default:
      return undefined;
  }
}

/**
 * What a DOCX session keeps for its lifetime: small, and holding no part
 * trees and no string that still points into a part's text (a string sliced
 * from a part keeps the whole part alive — see `ownString`). The body and
 * the notes are read, streamed and dropped inside `parseDocxSections`.
 */
export interface DocxIndex {
  documentPart: string;
  /** Footnote and endnote part paths, read by `parseDocxSections`. */
  notesParts: { footnotes: string | undefined; endnotes: string | undefined };
  styles: StyleData;
  numbering: NumberingData;
  theme: DrawingTheme;
  /** Document-part relationships: `r:id` → resolved internal part path. */
  relTargets: Map<string, string>;
  /** External hyperlink targets: `r:id` → URL. */
  hyperlinks: Map<string, string>;
  /** Word lays the document out as 2013 or later (see `wordLayoutShrinksSpaces`). */
  justifyShrinks: boolean;
  /** `w:evenAndOddHeaders` of the settings: even page numbers take their own header and footer. */
  evenAndOddHeaders: boolean;
  /** Word applies section line grids inside table cells only when this setting is enabled. */
  adjustLineHeightInTable: boolean;
  /**
   * The document title from its core properties, if it sets one — the label
   * of DOCX's single unit. The heading outline is not an index concern: it
   * comes from the parsed sections (`document.headings`).
   */
  title: string | undefined;
}

interface ParseContext {
  styles: StyleData;
  numbering: NumberingData;
  notes: Map<string, XmlElement>;
  theme: DrawingTheme;
  relTargets: Map<string, string>;
  hyperlinks: Map<string, string>;
  /** MIME lookup for image parts (`OpcPackage.mediaContentType`); the worker media route serves the bytes. */
  mediaContentType: (part: string) => string;
  fillOwner: FillOwner;
  /** The part's charts, read ahead by part path (`readPartCharts`): drawings parse synchronously. */
  charts: PartCharts;
  footnoteLabels: Map<string, string>;
  footnoteOrder: string[];
  /** List counters by list definition (`abstractNum`), see `advanceCounters`. */
  counters: Map<string, number[]>;
  /**
   * The table style's `w:rPr`/`w:pPr` for the cell being parsed, set only over a
   * cell's content and restored after (see `parseTable`). They layer *below* the
   * paragraph and character styles: a `Heading 2` paragraph in a header row keeps
   * its own colour over the region's, while a region's bold the paragraph style
   * does not state still applies (probe `docx-table-style`).
   */
  tableStyleText?: { rPr: XmlElement[]; pPr: XmlElement[] } | undefined;
  /** The numbering ids already used in this part: a `w:startOverride` applies at a num's first use. */
  startedNums: Set<string>;
  justifyShrinks: boolean;
  adjustLineHeightInTable: boolean;
  /** Drawings read without a `wp:docPr/@id`, counted over every part of one parse (see `DocumentSection.missingIds`). */
  missing: { drawings: number };
  /**
   * Body-level table-of-contents tracking. Set while streaming `w:body`
   * (see `parseBodySections`); tables and text boxes disable it over their
   * content, and header, footer and note parts never enable it, so only a
   * body-level TOC is ever refreshed and anything else keeps its cache.
   */
  toc?: TocTracking | undefined;
  /** Body bookmarks, present while parsing the body (see `DocxBookmarkIndex`). */
  bookmarks?: DocxBookmarkIndex | undefined;
  signal?: AbortSignal;
}

type PartCharts = Map<string, ReturnType<typeof readChartModel>>;

/**
 * Every chart a part's relationships point at, read once before the part's
 * drawings are (`parseDrawing` is synchronous). Word, like Excel, has no
 * colour map for chart DrawingML: `tx1` and `bg1` take the default one.
 */
async function readPartCharts(pkg: OpcPackage, part: string, theme: DrawingTheme, signal?: AbortSignal): Promise<PartCharts> {
  const charts: PartCharts = new Map();
  let chartTheme: DrawingTheme | undefined;
  for (const relationship of await pkg.relationships(part, signal)) {
    // `/chart`, and `/chartEx` for an Office 2016 chart (waterfall, treemap, ...).
    if (relationship.targetMode === "External" || !(relationship.type.endsWith("/chart") || relationship.type.endsWith("/chartEx"))) continue;
    const root = await pkg.readXmlIfPresent(relationship.target, signal);
    throwIfAborted(signal);
    if (root === undefined) continue;
    chartTheme ??= themeForColorMap(theme, undefined);
    charts.set(relationship.target, readChartModel(root, chartTheme, { owner: await chartFillOwner(pkg, relationship.target, signal) }));
  }
  return charts;
}

/** A part's own relationships: ids are per part, so a header, a footnote or the body each resolve theirs. */
interface PartRels {
  relationships: FillOwner["relationships"];
  /** `r:id` → internal part path. */
  relTargets: Map<string, string>;
  /** `r:id` → external hyperlink URL. */
  hyperlinks: Map<string, string>;
}

async function readPartRels(pkg: OpcPackage, part: string, signal?: AbortSignal): Promise<PartRels> {
  const relationships = await pkg.relationships(part, signal);
  const rels: PartRels = { relationships, relTargets: new Map(), hyperlinks: new Map() };
  for (const relationship of relationships) {
    if (relationship.targetMode !== "External") rels.relTargets.set(relationship.id, relationship.target);
    else if (relationship.type.endsWith("/hyperlink")) rels.hyperlinks.set(relationship.id, relationship.target);
  }
  return rels;
}

function partContext(
  index: DocxIndex,
  notes: Map<string, XmlElement>,
  rels: PartRels,
  mediaContentType: (part: string) => string,
  charts: PartCharts,
  missing: { drawings: number },
  signal?: AbortSignal,
): ParseContext {
  return {
    styles: index.styles,
    numbering: index.numbering,
    notes,
    theme: index.theme,
    relTargets: rels.relTargets,
    hyperlinks: rels.hyperlinks,
    mediaContentType,
    fillOwner: { relationships: rels.relationships, mediaContentType, inherited: false },
    charts,
    footnoteLabels: new Map(),
    footnoteOrder: [],
    counters: new Map(),
    startedNums: new Set(),
    justifyShrinks: index.justifyShrinks,
    adjustLineHeightInTable: index.adjustLineHeightInTable,
    missing,
    ...(signal === undefined ? {} : { signal }),
  };
}

/**
 * The parts a section's own header or footer references name, by `w:type`
 * (absent means `default`). Which one a page shows is decided at layout, from
 * `w:titlePg` and `w:evenAndOddHeaders` (see `headerFooterFor` in pagination).
 */
function headerFooterParts(
  sectPr: XmlElement | undefined,
  local: "headerReference" | "footerReference",
  relTargets: Map<string, string>,
): HeaderFooterParts | undefined {
  if (sectPr === undefined) return undefined;
  const parts: HeaderFooterParts = {};
  for (const ref of childElements(sectPr, local)) {
    const type = attr(ref, "type") ?? "default";
    if (type !== "default" && type !== "first" && type !== "even") continue;
    const target = relTargets.get(attr(ref, "id") ?? "");
    if (target !== undefined && parts[type] === undefined) parts[type] = target;
  }
  return Object.keys(parts).length > 0 ? parts : undefined;
}

/**
 * Index a DOCX without reading its body or inflating drawing media:
 * styles, numbering, theme, relationships and title. `document.xml` and the
 * notes are read by `parseDocxSections`; media bytes stay in the package
 * until a renderer asks for them through the media route.
 */
export async function indexDocx(pkg: OpcPackage, signal?: AbortSignal): Promise<DocxIndex> {
  throwIfAborted(signal);
  const documentPart = await pkg.officeDocument("docx", signal);
  throwIfAborted(signal);

  const relationships = await pkg.relationships(documentPart, signal);
  throwIfAborted(signal);
  const relTargets = new Map<string, string>();
  const hyperlinks = new Map<string, string>();
  let stylesPart: string | undefined;
  let numberingPart: string | undefined;
  let footnotesPart: string | undefined;
  let endnotesPart: string | undefined;
  let settingsPart: string | undefined;
  for (const relationship of relationships) {
    throwIfAborted(signal);
    if (relationship.targetMode === "External") {
      if (relationship.type.endsWith("/hyperlink")) hyperlinks.set(relationship.id, relationship.target);
      continue;
    }
    relTargets.set(relationship.id, relationship.target);
    if (relationship.type.endsWith("/styles")) stylesPart = relationship.target;
    else if (relationship.type.endsWith("/numbering")) numberingPart = relationship.target;
    else if (relationship.type.endsWith("/footnotes")) footnotesPart = relationship.target;
    else if (relationship.type.endsWith("/endnotes")) endnotesPart = relationship.target;
    else if (relationship.type.endsWith("/settings")) settingsPart = relationship.target;
  }

  const styles = parseStyles(
    stylesPart === undefined ? undefined : await pkg.readXmlIfPresent(stylesPart, signal),
    signal,
  );
  const numbering = parseNumbering(
    numberingPart === undefined ? undefined : await pkg.readXmlIfPresent(numberingPart, signal),
    signal,
  );
  const theme = await readDrawingTheme(pkg, signal);
  throwIfAborted(signal);
  const settings = settingsPart === undefined ? undefined : await pkg.readXmlIfPresent(settingsPart, signal);
  const justifyShrinks = wordLayoutShrinksSpaces(settings);
  const adjustLineHeightInTable = wordAdjustsLineHeightInTable(settings);
  const evenAndOddHeaders = onOffElement(settings === undefined ? undefined : firstChild(settings, "evenAndOddHeaders")) === true;
  throwIfAborted(signal);
  const title = await pkg.coreTitle(signal);
  throwIfAborted(signal);
  return {
    documentPart,
    notesParts: { footnotes: footnotesPart, endnotes: endnotesPart },
    styles,
    numbering,
    theme,
    relTargets,
    hyperlinks,
    justifyShrinks,
    evenAndOddHeaders,
    adjustLineHeightInTable,
    title: title === undefined ? undefined : ownString(title),
  };
}

/** `w:compatibilityMode` of Word 2013, the layout that shrinks justified spaces. */
const WORD_2013_MODE = 15;

/**
 * Whether Word lays this document out with its 2013 engine, which fits one
 * more word on a justified line by taking back part of each space (probe
 * docx-justify: up to about a quarter of a space, and about a third of the
 * word it fits; none left-aligned, none in Word 2010 mode). The mode is
 * `w:compat/w:compatSetting[@w:name="compatibilityMode"]`; a document
 * without it predates 2013.
 */
function wordLayoutShrinksSpaces(settings: XmlElement | undefined): boolean {
  const compat = settings === undefined ? undefined : firstChild(settings, "compat");
  if (compat === undefined) return false;
  for (const setting of childElements(compat, "compatSetting")) {
    if (attr(setting, "name") !== "compatibilityMode") continue;
    const mode = Number(attr(setting, "val"));
    return Number.isFinite(mode) && mode >= WORD_2013_MODE;
  }
  return false;
}

/** `w:adjustLineHeightInTable` opts table cells into the section's line grid. */
function wordAdjustsLineHeightInTable(settings: XmlElement | undefined): boolean {
  const compat = settings === undefined ? undefined : firstChild(settings, "compat");
  return compat !== undefined && onOffElement(firstChild(compat, "adjustLineHeightInTable")) === true;
}

/** Footnote and endnote entries by `f`/`e`-prefixed id (separators excluded), for one section parse. */
async function readNotes(pkg: OpcPackage, index: DocxIndex, signal?: AbortSignal): Promise<Map<string, XmlElement>> {
  const notes = new Map<string, XmlElement>();
  for (const [part, prefix, local] of [
    [index.notesParts.footnotes, "f", "footnote"],
    [index.notesParts.endnotes, "e", "endnote"],
  ] as const) {
    const parsed = part === undefined ? undefined : await pkg.readXmlIfPresent(part, signal);
    throwIfAborted(signal);
    for (const [id, note] of parseNotes(parsed, prefix, local, signal)) notes.set(id, note);
  }
  return notes;
}

type BlockDraft =
  | { node: FlowNode }
  | { listItem: { numId: string; item: DocxListItemNode } }
  /**
   * A table-of-contents field's drafts: `begin` opens a refreshable range
   * (carrying the live session), each `cached` wraps one cached result
   * paragraph (dropped on refresh, unwrapped on fallback), and `end` closes
   * the range. All markers are spliced out before list grouping, so the
   * public model never sees them.
   */
  | { toc: { kind: "begin"; session: TocSession } | { kind: "cached"; inner: { node: FlowNode } } | { kind: "end" } };

interface RunDraft {
  runs: TextRunModel[];
  blocks: Array<DocxImageNode | DocxChartNode | DocxTextBoxNode | DocxUnsupportedNode>;
  /** Pictures and charts anchored (`wp:anchor`) in the run: they belong to the paragraph, not to a position in its text. */
  floats: DocxDrawingNode[];
  /**
   * Where an explicit page break (`w:br w:type="page"`) fell in the run: how
   * many of `runs` came before it. The paragraph splits there, so text after
   * the break in the same run lands on the new page. Absent: no break.
   */
  pageBreakAt?: number;
  softBreak: boolean;
  /** Top-level field instruction chunk seen in this run, for TOC detection (see `parseParagraph`). */
  tocInstrChunk?: string | undefined;
  /** Whether this run saw the open field's `separate` at the top level. */
  tocSeparated?: boolean;
  /** A complex-field end with no local frame while a TOC is open: the TOC's own end. */
  tocClosed?: boolean;
}

/**
 * One open complex field (`fldChar` begin → separate → end). The frames span
 * the sibling runs of a paragraph, since Word splits a field across runs.
 */
interface FieldFrame {
  /** The live PAGE/NUMPAGES run, once its instruction has been read. */
  live?: TextRunModel;
  /** True between `separate` and `end`: the runs there are the cached result. */
  inResult: boolean;
  /** Text of the cached result runs, diverted off the flow while `live` is set. */
  cached: string;
}

/**
 * A supported table-of-contents field's selection and presentation, parsed
 * from its `w:instrText` (`{ TOC \h \o "1-2" }`). Levels are Word outline
 * levels 1-9; entry levels (`TOC1` styles) count from `minLevel`.
 */
interface TocInstruction {
  /** Entries link to their headings (`\h`); without it they are plain text with page numbers. */
  hyperlink: boolean;
  /** Outline levels collected (`\o "1-2"`; absent is every level). */
  minLevel: number;
  maxLevel: number;
  /** Entry/page separator (`\p "x"`); absent is Word's default tab with dot leaders. */
  separator?: string;
  /** Entry levels without page numbers (`\n`, optionally `"m-n"`); absent keeps them all. */
  omitPageMin: number;
  omitPageMax: number;
  /**
   * The named bookmark a field collects from (`\b Part1`), resolved against
   * `DocxBookmarkIndex.ranges` at splice time. A name the document does not
   * hold makes the field unsupportable, so it keeps its cache rather than
   * listing the whole document (see the refresh rule on `parseTocInstruction`).
   */
  bookmark?: string;
}

/** One cached TOC entry paragraph: its entry text and its cached page number. */
interface TocCachedEntry {
  entry: string;
  page: string;
}

/** A bookmark's half-open range in body block-index space: `[from, to)`. */
interface BookmarkRange {
  from: number;
  to: number;
}

/** One bookmark boundary, queued until its paragraph's block index is known. */
type BookmarkEvent =
  | { kind: "start"; id: string; name: string }
  | { kind: "end"; id: string };

/**
 * Where the body bookmarks sit, recorded while `w:body` streams past.
 *
 * A bookmark is a named label around a piece of content — `w:bookmarkStart`
 * with a name, `w:bookmarkEnd` with the same `w:id` — and it is what a TOC's
 * `\b` switch collects entries from. Word's own `_Toc…` bookmarks (the ones it
 * puts on headings so its contents lines can link) are ignored here: aared
 * mints its own fragments for those.
 *
 * Ranges are recorded in body **block-index** space, the index a block takes in
 * its section's draft list, because a heading is exactly one block. Word merges
 * several ranges sharing a name into one, so a name maps to a list. A range
 * that would cross a section break is dropped rather than paired across two
 * index spaces, so a `\b` field either resolves honestly or keeps its cache.
 */
interface DocxBookmarkIndex {
  ranges: Map<string, BookmarkRange[]>;
  /** Open `bookmarkStart`s by `w:id`, until their end arrives. */
  open: Map<string, { name: string; from: number; section: number }>;
  /** Boundaries seen inside the paragraph being parsed, drained by its caller. */
  events: BookmarkEvent[];
  /** Section the open ranges were recorded in, so a cross-section pair is dropped. */
  section: number;
}

/** One open table-of-contents field while its cached paragraphs stream past. */
interface TocSession {
  /**
   * The refresh plan, or undefined when the switches need the saved result.
   * Re-read as split instruction chunks accumulate; a session whose final
   * instruction is unsupported unwraps back to its cache at the splice.
   */
  instruction: TocInstruction | undefined;
  /** Instruction text so far, joined from its top-level chunks. */
  raw: string;
  /** Whether the instruction's `separate` has been seen. */
  separated: boolean;
  /** Cached entry paragraphs seen so far, for page-number carry-over on refresh. */
  cached: TocCachedEntry[];
  /** Tab stops of the first cached entry paragraph, reused for refreshed entries. */
  tabStops?: DocxTabStop[];
}

/** Body-level TOC tracking. Present only while parsing body content (never in tables, text boxes, headers or notes). */
interface TocTracking {
  /** The open TOC field, while its cached result paragraphs stream past. */
  active?: TocSession | undefined;  /** Count of `tocAnchor` fragments handed out, keeping them unique per document. */
  serial: number;
}

function pageBreakNode(strength: DocxPageBreakNode["strength"]): DocxPageBreakNode {
  return { type: "pageBreak", strength };
}

function footnoteLabel(ctx: ParseContext, key: string): string {
  const existing = ctx.footnoteLabels.get(key);
  if (existing !== undefined) return existing;
  const label = String(ctx.footnoteOrder.length + 1);
  ctx.footnoteLabels.set(key, label);
  ctx.footnoteOrder.push(key);
  return label;
}

function baseRunProps(ctx: ParseContext, styleId: string | undefined, charStyleId: string | undefined): ResolvedRunProps {
  const out: ResolvedRunProps = {};
  readRunProps(ctx.styles.defaultRPr, out, ctx.theme);
  // A table style's regions sit below the paragraph and character styles: they
  // supply what those leave unsaid, and lose to anything they do state.
  for (const rPr of ctx.tableStyleText?.rPr ?? []) readRunProps(rPr, out, ctx.theme);
  for (const style of paragraphAncestry(ctx.styles, styleId)) readRunProps(style.rPr, out, ctx.theme);
  for (const style of characterAncestry(ctx.styles, charStyleId)) readRunProps(style.rPr, out, ctx.theme);
  return out;
}

function finalizeRun(
  ctx: ParseContext,
  props: ResolvedRunProps,
  paragraphText: string,
  text: string,
  extra?: Partial<TextRunModel>,
): TextRunModel | undefined {
  if (text === "") return undefined;
  const run: TextRunModel = { text };
  const arabic = usesArabic(text);
  // Word styles a complex-script run (Arabic text, or any run marked `w:rtl`)
  // with `bCs`/`iCs` only and every other run with `b`/`i` only: a run with
  // just `<w:b/>` is not bold in Arabic, and `bCs` alone never bolds Latin text
  // (probe `docx-cs-bold`, Word 16: all eight b/bCs/i/iCs combinations, both scripts).
  const complex = arabic || props.rtl === true;
  const bold = complex ? props.boldCs : props.bold;
  const italic = complex ? props.italicCs : props.italic;
  if (bold !== undefined) run.bold = bold;
  if (italic !== undefined) run.italic = italic;
  const size = arabic && props.fontSizeCsPt !== undefined ? props.fontSizeCsPt : props.fontSizePt;
  if (size !== undefined) run.fontSizePt = size;
  const language = props.languageBidi ?? props.language;
  if (language !== undefined) run.language = language;
  if (props.color !== undefined) run.color = props.color;
  // Blank runs (tabs, breaks) follow the paragraph slot, like the PPTX rule.
  const familyText = text.trim() === "" ? paragraphText : text;
  Object.assign(run, resolveRunFamily(props, familyText, ctx.theme));
  if (props.underline === true) run.underline = true;
  if (props.strikethrough === true) run.strikethrough = true;
  if (props.verticalAlign !== undefined) run.verticalAlign = props.verticalAlign;
  if (props.highlight !== undefined) run.highlight = props.highlight;
  if (props.shading !== undefined) run.shading = props.shading;
  if (props.letterSpacingPt !== undefined && props.letterSpacingPt !== 0) run.letterSpacingPt = props.letterSpacingPt;
  if (props.caps !== undefined) run.caps = props.caps;
  if (extra !== undefined) Object.assign(run, extra);
  return run;
}

/** The `w:drawing` of an Office 2016 chart: inside `mc:AlternateContent/mc:Choice`, its graphic data a `cx:chart`. */
function chartExDrawing(alternate: XmlElement): XmlElement | undefined {
  for (const choice of alternate.children) {
    if (choice.local !== "Choice") continue;
    const drawing = firstChild(choice, "drawing");
    const placement = drawing === undefined ? undefined : firstChild(drawing, "inline") ?? firstChild(drawing, "anchor");
    const graphic = placement === undefined ? undefined : firstChild(placement, "graphic");
    const data = graphic === undefined ? undefined : firstChild(graphic, "graphicData");
    if (drawing !== undefined && data !== undefined && firstChild(data, "chart") !== undefined) return drawing;
  }
  return undefined;
}

/** One supported branch of `mc:AlternateContent`; never read Choice and Fallback together. */
function alternateDrawing(alternate: XmlElement): { drawing?: XmlElement; pict?: XmlElement } | undefined {
  for (const branch of alternate.children) {
    if (branch.local === "Choice") {
      const drawing = descendant(branch, "drawing");
      if (drawing !== undefined) return { drawing };
    }
  }
  const fallback = alternate.children.find((child) => child.local === "Fallback");
  if (fallback === undefined) return undefined;
  const drawing = descendant(fallback, "drawing");
  if (drawing !== undefined) return { drawing };
  const pict = descendant(fallback, "pict");
  return pict === undefined ? undefined : { pict };
}

function parseDrawing(drawing: XmlElement, ctx: ParseContext): DocxImageNode | DocxChartNode | DocxTextBoxNode | DocxUnsupportedNode | undefined {
  throwIfAborted(ctx.signal);
  const inline = firstChild(drawing, "inline");
  const anchor = firstChild(drawing, "anchor");
  const placement = inline ?? anchor;
  if (placement === undefined) return undefined;
  const extent = firstChild(placement, "extent");
  const widthEmu = extent === undefined ? undefined : numberAttr(extent, "cx", "wp:extent");
  const heightEmu = extent === undefined ? undefined : numberAttr(extent, "cy", "wp:extent");
  const overflow = firstChild(placement, "effectExtent");
  const effectInsets = overflow === undefined ? undefined : {
    left: emuToInches(numberAttr(overflow, "l", "wp:effectExtent") ?? 0),
    top: emuToInches(numberAttr(overflow, "t", "wp:effectExtent") ?? 0),
    right: emuToInches(numberAttr(overflow, "r", "wp:effectExtent") ?? 0),
    bottom: emuToInches(numberAttr(overflow, "b", "wp:effectExtent") ?? 0),
  };
  // `a:graphic` nests inside `wp:inline`/`wp:anchor`, never directly under
  // `w:drawing` — a direct-child lookup misses every inline picture.
  const graphic = firstChild(placement, "graphic");
  const graphicData = graphic === undefined ? undefined : firstChild(graphic, "graphicData");
  const uri = graphicData === undefined ? undefined : attr(graphicData, "uri");
  const pic = graphicData === undefined ? undefined : firstChild(graphicData, "pic");
  const chart = graphicData === undefined
    ? undefined
    : graphicData.children.find((child) => child.local === "chart");
  const docPr = firstChild(placement, "docPr");
  // Word calls a drawing with no `wp:docPr/@id` damaged; nothing here uses the id, so it is only counted.
  if (docPr === undefined || attr(docPr, "id") === undefined) ctx.missing.drawings++;
  const alt = docPr === undefined
    ? undefined
    : attr(docPr, "descr") || attr(docPr, "title") || undefined;
  const shape = graphicData === undefined ? undefined : descendant(graphicData, "wsp");
  const textbox = shape === undefined ? undefined : firstChild(shape, "txbx");
  const textContent = textbox === undefined ? undefined : firstChild(textbox, "txbxContent");
  if (textContent !== undefined && widthEmu !== undefined && heightEmu !== undefined) {
    // The shape that owns this text, then its own `spPr` and `bodyPr`. Deep
    // searches are wrong for both: `wps:bodyPr` follows `wps:txbx`, so a box
    // inside a box resolved the inner shape's insets, and the two `spPr` walks
    // were repeated for one answer. The first `wps:wsp` in document order is the
    // outer box, which is the one whose extent this branch paints.
    const shapeProperties = shape === undefined ? undefined : firstChild(shape, "spPr");
    const bodyProperties = shape === undefined ? undefined : firstChild(shape, "bodyPr");
    // Text-box content parses untracked, as table cells do: a TOC field
    // never lives there, and tracking one would corrupt the body's.
    const savedToc = ctx.toc;
    ctx.toc = undefined;
    let content: FlowNode[];
    try {
      content = groupListItems(parseBlocks(textContent, ctx));
    } finally {
      ctx.toc = savedToc;
    }
    const drawingTheme = themeForColorMap(ctx.theme, undefined);
    const fill = shapeProperties === undefined ? undefined : resolveDirectFill(shapeProperties, drawingTheme, ctx.fillOwner);
    const effects = readDrawingEffects(shapeProperties, drawingTheme);
    const border = drawingBorder(shapeProperties, drawingTheme);
    const presetGeometry = shapeProperties === undefined ? undefined : firstChild(shapeProperties, "prstGeom");
    const preset = presetGeometry === undefined ? undefined : attr(presetGeometry, "prst");
    const adjustments = drawingGeometryAdjustments(presetGeometry);
    const xfrm = shapeProperties === undefined ? undefined : firstChild(shapeProperties, "xfrm");
    const localExtent = xfrm === undefined ? undefined : firstChild(xfrm, "ext");
    const size = {
      width: (localExtent === undefined ? undefined : numberAttr(localExtent, "cx", "a:ext")) ?? widthEmu,
      height: (localExtent === undefined ? undefined : numberAttr(localExtent, "cy", "a:ext")) ?? heightEmu,
    };
    const geometry = geometryForPreset(preset, size, adjustments)
      ?? unresolvedGeometry(size, preset, adjustments, shapeProperties === undefined ? undefined : firstChild(shapeProperties, "custGeom"));
    // DOCX has no preset catalog/evaluator route. Keep the marked bounds
    // fallback, but do not expose PPTX's internal session transport field.
    delete geometry.deferred;
    const transform = drawingShapeTransform(xfrm);
    const verticalAnchor = bodyProperties === undefined ? undefined : attr(bodyProperties, "anchor");
    const wrapping = bodyProperties === undefined ? undefined : attr(bodyProperties, "wrap");
    // Word text stays at its w:sz/w:szCs sizes. normAutofit does not apply
    // PowerPoint's fontScale or line-spacing reduction to w:txbxContent.
    const resize = bodyProperties !== undefined && firstChild(bodyProperties, "spAutoFit") !== undefined;
    return {
      type: "textBox",
      content,
      width: emuToInches(widthEmu),
      height: emuToInches(heightEmu),
      direction: footnoteDirection(content),
      ...(bodyProperties === undefined ? {} : { insets: {
        left: emuToInches(numberAttr(bodyProperties, "lIns", "wps:bodyPr") ?? 91440),
        right: emuToInches(numberAttr(bodyProperties, "rIns", "wps:bodyPr") ?? 91440),
        top: emuToInches(numberAttr(bodyProperties, "tIns", "wps:bodyPr") ?? 45720),
        bottom: emuToInches(numberAttr(bodyProperties, "bIns", "wps:bodyPr") ?? 45720),
      } }),
      ...(fill == null ? {} : { fill }),
      ...(border === undefined ? {} : { border }),
      geometry,
      ...(transform === undefined ? {} : { transform }),
      ...(effects === undefined ? {} : { effects, ...(effectInsets === undefined ? {} : { effectInsets }) }),
      verticalAlign: verticalAnchor === "ctr" ? "middle" : verticalAnchor === "b" ? "bottom" : "top",
      wrapping: wrapping === "none" ? "none" : "square",
      autofit: resize ? "resizeShape" : "none",
      ...(anchor === undefined ? {} : { anchor: parseAnchor(anchor) }),
    };
  }
  if (chart !== undefined || (uri !== undefined && uri.endsWith("/chart"))) {
    // The shared chart reader already read the part (`readPartCharts`); a
    // chart it cannot draw, or one with no part or extent, is a marked fallback.
    const rid = chart === undefined ? undefined : attr(chart, "id");
    const target = rid === undefined ? undefined : ctx.relTargets.get(rid);
    const read = target === undefined ? undefined : ctx.charts.get(target);
    if (read === undefined || !("spec" in read) || widthEmu === undefined || heightEmu === undefined) {
      return { type: "unsupported", source: "chart", label: alt ?? "Chart" };
    }
    return {
      type: "chart",
      spec: read.spec,
      width: emuToInches(widthEmu),
      height: emuToInches(heightEmu),
      ...(alt === undefined ? {} : { alt }),
      ...(anchor === undefined ? {} : { anchor: parseAnchor(anchor) }),
    };
  }
  if (pic === undefined || widthEmu === undefined || heightEmu === undefined) {
    return { type: "unsupported", source: "drawing", label: alt ?? "Drawing" };
  }
  const blipFill = firstChild(pic, "blipFill");
  // An embedded SVG (`asvg:svgBlip`) paints first; the main `blip` stays as
  // its raster fallback. Part contexts hold no package, so existence is left
  // to the renderer's fallback.
  const source = resolveImageSource(blipFill === undefined ? undefined : firstChild(blipFill, "blip"), {
    relationships: ctx.fillOwner.relationships.map((relationship) => ({
      id: relationship.id,
      type: relationship.type,
      target: relationship.target,
      targetMode: relationship.targetMode,
    })),
    mediaContentType: ctx.mediaContentType,
  });
  if (source === undefined) return { type: "unsupported", source: "drawing", label: alt ?? "Drawing" };
  const effects = readDrawingEffects(firstChild(pic, "spPr"), themeForColorMap(ctx.theme, undefined));
  return {
    type: "image",
    mediaId: source.mediaId,
    contentType: source.contentType,
    ...(source.fallbackMediaId === undefined || source.fallbackContentType === undefined ? {} : {
      fallbackMediaId: source.fallbackMediaId,
      fallbackContentType: source.fallbackContentType,
    }),
    width: emuToInches(widthEmu),
    height: emuToInches(heightEmu),
    ...(alt === undefined ? {} : { alt }),
    ...(effects === undefined ? {} : { effects, ...(effectInsets === undefined ? {} : { effectInsets }) }),
    ...(anchor === undefined ? {} : { anchor: parseAnchor(anchor) }),
  };
}

/** First descendant with the requested local name, in XML order. */
function descendant(root: XmlElement, local: string): XmlElement | undefined {
  const stack = [...root.children].reverse();
  while (stack.length > 0) {
    const node = stack.pop()!;
    if (node.local === local) return node;
    for (let index = node.children.length - 1; index >= 0; index--) stack.push(node.children[index]!);
  }
  return undefined;
}

function drawingBorder(properties: XmlElement | undefined, theme: DrawingTheme): DocxParagraphBorder | undefined {
  const line = properties === undefined ? undefined : firstChild(properties, "ln");
  if (line === undefined || directFillProperty(line)?.local === "noFill") return undefined;
  const fill = directFillProperty(line);
  const color = fill?.local === "solidFill" ? resolveDrawingColor(fill, theme) : undefined;
  const width = numberAttr(line, "w", "a:ln");
  const dashElement = firstChild(line, "prstDash");
  const dash = dashElement === undefined ? undefined : attr(dashElement, "val");
  const style = dash === "dot" || dash === "sysDot" ? "dotted"
    : dash !== undefined && ["dash", "sysDash", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot", "sysDashDot", "sysDashDotDot"].includes(dash) ? "dashed" : "solid";
  return { color: color ?? "currentColor", widthPt: width === undefined ? 0.75 : width / 12700, spacePt: 0, style };
}

const ANCHOR_WRAPS = {
  wrapSquare: "square",
  wrapTight: "tight",
  wrapThrough: "through",
  wrapTopAndBottom: "topAndBottom",
  wrapNone: "none",
} as const;

/** Word's wrapping polygons live in a square of this many units on each side. */
const WRAP_POLYGON_UNITS = 21600;

function anchorPosition<Align extends string>(position: XmlElement | undefined, aligns: readonly Align[]): DocxAnchorPosition<Align> {
  const out: DocxAnchorPosition<Align> = { from: (position === undefined ? undefined : attr(position, "relativeFrom")) ?? "column" };
  const align = position === undefined ? undefined : firstChild(position, "align");
  const alignment = align === undefined ? undefined : (attr(align, "val") ?? align.text).trim();
  if (alignment !== undefined && (aligns as readonly string[]).includes(alignment)) out.align = alignment as Align;
  const offset = position === undefined ? undefined : firstChild(position, "posOffset");
  const emu = offset === undefined ? Number.NaN : Number(offset.text.trim());
  if (Number.isFinite(emu)) out.offset = emuToInches(emu);
  return out;
}

/**
 * A picture's placement from its `wp:anchor` (ECMA-376 20.4.2.3): exactly one
 * wrap element, a horizontal and a vertical position, the text distances
 * (Word writes them on the anchor; a wrap element may state its own), and for
 * tight and through wrapping the polygon in Word's 21600-unit square. Ground
 * truth: the anchors Word saves for each wrapping case (probe docx-float-wrap).
 */
function parseAnchor(anchor: XmlElement): DocxImageAnchor {
  const wrapName = Object.keys(ANCHOR_WRAPS).find((name) => firstChild(anchor, name) !== undefined) as keyof typeof ANCHOR_WRAPS | undefined;
  const wrapElement = wrapName === undefined ? undefined : firstChild(anchor, wrapName);
  const distance = (name: "distT" | "distB" | "distL" | "distR"): number => {
    const own = wrapElement === undefined ? undefined : numberAttr(wrapElement, name, `wp:${wrapName}`);
    return emuToInches(own ?? numberAttr(anchor, name, "wp:anchor") ?? 0);
  };
  const wrapText = wrapElement === undefined ? undefined : attr(wrapElement, "wrapText");
  const out: DocxImageAnchor = {
    wrap: wrapName === undefined ? "none" : ANCHOR_WRAPS[wrapName],
    behind: ["1", "true"].includes(attr(anchor, "behindDoc") ?? ""),
    wrapText: wrapText === "left" || wrapText === "right" || wrapText === "largest" ? wrapText : "bothSides",
    horizontal: anchorPosition(firstChild(anchor, "positionH"), ["left", "center", "right", "inside", "outside"] as const),
    vertical: anchorPosition(firstChild(anchor, "positionV"), ["top", "center", "bottom", "inside", "outside"] as const),
    distance: { top: distance("distT"), bottom: distance("distB"), left: distance("distL"), right: distance("distR") },
  };
  const polygon = wrapElement === undefined ? undefined : firstChild(wrapElement, "wrapPolygon");
  if (polygon !== undefined && (wrapName === "wrapTight" || wrapName === "wrapThrough")) {
    const points: Array<[number, number]> = [];
    for (const point of polygon.children) {
      if (point.local !== "start" && point.local !== "lineTo") continue;
      const x = Number(attr(point, "x"));
      const y = Number(attr(point, "y"));
      if (Number.isFinite(x) && Number.isFinite(y)) points.push([x / WRAP_POLYGON_UNITS, y / WRAP_POLYGON_UNITS]);
    }
    if (points.length >= 3) out.polygon = points;
  }
  return out;
}

function parseVmlPicture(pict: XmlElement, ctx: ParseContext): DocxImageNode | DocxUnsupportedNode | undefined {
  throwIfAborted(ctx.signal);
  const stack = [...pict.children];
  let imageData: XmlElement | undefined;
  let shape: XmlElement | undefined;
  while (stack.length > 0) {
    const current = stack.pop()!;
    if (current.local === "imagedata" && imageData === undefined) imageData = current;
    if ((current.local === "shape" || current.local === "rect") && shape === undefined) shape = current;
    for (const child of current.children) stack.push(child);
  }
  const rid = imageData === undefined ? undefined : attr(imageData, "id");
  const target = rid === undefined ? undefined : ctx.relTargets.get(rid);
  if (target === undefined) return { type: "unsupported", source: "drawing", label: "Picture" };
  const style = shape === undefined ? undefined : attr(shape, "style");
  const width = style === undefined ? undefined : /width:([\d.]+)pt/.exec(style);
  const height = style === undefined ? undefined : /height:([\d.]+)pt/.exec(style);
  if (width === null || width === undefined || height === null || height === undefined) {
    return { type: "unsupported", source: "drawing", label: "Picture" };
  }
  return {
    type: "image",
    mediaId: target,
    contentType: ctx.mediaContentType(target),
    width: ptToInches(Number(width[1])),
    height: ptToInches(Number(height[1])),
  };
}

function parseVmlTextBox(pict: XmlElement, ctx: ParseContext): DocxTextBoxNode | undefined {
  const content = descendant(pict, "txbxContent");
  if (content === undefined) return undefined;
  const shape = descendant(pict, "shape") ?? descendant(pict, "rect");
  const style = shape === undefined ? "" : attr(shape, "style") ?? "";
  const cssLength = (name: string, fallbackPt: number): number => {
    const match = new RegExp(`(?:^|;)\\s*${name}:\\s*(-?[\\d.]+)(pt|in|px)?`, "i").exec(style);
    if (match === null) return ptToInches(fallbackPt);
    const value = Number(match[1]);
    return match[2]?.toLowerCase() === "in" ? value : match[2]?.toLowerCase() === "px" ? value / 96 : ptToInches(value);
  };
  const textbox = descendant(pict, "textbox");
  const insets = (textbox === undefined ? undefined : attr(textbox, "inset") ?? "7.2pt,3.6pt,7.2pt,3.6pt")
    ?? "7.2pt,3.6pt,7.2pt,3.6pt";
  const insetValues = insets
    .split(",").map((part) => Number.parseFloat(part));
  const inches = (value: number | undefined, fallback: number): number => value === undefined || !Number.isFinite(value) ? fallback : ptToInches(value);
  // Legacy text-box content parses untracked, as drawings above do.
  const savedToc = ctx.toc;
  ctx.toc = undefined;
  let nodes: FlowNode[];
  try {
    nodes = groupListItems(parseBlocks(content, ctx));
  } finally {
    ctx.toc = savedToc;
  }
  const filled = shape === undefined ? undefined : attr(shape, "filled");
  const vmlColor = (value: string | undefined): string | undefined => {
    const channels = value === undefined ? undefined : parseHexColor(value);
    return channels === undefined ? undefined : toHexColor(channels);
  };
  const fill = filled === "f" || filled === "false" || filled === "0" ? undefined : vmlColor(shape === undefined ? undefined : attr(shape, "fillcolor"));
  const stroke = shape === undefined ? undefined : attr(shape, "stroked");
  const borderColor = vmlColor(shape === undefined ? undefined : attr(shape, "strokecolor"));
  const weight = shape === undefined ? undefined : attr(shape, "strokeweight");
  const borderWidth = weight === undefined ? 0.75 : Number.parseFloat(weight);
  const left = /(?:^|;)\s*margin-left:\s*(-?[\d.]+)(pt|in|px)?/i.exec(style);
  const top = /(?:^|;)\s*margin-top:\s*(-?[\d.]+)(pt|in|px)?/i.exec(style);
  const coordInches = (match: RegExpExecArray | null): number | undefined => {
    if (match === null) return undefined;
    const value = Number(match[1]);
    return match[2]?.toLowerCase() === "in" ? value : match[2]?.toLowerCase() === "px" ? value / 96 : ptToInches(value);
  };
  const absolute = /(?:^|;)\s*position:\s*absolute\b/i.test(style);
  const leftOffset = absolute ? coordInches(left) : undefined;
  const topOffset = absolute ? coordInches(top) : undefined;
  return {
    type: "textBox",
    content: nodes,
    width: cssLength("width", 216),
    height: cssLength("height", 72),
    direction: footnoteDirection(nodes),
      insets: { left: inches(insetValues[0], 0.1), top: inches(insetValues[1], 0.05), right: inches(insetValues[2], 0.1), bottom: inches(insetValues[3], 0.05) },
    ...(fill === undefined ? {} : { fill: { kind: "solid" as const, color: fill } }),
    ...(stroke === "f" || stroke === "false" || borderColor === undefined ? {} : { border: { color: borderColor, widthPt: Number.isFinite(borderWidth) ? borderWidth : 0.75, spacePt: 0, style: "solid" as const } }),
    legacy: true,
    ...(leftOffset === undefined && topOffset === undefined ? {} : { anchor: {
      wrap: "square" as const, behind: false, wrapText: "bothSides" as const,
      horizontal: { from: "column", ...(leftOffset === undefined ? {} : { offset: leftOffset }) },
      vertical: { from: "paragraph", ...(topOffset === undefined ? {} : { offset: topOffset }) },
      distance: { top: 0, bottom: 0, left: 0, right: 0 },
    } }),
  };
}

function parseOleObject(object: XmlElement): DocxUnsupportedNode {
  const progId = attr(object, "progId") ?? attr(object, "progID");
  return { type: "unsupported", source: "ole", label: progId ?? "Embedded object" };
}

/** Collect one run's text plus any drawing blocks it carries. */
function parseRunContent(
  run: XmlElement,
  ctx: ParseContext,
  props: ResolvedRunProps,
  paragraphText: string,
  href: string | undefined,
  draft: RunDraft,
  fields: FieldFrame[],
): void {
  throwIfAborted(ctx.signal);
  let pending = "";
  const flush = (): void => {
    if (pending === "") return;
    const finalized = finalizeRun(ctx, props, paragraphText, pending, href === undefined ? undefined : { hyperlink: href });
    pending = "";
    if (finalized !== undefined) draft.runs.push(finalized);
  };
  // The cached result of a live PAGE/NUMPAGES field is not flow content: the
  // field evaluates per page, so painting the cache too prints "Page 21 of 42".
  const liveResult = (): FieldFrame | undefined => {
    const frame = fields[fields.length - 1];
    return frame !== undefined && frame.live !== undefined && frame.inResult ? frame : undefined;
  };
  for (const child of run.children) {
    throwIfAborted(ctx.signal);
    switch (child.local) {
      case "t": {
        const frame = liveResult();
        if (frame === undefined) pending += child.text;
        else frame.cached += child.text;
        break;
      }
      case "tab":
        if (liveResult() === undefined) pending += "\t";
        break;
      case "br":
      case "cr": {
        // Only an explicit page break is a pagination signal; every other
        // break (including `w:cr` and text-wrapping breaks) is a line break
        // in flow HTML.
        if (child.local === "br" && attr(child, "type") === "page") {
          flush();
          draft.pageBreakAt ??= draft.runs.length;
        } else {
          pending += "\n";
        }
        break;
      }
      case "noBreakHyphen":
        pending += NON_BREAKING_HYPHEN;
        break;
      case "sym": {
        const font = attr(child, "font");
        const char = attr(child, "char");
        const code = char === undefined ? Number.NaN : Number.parseInt(char, 16);
        if (font !== undefined && Number.isInteger(code)) {
          flush();
          const symbol = finalizeRun(
            ctx,
            { ...props, ascii: font, hAnsi: font },
            paragraphText,
            String.fromCodePoint(code),
            href === undefined ? undefined : { hyperlink: href },
          );
          if (symbol !== undefined) draft.runs.push(symbol);
        }
        break;
      }
      case "AlternateContent":
      case "drawing": {
        // An Office 2016 chart is the Choice's `w:drawing`; every other alternate keeps being skipped.
        const chosen = child.local === "drawing" ? { drawing: child } : alternateDrawing(child);
        const drawing = child.local === "drawing" ? child : chartExDrawing(child) ?? chosen?.drawing;
        const pict = child.local === "AlternateContent" && drawing === undefined ? chosen?.pict : undefined;
        if (drawing === undefined && pict === undefined) break;
        flush();
        if (pict !== undefined) {
          const textBox = parseVmlTextBox(pict, ctx);
          if (textBox?.anchor !== undefined) draft.floats.push(textBox);
          else if (textBox !== undefined) draft.blocks.push(textBox);
          if (descendant(pict, "imagedata") !== undefined) draft.blocks.push(parseVmlPicture(pict, ctx) ?? { type: "unsupported", source: "drawing", label: "Picture" });
          break;
        }
        const image = parseDrawing(drawing!, ctx);
        if (image === undefined) break;
        // An anchored picture or chart is placed relative to its paragraph and
        // never splits the text it sits in.
        if (image.type !== "unsupported" && image.anchor !== undefined) {
          draft.floats.push(image);
          break;
        }
        // A picture in a paragraph that also carries text stays on that line
        // (a header logo beside its wordmark); a picture-only paragraph keeps
        // the block image with its host alignment.
        if (image.type === "image" && drawing !== undefined && paragraphText.trim() !== "" && firstChild(drawing, "inline") !== undefined) {
          draft.runs.push({
            text: "",
            image: {
              mediaId: image.mediaId,
              contentType: image.contentType,
              width: image.width,
              height: image.height,
              ...(image.alt === undefined ? {} : { alt: image.alt }),
              ...(image.effects === undefined ? {} : { effects: image.effects,
                ...(image.effectInsets === undefined ? {} : { effectInsets: image.effectInsets }) }),
            },
            ...(href === undefined ? {} : { hyperlink: href }),
          });
        } else {
          draft.blocks.push(image);
        }
        break;
      }
      case "pict": {
        flush();
        const textBox = parseVmlTextBox(child, ctx);
        if (textBox?.anchor !== undefined) draft.floats.push(textBox);
        else if (textBox !== undefined) draft.blocks.push(textBox);
        if (descendant(child, "imagedata") !== undefined) {
          const picture = parseVmlPicture(child, ctx);
          if (picture !== undefined) draft.blocks.push(picture);
        }
        break;
      }
      case "object":
      case "oleObj":
        flush();
        draft.blocks.push(parseOleObject(child));
        break;
      case "footnoteReference":
      case "endnoteReference": {
        const id = attr(child, "id");
        if (id === undefined) break;
        const key = `${child.local === "footnoteReference" ? "f" : "e"}${id}`;
        if (!ctx.notes.has(key)) break;
        flush();
        const marker = finalizeRun(ctx, props, paragraphText, footnoteLabel(ctx, key), {
          verticalAlign: "superscript",
          footnoteRef: key,
          ...(href === undefined ? {} : { hyperlink: href }),
        });
        if (marker !== undefined) draft.runs.push(marker);
        break;
      }
      case "footnoteRef":
      case "endnoteRef":
      case "separator":
      case "continuationSeparator":
      case "delText":
      case "rPr":
        break;
      case "fldChar": {
        flush();
        const type = attr(child, "fldCharType");
        if (type === "begin") {
          fields.push({ inResult: false, cached: "" });
        } else if (type === "separate") {
          const frame = fields[fields.length - 1];
          if (frame !== undefined) frame.inResult = true;
          // The top-level `separate` ends the field instruction: later
          // instruction chunks belong to nested fields, not the TOC. The
          // opener's own `separate` arrives before its session exists, so the
          // flag is recorded per paragraph and handed to the session.
          if (fields.length === 1 && ctx.toc !== undefined) {
            draft.tocSeparated = true;
            if (ctx.toc.active !== undefined) ctx.toc.active.separated = true;
          }
        } else if (type === "end") {
          // An end with no local frame while a TOC is open is the TOC's own
          // end: its cached paragraphs span several runs, so no local frame
          // is open for it. Nested `PAGEREF` ends always pop a local frame.
          if (fields.length === 0 && ctx.toc?.active !== undefined) draft.tocClosed = true;
          const frame = fields.pop();
          // Keep the cache on the live run: it is the flow-mode fallback text.
          if (frame?.live !== undefined && frame.cached !== "") frame.live.text = frame.cached;
        }
        break;
      }
      case "lastRenderedPageBreak":
        // Stale Word layout hint: "a page began here last time Word laid this
        // out". Only meaningful before any text; `emitDraft` decides.
        flush();
        if (draft.runs.length === 0) draft.softBreak = true;
        break;
      case "instrText": {
        // Layout pagination substitutes the real PAGE/NUMPAGES per page; static
        // flow falls back to the field's cached result ("1" when it has none).
        const field = pageField(child.text);
        if (field !== undefined) {
          flush();
          const marker = finalizeRun(ctx, props, paragraphText, "1", { field });
          if (marker !== undefined) {
            draft.runs.push(marker);
            const frame = fields[fields.length - 1];
            if (frame !== undefined && !frame.inResult && frame.live === undefined) frame.live = marker;
          }
        } else if (ctx.toc !== undefined && fields.length === 1) {
          const active = ctx.toc.active;
          if (active !== undefined && !active.separated) {
            // A long TOC instruction can split across runs: accumulate
            // top-level chunks until its `separate` and re-read the switches.
            // A late unsupported switch flips the session back to fallback
            // (the cache stays held until the splice decides).
            active.raw += child.text ?? "";
            active.instruction = parseTocInstruction(active.raw);
          } else if (active === undefined) {
            // A candidate opener: `parseParagraph` joins the paragraph's
            // chunks (an instruction can split mid-word) and opens only when
            // the whole names a TOC. Nested `PAGEREF` instructions arrive with
            // a TOC already open and stay ordinary cached text.
            draft.tocInstrChunk = child.text ?? "";
          }
        }
        break;
      }
      default:
        break;
    }
  }
  flush();
}

/** The live page field a field instruction names (`PAGE`, `NUMPAGES`), if it is one. */
function pageField(instruction: string | undefined): "PAGE" | "NUMPAGES" | undefined {
  const name = fieldTokens(instruction ?? "")[0]?.toUpperCase();
  return name === "PAGE" || name === "NUMPAGES" ? name : undefined;
}

/** Whether a field instruction opens a table of contents (`TOC ...`), before its switches are understood. */
function isTocInstruction(instruction: string | undefined): boolean {
  if (instruction === undefined) return false;
  const first = fieldTokens(instruction)[0];
  return first !== undefined && first.toUpperCase() === "TOC";
}

/**
 * Split a field instruction into whitespace-separated tokens, keeping `"quoted
 * ranges"` whole. The one tokenizer every field instruction is read through
 * (`pageField`, `isTocInstruction`, `parseTocInstruction`), so a switch quoted
 * as a range is never split at its hyphen wherever it is read.
 */
function fieldTokens(instruction: string): string[] {
  const tokens: string[] = [];
  const pattern = /"([^"]*)"|(\S+)/g;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(instruction)) !== null) tokens.push(match[1] ?? match[2]!);
  return tokens;
}

/** An outline-level range (`"1-2"` or `"3"`), clamped to Word's levels 1-9. */
function levelRange(token: string | undefined): { min: number; max: number } | undefined {
  if (token === undefined) return undefined;
  const single = /^\s*(\d+)\s*$/.exec(token);
  if (single !== null) {
    const level = Math.min(9, Math.max(1, Number(single[1])));
    return { min: level, max: level };
  }
  const span = /^\s*(\d+)\s*-\s*(\d+)\s*$/.exec(token);
  if (span === null) return undefined;
  const from = Math.min(9, Math.max(1, Number(span[1])));
  const to = Math.min(9, Math.max(1, Number(span[2])));
  return { min: Math.min(from, to), max: Math.max(from, to) };
}

/**
 * **The refresh rule, in one place.** A field is refreshed only when every
 * switch that could change its result is understood. Anything else keeps the
 * whole saved result, because a refresh that silently drops what an
 * unread switch selects is worse for a reader than a table that is merely
 * stale: the stale one is at least the reader's own structure, and it says so.
 * Two ways to honour this, both used here: `parseTocInstruction` returns
 * `undefined` for a field it cannot fully evaluate (so the splice unwraps the
 * cache), and `resolveTocRanges` unwraps a session whose instruction went
 * undefined mid-stream, which is how a late switch in a split instruction
 * falls back.
 *
 * Adding a switch therefore means saying what it does. If its effect cannot be
 * honoured, it belongs in the `else` below with the others — not parsed and
 * then ignored.
 *
 * A TOC field's instruction as a refresh plan, or undefined when the field
 * must keep its saved result. Heading outline levels (`\o`), directly applied
 * outline formatting (`\u`, folded into the same outline levels the parser
 * already resolves), the bookmark range a field collects from (`\b`, see
 * `docxBookmarkRanges`) and presentation switches (`\h`, `\n`, `\p`, `\w`,
 * `\x`; `\z` is accepted and ignored — flow mode has no web-layout to hide
 * numbers in) refresh; caption, sequence, TC-entry and chapter-number switches
 * (`\a`, `\c`, `\d`, `\f`, `\l`, `\s`) and anything unrecognized keep the
 * cache. Switch names match case-insensitively, as Word reads them (probe
 * `numbered-heading-probe`: `{ TOC \h \o "1-2" }`).
 */
function parseTocInstruction(instruction: string): TocInstruction | undefined {
  const tokens = fieldTokens(instruction);
  if (tokens.length === 0 || tokens[0]!.toUpperCase() !== "TOC") return undefined;
  const out: TocInstruction = { hyperlink: false, minLevel: 1, maxLevel: 9, omitPageMin: 10, omitPageMax: 0 };
  let index = 1;
  while (index < tokens.length) {
    const flag = tokens[index]!.toLowerCase();
    if (flag === "\\h" || flag === "\\u" || flag === "\\w" || flag === "\\x" || flag === "\\z") {
      if (flag === "\\h") out.hyperlink = true;
      index++;
    } else if (flag === "\\o") {
      const range = levelRange(tokens[index + 1]);
      if (range !== undefined) {
        out.minLevel = range.min;
        out.maxLevel = range.max;
        index += 2;
      } else {
        index++;
      }
    } else if (flag === "\\n") {
      const range = levelRange(tokens[index + 1]);
      if (range !== undefined) {
        out.omitPageMin = range.min;
        out.omitPageMax = range.max;
        index += 2;
      } else {
        out.omitPageMin = 1;
        out.omitPageMax = 9;
        index++;
      }
    } else if (flag === "\\b") {
      // A `\b` with no name cannot be honoured, and guessing would list the
      // whole document: the field falls back like any unread switch.
      const name = tokens[index + 1];
      if (name === undefined) return undefined;
      out.bookmark = name;
      index += 2;
    } else if (flag === "\\p") {
      const separator = tokens[index + 1];
      if (separator !== undefined && [...separator].length === 1) {
        out.separator = separator;
        index += 2;
      } else {
        index++;
      }
    } else {
      return undefined;
    }
  }
  return out;
}

function paragraphTextOf(paragraph: XmlElement): string {
  const parts: string[] = [];
  const stack = [paragraph];
  while (stack.length > 0) {
    const current = stack.pop()!;
    if (current.local === "del" || current.local === "moveFrom" || current.local === "txbxContent") continue;
    if (current.local === "t") parts.push(current.text);
    for (let index = current.children.length - 1; index >= 0; index--) stack.push(current.children[index]!);
  }
  return parts.join("");
}

/** Paragraph content children: tracked-changes-as-final (`w:ins`/`w:moveTo` in, `w:del`/`w:moveFrom` out). */
function contentChildren(parent: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  for (const child of parent.children) {
    if (child.local === "del" || child.local === "moveFrom") continue;
    if (child.local === "ins" || child.local === "moveTo") {
      out.push(...contentChildren(child));
      continue;
    }
    out.push(child);
  }
  return out;
}

interface ParagraphDraft {
  nodes: FlowNode[];
  numId?: string;
  ilvl?: number;
  /** Source paragraph style id, for heading numbering run resolution. */
  styleId?: string;
  /**
   * Table-of-contents boundary this paragraph carries: `open` saw the TOC
   * instruction, `close` saw the TOC's own end, `inside` sits between the
   * two, and `single` is one `w:fldSimple` TOC field alone in its paragraph.
   */
  tocBoundary?: "open" | "close" | "inside" | "single";
  /** The TOC session this paragraph opens, closes, or alone constitutes. */
  tocSession?: TocSession | undefined;
}

/** Move one run's content into the open paragraph, splitting on pagination signals found in the run. */
function emitDraft(
  draft: RunDraft,
  runs: TextRunModel[],
  pieces: FlowNode[],
  floats: DocxDrawingNode[],
  flushRuns: () => void,
  align: TextAlign | undefined,
  direction: TextDirection,
): void {
  // A rendered-page-break hint is a preferred break *before* the paragraph,
  // and only when no text of the paragraph precedes it: mid-paragraph it is
  // Word's line-level layout cache and must never split the text.
  if (draft.softBreak && draft.pageBreakAt === undefined && runs.length === 0 && pieces.length === 0) {
    pieces.push(pageBreakNode("preferred"));
  }
  const breakAt = draft.pageBreakAt ?? draft.runs.length;
  runs.push(...draft.runs.slice(0, breakAt));
  floats.push(...draft.floats);
  for (const block of draft.blocks) {
    // Inline images and charts inherit the host paragraph's alignment (a
    // centered source paragraph centers its image); anchored ones never reach here.
    if ((block.type === "image" || block.type === "chart") && align !== undefined) block.align = align;
    // There is no text in a picture to justify, so Word leaves it on the
    // paragraph's start side: a justified right-to-left paragraph's picture
    // sits at its right edge (measured on the Word-authored coverage page).
    if (block.type === "image" || block.type === "chart") {
      if (block.align === undefined || block.align === "justify") {
        block.align = direction === "rtl" ? "right" : "left";
      }
    }
    flushRuns();
    pieces.push(block);
  }
  if (draft.pageBreakAt !== undefined) {
    flushRuns();
    pieces.push(pageBreakNode("mandatory"));
    runs.push(...draft.runs.slice(breakAt));
  }
}

/**
 * The line `w:beforeLines`/`w:afterLines` count in: always 12pt, whatever the
 * font, size, script or line spacing. With a line grid on
 * (`w:docGrid w:type="lines"`) Word counts the grid pitch instead, but aared
 * has no line grid, so 12pt stands there too. Ground truth: Word's own PDF
 * export (probe `line-spacing-units`).
 */
const SPACING_LINE_PT = 12;

/**
 * Space before or after a paragraph. The line value and the point value
 * inherit separately through docDefaults, styles and direct formatting; a line
 * value above zero wins, even over the paragraph's own points, and zero means
 * "in points". Ground truth: Word reading hand-edited XML (probe
 * `line-spacing-units`) — a style's 1 line beat a paragraph's own 6pt,
 * `beforeLines="0" before="120"` painted 6pt, and `beforeLines="50"` alone 6pt.
 * Word itself always saves both, so its own files agree either way.
 */
function paragraphSpacePt(lines: number | undefined, points: number | undefined): number | undefined {
  return lines !== undefined && lines > 0 ? lines * SPACING_LINE_PT : points;
}

function layoutFor(props: ResolvedParaProps, direction: TextDirection, justifyShrinks: boolean): DocxParagraphLayout {
  const layout: DocxParagraphLayout = {};
  if (props.fontSizePt !== undefined) layout.fontSizePt = props.fontSizePt;
  const align = alignFor(props.jc, direction);
  if (align !== undefined) layout.align = align;
  if (align === "justify" && justifyShrinks) layout.justifyShrink = true;
  if (props.lineSpacing !== undefined) layout.lineSpacing = props.lineSpacing;
  if (props.lineSpacingPt !== undefined) layout.lineSpacingPt = props.lineSpacingPt;
  if (props.lineSpacingAtLeast === true) layout.lineSpacingAtLeast = true;
  const spaceBeforePt = paragraphSpacePt(props.spaceBeforeLines, props.spaceBeforePt);
  const spaceAfterPt = paragraphSpacePt(props.spaceAfterLines, props.spaceAfterPt);
  if (spaceBeforePt !== undefined) layout.spaceBeforePt = spaceBeforePt;
  if (spaceAfterPt !== undefined) layout.spaceAfterPt = spaceAfterPt;
  if (props.spaceBeforeLines !== undefined) layout.spaceBeforeLines = props.spaceBeforeLines;
  if (props.spaceAfterLines !== undefined) layout.spaceAfterLines = props.spaceAfterLines;
  if (props.snapToGrid !== undefined) layout.snapToGrid = props.snapToGrid;
  if (props.leftIndent !== undefined) layout.leftIndent = props.leftIndent;
  if (props.rightIndent !== undefined) layout.rightIndent = props.rightIndent;
  if (props.firstLineIndent !== undefined) layout.firstLineIndent = props.firstLineIndent;
  if (props.shading !== undefined) layout.shading = props.shading;
  if (props.borders !== undefined && Object.keys(props.borders).length > 0) layout.borders = props.borders;
  if (props.tabs !== undefined && props.tabs.length > 0) layout.tabStops = props.tabs;
  if (props.keepNext === true) layout.keepNext = true;
  return layout;
}

/**
 * A paragraph's effective properties. Precedence: docDefaults < style chain
 * (base first) < direct formatting, each level read over the one before, so
 * direct formatting follows the same rules as a style: tab stops accumulate
 * (a `clear` removes an inherited one), border sides compose, and a
 * `w:spacing` line rule replaces the inherited one.
 */
function resolveParaProps(
  styles: StyleData,
  directPPr: XmlElement | undefined,
  tablePPr: XmlElement[] = [],
): { merged: ResolvedParaProps; styleId: string | undefined } {
  const directStyle = directPPr === undefined ? undefined : firstChild(directPPr, "pStyle");
  // A paragraph naming no style takes the document's default paragraph style
  // (`Normal`), as Word does.
  const styleId = (directStyle === undefined ? undefined : attr(directStyle, "val")) ?? styles.defaultParagraphStyleId;
  const merged: ResolvedParaProps = {};
  if (styles.defaultPPr !== undefined) readParaProps(styles.defaultPPr, merged);
  // Below the paragraph style, above nothing else (see `baseRunProps`).
  for (const pPr of tablePPr) readParaProps(pPr, merged);
  for (const style of paragraphAncestry(styles, styleId)) readParaProps(style.pPr, merged);
  readParaProps(directPPr, merged);
  return { merged, styleId };
}

/**
 * The outline level (1-9) a paragraph carries: its style's heading level, else
 * its resolved `w:outlineLvl` (0-8; level 9 is Word's "body text"). The one
 * rule that makes a paragraph a heading node; the public outline
 * (`document.headings`) is read back from those nodes, never re-derived.
 */
function outlineLevel(styleId: string | undefined, merged: ResolvedParaProps, styles: StyleData): number | undefined {
  return headingLevelForStyle(styleId, styles)
    ?? (merged.outlineLvl !== undefined && merged.outlineLvl <= 8 ? merged.outlineLvl + 1 : undefined);
}

function parseParagraph(paragraph: XmlElement, ctx: ParseContext): ParagraphDraft {
  throwIfAborted(ctx.signal);
  const pPr = firstChild(paragraph, "pPr");
  const { merged, styleId } = resolveParaProps(ctx.styles, pPr, ctx.tableStyleText?.pPr ?? []);

  const direction: TextDirection = merged.bidi === true ? "rtl" : "ltr";
  const text = paragraphTextOf(paragraph);
  const paraAlign = alignFor(merged.jc, direction);
  // The paragraph's default run size (its style chain's), for lines whose runs state none.
  const defaultRun = baseRunProps(ctx, styleId, undefined);
  const defaultSize = defaultRun.fontSizePt ?? defaultRun.fontSizeCsPt;
  if (defaultSize !== undefined) merged.fontSizePt = defaultSize;

  const headingLevel = outlineLevel(styleId, merged, ctx.styles);
  const pieces: FlowNode[] = [];
  const floats: DocxDrawingNode[] = [];
  const fields: FieldFrame[] = [];
  /** Top-level field instruction chunks seen in this paragraph, joined to detect a TOC opener. */
  let tocInstrChunks: string[] | undefined;
  /** Whether this paragraph saw a top-level `separate`. */
  let tocParaSeparated = false;
  /** Whether this paragraph saw the open TOC's own end. */
  let tocClosed = false;
  /** Supported lone simple TOC field in this paragraph, unless a complex TOC is already open. */
  let tocSingle: TocInstruction | undefined;
  /** The open paragraph's runs; `flushRuns` moves them out into a node, so a holder of this array keeps filling the next one. */
  const runs: TextRunModel[] = [];
  const flushRuns = (): void => {
    if (runs.length === 0) return;
    const taken = runs.splice(0);
    const layout = layoutFor(merged, direction, ctx.justifyShrinks);
    if (headingLevel === undefined) {
      pieces.push({ type: "paragraph", runs: taken, direction, ...layout, ...(styleId === undefined ? {} : { styleId }) });
    } else {
      pieces.push({ type: "heading", level: headingLevel, text: runsText(taken), direction, runs: taken, ...layout, ...(styleId === undefined ? {} : { styleId }) });
    }
  };

  /**
   * Read one `w:r` into the open paragraph. `styleOf` maps the run's own
   * `rStyle` to the character style it takes: a run keeps it unless the
   * caller says otherwise (a hyperlink's runs fall back to the Hyperlink style).
   */
  const readRun = (run: XmlElement, href: string | undefined, styleOf?: (own: string | undefined) => string | undefined): void => {
    const direct = firstChild(run, "rPr");
    const rStyle = direct === undefined ? undefined : firstChild(direct, "rStyle");
    const own = rStyle === undefined ? undefined : attr(rStyle, "val");
    const props = baseRunProps(ctx, styleId, styleOf === undefined ? own : styleOf(own));
    readRunProps(direct, props, ctx.theme);
    const draft: RunDraft = { runs: [], blocks: [], floats: [], softBreak: false };
    parseRunContent(run, ctx, props, text, href, draft, fields);
    if (draft.tocInstrChunk !== undefined) (tocInstrChunks ??= []).push(draft.tocInstrChunk);
    if (draft.tocSeparated) tocParaSeparated = true;
    if (draft.tocClosed) tocClosed = true;
    // Word orders a run without `w:rtl` left to right inside a right-to-left
    // paragraph (probe run-direction).
    for (const piece of draft.runs) {
      if (isLtrRunInRtl(piece.text, direction === "rtl", props.rtl === true)) piece.ltrRun = true;
    }
    emitDraft(draft, runs, pieces, floats, flushRuns, paraAlign, direction);
  };
  for (const child of contentChildren(paragraph)) {
    throwIfAborted(ctx.signal);
    if (child.local === "pPr") continue;
    // A bookmark may sit between runs inside a paragraph; its block index is
    // this paragraph's, applied by the caller once its drafts are pushed.
    if (child.local === "bookmarkStart") {
      bookmarkBoundary(ctx, 0, attr(child, "name"), attr(child, "id"), true);
      continue;
    }
    if (child.local === "bookmarkEnd") {
      bookmarkBoundary(ctx, 0, undefined, attr(child, "id"), true);
      continue;
    }
    if (child.local === "r") {
      readRun(child, undefined);
      continue;
    }
    if (child.local === "hyperlink") {
      const tocParagraph = styleId !== undefined && /^toc\s?\d$/i.test(styleId);
      const rid = attr(child, "id");
      const anchor = attr(child, "anchor");
      const href = rid !== undefined && ctx.hyperlinks.get(rid) !== undefined
        ? ctx.hyperlinks.get(rid)!
        : anchor !== undefined
          ? `#${anchor}`
          : undefined;
      // A TOC's own `\h` links carry the Hyperlink style, but Word paints
      // them as plain TOC text: the entry keeps its paragraph formatting.
      const linkStyle = (own: string | undefined): string | undefined => (tocParagraph ? undefined : own ?? "Hyperlink");
      for (const run of contentChildren(child)) {
        if (run.local === "r") readRun(run, href, linkStyle);
      }
      continue;
    }
    if (child.local === "sdt") {
      // A content control's runs are ordinary runs: they keep their own character style.
      const content = firstChild(child, "sdtContent");
      if (content === undefined) continue;
      for (const nested of contentChildren(content)) {
        if (nested.local === "r") readRun(nested, undefined);
      }
      continue;
    }
    if (child.local === "fldSimple") {
      const field = pageField(attr(child, "instr"));
      if (field !== undefined) {
        const cached = paragraphTextOf(child);
        const marker = finalizeRun(ctx, baseRunProps(ctx, styleId, undefined), text, cached === "" ? "1" : cached, { field });
        if (marker !== undefined) runs.push(marker);
      } else {
        // Any other simple field (DATE, TITLE, REF…) shows the result Word cached in it.
        // A lone simple TOC field with supported switches refreshes like a
        // one-paragraph complex TOC; anything else keeps its cache.
        const toc = ctx.toc?.active === undefined ? parseTocInstruction(attr(child, "instr") ?? "") : undefined;
        for (const run of contentChildren(child)) {
          if (run.local === "r") readRun(run, undefined);
        }
        if (toc !== undefined) tocSingle = toc;
      }
      continue;
    }
  }
  flushRuns();

  const emptyParagraph = pieces.length === 0 && floats.length === 0;
  const sectionBreak = pPr === undefined ? undefined : firstChild(pPr, "sectPr");
  // In Word 16, this empty paragraph with both properties adds no standalone
  // body line or page; the following section transition owns the page change.
  // The coverage fixture's Word export and paired captures verify the result.
  const emptySectionEndingParagraph = sectionBreak !== undefined && merged.pageBreakBefore === true && emptyParagraph;
  const filtered = pieces;
  if (emptyParagraph && !emptySectionEndingParagraph) {
    // An empty paragraph is a full line tall, at the size of its paragraph mark
    // (`w:pPr/w:rPr`), plus its own spacing (probe `docx-empty-paragraph`, Word
    // 16: one empty paragraph moves the next text down by exactly one line, a
    // 24pt mark by one 24pt line, 12pt before and after by those 24pt more). A
    // table cell's lone empty paragraph likewise gives the row its line.
    const markProps = baseRunProps(ctx, styleId, undefined);
    readRunProps(pPr === undefined ? undefined : firstChild(pPr, "rPr"), markProps, ctx.theme);
    const markSize = markProps.rtl === true ? markProps.fontSizeCsPt ?? markProps.fontSizePt : markProps.fontSizePt;
    if (markSize !== undefined) merged.fontSizePt = markSize;
    // Never a heading: an empty one would be an empty entry in the outline.
    filtered.push({ type: "paragraph", runs: [], direction, ...layoutFor(merged, direction, ctx.justifyShrinks), ...(styleId === undefined ? {} : { styleId }) });
  }
  if (floats.length > 0) {
    if (merged.numId !== undefined) {
      // A list item is not a paragraph node and cannot carry them: its pictures
      // and charts fall back to plain blocks ahead of the item.
      filtered.unshift(...floats.map((float): DocxDrawingNode => {
        const block = { ...float };
        delete block.anchor;
        return block;
      }));
    } else {
      const host = filtered.find((node) => node.type === "paragraph" || node.type === "heading");
      if (host?.type === "paragraph" || host?.type === "heading") host.floats = floats;
      // A paragraph holding only pictures is still one line tall in Word.
      else filtered.push({ type: "paragraph", runs: [], direction, ...layoutFor(merged, direction, ctx.justifyShrinks), floats });
    }
  }
  if (merged.pageBreakBefore === true && !emptySectionEndingParagraph) {
    // Inherited `w:pageBreakBefore` (direct or via the style chain) forces
    // the paragraph onto a new page in layout mode. Word ignores that hint on
    // an empty section-ending paragraph: the section transition itself starts
    // the next section on its authored page.
    filtered.unshift(pageBreakNode("mandatory"));
  }
  // Table-of-contents boundaries. A complex TOC opens a session its cached
  // paragraphs join; the paragraph holding its end closes it. A lone simple
  // TOC field is a session of its own paragraph. A paragraph can only open
  // when tracking is on and nothing is open (see `parseRunContent`); a close
  // without an open session is impossible for the same reason.
  let tocBoundary: ParagraphDraft["tocBoundary"];
  let tocSession: TocSession | undefined;
  if (ctx.toc !== undefined) {
    // A paragraph can only open when nothing is open; its chunks joined catch
    // an instruction split across runs (even mid-word). A lone simple TOC
    // field is a session of its own paragraph.
    const opener = tocInstrChunks !== undefined && ctx.toc.active === undefined ? tocInstrChunks.join("") : undefined;
    const opened = opener !== undefined && isTocInstruction(opener) ? opener : undefined;
    if (tocSingle !== undefined && opened === undefined && ctx.toc.active === undefined) {
      tocBoundary = "single";
      tocSession = { instruction: tocSingle, raw: "", separated: true, cached: [] };
    } else if (opened !== undefined) {
      tocSession = { instruction: parseTocInstruction(opened), raw: opened, separated: tocParaSeparated, cached: [] };
      if (tocClosed) {
        tocBoundary = "single";
      } else {
        tocBoundary = "open";
        ctx.toc.active = tocSession;
      }
    } else if (tocClosed) {
      tocBoundary = "close";
      tocSession = ctx.toc.active;
      ctx.toc.active = undefined;
    } else if (ctx.toc.active !== undefined) {
      tocBoundary = "inside";
    }
  }
  return {
    nodes: filtered,
    ...(merged.numId === undefined ? {} : { numId: merged.numId }),
    ilvl: merged.ilvl ?? 0,
    ...(styleId === undefined ? {} : { styleId }),
    ...(tocBoundary === undefined ? {} : { tocBoundary }),
    ...(tocSession === undefined ? {} : { tocSession }),
  };
}

/**
 * Count one numbered paragraph. Word counts per list definition (`w:abstractNum`), not per
 * `w:numId`: numbering ids that share a definition continue one sequence, and a `w:num` with a
 * `w:startOverride` starts its level over the first time it is used (probe `docx-list-counters`,
 * Word 16, which also showed an empty numbered paragraph and a numbered heading take a number).
 */
function advanceCounters(ctx: ParseContext, numId: string, ilvl: number): { level: NumberingLevel; knownLevels: Map<number, NumberingLevel>; counters: number[] } {
  const abstractId = ctx.numbering.nums.get(numId);
  const levels = abstractId === undefined ? undefined : ctx.numbering.abstracts.get(abstractId);
  const fallback: NumberingLevel = { numFmt: "bullet", lvlText: BULLET, start: 1 };
  const level = levels?.get(ilvl) ?? fallback;
  const knownLevels = levels ?? new Map<number, NumberingLevel>([[ilvl, level]]);
  const key = abstractId ?? `num:${numId}`;
  let counters = ctx.counters.get(key);
  if (counters === undefined) {
    counters = [];
    ctx.counters.set(key, counters);
  }
  if (!ctx.startedNums.has(numId)) {
    ctx.startedNums.add(numId);
    for (const overridden of ctx.numbering.startOverrides.keys()) {
      if (!overridden.startsWith(`${numId}:`)) continue;
      const from = Number(overridden.slice(numId.length + 1));
      if (Number.isInteger(from)) counters.length = Math.min(counters.length, from);
    }
  }
  counters[ilvl] = (counters[ilvl] ?? (startFor(ctx.numbering, numId, ilvl, level) - 1)) + 1;
  // Deeper levels start over, at their own start value.
  counters.length = ilvl + 1;
  return { level, knownLevels, counters };
}

/**
 * One list paragraph as a numbered item. Counters advance as the paragraph is
 * parsed, so ordinals follow document order across the body, table cells and
 * every other place a numbering id is used.
 */
function listItemFor(ctx: ParseContext, numId: string, ilvl: number, paragraph: DocxParagraphNode): DocxListItemNode {
  const { level, knownLevels, counters } = advanceCounters(ctx, numId, ilvl);
  const authored = markerFor(level, knownLevels, ilvl, counters);
  const { marker, markerRun } = symbolMarker(authored, markerRunFor(ctx, level, paragraph.styleId, authored));
  const item: DocxListItemNode = {
    level: ilvl,
    ordered: isOrderedFormat(level.numFmt),
    marker,
    ...(markerRun === undefined ? {} : { markerRun }),
    ...(level.suffix === undefined || level.suffix === "tab" ? {} : { markerSuffix: level.suffix }),
    paragraph,
  };
  // Numbering indents apply only when the paragraph sets none of its own:
  // direct `w:ind` always wins over the level default.
  if (paragraph.leftIndent === undefined && level.indentLeft !== undefined) {
    paragraph.leftIndent = level.indentLeft;
  }
  if (paragraph.firstLineIndent === undefined && level.indentHanging !== undefined) {
    paragraph.firstLineIndent = -level.indentHanging;
  }
  return item;
}

/** Runs of consecutive items of one numbering id become one `list` block. */
function groupListItems(drafts: BlockDraft[]): FlowNode[] {
  const out: FlowNode[] = [];
  let open: { numId: string; items: DocxListItemNode[] } | undefined;
  const feed = (draft: { node: FlowNode } | { listItem: { numId: string; item: DocxListItemNode } }): void => {
    if (!("listItem" in draft)) {
      open = undefined;
      out.push(draft.node);
      return;
    }
    const { numId, item } = draft.listItem;
    if (open === undefined || open.numId !== numId) {
      open = { numId, items: [] };
      out.push({ type: "list", items: open.items });
    }
    open.items.push(item);
  };
  for (const draft of drafts) {
    // TOC markers are spliced out before grouping (see `resolveTocRanges`);
    // a surviving one only bounds a list, and a surviving cached entry feeds
    // like the paragraph it wraps.
    if ("toc" in draft) {
      open = undefined;
      if (draft.toc.kind === "cached") feed(draft.toc.inner);
      continue;
    }
    feed(draft);
  }
  return out;
}

/** A heading available to refreshed TOC entries, in document order. */
interface TocHeading {
  node: DocxHeadingNode;
  /** Outline index for top-level body headings; headings in table cells have none. */
  outlineIndex: number | undefined;
  /** The block this heading occupies in its section's drafts, for a `\b` range. */
  block: number;
}

/**
 * Body headings in document order: top-level drafts take outline indexes (the
 * same index space `headingOutline` reads back later), table-cell headings
 * join without one. Footnote, text-box and cached-TOC drafts never
 * contribute; headers and footers parse without TOC tracking at all.
 *
 * `block` is the draft index a heading occupies, the same index space
 * `DocxBookmarkIndex` records bookmark ranges in — a table counts as one block,
 * so its cell headings share the table's position.
 */
function collectTocHeadings(drafts: BlockDraft[], out: TocHeading[], outline: { count: number }): void {
  for (const [block, draft] of drafts.entries()) {
    if ("toc" in draft || "listItem" in draft) continue;
    const node = draft.node;
    if (node.type === "heading") {
      out.push({ node, outlineIndex: outline.count++, block });
    } else if (node.type === "table") {
      for (const row of node.rows) for (const cell of row.cells) collectTocCellHeadings(cell.nodes, out, block);
    }
  }
}

function collectTocCellHeadings(nodes: FlowNode[], out: TocHeading[], block: number): void {
  for (const node of nodes) {
    if (node.type === "heading") out.push({ node, outlineIndex: undefined, block });
    else if (node.type === "table") {
      for (const row of node.rows) for (const cell of row.cells) collectTocCellHeadings(cell.nodes, out, block);
    }
  }
}

/** The fragment a refreshed TOC entry links at, assigning table-cell targets their own. */
function tocAnchorFor(ctx: ParseContext, node: DocxHeadingNode): string {
  if (node.tocAnchor === undefined) {
    ctx.toc!.serial++;
    node.tocAnchor = `aared-toc-${ctx.toc!.serial}`;
  }
  return node.tocAnchor;
}

/** A refreshed entry's cached page number: its display text matched against the saved entries. */
function cachedTocPage(session: TocSession, displayText: string): string {
  const wanted = displayText.trim();
  for (const entry of session.cached) {
    if (entry.entry === wanted) return entry.page;
  }
  // Entries added since the TOC was last saved have no cache: flow mode
  // shows no number, layout mode resolves the live one.
  return "";
}

/**
 * Paragraph properties for a block aared synthesizes rather than parses, which
 * takes only a style id: docDefaults, then that style's chain. The direct
 * `w:pPr` half of `resolveParaProps` has no element to read, so this is the
 * same inheritance with one fewer level.
 */
function styleOnlyParaProps(ctx: ParseContext, styleId: string): ResolvedParaProps {
  const merged: ResolvedParaProps = {};
  if (ctx.styles.defaultPPr !== undefined) readParaProps(ctx.styles.defaultPPr, merged);
  for (const style of paragraphAncestry(ctx.styles, styleId)) readParaProps(style.pPr, merged);
  return merged;
}

/** What a refreshed entry's own style says about text: runs, size, and the paragraph's direction. */
interface TocEntryStyle {
  id: string;
  direction: TextDirection;
  fontSizePt: number | undefined;
  /** The style chain's run defaults, so entry text resolves its script slot like any run. */
  runProps: ResolvedRunProps;
  layout: DocxParagraphLayout;
}

/**
 * A TOC entry's `TOCn` style, resolved like any other paragraph's. Word reads an
 * entry's spacing, indent and direction from that style, never from the heading
 * it lists: an entry for a right-to-left heading in a left-to-right TOC still
 * paints left to right. The entry is also *set* in that style — Word's cached
 * entries carry no size or typeface of their own but a heading's emphasis, and
 * a `Heading 1`'s bold is style-derived, so an entry in bold would be a
 * difference on every document with a styled heading.
 */
function tocEntryStyle(ctx: ParseContext, entryLevel: number): TocEntryStyle {
  const id = `TOC${entryLevel}`;
  const props = styleOnlyParaProps(ctx, id);
  const direction: TextDirection = props.bidi === true ? "rtl" : "ltr";
  const runProps = baseRunProps(ctx, id, undefined);
  const fontSizePt = runProps.fontSizePt ?? runProps.fontSizeCsPt;
  if (fontSizePt !== undefined) props.fontSizePt = fontSizePt;
  return { id, direction, fontSizePt, runProps, layout: layoutFor(props, direction, ctx.justifyShrinks) };
}

/**
 * Whether a session's switches can all be honoured against this document. A
 * `\b` naming a bookmark the body does not hold cannot: the caller keeps the
 * cached result rather than collect the whole document into a field that asked
 * for a chapter.
 */
function tocFieldSupported(session: TocSession, ctx: ParseContext): boolean {
  const instruction = session.instruction;
  if (instruction === undefined) return false;
  if (instruction.bookmark === undefined) return true;
  const ranges = ctx.bookmarks?.ranges.get(instruction.bookmark);
  return ranges !== undefined && ranges.length > 0;
}

/**
 * Fresh entry drafts for one TOC session from the document headings. Levels
 * outside the instruction's range are skipped, and a `\b` field keeps only the
 * headings inside its bookmark's ranges (Word merges several ranges of one name
 * into a single list, so a heading inside any of them is listed); each entry
 * takes the `TOCn` style of its entry level — resolved for its spacing,
 * indents, direction and text size — a right dot-leader tab (the cached
 * entries' own direct formatting, else the section body width), and runs of the
 * heading's display text plus a live `PAGEREF` page number: omitted with its gap
 * when the entry level takes no number (`\n`), or joined by the authored
 * separator instead of a tab (`\p`).
 */
function tocEntryDrafts(session: TocSession, headings: TocHeading[], ctx: ParseContext, bodyWidthIn: number): BlockDraft[] {
  // The caller only refreshes supported sessions (see `tocFieldSupported`).
  const instruction = session.instruction!;
  const ranges = instruction.bookmark === undefined ? undefined : ctx.bookmarks?.ranges.get(instruction.bookmark);
  const inRange = (heading: TocHeading): boolean =>
    ranges === undefined || ranges.some((range) => heading.block >= range.from && heading.block < range.to);
  const entries = headings.filter(
    (heading) => heading.node.level >= instruction.minLevel && heading.node.level <= instruction.maxLevel && inRange(heading),
  );
  if (entries.length === 0) {
    return [{ node: { type: "paragraph", runs: [{ text: "No table of contents entries found." }], direction: "ltr" } }];
  }
  const tabStops = session.tabStops ?? [{ pos: bodyWidthIn, align: "right" as const, leader: "dot" as const }];
  const styles = new Map<number, TocEntryStyle>();
  const drafts: BlockDraft[] = [];
  for (const heading of entries) {
    const node = heading.node;
    const entryLevel = node.level - instruction.minLevel + 1;
    let style = styles.get(entryLevel);
    if (style === undefined) {
      style = tocEntryStyle(ctx, entryLevel);
      styles.set(entryLevel, style);
    }
    const omitPage = entryLevel >= instruction.omitPageMin && entryLevel <= instruction.omitPageMax;
    const hyperlink = !instruction.hyperlink
      ? undefined
      : heading.outlineIndex !== undefined
        ? `#aared-heading-${heading.outlineIndex + 1}`
        : `#${tocAnchorFor(ctx, node)}`;
    const link = hyperlink === undefined ? {} : { hyperlink };
    const runs: TextRunModel[] = [];
    /**
     * One entry run: the heading run's text in the entry style's own
     * formatting. A field, a link, a picture and a footnote marker never join
     * an entry; a tab or line break inside one (`\w`, `\x`) is entry text.
     */
    const addText = (text: string, extra?: Partial<TextRunModel>): void => {
      const run = finalizeRun(ctx, style.runProps, text, text, { ...link, ...extra });
      if (run !== undefined) runs.push(run);
    };
    if (node.marker !== undefined) addText(node.marker + (node.markerSuffix === "nothing" ? "" : " "));
    for (const run of node.runs) {
      if (run.image !== undefined || run.footnoteRef !== undefined) continue;
      addText(run.text, run.ltrRun === true ? { ltrRun: true } : undefined);
    }
    const displayText = runs.map((run) => run.text).join("");
    if (!omitPage) {
      addText(instruction.separator ?? "\t");
      // Always a run, even with no saved number: layout mode fills it from the
      // page it laid the target out on, and an entry added since the TOC was
      // last saved is exactly the entry that has nothing cached to fall back
      // to. A run that vanished when empty would print no number at all.
      runs.push({
        text: cachedTocPage(session, displayText),
        field: "PAGEREF",
        ...(style.fontSizePt === undefined ? {} : { fontSizePt: style.fontSizePt }),
        ...link,
      });
    }
    drafts.push({
      node: {
        type: "paragraph",
        runs,
        direction: style.direction,
        tabStops,
        styleId: style.id,
        ...style.layout,
      },
    });
  }
  return drafts;
}

/** Drop TOC markers over `[from, to)`, keeping cached entries as the paragraphs they wrap. Returns the kept count. */
function unwrapTocCached(drafts: BlockDraft[], from: number, to: number): number {
  const kept: BlockDraft[] = [];
  for (let index = from; index < to; index++) {
    const draft = drafts[index]!;
    if ("toc" in draft) {
      if (draft.toc.kind === "cached") kept.push(draft.toc.inner);
    } else {
      kept.push(draft);
    }
  }
  drafts.splice(from, to - from, ...kept);
  return kept.length;
}

/**
 * Replace each balanced TOC range in one section's drafts with refreshed
 * entries. A TOC without a bookmark switch spans the document, so every
 * section splices from the same heading list; a TOC never spans sections, so
 * an unbalanced range keeps its cache (as does a stray end elsewhere).
 */
function resolveTocRanges(drafts: BlockDraft[], headings: TocHeading[], ctx: ParseContext, bodyWidthIn: number): void {
  let index = 0;
  while (index < drafts.length) {
    const draft = drafts[index]!;
    if (!("toc" in draft) || draft.toc.kind !== "begin") {
      index++;
      continue;
    }
    const session = draft.toc.session;
    let end = -1;
    for (let scan = index + 1; scan < drafts.length; scan++) {
      const candidate = drafts[scan]!;
      if (!("toc" in candidate)) continue;
      // Cached entries sit between the markers; a second begin means the
      // first never balanced — never refresh half a field.
      if (candidate.toc.kind === "begin") break;
      if (candidate.toc.kind === "end") {
        end = scan;
        break;
      }
    }
    if (end === -1) {
      unwrapTocCached(drafts, index, drafts.length);
      break;
    }
    if (!tocFieldSupported(session, ctx)) {
      // A switch we cannot honour, found in the opener, in a later instruction
      // chunk, or as a `\b` naming a bookmark this document has no range for:
      // keep the saved result rather than refresh half-understood.
      index += unwrapTocCached(drafts, index, end + 1);
      continue;
    }
    const replacement = tocEntryDrafts(session, headings, ctx, bodyWidthIn);
    drafts.splice(index, end - index + 1, ...replacement);
    index += replacement.length;
  }
  // A stray end (its begin lived in another section) drops; surviving cached
  // entries unwrap to the paragraphs they wrap.
  for (let scan = drafts.length - 1; scan >= 0; scan--) {
    const draft = drafts[scan]!;
    if (!("toc" in draft)) continue;
    if (draft.toc.kind === "cached") drafts.splice(scan, 1, draft.toc.inner);
    else drafts.splice(scan, 1);
  }
}

function tableEdge(edge: XmlElement): DocxTableEdge | undefined {
  const val = attr(edge, "val");
  if (val === undefined) return undefined;
  if (val === "nil" || val === "none") return "none";
  return paragraphBorder(edge);
}

function readEdges<K extends string>(parent: XmlElement | undefined, names: readonly K[]): Partial<Record<K, DocxTableEdge>> | undefined {
  if (parent === undefined) return undefined;
  const out: Partial<Record<K, DocxTableEdge>> = {};
  for (const edge of parent.children) {
    const name = names.find((candidate) => candidate === edge.local || (candidate === "left" && edge.local === "start") || (candidate === "right" && edge.local === "end"));
    if (name === undefined) continue;
    const resolved = tableEdge(edge);
    if (resolved !== undefined) out[name] = resolved;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

function readMargins(parent: XmlElement | undefined): InsetsModel | undefined {
  if (parent === undefined) return undefined;
  const out: InsetsModel = {};
  for (const edge of parent.children) {
    const value = numberAttr(edge, "w", "w:tcMar");
    if (value === undefined) continue;
    const inches = twipsToInches(value);
    if (edge.local === "top") out.top = inches;
    else if (edge.local === "bottom") out.bottom = inches;
    else if (edge.local === "left" || edge.local === "start") out.left = inches;
    else if (edge.local === "right" || edge.local === "end") out.right = inches;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/** Word's Normal Table default cell padding: 0.075in at the sides. */
const NORMAL_TABLE_MARGINS: InsetsModel = { top: 0, bottom: 0, left: twipsToInches(108), right: twipsToInches(108) };

/** Direct formatting over the style's: a side the direct edge states wins. */
function mergeTableBorders<K extends keyof DocxTableBorders>(
  base: Partial<DocxTableBorders> | undefined,
  over: Partial<DocxTableBorders> | undefined,
): Partial<DocxTableBorders> | undefined {
  if (base === undefined) return over;
  if (over === undefined) return base;
  const merged: Partial<DocxTableBorders> = { ...base };
  for (const key of Object.keys(over) as K[]) {
    const edge = over[key];
    if (edge !== undefined) merged[key] = edge;
  }
  return Object.keys(merged).length > 0 ? merged : undefined;
}

/**
 * Fill each cell's shading and borders from its table style's regions, under the
 * direct formatting already read. Walks the finished rows so every cell's spans
 * are final.
 *
 * Unlike DrawingML's `a:tcBdr`, Word's `w:tcBorders` names the four sides of the
 * cell it applies to and has no inside/outside distinction, so a region simply
 * overlays those four sides wherever it applies.
 */
function applyTableStyleCells(rows: DocxTableRowNode[], ctx: ParseContext, styleId: string | undefined, flags: TableStyleFlags): void {
  if (styleId === undefined) return;
  for (const [rowIndex, row] of rows.entries()) {
    let column = 0;
    for (const cell of row.cells) {
      const columnSpan = cell.columnSpan ?? 1;
      const rowSpan = cell.rowSpan ?? 1;
      const use = tableCellRegions(flags, rowIndex, rowIndex + rowSpan - 1, column, column + columnSpan - 1);
      const styleBorders: NonNullable<DocxTableCellNode["borders"]> = {};
      let fill: string | undefined;
      for (const region of tableStyleRegionsFor(ctx.styles, styleId, use)) {
        const tcPr = region.tcPr;
        if (tcPr === undefined) continue;
        const shading = firstChild(tcPr, "shd");
        const colour = shading === undefined ? undefined : hexColor(attr(shading, "fill"));
        if (colour !== undefined) fill = colour;
        const edges = readEdges(firstChild(tcPr, "tcBorders"), ["top", "right", "bottom", "left"] as const);
        if (edges === undefined) continue;
        for (const side of ["top", "right", "bottom", "left"] as const) {
          const edge = edges[side];
          if (edge !== undefined) styleBorders[side] = edge;
        }
      }
      // Direct formatting over the style (probe `docx-table-style`: a header
      // cell with its own `w:shd` painted that fill, not the region's).
      if (fill !== undefined && cell.shading === undefined) cell.shading = fill;
      const merged = mergeTableBorders(styleBorders, cell.borders);
      if (merged !== undefined) cell.borders = merged;
      column += columnSpan;
    }
  }
}

/**
 * A list marker's own formatting: the paragraph's run defaults with the numbering
 * level's `w:rPr` on top, resolved like any run so colour, weight, size and
 * family come out the same way. Absent when the level authors no formatting.
 */
function markerRunFor(ctx: ParseContext, level: NumberingLevel, styleId: string | undefined, marker: string): TextRunModel | undefined {
  if (level.rPr === undefined) return undefined;
  const props = baseRunProps(ctx, styleId, undefined);
  readRunProps(level.rPr, props, ctx.theme);
  return finalizeRun(ctx, props, marker, marker);
}

/**
 * A Wingdings or Symbol marker (Word's default bullet is Symbol's private-use
 * U+F0B7) as its Unicode shape, with the symbol font dropped so it paints in
 * any font that has it (`symbolFontStandIn`). Anything else is unchanged.
 */
function symbolMarker(marker: string, markerRun: TextRunModel | undefined): { marker: string; markerRun: TextRunModel | undefined } {
  const standIn = symbolFontStandIn(markerRun?.fontFamily, marker);
  if (standIn === undefined || markerRun === undefined) return { marker, markerRun };
  const run: TextRunModel = { ...markerRun, text: standIn };
  delete run.fontFamily;
  return { marker: standIn, markerRun: run };
}

/**
 * A cell's parse result: the cell (absent for a vertical-merge continuation),
 * the grid columns it covers, and whether Word's own saved layout put a page
 * break right before its first paragraph.
 */
interface TableCellParse {
  cell: DocxTableCellNode | undefined;
  columnSpan: number;
  leadingPageBreak: boolean;
}

/**
 * Where a cell sits, for the table style's region resolver. The cell's own span
 * is not known yet (a vertical merge grows it at the continuation below), so
 * text properties - which do not depend on a span - are what resolve during the
 * cell's own parse.
 */
interface TableCellPosition {
  styleId: string | undefined;
  flags: TableStyleFlags;
  rowIndex: number;
  column: number;
}

function parseTableCell(
  cell: XmlElement,
  ctx: ParseContext,
  origins: Array<DocxTableCellNode | undefined>,
  column: number,
  gridColumns: number,
  position: TableCellPosition,
): TableCellParse {
  throwIfAborted(ctx.signal);
  // A TOC field never lives in a table cell: parsing one there would corrupt
  // body-level tracking, so cells parse untracked (their headings still join
  // a body TOC through the later draft traversal).
  const savedToc = ctx.toc;
  ctx.toc = undefined;
  // The cell's style regions supply the text properties the paragraph and
  // character styles leave unsaid (see `ParseContext.tableStyleText`).
  const savedStyleText = ctx.tableStyleText;
  ctx.tableStyleText = position.styleId === undefined
    ? undefined
    : tableStyleTextPropsFor(ctx.styles, position.styleId, tableCellRegions(position.flags, position.rowIndex, position.rowIndex, position.column, position.column));
  try {
    return parseTableCellInner(cell, ctx, origins, column, gridColumns);
  } finally {
    ctx.toc = savedToc;
    ctx.tableStyleText = savedStyleText;
  }
}

function parseTableCellInner(
  cell: XmlElement,
  ctx: ParseContext,
  origins: Array<DocxTableCellNode | undefined>,
  column: number,
  gridColumns: number,
): TableCellParse {
  throwIfAborted(ctx.signal);
  const tcPr = firstChild(cell, "tcPr");
  const width = tcPr === undefined ? undefined : firstChild(tcPr, "tcW");
  const widthValue = width === undefined ? undefined : numberAttr(width, "w", "w:tcW");
  const span = tcPr === undefined ? undefined : firstChild(tcPr, "gridSpan");
  // Cut to the columns left in the table: the renderers loop over the span.
  const columnSpan = clampSpan(span === undefined ? undefined : numberAttr(span, "val", "w:gridSpan"), gridColumns - column) ?? 1;
  const merge = tcPr === undefined ? undefined : firstChild(tcPr, "vMerge");
  const mergeValue = merge === undefined ? undefined : (attr(merge, "val") ?? "continue");
  if (mergeValue === "continue") {
    // Vertically merged continuation: the cell above grows, this grid
    // position emits nothing. A continuation with no origin (malformed
    // grid) falls through as a normal cell rather than vanishing.
    const origin = origins[column];
    if (origin !== undefined) {
      origin.rowSpan = (origin.rowSpan ?? 1) + 1;
      return { cell: undefined, columnSpan, leadingPageBreak: false };
    }
  }
  const vAlign = tcPr === undefined ? undefined : firstChild(tcPr, "vAlign");
  const vAlignValue = vAlign === undefined ? undefined : attr(vAlign, "val");
  const shading = tcPr === undefined ? undefined : firstChild(tcPr, "shd");
  const fill = shading === undefined ? undefined : hexColor(attr(shading, "fill"));
  const cellMargins = readMargins(tcPr === undefined ? undefined : firstChild(tcPr, "tcMar"));
  const cellBorders = readEdges(tcPr === undefined ? undefined : firstChild(tcPr, "tcBorders"), ["top", "right", "bottom", "left"] as const);
  // Tables break between rows, never inside a cell, so a rendered-page-break
  // hint in a cell's own content has nowhere to go and is dropped — except
  // one that leads the cell's very first block, before anything else:
  // that is Word's own signal that the row itself started a fresh page
  // (`DocxTableRowNode.pageBreakBefore`, set by the caller for column 0).
  const isHint = (draft: BlockDraft): boolean => "node" in draft && draft.node.type === "pageBreak" && draft.node.strength === "preferred";
  const blocks = parseBlocks(cell, ctx);
  const leadingPageBreak = blocks.length > 0 && isHint(blocks[0]!);
  const cellNode: DocxTableCellNode = {
    nodes: groupListItems(blocks.filter((draft) => !isHint(draft))),
    ...(vAlignValue === "center"
      ? { verticalAlign: "middle" as const }
      : vAlignValue === "bottom"
        ? { verticalAlign: "bottom" as const }
        : {}),
    ...(fill === undefined ? {} : { shading: fill }),
    ...(cellMargins === undefined ? {} : { margins: cellMargins }),
    ...(cellBorders === undefined ? {} : { borders: cellBorders }),
    ...(columnSpan <= 1 ? {} : { columnSpan }),
    ...(widthValue === undefined ? {} : { width: twipsToInches(widthValue) }),
  };
  origins[column] = cellNode;
  for (let index = 1; index < columnSpan; index++) origins[column + index] = undefined;
  return { cell: cellNode, columnSpan, leadingPageBreak };
}

function parseTable(table: XmlElement, ctx: ParseContext): FlowNode | undefined {
  throwIfAborted(ctx.signal);
  const tblPr = firstChild(table, "tblPr");
  // `w:bidiVisual` (ECMA-376 17.4.1) lays the logical columns out right to
  // left. Word keeps writing `w:left`/`w:right` in such a table: its borders
  // mean the start/end sides, its cell margins stay physical (probe
  // docx-bidi-visual: a `w:left` border paints on the right, a `w:left`
  // margin stays on the left; Notes §43). DocxTableNode documents both.
  const bidi = tblPr === undefined ? undefined : firstChild(tblPr, "bidiVisual");
  const direction: TextDirection = onOffElement(bidi) === true ? "rtl" : "ltr";
  const grid = firstChild(table, "tblGrid");
  const columns = grid === undefined
    ? []
    : childElements(grid, "gridCol").map((column) => twipsToInches(numberAttr(column, "w", "w:gridCol") ?? 0));
  // The style a table draws from, and the `w:tblLook` switches that say which of
  // its regions apply. Word authors no per-cell shading for a style-driven
  // table (probe `docx-table-style`), so without this it paints plain.
  const tblStyle = tblPr === undefined ? undefined : firstChild(tblPr, "tblStyle");
  const styleId = tblStyle === undefined ? undefined : attr(tblStyle, "val");
  const look = tblPr === undefined ? undefined : firstChild(tblPr, "tblLook");
  const lookVal = tableLookFromVal(look === undefined ? undefined : attr(look, "val"));
  // Word writes both forms and they agree; the explicit attributes win, and the
  // legacy bitmask is the only source in files that carry just `w:val`.
  const lookOn = (name: string): boolean => {
    const value = look === undefined ? undefined : attr(look, name);
    if (value !== undefined) return onOffValue(value) === true;
    return lookVal?.[name as "firstRow"] ?? false;
  };
  const rowElements = childElements(table, "tr");
  const columnCount = Math.max(
    columns.length,
    ...rowElements.map((row) => childElements(row, "tc").length),
  );
  const flags: TableStyleFlags = {
    rows: rowElements.length,
    columns: Math.max(columnCount, 1),
    firstRow: lookOn("firstRow"),
    lastRow: lookOn("lastRow"),
    firstColumn: lookOn("firstColumn"),
    lastColumn: lookOn("lastColumn"),
    bandRow: !(lookOn("noHBand")),
    bandColumn: !(lookOn("noVBand")),
  };
  const rows: DocxTableRowNode[] = [];
  const origins: Array<DocxTableCellNode | undefined> = [];
  for (const [rowIndex, row] of rowElements.entries()) {
    throwIfAborted(ctx.signal);
    // `w:trHeight` is a child of `w:trPr`, not of `w:tr` (ECMA-376 17.4.81):
    // reading it off the row can never match, so every authored row height a
    // Word file saves was silently dropped. `w:hRule` is not read, because the
    // consumer treats an authored height as a floor (`atLeast`, Word's
    // default) and takes the larger of content and height — see
    // `pagination.ts` `foldTableHeightPx` and `DocxTableRowNode.height`.
    const trPr = firstChild(row, "trPr");
    const height = trPr === undefined ? undefined : firstChild(trPr, "trHeight");
    const heightValue = height === undefined ? undefined : numberAttr(height, "val", "w:trHeight");
    const cells: DocxTableCellNode[] = [];
    let column = 0;
    // Word's own saved break leading the row's first cell means the whole
    // row started a fresh page: any cell can carry the marker (all of a
    // row's cells share its page), but only the first is checked — the
    // common, unambiguous case parsed so far.
    let pageBreakBefore = false;
    const rowCells = childElements(row, "tc");
    const gridColumns = Math.max(columns.length, rowCells.length);
    for (const cell of rowCells) {
      throwIfAborted(ctx.signal);
      const { cell: parsed, columnSpan, leadingPageBreak } = parseTableCell(cell, ctx, origins, column, gridColumns, {
        styleId,
        flags,
        rowIndex,
        column,
      });
      if (column === 0 && leadingPageBreak) pageBreakBefore = true;
      if (parsed !== undefined) cells.push(parsed);
      column += columnSpan;
    }
    rows.push({
      ...(heightValue === undefined ? {} : { height: twipsToInches(heightValue) }),
      ...(pageBreakBefore ? { pageBreakBefore: true as const } : {}),
      cells,
    });
  }
  if (rows.length === 0) return undefined;
  const tblW = tblPr === undefined ? undefined : firstChild(tblPr, "tblW");
  const widthValue = tblW === undefined || (attr(tblW, "type") ?? "dxa") !== "dxa" ? undefined : numberAttr(tblW, "w", "w:tblW");
  // The style's own `w:tblBorders` and `w:tblCellMar` are the base the direct
  // ones layer over: Word paints `Table Grid`'s borders from `styles.xml`, so a
  // table with no `w:tblBorders` of its own is not unbordered (probe
  // `docx-table-style`).
  let styleBorders: Partial<DocxTableBorders> | undefined;
  const styleMargins: InsetsModel = {};
  // Each derived style overrides only the edges/sides it states; taking
  // just the final style drops the properties inherited through basedOn.
  for (const style of tableStyleAncestry(ctx.styles, styleId)) {
    const properties = style.whole.tblPr;
    if (properties === undefined) continue;
    styleBorders = mergeTableBorders(
      styleBorders,
      readEdges(firstChild(properties, "tblBorders"), ["top", "right", "bottom", "left", "insideH", "insideV"] as const),
    );
    Object.assign(styleMargins, readMargins(firstChild(properties, "tblCellMar")));
  }
  const borders = mergeTableBorders(
    styleBorders,
    readEdges(tblPr === undefined ? undefined : firstChild(tblPr, "tblBorders"), ["top", "right", "bottom", "left", "insideH", "insideV"] as const),
  );
  const cellMargins = {
    ...NORMAL_TABLE_MARGINS,
    ...styleMargins,
    ...readMargins(tblPr === undefined ? undefined : firstChild(tblPr, "tblCellMar")),
  };
  // Cell shading and borders the style resolved to, applied once every row's
  // spans are known: which of a region's edges a cell draws as outer or inside
  // depends on the span, and a vertical merge's `rowSpan` is only final at the
  // continuation below it.
  applyTableStyleCells(rows, ctx, styleId, flags);
  // Measured from the start side in both directions (probe docx-bidi-visual:
  // a 1in indent moved a right-to-left table in from the right margin).
  const tblInd = tblPr === undefined ? undefined : firstChild(tblPr, "tblInd");
  const indentValue = tblInd === undefined || (attr(tblInd, "type") ?? "dxa") !== "dxa" ? undefined : numberAttr(tblInd, "w", "w:tblInd");
  return {
    type: "table",
    direction,
    columns,
    rows,
    ...(widthValue === undefined || widthValue <= 0 ? {} : { width: twipsToInches(widthValue) }),
    ...(borders === undefined ? {} : { borders }),
    cellMargins,
    ...(indentValue === undefined || indentValue === 0 ? {} : { indent: twipsToInches(indentValue) }),
  };
}

/** One bookmark boundary at body block `at` (half-open): start opens a range, end closes it. */
function bookmarkBoundary(ctx: ParseContext, at: number, name: string | undefined, id: string | undefined, queue: boolean): void {
  const index = ctx.bookmarks;
  if (index === undefined || id === undefined || id === "") return;
  if (name !== undefined) {
    if (queue) index.events.push({ kind: "start", id, name });
    else {
      index.open.set(id, { name, from: at, section: index.section });
    }
    return;
  }
  if (queue) {
    index.events.push({ kind: "end", id });
    return;
  }
  const start = index.open.get(id);
  if (start === undefined) return;
  index.open.delete(id);
  // A start and its end in different sections would pair two index spaces.
  if (start.section !== index.section) return;
  const ranges = index.ranges.get(start.name);
  if (ranges === undefined) index.ranges.set(start.name, [{ from: start.from, to: at }]);
  else ranges.push({ from: start.from, to: at });
}

/**
 * Apply the boundaries queued by the paragraph that filled body blocks
 * `[at, after)`. A bookmark written inside a paragraph spans exactly that
 * paragraph's blocks — the right granularity when what a `\b` TOC collects is
 * headings, and a heading is one block — so its start takes `at` and its end
 * `after`. (A bookmark written *between* paragraphs is not queued: it is
 * recorded against the boundary index directly, which is already half-open.)
 */
function applyBookmarkEvents(ctx: ParseContext, at: number, after: number): void {
  const index = ctx.bookmarks;
  if (index === undefined) return;
  if (index.events.length === 0) return;
  for (const event of index.events) {
    if (event.kind === "start") {
      index.open.set(event.id, { name: event.name, from: at, section: index.section });
      continue;
    }
    const start = index.open.get(event.id);
    if (start === undefined || start.section !== index.section) continue;
    index.open.delete(event.id);
    const ranges = index.ranges.get(start.name);
    if (ranges === undefined) index.ranges.set(start.name, [{ from: start.from, to: after }]);
    else ranges.push({ from: start.from, to: after });
  }
  index.events = [];
}

function parseBlocks(parent: XmlElement, ctx: ParseContext): BlockDraft[] {  throwIfAborted(ctx.signal);
  const drafts: BlockDraft[] = [];
  for (const child of parent.children) {
    throwIfAborted(ctx.signal);
    if (child.local === "p") {
      const at = drafts.length;
      pushParagraphDrafts(drafts, parseParagraph(child, ctx), ctx);
      applyBookmarkEvents(ctx, at, drafts.length);
      continue;
    }
    if (child.local === "bookmarkStart") {
      bookmarkBoundary(ctx, drafts.length, attr(child, "name"), attr(child, "id"), false);
      continue;
    }
    if (child.local === "bookmarkEnd") {
      bookmarkBoundary(ctx, drafts.length, undefined, attr(child, "id"), false);
      continue;
    }
    if (child.local === "tbl") {
      const table = parseTable(child, ctx);
      if (table !== undefined) drafts.push({ node: table });
      continue;
    }
    if (child.local === "sdt") {
      const content = firstChild(child, "sdtContent");
      if (content !== undefined) drafts.push(...parseBlocks(content, ctx));
    }
  }
  return drafts;
}

/**
 * A numbered heading takes its number in Word's sequence and paints its
 * marker, as a list item paints its own (probe `numbered-heading-probe`,
 * Word 16: `ListString` `1.` for the Heading 1 paragraph, painted `1. Project
 * Overview` with the marker at the paragraph start and the text after the
 * level's marker gap). The marker leads `text` so outlines, titles and
 * refreshed TOC entries read what Word shows; `runs` keep the authored text
 * alone. A suppressed level (`w:numFmt w:val="none"`) still takes its number
 * but paints no marker.
 */
function numberHeading(ctx: ParseContext, parsed: ParagraphDraft, node: DocxHeadingNode): void {
  const numId = parsed.numId!;
  const ilvl = parsed.ilvl ?? 0;
  const { level, knownLevels, counters } = advanceCounters(ctx, numId, ilvl);
  // Numbering indents apply only when the paragraph sets none of its own,
  // exactly as for list items.
  if (node.leftIndent === undefined && level.indentLeft !== undefined) node.leftIndent = level.indentLeft;
  if (node.firstLineIndent === undefined && level.indentHanging !== undefined) node.firstLineIndent = -level.indentHanging;
  const authored = markerFor(level, knownLevels, ilvl, counters);
  if (authored === "") return;
  const { marker, markerRun } = symbolMarker(authored, markerRunFor(ctx, level, parsed.styleId, authored));
  node.marker = marker;
  if (markerRun !== undefined) node.markerRun = markerRun;
  if (level.suffix === "space" || level.suffix === "nothing") node.markerSuffix = level.suffix;
  node.text = marker + (level.suffix === "nothing" ? "" : " ") + node.text;
}

/**
 * Remember one cached TOC paragraph's entry text and page number for
 * carry-over: a refreshed entry matching a cached one keeps its page number
 * in flow mode (layout mode resolves the live one). The page number is
 * whatever follows the last tab, as Word writes entry, tab, `PAGEREF`.
 */
function recordTocCache(session: TocSession, nodes: FlowNode[]): void {
  for (const node of nodes) {
    if (node.type !== "paragraph" && node.type !== "heading") continue;
    if (session.tabStops === undefined && node.tabStops !== undefined) session.tabStops = node.tabStops;
    const joined = node.runs.map((run) => run.text).join("");
    // The entry text runs up to the gap that introduces the page number: the
    // last tab (Word's default, or the separator a `\p` switch authors). A
    // cached entry's `PAGEREF` is not a live field, so its result is plain
    // text after that gap.
    const tab = joined.lastIndexOf("\t");
    const entry = (tab === -1 ? joined : joined.slice(0, tab)).trim();
    const page = (tab === -1 ? "" : joined.slice(tab + 1)).trim();
    if (entry !== "") session.cached.push({ entry, page });
  }
}

/** A parsed paragraph's nodes as drafts: the first paragraph of a numbered one is the list item. */
function pushParagraphDrafts(drafts: BlockDraft[], parsed: ParagraphDraft, ctx: ParseContext): void {
  if (parsed.tocBoundary === "open" || parsed.tocBoundary === "single") {
    // The opening paragraph's own nodes are the first cached entry: record
    // them for page-number carry-over, then hold everything for the splice.
    recordTocCache(parsed.tocSession!, parsed.nodes);
    drafts.push({ toc: { kind: "begin", session: parsed.tocSession! } });
    for (const node of parsed.nodes) drafts.push({ toc: { kind: "cached", inner: { node } } });
    if (parsed.tocBoundary === "single") drafts.push({ toc: { kind: "end" } });
    return;
  }
  if (parsed.tocBoundary === "inside") {
    recordTocCache(ctx.toc!.active!, parsed.nodes);
    for (const node of parsed.nodes) drafts.push({ toc: { kind: "cached", inner: { node } } });
    return;
  }
  let listed = false;
  const pushNode = (node: FlowNode): void => {
    if (!listed && parsed.numId !== undefined && node.type === "paragraph") {
      drafts.push({ listItem: { numId: parsed.numId, item: listItemFor(ctx, parsed.numId, parsed.ilvl ?? 0, node) } });
      listed = true;
    } else if (!listed && parsed.numId !== undefined && node.type === "heading") {
      numberHeading(ctx, parsed, node);
      listed = true;
      drafts.push({ node });
    } else {
      drafts.push({ node });
    }
  };
  if (parsed.tocBoundary === "close") {
    // The closing paragraph is usually the empty field end alone; a paragraph
    // carrying real content past the end keeps it as body content.
    const onlyEmpty = parsed.nodes.length === 1
      && parsed.nodes[0]?.type === "paragraph"
      && parsed.nodes[0].runs.length === 0
      && parsed.nodes[0].floats === undefined;
    if (!onlyEmpty) for (const node of parsed.nodes) pushNode(node);
    drafts.push({ toc: { kind: "end" } });
    return;
  }
  for (const node of parsed.nodes) pushNode(node);
}

function footnoteDirection(nodes: FlowNode[]): TextDirection {
  for (const node of nodes) {
    if (node.type === "paragraph" || node.type === "heading") return node.direction;
  }
  return "ltr";
}

/**
 * Inflate and parse the document's sections: reads the notes, streams the
 * body, and keeps nothing afterwards. `root` order is document order
 * within each section and each entry is an independent virtualization block
 * (see the `DocumentSection` contract in `core/model.ts`). Footnote entries
 * resolve onto the last section in reference order so flow rendering keeps a
 * single end-of-document list; layout pagination reassigns them per page.
 */
export async function parseDocxSections(
  pkg: OpcPackage,
  index: DocxIndex,
  signal?: AbortSignal,
): Promise<DocumentSection[]> {
  throwIfAborted(signal);
  const mediaContentType = (part: string): string => pkg.mediaContentType(part);
  // Notes first: a footnote reference in the body becomes a marker only when its note exists.
  const notes = await readNotes(pkg, index, signal);
  const charts = await readPartCharts(pkg, index.documentPart, index.theme, signal);
  const missing = { drawings: 0 };
  const bodyRels = await readPartRels(pkg, index.documentPart, signal);
  const ctx = partContext(index, notes, bodyRels, mediaContentType, charts, missing, signal);
  // Body-level table-of-contents fields refresh from the parsed headings;
  // tracking stays off everywhere else (see `ParseContext.toc`). Body
  // bookmarks are recorded in the same pass, for a `\b` field's range.
  ctx.toc = { serial: 0 };
  ctx.bookmarks = { ranges: new Map(), open: new Map(), events: [], section: 0 };
  const drafts = await parseBodySections(pkg, index.documentPart, ctx);
  // Refresh body TOCs from the document headings before list grouping: every
  // section splices from the same document-order heading list, and refreshed
  // entries are plain paragraphs that group like any other.
  const tocCounter = { count: 0 };
  const tocHeadings: TocHeading[] = [];
  for (const draft of drafts) collectTocHeadings(draft.drafts, tocHeadings, tocCounter);
  for (const draft of drafts) {
    const page = draft.properties.page;
    resolveTocRanges(draft.drafts, tocHeadings, ctx, twipsToInches(page.widthTwips - page.margins.left - page.margins.right));
  }
  const sections: DocumentSection[] = [];
  // Sections usually inherit the same header/footer parts: parse each once
  // and share the (read-only) node list between the sections using it.
  const headerFooters = new Map<string, FlowNode[]>();
  const headerFooter = async (part: string): Promise<FlowNode[]> => {
    let nodes = headerFooters.get(part);
    if (nodes === undefined) {
      nodes = await parseHeaderFooterPart(pkg, index, notes, part, mediaContentType, missing, signal);
      headerFooters.set(part, nodes);
    }
    return nodes;
  };
  // A section that names no header or footer of a kind is linked to the
  // previous one and shows the same, each kind (default, first, even) on its
  // own (ECMA-376 17.10.5; probe `docx-first-header`: a section with its own
  // default but no first header showed the earlier section's first header).
  const headerParts: HeaderFooterParts = {};
  const footerParts: HeaderFooterParts = {};
  for (const draft of drafts) {
    throwIfAborted(signal);
    const root = groupListItems(draft.drafts);
    const properties: SectionProperties = index.evenAndOddHeaders ? { ...draft.properties, evenAndOddPages: true } : draft.properties;
    const section: DocumentSection = { properties, root };
    for (const kind of ["default", "first", "even"] as const) {
      const ownHeader = draft.properties.headerParts?.[kind];
      const ownFooter = draft.properties.footerParts?.[kind];
      if (ownHeader !== undefined) headerParts[kind] = ownHeader;
      if (ownFooter !== undefined) footerParts[kind] = ownFooter;
      const headerPart = headerParts[kind];
      const footerPart = footerParts[kind];
      const header = headerPart === undefined ? undefined : await headerFooter(headerPart);
      const footer = footerPart === undefined ? undefined : await headerFooter(footerPart);
      if (header !== undefined && header.length > 0) section[kind === "default" ? "header" : kind === "first" ? "firstHeader" : "evenHeader"] = header;
      if (footer !== undefined && footer.length > 0) section[kind === "default" ? "footer" : kind === "first" ? "firstFooter" : "evenFooter"] = footer;
    }
    const firstHeading = root.find((node): node is DocxHeadingNode => node.type === "heading");
    if (firstHeading !== undefined) section.title = firstHeading.text;
    sections.push(section);
  }
  const last = sections[sections.length - 1]!;
  // Counted after the notes below are parsed too: the document's total sits on its first section.
  // A note's relationships live with its own part (`footnotes.xml`, `endnotes.xml`), not the body's.
  const noteContexts = new Map<string, ParseContext>();
  for (const key of ctx.footnoteOrder) {
    throwIfAborted(signal);
    const note = ctx.notes.get(key);
    if (note === undefined) continue;
    const notesPart = key.startsWith("f") ? index.notesParts.footnotes : index.notesParts.endnotes;
    let noteCtx = noteContexts.get(notesPart ?? "");
    if (noteCtx === undefined) {
      const rels = notesPart === undefined ? bodyRels : await readPartRels(pkg, notesPart, signal);
      const noteCharts = notesPart === undefined ? charts : await readPartCharts(pkg, notesPart, index.theme, signal);
      noteCtx = partContext(index, notes, rels, mediaContentType, noteCharts, missing, signal);
      noteCtx.footnoteLabels = ctx.footnoteLabels;
      noteCtx.footnoteOrder = ctx.footnoteOrder;
      noteContexts.set(notesPart ?? "", noteCtx);
    }
    const nodes = groupListItems(parseBlocks(note, noteCtx));
    last.root.push({
      type: "footnote",
      id: key,
      label: ctx.footnoteLabels.get(key) ?? "?",
      nodes,
      direction: footnoteDirection(nodes),
    });
  }
  throwIfAborted(signal);
  if (missing.drawings > 0) sections[0]!.missingIds = missing.drawings;
  return sections;
}

async function parseHeaderFooterPart(
  pkg: OpcPackage,
  index: DocxIndex,
  notes: Map<string, XmlElement>,
  part: string,
  mediaContentType: (part: string) => string,
  missing: { drawings: number },
  signal?: AbortSignal,
): Promise<FlowNode[]> {
  const partRoot = await pkg.readXmlIfPresent(part, signal);
  throwIfAborted(signal);
  if (partRoot === undefined) return [];
  const rels = await readPartRels(pkg, part, signal);
  throwIfAborted(signal);
  const charts = await readPartCharts(pkg, part, index.theme, signal);
  const partCtx = partContext(index, notes, rels, mediaContentType, charts, missing, signal);
  return groupListItems(parseBlocks(partRoot, partCtx));
}

/** Word default page geometry (Letter) used when `w:sectPr` omits a value. */
const DEFAULT_PAGE_TWIPS = { width: 12240, height: 15840 };
const DEFAULT_MARGINS_TWIPS = { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 };

function sectionPropertiesFrom(
  sectPr: XmlElement | undefined,
  relTargets: Map<string, string>,
  adjustLineHeightInTable: boolean,
): SectionProperties {
  const typeChild = sectPr === undefined ? undefined : firstChild(sectPr, "type");
  const typeValue = typeChild === undefined ? undefined : attr(typeChild, "val");
  // An absent type means `nextPage` for an embedded break. `nextColumn` has
  // no page meaning (multi-column layout is out of scope) and degrades to
  // `continuous`, keeping content on the open page.
  const breakType: SectionBreakType =
    typeValue === "continuous" || typeValue === "nextColumn"
      ? "continuous"
      : typeValue === "evenPage"
        ? "evenPage"
        : typeValue === "oddPage"
          ? "oddPage"
          : "nextPage";
  const sizeChild = sectPr === undefined ? undefined : firstChild(sectPr, "pgSz");
  const width = sizeChild === undefined
    ? DEFAULT_PAGE_TWIPS.width
    : numberAttr(sizeChild, "w", "w:pgSz") ?? DEFAULT_PAGE_TWIPS.width;
  const height = sizeChild === undefined
    ? DEFAULT_PAGE_TWIPS.height
    : numberAttr(sizeChild, "h", "w:pgSz") ?? DEFAULT_PAGE_TWIPS.height;
  const orient = sizeChild === undefined ? undefined : attr(sizeChild, "orient");
  const orientation = orient === "landscape" || (orient !== "portrait" && width > height) ? "landscape" : "portrait";
  const marChild = sectPr === undefined ? undefined : firstChild(sectPr, "pgMar");
  const mar = (name: string, fallback: number): number =>
    marChild === undefined ? fallback : numberAttr(marChild, name, "w:pgMar") ?? fallback;
  const headerParts = headerFooterParts(sectPr, "headerReference", relTargets);
  const footerParts = headerFooterParts(sectPr, "footerReference", relTargets);
  const numType = sectPr === undefined ? undefined : firstChild(sectPr, "pgNumType");
  const pageNumberStart = numType === undefined ? undefined : numberAttr(numType, "start", "w:pgNumType");
  const docGrid = sectPr === undefined ? undefined : firstChild(sectPr, "docGrid");
  const gridType = docGrid === undefined ? undefined : attr(docGrid, "type");
  const gridLinePitch = docGrid === undefined ? undefined : numberAttr(docGrid, "linePitch", "w:docGrid");
  return {
    breakType,
    page: {
      widthTwips: width,
      heightTwips: height,
      orientation,
      margins: {
        top: mar("top", DEFAULT_MARGINS_TWIPS.top),
        right: mar("right", DEFAULT_MARGINS_TWIPS.right),
        bottom: mar("bottom", DEFAULT_MARGINS_TWIPS.bottom),
        left: mar("left", DEFAULT_MARGINS_TWIPS.left),
        header: mar("header", DEFAULT_MARGINS_TWIPS.header),
        footer: mar("footer", DEFAULT_MARGINS_TWIPS.footer),
        gutter: mar("gutter", DEFAULT_MARGINS_TWIPS.gutter),
      },
    },
    ...(docGrid === undefined ? {} : {
      docGrid: {
        type: gridType ?? "default",
        ...(gridLinePitch !== undefined ? { linePitch: gridLinePitch } : gridType === "lines" ? { linePitch: 360 } : {}),
        ...(adjustLineHeightInTable ? { adjustLineHeightInTable: true as const } : {}),
      },
    }),
    ...(headerParts === undefined ? {} : { headerParts }),
    ...(footerParts === undefined ? {} : { footerParts }),
    ...(sectPr !== undefined && onOffElement(firstChild(sectPr, "titlePg")) === true ? { titlePage: true as const } : {}),
    ...(pageNumberStart === undefined ? {} : { pageNumberStart }),
  };
}

interface SectionDraft {
  properties: SectionProperties;
  drafts: BlockDraft[];
}

/**
 * Split the body into per-section block drafts, streaming `w:body`'s
 * children so the whole body never exists as one tree: each child is parsed
 * and dropped. A paragraph-embedded `w:pPr/w:sectPr` closes the section that
 * precedes it (the paragraph's own content belongs to the closing section);
 * the last body-level `sectPr` governs the last section.
 */
async function parseBodySections(pkg: OpcPackage, documentPart: string, ctx: ParseContext): Promise<SectionDraft[]> {
  throwIfAborted(ctx.signal);
  const sections: SectionDraft[] = [];
  let current: BlockDraft[] = [];
  let trailing: SectionProperties | undefined;
  const root = await pkg.readXmlIfPresent(documentPart, ctx.signal, {
    streamChildrenOf: "body",
    onChild: (child) => {
      throwIfAborted(ctx.signal);
      switch (child.local) {
        case "p": {
          const at = current.length;
          pushParagraphDrafts(current, parseParagraph(child, ctx), ctx);
          applyBookmarkEvents(ctx, at, current.length);
          const pPr = firstChild(child, "pPr");
          const embedded = pPr === undefined ? undefined : firstChild(pPr, "sectPr");
          if (embedded !== undefined) {
            sections.push({ drafts: current, properties: sectionPropertiesFrom(embedded, ctx.relTargets, ctx.adjustLineHeightInTable) });
            current = [];
            if (ctx.bookmarks !== undefined) ctx.bookmarks.section++;
          }
          return;
        }
        case "bookmarkStart":
          bookmarkBoundary(ctx, current.length, attr(child, "name"), attr(child, "id"), false);
          return;
        case "bookmarkEnd":
          bookmarkBoundary(ctx, current.length, undefined, attr(child, "id"), false);
          return;
        case "tbl": {
          const table = parseTable(child, ctx);
          if (table !== undefined) current.push({ node: table });
          return;
        }
        case "sdt": {
          const content = firstChild(child, "sdtContent");
          if (content !== undefined) current.push(...parseBlocks(content, ctx));
          return;
        }
        case "sectPr":
          trailing = sectionPropertiesFrom(child, ctx.relTargets, ctx.adjustLineHeightInTable);
          return;
        default:
          // bookmarkStart/End and anything else at body level: skipped.
          return;
      }
    },
  });
  if (root === undefined || firstChild(root, "body") === undefined) {
    throw damagedDocument(`document "${documentPart}" is missing w:body`);
  }
  sections.push({ drafts: current, properties: trailing ?? sectionPropertiesFrom(undefined, ctx.relTargets, ctx.adjustLineHeightInTable) });
  return sections;
}
