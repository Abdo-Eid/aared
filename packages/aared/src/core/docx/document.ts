/**
 * Word `document.xml` reader for the DOCX route: body → resolved
 * `DocumentSection` flow blocks. Headless, plain data — DOM painting belongs to
 * `aared/docx`, never here.
 *
 * Coverage: style + numbering inheritance, bidirectional paragraphs
 * (mixed Arabic/Latin runs in one paragraph), heading outline, tables
 * (logical order, `bidiVisual`, spans, shading), lists (bullet + ordered,
 * nested levels), inline images and `wp:anchor` pictures with full placement, footnotes and
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
import { parseHexDigits, toHexColor } from "../color.js";
import { runsText } from "../search.js";
import { readDrawingTheme, themeFontFamily } from "../theme.js";
import { isLtrRunInRtl, usesArabic } from "../typography.js";
import type { DrawingTheme } from "../theme.js";
import { damagedDocument, throwIfAborted } from "../errors.js";
import { TWIPS_PER_PT, emuToInches, ptToInches, twipsToInches } from "../units.js";
import type {
  DocxAnchorPosition,
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
  DocxUnsupportedNode,
  DocumentSection,
  TextAlign,
  FlowNode,
  SectionBreakType,
  SectionProperties,
  TextDirection,
  TextRunModel,
} from "../model.js";
import { characterAncestry, headingLevelForStyle, paragraphAncestry, parseStyles } from "./styles.js";
import type { StyleData } from "./styles.js";
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
      case "bCs":
        out.bold = onOffValue(attr(child, "val"));
        break;
      case "i":
      case "iCs":
        out.italic = onOffValue(attr(child, "val"));
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
        const themeColor = attr(child, "themeColor");
        const themed = themeColor === undefined ? undefined : theme.colors[themeColor];
        if (themed !== undefined) {
          out.color = themed;
        } else {
          const direct = hexColor(attr(child, "val"));
          if (direct !== undefined) out.color = direct;
        }
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
        const on = onOffValue(attr(child, "val")) !== false;
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
  footnoteLabels: Map<string, string>;
  footnoteOrder: string[];
  counters: Map<string, number[]>;
  justifyShrinks: boolean;
  signal?: AbortSignal;
}

function partContext(
  index: DocxIndex,
  notes: Map<string, XmlElement>,
  relTargets: Map<string, string>,
  mediaContentType: (part: string) => string,
  signal?: AbortSignal,
): ParseContext {
  return {
    styles: index.styles,
    numbering: index.numbering,
    notes,
    theme: index.theme,
    relTargets,
    hyperlinks: index.hyperlinks,
    mediaContentType,
    footnoteLabels: new Map(),
    footnoteOrder: [],
    counters: new Map(),
    justifyShrinks: index.justifyShrinks,
    ...(signal === undefined ? {} : { signal }),
  };
}

function headerFooterParts(
  sectPr: XmlElement | undefined,
  local: "headerReference" | "footerReference",
  relTargets: Map<string, string>,
): string[] {
  if (sectPr === undefined) return [];
  const refs = childElements(sectPr, local);
  const rank = (type: string | undefined): number => (type === "default" ? 0 : type === "first" ? 1 : type === "even" ? 2 : 3);
  return refs
    .map((ref) => ({ type: attr(ref, "type"), target: relTargets.get(attr(ref, "id") ?? "") }))
    .filter((ref): ref is { type: string | undefined; target: string } => ref.target !== undefined)
    .sort((a, b) => rank(a.type) - rank(b.type))
    .map((ref) => ref.target);
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
  const justifyShrinks = wordLayoutShrinksSpaces(
    settingsPart === undefined ? undefined : await pkg.readXmlIfPresent(settingsPart, signal),
  );
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

type BlockDraft = { node: FlowNode } | { listItem: { numId: string; ilvl: number; paragraph: DocxParagraphNode } };

interface RunDraft {
  runs: TextRunModel[];
  blocks: Array<DocxImageNode | DocxUnsupportedNode>;
  /** Pictures anchored (`wp:anchor`) in the run: they belong to the paragraph, not to a position in its text. */
  floats: DocxImageNode[];
  /**
   * Pagination signals seen anywhere in the run. They split the paragraph at
   * the run boundary (a break mid-run with text on both sides is vanishingly
   * rare; the text stays together and the break lands after the run).
   */
  pageBreak: boolean;
  softBreak: boolean;
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
  if (props.bold !== undefined) run.bold = props.bold;
  if (props.italic !== undefined) run.italic = props.italic;
  const arabic = usesArabic(text);
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

function parseDrawing(drawing: XmlElement, ctx: ParseContext): DocxImageNode | DocxUnsupportedNode | undefined {
  throwIfAborted(ctx.signal);
  const inline = firstChild(drawing, "inline");
  const anchor = firstChild(drawing, "anchor");
  const placement = inline ?? anchor;
  if (placement === undefined) return undefined;
  const extent = firstChild(placement, "extent");
  const widthEmu = extent === undefined ? undefined : numberAttr(extent, "cx", "wp:extent");
  const heightEmu = extent === undefined ? undefined : numberAttr(extent, "cy", "wp:extent");
  // `a:graphic` nests inside `wp:inline`/`wp:anchor`, never directly under
  // `w:drawing` — a direct-child lookup misses every inline picture.
  const graphic = firstChild(placement, "graphic") ?? firstChild(drawing, "graphic");
  const graphicData = graphic === undefined ? undefined : firstChild(graphic, "graphicData");
  const uri = graphicData === undefined ? undefined : attr(graphicData, "uri");
  const pic = graphicData === undefined ? undefined : firstChild(graphicData, "pic");
  const chart = graphicData === undefined
    ? undefined
    : graphicData.children.find((child) => child.local === "chart");
  const docPr = firstChild(placement, "docPr");
  const alt = docPr === undefined
    ? undefined
    : attr(docPr, "descr") || attr(docPr, "title") || undefined;
  if (chart !== undefined || (uri !== undefined && uri.endsWith("/chart"))) {
    // Chart drawing is PPTX-owned (the shared chart-to-SVG engine); DOCX
    // only adapts the reference into a marked fallback, never a duplicate
    // chart engine.
    const title = alt ?? "Chart";
    return { type: "unsupported", source: "chart", label: title };
  }
  if (pic === undefined || widthEmu === undefined || heightEmu === undefined) {
    return { type: "unsupported", source: "drawing", label: alt ?? "Drawing" };
  }
  const blipFill = firstChild(pic, "blipFill");
  const blip = blipFill === undefined ? undefined : firstChild(blipFill, "blip");
  const embed = blip === undefined ? undefined : attr(blip, "embed");
  const target = embed === undefined ? undefined : ctx.relTargets.get(embed);
  if (target === undefined) return { type: "unsupported", source: "drawing", label: alt ?? "Drawing" };
  return {
    type: "image",
    mediaId: target,
    contentType: ctx.mediaContentType(target),
    width: emuToInches(widthEmu),
    height: emuToInches(heightEmu),
    ...(alt === undefined ? {} : { alt }),
    ...(anchor === undefined ? {} : { anchor: parseAnchor(anchor) }),
  };
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
          draft.pageBreak = true;
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
      case "drawing": {
        flush();
        const image = parseDrawing(child, ctx);
        if (image === undefined) break;
        // An anchored picture is placed relative to its paragraph and never
        // splits the text it sits in.
        if (image.type === "image" && image.anchor !== undefined) {
          draft.floats.push(image);
          break;
        }
        // A picture in a paragraph that also carries text stays on that line
        // (a header logo beside its wordmark); a picture-only paragraph keeps
        // the block image with its host alignment.
        if (image.type === "image" && paragraphText.trim() !== "" && firstChild(child, "inline") !== undefined) {
          draft.runs.push({
            text: "",
            image: {
              mediaId: image.mediaId,
              contentType: image.contentType,
              width: image.width,
              height: image.height,
              ...(image.alt === undefined ? {} : { alt: image.alt }),
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
        const picture = parseVmlPicture(child, ctx);
        if (picture !== undefined) draft.blocks.push(picture);
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
        } else if (type === "end") {
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
  const name = instruction?.trim().split(/\s+/)[0]?.toUpperCase();
  return name === "PAGE" || name === "NUMPAGES" ? name : undefined;
}

function paragraphTextOf(paragraph: XmlElement): string {
  const parts: string[] = [];
  const stack = [paragraph];
  while (stack.length > 0) {
    const current = stack.pop()!;
    if (current.local === "del" || current.local === "moveFrom") continue;
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
}

/** Move one run's content into the open paragraph, splitting on pagination signals found in the run. */
function emitDraft(
  draft: RunDraft,
  runs: TextRunModel[],
  pieces: FlowNode[],
  floats: DocxImageNode[],
  flushRuns: () => void,
  align: TextAlign | undefined,
): void {
  // A rendered-page-break hint is a preferred break *before* the paragraph,
  // and only when no text of the paragraph precedes it: mid-paragraph it is
  // Word's line-level layout cache and must never split the text.
  if (draft.softBreak && !draft.pageBreak && runs.length === 0 && pieces.length === 0) {
    pieces.push(pageBreakNode("preferred"));
  }
  runs.push(...draft.runs);
  floats.push(...draft.floats);
  for (const block of draft.blocks) {
    // Inline images inherit the host paragraph's alignment (a centered
    // source paragraph centers its image); anchored ones never reach here.
    if (block.type === "image" && align !== undefined) block.align = align;
    flushRuns();
    pieces.push(block);
  }
  if (draft.pageBreak) {
    flushRuns();
    pieces.push(pageBreakNode("mandatory"));
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
 * Every field of `DocxParagraphLayout`. A record over the interface's keys, so
 * adding a layout field without listing it here fails to compile instead of
 * silently dropping it when a paragraph becomes a heading.
 */
const LAYOUT_FIELDS: Record<keyof DocxParagraphLayout, true> = {
  fontSizePt: true,
  align: true,
  justifyShrink: true,
  tabStops: true,
  lineSpacing: true,
  lineSpacingPt: true,
  lineSpacingAtLeast: true,
  spaceBeforePt: true,
  spaceAfterPt: true,
  leftIndent: true,
  rightIndent: true,
  firstLineIndent: true,
  shading: true,
  borders: true,
  keepNext: true,
};
const LAYOUT_KEYS = Object.keys(LAYOUT_FIELDS) as Array<keyof DocxParagraphLayout>;

/** The layout a paragraph or heading node carries, present fields only. */
function layoutOf(node: DocxParagraphLayout): DocxParagraphLayout {
  const layout: Record<string, unknown> = {};
  for (const key of LAYOUT_KEYS) if (node[key] !== undefined) layout[key] = node[key];
  return layout as DocxParagraphLayout;
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
  defaultStyleId?: string,
): { merged: ResolvedParaProps; styleId: string | undefined } {
  const directStyle = directPPr === undefined ? undefined : firstChild(directPPr, "pStyle");
  // A paragraph naming no style takes the caller's (a table style's), else the
  // document's default paragraph style (`Normal`), as Word does.
  const styleId = (directStyle === undefined ? undefined : attr(directStyle, "val")) ?? defaultStyleId ?? styles.defaultParagraphStyleId;
  const merged: ResolvedParaProps = {};
  if (styles.defaultPPr !== undefined) readParaProps(styles.defaultPPr, merged);
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

function parseParagraph(paragraph: XmlElement, ctx: ParseContext, defaultStyleId?: string): ParagraphDraft {
  throwIfAborted(ctx.signal);
  const { merged, styleId } = resolveParaProps(ctx.styles, firstChild(paragraph, "pPr"), defaultStyleId);

  const direction: TextDirection = merged.bidi === true ? "rtl" : "ltr";
  const text = paragraphTextOf(paragraph);
  const paraAlign = alignFor(merged.jc, direction);
  const defaultSize = mergedFontSize(ctx, styleId);
  if (defaultSize !== undefined && merged.fontSizePt === undefined) merged.fontSizePt = defaultSize;

  const pieces: FlowNode[] = [];
  const floats: DocxImageNode[] = [];
  const fields: FieldFrame[] = [];
  let runs: TextRunModel[] = [];
  const flushRuns = (): void => {
    if (runs.length === 0) return;
    const node: DocxParagraphNode = { type: "paragraph", runs, direction, ...layoutFor(merged, direction, ctx.justifyShrinks) };
    if (styleId !== undefined) node.styleId = styleId;
    pieces.push(node);
    runs = [];
  };

  /**
   * Read one `w:r` into the open paragraph. `styleOf` maps the run's own
   * `rStyle` to the character style it takes: most runs keep it, a
   * hyperlink's runs fall back to the Hyperlink style.
   */
  const readRun = (run: XmlElement, href: string | undefined, styleOf: (own: string | undefined) => string | undefined): void => {
    const direct = firstChild(run, "rPr");
    const rStyle = direct === undefined ? undefined : firstChild(direct, "rStyle");
    const props = baseRunProps(ctx, styleId, styleOf(rStyle === undefined ? undefined : attr(rStyle, "val")));
    readRunProps(direct, props, ctx.theme);
    const draft: RunDraft = { runs: [], blocks: [], floats: [], pageBreak: false, softBreak: false };
    parseRunContent(run, ctx, props, text, href, draft, fields);
    // Word orders a run without `w:rtl` left to right inside a right-to-left
    // paragraph (probe run-direction).
    for (const piece of draft.runs) {
      if (isLtrRunInRtl(piece.text, direction === "rtl", props.rtl === true)) piece.ltrRun = true;
    }
    emitDraft(draft, runs, pieces, floats, flushRuns, paraAlign);
  };
  const ownStyle = (own: string | undefined): string | undefined => own;

  for (const child of contentChildren(paragraph)) {
    throwIfAborted(ctx.signal);
    if (child.local === "pPr") continue;
    if (child.local === "r") {
      readRun(child, undefined, ownStyle);
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
        if (nested.local === "r") readRun(nested, undefined, ownStyle);
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
        for (const run of contentChildren(child)) {
          if (run.local === "r") readRun(run, undefined, ownStyle);
        }
      }
      continue;
    }
  }
  flushRuns();

  const headingLevel = outlineLevel(styleId, merged, ctx.styles);
  const nodes: FlowNode[] = pieces.map((piece): FlowNode => {
    if (piece.type !== "paragraph" || headingLevel === undefined) return piece;
    return {
      type: "heading",
      level: headingLevel,
      text: runsText(piece.runs),
      direction: piece.direction,
      runs: piece.runs,
      ...layoutOf(piece),
    };
  });

  const filtered = nodes.filter((node) => node.type !== "paragraph" || node.runs.length > 0);
  if (floats.length > 0) {
    if (merged.numId !== undefined) {
      // A list item is not a paragraph node and cannot carry them: its pictures
      // fall back to plain block images ahead of the item.
      filtered.unshift(...floats.map(({ anchor: _anchor, ...image }): DocxImageNode => image));
    } else {
      const host = filtered.find((node) => node.type === "paragraph" || node.type === "heading");
      if (host?.type === "paragraph" || host?.type === "heading") host.floats = floats;
      // A paragraph holding only pictures is still one line tall in Word.
      else filtered.push({ type: "paragraph", runs: [], direction, ...layoutFor(merged, direction, ctx.justifyShrinks), floats });
    }
  }
  if (merged.pageBreakBefore === true) {
    // Inherited `w:pageBreakBefore` (direct or via the style chain) forces
    // the paragraph onto a new page in layout mode.
    filtered.unshift(pageBreakNode("mandatory"));
  }
  return {
    nodes: filtered,
    ...(merged.numId === undefined ? {} : { numId: merged.numId }),
    ilvl: merged.ilvl ?? 0,
  };
}

function mergedFontSize(ctx: ParseContext, styleId: string | undefined): number | undefined {
  const props: ResolvedRunProps = {};
  readRunProps(ctx.styles.defaultRPr, props, ctx.theme);
  for (const style of paragraphAncestry(ctx.styles, styleId)) readRunProps(style.rPr, props, ctx.theme);
  return props.fontSizePt ?? props.fontSizeCsPt;
}

function groupListItems(drafts: BlockDraft[], ctx: ParseContext): FlowNode[] {
  throwIfAborted(ctx.signal);
  const out: FlowNode[] = [];
  let open: { numId: string; items: DocxListItemNode[] } | undefined;
  const close = (): void => {
    if (open !== undefined && open.items.length > 0) out.push({ type: "list", items: open.items });
    open = undefined;
  };
  for (const draft of drafts) {
    throwIfAborted(ctx.signal);
    if (!("listItem" in draft)) {
      close();
      out.push(draft.node);
      continue;
    }
    const { numId, ilvl, paragraph } = draft.listItem;
    const abstractId = ctx.numbering.nums.get(numId);
    const levels = abstractId === undefined ? undefined : ctx.numbering.abstracts.get(abstractId);
    const fallback: NumberingLevel = { numFmt: "bullet", lvlText: BULLET, start: 1 };
    const level = levels?.get(ilvl) ?? fallback;
    const knownLevels = levels ?? new Map<number, NumberingLevel>([[ilvl, level]]);
    if (open === undefined || open.numId !== numId) {
      close();
      open = { numId, items: [] };
    }
    let counters = ctx.counters.get(numId);
    if (counters === undefined) {
      counters = [];
      ctx.counters.set(numId, counters);
    }
    counters[ilvl] = (counters[ilvl] ?? (startFor(ctx.numbering, numId, ilvl, level) - 1)) + 1;
    for (let depth = ilvl + 1; depth < counters.length; depth++) counters[depth] = 0;
    const ordered = isOrderedFormat(level.numFmt);
    const authored = markerFor(level, knownLevels, ilvl, counters);
    const { marker, markerRun } = symbolMarker(authored, markerRunFor(ctx, level, paragraph, authored));
    const item: DocxListItemNode = {
      level: ilvl,
      ordered,
      marker,
      ...(markerRun === undefined ? {} : { markerRun }),
      ...(level.suffix === undefined || level.suffix === "tab" ? {} : { markerSuffix: level.suffix }),
      paragraph,
    };
    // Numbering indents apply only when the paragraph sets none of its own:
    // direct `w:ind` always wins over the level default.
    if (item.paragraph.leftIndent === undefined && level.indentLeft !== undefined) {
      item.paragraph.leftIndent = level.indentLeft;
    }
    if (item.paragraph.firstLineIndent === undefined && level.indentHanging !== undefined) {
      item.paragraph.firstLineIndent = -level.indentHanging;
    }
    open.items.push(item);
  }
  close();
  return out;
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

/**
 * A list marker's own formatting: the paragraph's run defaults with the numbering
 * level's `w:rPr` on top, resolved like any run so colour, weight, size and
 * family come out the same way. Absent when the level authors no formatting.
 */
function markerRunFor(ctx: ParseContext, level: NumberingLevel, paragraph: DocxParagraphNode, marker: string): TextRunModel | undefined {
  if (level.rPr === undefined) return undefined;
  const props = baseRunProps(ctx, paragraph.styleId, undefined);
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

/** A cell's parse result plus whether Word's own saved layout put a page break right before its first paragraph. */
interface TableCellParse {
  cell: DocxTableCellNode | undefined;
  leadingPageBreak: boolean;
}

function parseTableCell(cell: XmlElement, ctx: ParseContext, origins: Array<DocxTableCellNode | undefined>, column: number): TableCellParse {
  throwIfAborted(ctx.signal);
  const tcPr = firstChild(cell, "tcPr");
  const width = tcPr === undefined ? undefined : firstChild(tcPr, "tcW");
  const widthValue = width === undefined ? undefined : numberAttr(width, "w", "w:tcW");
  const span = tcPr === undefined ? undefined : firstChild(tcPr, "gridSpan");
  const spanValue = span === undefined ? undefined : numberAttr(span, "val", "w:gridSpan");
  const columnSpan = spanValue === undefined || spanValue < 1 ? 1 : Math.floor(spanValue);
  const merge = tcPr === undefined ? undefined : firstChild(tcPr, "vMerge");
  const mergeValue = merge === undefined ? undefined : (attr(merge, "val") ?? "continue");
  if (mergeValue === "continue") {
    // Vertically merged continuation: the cell above grows, this grid
    // position emits nothing. A continuation with no origin (malformed
    // grid) falls through as a normal cell rather than vanishing.
    const origin = origins[column];
    if (origin !== undefined) {
      origin.rowSpan = (origin.rowSpan ?? 1) + 1;
      return { cell: undefined, leadingPageBreak: false };
    }
  }
  const vAlign = tcPr === undefined ? undefined : firstChild(tcPr, "vAlign");
  const vAlignValue = vAlign === undefined ? undefined : attr(vAlign, "val");
  const shading = tcPr === undefined ? undefined : firstChild(tcPr, "shd");
  const fill = shading === undefined ? undefined : hexColor(attr(shading, "fill"));
  const cellMargins = readMargins(tcPr === undefined ? undefined : firstChild(tcPr, "tcMar"));
  const cellBorders = readEdges(tcPr === undefined ? undefined : firstChild(tcPr, "tcBorders"), ["top", "right", "bottom", "left"] as const);
  const drafts: BlockDraft[] = [];
  // Tables break between rows, never inside a cell, so a rendered-page-break
  // hint in a cell's own content has nowhere to go and is dropped — except
  // one that leads the cell's very first paragraph, before anything else:
  // that is Word's own signal that the row itself started a fresh page
  // (`DocxTableRowNode.pageBreakBefore`, set by the caller for column 0).
  let leadingPageBreak = false;
  for (const child of cell.children) {
    throwIfAborted(ctx.signal);
    if (child.local === "tcPr") continue;
    if (child.local === "p") {
      const parsed = parseParagraph(child, ctx);
      for (const [index, node] of parsed.nodes.entries()) {
        if (node.type === "pageBreak" && node.strength === "preferred") {
          if (drafts.length === 0 && index === 0) leadingPageBreak = true;
          continue;
        }
        drafts.push({ node });
      }
      if (parsed.numId !== undefined) {
        const last = drafts[drafts.length - 1];
        if (last !== undefined && "node" in last && last.node.type === "paragraph") {
          drafts[drafts.length - 1] = { listItem: { numId: parsed.numId, ilvl: parsed.ilvl ?? 0, paragraph: last.node } };
        }
      }
      continue;
    }
    if (child.local === "tbl") {
      const table = parseTable(child, ctx);
      if (table !== undefined) drafts.push({ node: table });
    }
  }
  const cellNode: DocxTableCellNode = {
    nodes: groupListItems(drafts, ctx),
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
  return { cell: cellNode, leadingPageBreak };
}

function parseTable(table: XmlElement, ctx: ParseContext): FlowNode | undefined {
  throwIfAborted(ctx.signal);
  const tblPr = firstChild(table, "tblPr");
  // `w:bidiVisual` (ECMA-376 17.4.1) lays the logical columns out right to
  // left. Word keeps writing `w:left`/`w:right` in such a table: its borders
  // mean the start/end sides, its cell margins stay physical (probe
  // docx-bidi-visual: a `w:left` border paints on the right, a `w:left`
  // margin stays on the left; DECISION_LOG §43). DocxTableNode documents both.
  const bidi = tblPr === undefined ? undefined : firstChild(tblPr, "bidiVisual");
  const direction: TextDirection = onOffElement(bidi) === true ? "rtl" : "ltr";
  const grid = firstChild(table, "tblGrid");
  const columns = grid === undefined
    ? []
    : childElements(grid, "gridCol").map((column) => twipsToInches(numberAttr(column, "w", "w:gridCol") ?? 0));
  const rows: DocxTableRowNode[] = [];
  const origins: Array<DocxTableCellNode | undefined> = [];
  for (const row of childElements(table, "tr")) {
    throwIfAborted(ctx.signal);
    const height = firstChild(row, "trHeight");
    const heightValue = height === undefined ? undefined : numberAttr(height, "val", "w:trHeight");
    const cells: DocxTableCellNode[] = [];
    let column = 0;
    // Word's own saved break leading the row's first cell means the whole
    // row started a fresh page: any cell can carry the marker (all of a
    // row's cells share its page), but only the first is checked — the
    // common, unambiguous case parsed so far.
    let pageBreakBefore = false;
    for (const cell of childElements(row, "tc")) {
      throwIfAborted(ctx.signal);
      const { cell: parsed, leadingPageBreak } = parseTableCell(cell, ctx, origins, column);
      if (column === 0 && leadingPageBreak) pageBreakBefore = true;
      const span = parsed?.columnSpan ?? 1;
      if (parsed !== undefined) cells.push(parsed);
      column += span;
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
  const borders = readEdges(tblPr === undefined ? undefined : firstChild(tblPr, "tblBorders"), ["top", "right", "bottom", "left", "insideH", "insideV"] as const) as DocxTableBorders | undefined;
  const cellMargins = { ...NORMAL_TABLE_MARGINS, ...readMargins(tblPr === undefined ? undefined : firstChild(tblPr, "tblCellMar")) };
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

function parseBlocks(parent: XmlElement, ctx: ParseContext): BlockDraft[] {
  throwIfAborted(ctx.signal);
  const drafts: BlockDraft[] = [];
  for (const child of parent.children) {
    throwIfAborted(ctx.signal);
    if (child.local === "p") {
      pushParagraphDrafts(drafts, parseParagraph(child, ctx));
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
      continue;
    }
    if (child.local === "sectPr" || child.local === "bookmarkStart" || child.local === "bookmarkEnd") continue;
  }
  return drafts;
}

function pushParagraphDrafts(drafts: BlockDraft[], parsed: ParagraphDraft): void {
  let listed = false;
  for (const node of parsed.nodes) {
    if (!listed && parsed.numId !== undefined && node.type === "paragraph") {
      drafts.push({ listItem: { numId: parsed.numId, ilvl: parsed.ilvl ?? 0, paragraph: node } });
      listed = true;
    } else {
      drafts.push({ node });
    }
  }
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
  const ctx = partContext(index, notes, new Map(index.relTargets), mediaContentType, signal);
  const drafts = await parseBodySections(pkg, index.documentPart, ctx);
  const sections: DocumentSection[] = [];
  // Sections usually inherit the same header/footer parts: parse each once
  // and share the (read-only) node list between the sections using it.
  const headerFooters = new Map<string, FlowNode[]>();
  const headerFooter = async (part: string): Promise<FlowNode[]> => {
    let nodes = headerFooters.get(part);
    if (nodes === undefined) {
      nodes = await parseHeaderFooterPart(pkg, index, notes, part, mediaContentType, signal);
      headerFooters.set(part, nodes);
    }
    return nodes;
  };
  for (const draft of drafts) {
    throwIfAborted(signal);
    const root = groupListItems(draft.drafts, ctx);
    const section: DocumentSection = { properties: draft.properties, root };
    // Headers/footers render from the section's own references: first
    // referenced part wins, the rest are first/even-page alternates.
    const headerPart = draft.properties.headerRefs[0];
    if (headerPart !== undefined) {
      const header = await headerFooter(headerPart);
      if (header.length > 0) section.header = header;
    }
    const footerPart = draft.properties.footerRefs[0];
    if (footerPart !== undefined) {
      const footer = await headerFooter(footerPart);
      if (footer.length > 0) section.footer = footer;
    }
    const firstHeading = root.find((node): node is DocxHeadingNode => node.type === "heading");
    if (firstHeading !== undefined) section.title = firstHeading.text;
    sections.push(section);
  }
  const last = sections[sections.length - 1]!;
  for (const key of ctx.footnoteOrder) {
    throwIfAborted(signal);
    const note = ctx.notes.get(key);
    if (note === undefined) continue;
    const noteCtx = partContext(index, notes, new Map(index.relTargets), mediaContentType, signal);
    noteCtx.footnoteLabels = ctx.footnoteLabels;
    noteCtx.footnoteOrder = ctx.footnoteOrder;
    const nodes = groupListItems(parseBlocks(note, noteCtx), noteCtx);
    last.root.push({
      type: "footnote",
      id: key,
      label: ctx.footnoteLabels.get(key) ?? "?",
      nodes,
      direction: footnoteDirection(nodes),
    });
  }
  throwIfAborted(signal);
  return sections;
}

async function parseHeaderFooterPart(
  pkg: OpcPackage,
  index: DocxIndex,
  notes: Map<string, XmlElement>,
  part: string,
  mediaContentType: (part: string) => string,
  signal?: AbortSignal,
): Promise<FlowNode[]> {
  const partRoot = await pkg.readXmlIfPresent(part, signal);
  throwIfAborted(signal);
  if (partRoot === undefined) return [];
  const partRels = await pkg.relationships(part, signal);
  throwIfAborted(signal);
  const targets = new Map<string, string>();
  for (const relationship of partRels) {
    if (relationship.targetMode !== "Internal") continue;
    targets.set(relationship.id, relationship.target);
  }
  const partCtx = partContext(index, notes, targets, mediaContentType, signal);
  return groupListItems(parseBlocks(partRoot, partCtx), partCtx);
}

/** Word default page geometry (Letter) used when `w:sectPr` omits a value. */
const DEFAULT_PAGE_TWIPS = { width: 12240, height: 15840 };
const DEFAULT_MARGINS_TWIPS = { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 };

function sectionPropertiesFrom(sectPr: XmlElement | undefined, relTargets: Map<string, string>): SectionProperties {
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
    headerRefs: headerFooterParts(sectPr, "headerReference", relTargets),
    footerRefs: headerFooterParts(sectPr, "footerReference", relTargets),
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
          pushParagraphDrafts(current, parseParagraph(child, ctx));
          const pPr = firstChild(child, "pPr");
          const embedded = pPr === undefined ? undefined : firstChild(pPr, "sectPr");
          if (embedded !== undefined) {
            sections.push({ drafts: current, properties: sectionPropertiesFrom(embedded, ctx.relTargets) });
            current = [];
          }
          return;
        }
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
          trailing = sectionPropertiesFrom(child, ctx.relTargets);
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
  sections.push({ drafts: current, properties: trailing ?? sectionPropertiesFrom(undefined, ctx.relTargets) });
  return sections;
}
