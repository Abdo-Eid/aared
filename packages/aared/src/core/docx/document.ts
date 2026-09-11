/**
 * Word `document.xml` reader for the DOCX route: body → resolved
 * `DocumentSection` flow blocks. Headless, plain data — DOM painting belongs to
 * `aared/docx`, never here.
 *
 * Coverage: style + numbering inheritance, bidirectional paragraphs
 * (mixed Arabic/Latin runs in one paragraph), heading outline, tables
 * (logical order, `tblBidiVisual`, spans, shading), lists (bullet + ordered,
 * nested levels), inline images and basic `wp:anchor` floats, footnotes and
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

import { attr, childElements, firstChild, parseXml } from "../xml.js";
import type { XmlElement } from "../xml.js";
import type { OpcPackage } from "../opc.js";
import { readDrawingTheme, usesArabic } from "../theme.js";
import type { DrawingTheme } from "../theme.js";
import { throwIfAborted } from "../zip.js";
import type {
  DocxHeadingNode,
  DocxImageNode,
  DocxListItemNode,
  DocxPageBreakNode,
  DocxParagraphBorder,
  DocxParagraphBorders,
  DocxParagraphLayout,
  DocxParagraphNode,
  DocxTableCellNode,
  DocxTableRowNode,
  DocxUnsupportedNode,
  DocumentSection,
  DocxTextAlign,
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

const DXA_PER_INCH = 1440;
const EMU_PER_INCH = 914400;
const TWENTIETHS_PER_POINT = 20;
const SINGLE_LINE = 240;

function dxaToInches(value: number): number {
  return value / DXA_PER_INCH;
}

function emuToInches(value: number): number {
  return value / EMU_PER_INCH;
}

function numberAttr(element: XmlElement, name: string, context: string): number | undefined {
  const value = attr(element, name);
  if (value === undefined) return undefined;
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) throw new Error(`aared/core: invalid ${name}="${value}" on ${context}`);
  return parsed;
}

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

function hexColor(value: string | undefined): string | undefined {
  if (value === undefined || value === "auto") return undefined;
  if (!/^[0-9a-fA-F]{6}$/.test(value)) return undefined;
  return `#${value.toUpperCase()}`;
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
      default:
        break;
    }
  }
}

/** Resolve a `majorHAnsi`-style theme token through the shared DrawingML theme. */
function themeTokenFamily(token: string | undefined, theme: DrawingTheme): string | undefined {
  if (token === undefined) return undefined;
  const lower = token.toLowerCase();
  const major = lower.startsWith("major");
  const slot = lower.endsWith("bidi") ? "cs" : "latin";
  const family = major ? theme.fonts.major : theme.fonts.minor;
  if (slot === "cs") return family.complexScript === "" ? undefined : family.complexScript;
  return family.latin === "" ? undefined : family.latin;
}

function resolveRunFamily(props: ResolvedRunProps, text: string, theme: DrawingTheme): Pick<TextRunModel, "fontFamily" | "latinTypeface" | "complexScriptTypeface"> {
  const ascii = themeTokenFamily(props.asciiTheme, theme) ?? props.ascii ?? themeTokenFamily(props.hAnsiTheme, theme) ?? props.hAnsi;
  const cs = themeTokenFamily(props.csTheme, theme) ?? props.cs;
  const out: Pick<TextRunModel, "fontFamily" | "latinTypeface" | "complexScriptTypeface"> = {};
  if (ascii !== undefined) out.latinTypeface = ascii;
  if (cs !== undefined) out.complexScriptTypeface = cs;
  // Whitespace follows its paragraph's script slot (same rule as the PPTX
  // parser): an inherited run must not lose its family merely because its
  // own text has no script characters. Callers pass paragraph text for
  // blank runs; see `parseRun`.
  const arabic = usesArabic(text) || props.rtl === true || props.hintCs === true;
  const preferred = arabic
    ? cs ?? ascii ?? theme.fonts.minor.complexScript
    : ascii ?? cs ?? theme.fonts.minor.latin;
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
  lineSpacing?: number;
  lineSpacingPt?: number;
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndent?: number;
  outlineLvl?: number;
  numId?: string;
  ilvl?: number;
  shading?: string;
  pageBreakBefore?: boolean;
  borders?: DocxParagraphBorders;
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
      case "spacing": {
        const before = numberAttr(child, "before", "w:spacing");
        const after = numberAttr(child, "after", "w:spacing");
        const line = numberAttr(child, "line", "w:spacing");
        if (before !== undefined) out.spaceBeforePt = before / TWENTIETHS_PER_POINT;
        if (after !== undefined) out.spaceAfterPt = after / TWENTIETHS_PER_POINT;
        if (line !== undefined) {
          // `w:lineRule="auto"` (or absent) is a 240ths multiplier;
          // `exact`/`atLeast` are twips. The authored value is preserved in
          // the model; the renderer clamps its CSS box (see HANDOFF).
          const rule = attr(child, "lineRule") ?? "auto";
          if (rule === "auto") out.lineSpacing = line / SINGLE_LINE;
          else out.lineSpacingPt = line / TWENTIETHS_PER_POINT;
        }
        break;
      }
      case "ind": {
        const left = numberAttr(child, "left", "w:ind");
        const right = numberAttr(child, "right", "w:ind");
        const firstLine = numberAttr(child, "firstLine", "w:ind");
        const hanging = numberAttr(child, "hanging", "w:ind");
        if (left !== undefined) out.leftIndent = dxaToInches(left);
        if (right !== undefined) out.rightIndent = dxaToInches(right);
        if (firstLine !== undefined) out.firstLineIndent = dxaToInches(firstLine);
        else if (hanging !== undefined) out.firstLineIndent = -dxaToInches(hanging);
        break;
      }
      case "outlineLvl": {
        const value = numberAttr(child, "outlineLvl", "w:outlineLvl");
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
      default:
        break;
    }
  }
}

function alignFor(jc: string | undefined, direction: TextDirection): DocxParagraphLayout["align"] {
  switch (jc) {
    case "left":
    case undefined:
      return undefined;
    case "right":
      return "right";
    case "center":
      return "center";
    case "both":
    case "distribute":
      return "justify";
    case "start":
      return direction === "rtl" ? "right" : undefined;
    case "end":
      return direction === "rtl" ? undefined : "right";
    default:
      return undefined;
  }
}

export interface DocxIndex {
  documentPart: string;
  body: XmlElement;
  styles: StyleData;
  numbering: NumberingData;
  /** Footnote/endnote entries by `f`/`e`-prefixed id; separators excluded. */
  notes: Map<string, XmlElement>;
  theme: DrawingTheme;
  /** Document-part relationships: `r:id` → resolved internal part path. */
  relTargets: Map<string, string>;
  /** External hyperlink targets: `r:id` → URL. */
  hyperlinks: Map<string, string>;
}

interface ParseContext {
  styles: StyleData;
  numbering: NumberingData;
  notes: Map<string, XmlElement>;
  theme: DrawingTheme;
  relTargets: Map<string, string>;
  hyperlinks: Map<string, string>;
  /** MIME lookup for image parts; the worker media route serves the bytes. */
  contentTypeOf: (part: string) => string | undefined;
  footnoteLabels: Map<string, string>;
  footnoteOrder: string[];
  counters: Map<string, number[]>;
  signal?: AbortSignal;
}

function partContext(
  index: DocxIndex,
  relTargets: Map<string, string>,
  contentTypeOf: (part: string) => string | undefined,
  signal?: AbortSignal,
): ParseContext {
  return {
    styles: index.styles,
    numbering: index.numbering,
    notes: index.notes,
    theme: index.theme,
    relTargets,
    hyperlinks: index.hyperlinks,
    contentTypeOf,
    footnoteLabels: new Map(),
    footnoteOrder: [],
    counters: new Map(),
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

async function rootForPart(pkg: OpcPackage, part: string, signal?: AbortSignal): Promise<XmlElement | undefined> {
  throwIfAborted(signal);
  if (!pkg.hasPart(part)) return undefined;
  const root = parseXml(await pkg.readText(part, signal));
  throwIfAborted(signal);
  return root;
}

/**
 * Index a DOCX without inflating drawing media: document/styles/numbering/
 * notes/theme plus the header/footer part order. Media bytes stay in the
 * package until a renderer asks for them through the media route.
 */
export async function indexDocx(pkg: OpcPackage, signal?: AbortSignal): Promise<DocxIndex> {
  throwIfAborted(signal);
  const packageRelationships = await pkg.packageRelationships(signal);
  throwIfAborted(signal);
  const office = packageRelationships.find((relationship) => relationship.type.endsWith("/officeDocument"));
  if (office === undefined || office.targetMode === "External") {
    throw new Error("aared/core: DOCX has no internal officeDocument relationship");
  }
  const contentType = pkg.contentTypeOf(office.target);
  if (
    contentType !== "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    && contentType !== "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
  ) {
    throw new Error("aared/core: package is not a Word document");
  }
  const documentPart = office.target;
  const documentRoot = await rootForPart(pkg, documentPart, signal);
  const body = documentRoot === undefined ? undefined : firstChild(documentRoot, "body");
  if (body === undefined) throw new Error(`aared/core: document "${documentPart}" is missing w:body`);
  throwIfAborted(signal);

  const relationships = await pkg.relationships(documentPart, signal);
  throwIfAborted(signal);
  const relTargets = new Map<string, string>();
  const hyperlinks = new Map<string, string>();
  let stylesPart: string | undefined;
  let numberingPart: string | undefined;
  let footnotesPart: string | undefined;
  let endnotesPart: string | undefined;
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
  }

  const styles = parseStyles(
    stylesPart === undefined ? undefined : await rootForPart(pkg, stylesPart, signal),
    signal,
  );
  const numbering = parseNumbering(
    numberingPart === undefined ? undefined : await rootForPart(pkg, numberingPart, signal),
    signal,
  );
  const notes = new Map<string, XmlElement>();
  for (const [part, prefix, local] of [
    [footnotesPart, "f", "footnote"],
    [endnotesPart, "e", "endnote"],
  ] as const) {
    const parsed = part === undefined ? undefined : await rootForPart(pkg, part, signal);
    throwIfAborted(signal);
    for (const [id, note] of parseNotes(parsed, prefix, local, signal)) notes.set(id, note);
  }
  const theme = await readDrawingTheme(pkg, signal);
  throwIfAborted(signal);
  return {
    documentPart,
    body,
    styles,
    numbering,
    notes,
    theme,
    relTargets,
    hyperlinks,
  };
}

type BlockDraft = { node: FlowNode } | { listItem: { numId: string; ilvl: number; paragraph: DocxParagraphNode } };

interface RunDraft {
  runs: TextRunModel[];
  blocks: Array<DocxImageNode | DocxUnsupportedNode>;
  /**
   * Pagination signals seen anywhere in the run. They split the paragraph at
   * the run boundary (a break mid-run with text on both sides is vanishingly
   * rare; the text stays together and the break lands after the run).
   */
  pageBreak: boolean;
  softBreak: boolean;
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
  let float: DocxImageNode["float"];
  if (anchor !== undefined) {
    const positionH = firstChild(anchor, "positionH");
    const align = positionH === undefined ? undefined : firstChild(positionH, "align");
    const side = align === undefined ? undefined : attr(align, "val") ?? align.text;
    if (side === "left") float = "left";
    else if (side === "right") float = "right";
  }
  return {
    type: "image",
    mediaId: target,
    contentType: ctx.contentTypeOf(target) ?? "application/octet-stream",
    width: emuToInches(widthEmu),
    height: emuToInches(heightEmu),
    ...(alt === undefined ? {} : { alt }),
    ...(float === undefined ? {} : { float }),
  };
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
    contentType: ctx.contentTypeOf(target) ?? "application/octet-stream",
    width: Number(width[1]) / 72,
    height: Number(height[1]) / 72,
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
): void {
  throwIfAborted(ctx.signal);
  let pending = "";
  const flush = (): void => {
    if (pending === "") return;
    const finalized = finalizeRun(ctx, props, paragraphText, pending, href === undefined ? undefined : { hyperlink: href });
    pending = "";
    if (finalized !== undefined) draft.runs.push(finalized);
  };
  for (const child of run.children) {
    throwIfAborted(ctx.signal);
    switch (child.local) {
      case "t":
        pending += child.text;
        break;
      case "tab":
        pending += "\t";
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
        pending += "‑";
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
        if (image !== undefined) draft.blocks.push(image);
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
      case "fldChar":
      case "delText":
      case "rPr":
        break;
      case "lastRenderedPageBreak":
        // Stale Word layout hint: a preferred (not mandatory) split point.
        flush();
        draft.softBreak = true;
        break;
      case "instrText": {
        const instruction = child.text.trim().split(/\s+/)[0]?.toUpperCase();
        // Static flow renders PAGE/NUMPAGES as the single flowing page
        // fiction; layout pagination substitutes the real values per page.
        if (instruction === "PAGE" || instruction === "NUMPAGES") {
          flush();
          const field = finalizeRun(ctx, props, paragraphText, "1", { field: instruction });
          if (field !== undefined) draft.runs.push(field);
        }
        break;
      }
      default:
        break;
    }
  }
  flush();
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
  flushRuns: () => void,
  align: DocxTextAlign | undefined,
): void {
  runs.push(...draft.runs);
  for (const block of draft.blocks) {
    // Inline images inherit the host paragraph's alignment (a centered
    // source paragraph centers its image); floats ignore it, as in Word.
    if (block.type === "image" && block.float === undefined && align !== undefined) block.align = align;
    flushRuns();
    pieces.push(block);
  }
  if (draft.pageBreak || draft.softBreak) {
    flushRuns();
    pieces.push(pageBreakNode(draft.pageBreak ? "mandatory" : "preferred"));
  }
}

function layoutFor(props: ResolvedParaProps, direction: TextDirection): DocxParagraphLayout {
  const layout: DocxParagraphLayout = {};
  if (props.fontSizePt !== undefined) layout.fontSizePt = props.fontSizePt;
  const align = alignFor(props.jc, direction);
  if (align !== undefined) layout.align = align;
  if (props.lineSpacing !== undefined) layout.lineSpacing = props.lineSpacing;
  if (props.lineSpacingPt !== undefined) layout.lineSpacingPt = props.lineSpacingPt;
  if (props.spaceBeforePt !== undefined) layout.spaceBeforePt = props.spaceBeforePt;
  if (props.spaceAfterPt !== undefined) layout.spaceAfterPt = props.spaceAfterPt;
  if (props.leftIndent !== undefined) layout.leftIndent = props.leftIndent;
  if (props.rightIndent !== undefined) layout.rightIndent = props.rightIndent;
  if (props.firstLineIndent !== undefined) layout.firstLineIndent = props.firstLineIndent;
  if (props.shading !== undefined) layout.shading = props.shading;
  if (props.borders !== undefined && Object.keys(props.borders).length > 0) layout.borders = props.borders;
  return layout;
}

function applyLayout(node: DocxParagraphNode | DocxHeadingNode, layout: DocxParagraphLayout): void {
  if (layout.fontSizePt !== undefined) node.fontSizePt = layout.fontSizePt;
  if (layout.align !== undefined) node.align = layout.align;
  if (layout.lineSpacing !== undefined) node.lineSpacing = layout.lineSpacing;
  if (layout.lineSpacingPt !== undefined) node.lineSpacingPt = layout.lineSpacingPt;
  if (layout.spaceBeforePt !== undefined) node.spaceBeforePt = layout.spaceBeforePt;
  if (layout.spaceAfterPt !== undefined) node.spaceAfterPt = layout.spaceAfterPt;
  if (layout.leftIndent !== undefined) node.leftIndent = layout.leftIndent;
  if (layout.rightIndent !== undefined) node.rightIndent = layout.rightIndent;
  if (layout.firstLineIndent !== undefined) node.firstLineIndent = layout.firstLineIndent;
  if (layout.shading !== undefined) node.shading = layout.shading;
  if (layout.borders !== undefined) node.borders = layout.borders;
}

function parseParagraph(paragraph: XmlElement, ctx: ParseContext, defaultStyleId?: string): ParagraphDraft {
  throwIfAborted(ctx.signal);
  const directPPr = firstChild(paragraph, "pPr");
  const directProps: ResolvedParaProps = {};
  readParaProps(directPPr, directProps);
  const styleId = directProps.styleId ?? defaultStyleId;
  // Precedence: docDefaults < style chain (base first) < direct formatting.
  const merged: ResolvedParaProps = {};
  if (ctx.styles.defaultPPr !== undefined) readParaProps(ctx.styles.defaultPPr, merged);
  for (const style of paragraphAncestry(ctx.styles, styleId)) readParaProps(style.pPr, merged);
  for (const [key, value] of Object.entries(directProps)) {
    if (value !== undefined) (merged as Record<string, unknown>)[key] = value;
  }

  const direction: TextDirection = merged.bidi === true ? "rtl" : "ltr";
  const text = paragraphTextOf(paragraph);
  const paraAlign = alignFor(merged.jc, direction);
  const defaultSize = mergedFontSize(ctx, styleId);
  if (defaultSize !== undefined && merged.fontSizePt === undefined) merged.fontSizePt = defaultSize;

  const pieces: FlowNode[] = [];
  let runs: TextRunModel[] = [];
  const flushRuns = (): void => {
    if (runs.length === 0) return;
    const node: DocxParagraphNode = { type: "paragraph", runs, direction };
    applyLayout(node, layoutFor(merged, direction));
    if (styleId !== undefined) node.styleId = styleId;
    pieces.push(node);
    runs = [];
  };

  for (const child of contentChildren(paragraph)) {
    throwIfAborted(ctx.signal);
    if (child.local === "pPr") continue;
    if (child.local === "r") {
      const direct = firstChild(child, "rPr");
      const charStyle = direct === undefined ? undefined : firstChild(direct, "rStyle");
      const charStyleId = charStyle === undefined ? undefined : attr(charStyle, "val");
      const props = baseRunProps(ctx, styleId, charStyleId);
      readRunProps(direct, props, ctx.theme);
      const draft: RunDraft = { runs: [], blocks: [], pageBreak: false, softBreak: false };
      parseRunContent(child, ctx, props, text, undefined, draft);
      emitDraft(draft, runs, pieces, flushRuns, paraAlign);
      continue;
    }
    if (child.local === "hyperlink") {
      const rid = attr(child, "id");
      const anchor = attr(child, "anchor");
      const href = rid !== undefined && ctx.hyperlinks.get(rid) !== undefined
        ? ctx.hyperlinks.get(rid)!
        : anchor !== undefined
          ? `#${anchor}`
          : undefined;
      for (const run of contentChildren(child)) {
        if (run.local !== "r") continue;
        const direct = firstChild(run, "rPr");
        const charStyle = direct === undefined ? undefined : firstChild(direct, "rStyle");
        const charStyleId = charStyle === undefined ? undefined : attr(charStyle, "val");
        const props = baseRunProps(ctx, styleId, charStyleId ?? "Hyperlink");
        readRunProps(direct, props, ctx.theme);
        const draft: RunDraft = { runs: [], blocks: [], pageBreak: false, softBreak: false };
        parseRunContent(run, ctx, props, text, href, draft);
        emitDraft(draft, runs, pieces, flushRuns, paraAlign);
      }
      continue;
    }
    if (child.local === "sdt") {
      const content = firstChild(child, "sdtContent");
      if (content === undefined) continue;
      for (const nested of contentChildren(content)) {
        if (nested.local !== "r") continue;
        const direct = firstChild(nested, "rPr");
        const props = baseRunProps(ctx, styleId, undefined);
        readRunProps(direct, props, ctx.theme);
        const draft: RunDraft = { runs: [], blocks: [], pageBreak: false, softBreak: false };
        parseRunContent(nested, ctx, props, text, undefined, draft);
        emitDraft(draft, runs, pieces, flushRuns, paraAlign);
      }
      continue;
    }
    if (child.local === "fldSimple") {
      const instruction = attr(child, "instr")?.trim().split(/\s+/)[0]?.toUpperCase();
      if (instruction === "PAGE" || instruction === "NUMPAGES") {
        const marker = finalizeRun(ctx, baseRunProps(ctx, styleId, undefined), text, "1", { field: instruction });
        if (marker !== undefined) runs.push(marker);
      }
      continue;
    }
  }
  flushRuns();

  const headingLevel = headingLevelForStyle(styleId, ctx.styles)
    ?? (merged.outlineLvl !== undefined && merged.outlineLvl <= 8 ? merged.outlineLvl + 1 : undefined);
  const nodes: FlowNode[] = pieces.map((piece): FlowNode => {
    if (piece.type !== "paragraph" || headingLevel === undefined) return piece;
    const heading: DocxHeadingNode = {
      type: "heading",
      level: headingLevel,
      text: piece.runs.map((run) => run.text).join(""),
      direction: piece.direction,
      runs: piece.runs,
    };
    applyLayout(heading, paragraphLayoutOf(piece));
    return heading;
  });

  const filtered = nodes.filter((node) => node.type !== "paragraph" || node.runs.length > 0);
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

function paragraphLayoutOf(piece: DocxParagraphNode): DocxParagraphLayout {
  const layout: DocxParagraphLayout = {};
  if (piece.fontSizePt !== undefined) layout.fontSizePt = piece.fontSizePt;
  if (piece.align !== undefined) layout.align = piece.align;
  if (piece.lineSpacing !== undefined) layout.lineSpacing = piece.lineSpacing;
  if (piece.lineSpacingPt !== undefined) layout.lineSpacingPt = piece.lineSpacingPt;
  if (piece.spaceBeforePt !== undefined) layout.spaceBeforePt = piece.spaceBeforePt;
  if (piece.spaceAfterPt !== undefined) layout.spaceAfterPt = piece.spaceAfterPt;
  if (piece.leftIndent !== undefined) layout.leftIndent = piece.leftIndent;
  if (piece.rightIndent !== undefined) layout.rightIndent = piece.rightIndent;
  if (piece.firstLineIndent !== undefined) layout.firstLineIndent = piece.firstLineIndent;
  if (piece.shading !== undefined) layout.shading = piece.shading;
  if (piece.borders !== undefined) layout.borders = piece.borders;
  return layout;
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
    const fallback: NumberingLevel = { numFmt: "bullet", lvlText: "•", start: 1 };
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
    const marker = markerFor(level, knownLevels, ilvl, counters);
    const item: DocxListItemNode = { level: ilvl, ordered, marker, paragraph };
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

function parseTableCell(cell: XmlElement, ctx: ParseContext, origins: Array<DocxTableCellNode | undefined>, column: number): DocxTableCellNode | undefined {
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
      return undefined;
    }
  }
  const vAlign = tcPr === undefined ? undefined : firstChild(tcPr, "vAlign");
  const vAlignValue = vAlign === undefined ? undefined : attr(vAlign, "val");
  const shading = tcPr === undefined ? undefined : firstChild(tcPr, "shd");
  const fill = shading === undefined ? undefined : hexColor(attr(shading, "fill"));
  const drafts: BlockDraft[] = [];
  for (const child of cell.children) {
    throwIfAborted(ctx.signal);
    if (child.local === "tcPr") continue;
    if (child.local === "p") {
      const parsed = parseParagraph(child, ctx);
      for (const node of parsed.nodes) drafts.push({ node });
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
    ...(columnSpan <= 1 ? {} : { columnSpan }),
    ...(widthValue === undefined ? {} : { width: dxaToInches(widthValue) }),
  };
  origins[column] = cellNode;
  for (let index = 1; index < columnSpan; index++) origins[column + index] = undefined;
  return cellNode;
}

function parseTable(table: XmlElement, ctx: ParseContext): FlowNode | undefined {
  throwIfAborted(ctx.signal);
  const tblPr = firstChild(table, "tblPr");
  const bidi = tblPr === undefined ? undefined : firstChild(tblPr, "tblBidiVisual");
  const direction: TextDirection = onOffElement(bidi) === true ? "rtl" : "ltr";
  const grid = firstChild(table, "tblGrid");
  const columns = grid === undefined
    ? []
    : childElements(grid, "gridCol").map((column) => dxaToInches(numberAttr(column, "w", "w:gridCol") ?? 0));
  const rows: DocxTableRowNode[] = [];
  const origins: Array<DocxTableCellNode | undefined> = [];
  for (const row of childElements(table, "tr")) {
    throwIfAborted(ctx.signal);
    const height = firstChild(row, "trHeight");
    const heightValue = height === undefined ? undefined : numberAttr(height, "val", "w:trHeight");
    const cells: DocxTableCellNode[] = [];
    let column = 0;
    for (const cell of childElements(row, "tc")) {
      throwIfAborted(ctx.signal);
      const parsed = parseTableCell(cell, ctx, origins, column);
      const span = parsed?.columnSpan ?? 1;
      if (parsed !== undefined) cells.push(parsed);
      column += span;
    }
    rows.push({
      ...(heightValue === undefined ? {} : { height: dxaToInches(heightValue) }),
      cells,
    });
  }
  if (rows.length === 0) return undefined;
  return { type: "table", direction, columns, rows };
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
 * Inflate and parse the document's sections. `root` order is document order
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
  const contentTypeOf = (part: string): string | undefined => pkg.contentTypeOf(part);
  const ctx = partContext(index, new Map(index.relTargets), contentTypeOf, signal);
  const drafts = parseBodySections(index.body, ctx);
  const sections: DocumentSection[] = [];
  for (const draft of drafts) {
    throwIfAborted(signal);
    const root = groupListItems(draft.drafts, ctx);
    const section: DocumentSection = { properties: draft.properties, root };
    // Headers/footers render from the section's own references: first
    // referenced part wins, the rest are first/even-page alternates.
    const headerPart = draft.properties.headerRefs[0];
    if (headerPart !== undefined) {
      const header = await parseHeaderFooterPart(pkg, index, headerPart, contentTypeOf, signal);
      if (header.length > 0) section.header = header;
    }
    const footerPart = draft.properties.footerRefs[0];
    if (footerPart !== undefined) {
      const footer = await parseHeaderFooterPart(pkg, index, footerPart, contentTypeOf, signal);
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
    const noteCtx = partContext(index, new Map(index.relTargets), contentTypeOf, signal);
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
  part: string,
  contentTypeOf: (part: string) => string | undefined,
  signal?: AbortSignal,
): Promise<FlowNode[]> {
  const partRoot = await rootForPart(pkg, part, signal);
  throwIfAborted(signal);
  if (partRoot === undefined) return [];
  const partRels = await pkg.relationships(part, signal);
  throwIfAborted(signal);
  const targets = new Map<string, string>();
  for (const relationship of partRels) {
    if (relationship.targetMode !== "Internal") continue;
    targets.set(relationship.id, relationship.target);
  }
  const partCtx = partContext(index, targets, contentTypeOf, signal);
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
 * Split the body into per-section block drafts. A paragraph-embedded
 * `w:pPr/w:sectPr` closes the section that precedes it (the paragraph's own
 * content belongs to the closing section); the trailing body `sectPr`
 * governs the last section.
 */
function parseBodySections(body: XmlElement, ctx: ParseContext): SectionDraft[] {
  throwIfAborted(ctx.signal);
  const sections: SectionDraft[] = [];
  let current: BlockDraft[] = [];
  for (const child of body.children) {
    throwIfAborted(ctx.signal);
    if (child.local === "p") {
      pushParagraphDrafts(current, parseParagraph(child, ctx));
      const pPr = firstChild(child, "pPr");
      const embedded = pPr === undefined ? undefined : firstChild(pPr, "sectPr");
      if (embedded !== undefined) {
        sections.push({ drafts: current, properties: sectionPropertiesFrom(embedded, ctx.relTargets) });
        current = [];
      }
      continue;
    }
    if (child.local === "tbl") {
      const table = parseTable(child, ctx);
      if (table !== undefined) current.push({ node: table });
      continue;
    }
    if (child.local === "sdt") {
      const content = firstChild(child, "sdtContent");
      if (content !== undefined) current.push(...parseBlocks(content, ctx));
      continue;
    }
    if (child.local === "sectPr" || child.local === "bookmarkStart" || child.local === "bookmarkEnd") continue;
  }
  let trailing: XmlElement | undefined;
  for (const child of body.children) {
    throwIfAborted(ctx.signal);
    if (child.local === "sectPr") trailing = child;
  }
  sections.push({ drafts: current, properties: sectionPropertiesFrom(trailing, ctx.relTargets) });
  return sections;
}
