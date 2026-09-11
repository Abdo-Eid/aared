/**
 * Focused PPTX reader for the Arabic/table gate fixture.
 *
 * This is deliberately headless and returns only plain data. It reads
 * PresentationML for slide order and DrawingML for text/table content;
 * DOM/SVG painting belongs to `aared/pptx`, never here.
 */

import { attr, childElements, firstChild, parseXml, textContent } from "./xml.js";
import type { XmlElement } from "./xml.js";
import type { OpcPackage } from "./opc.js";
import { readDrawingTheme, resolveDrawingColor, resolveThemeTypeface, usesArabic } from "./theme.js";
import type { DrawingTheme } from "./theme.js";
import { throwIfAborted } from "./zip.js";
import { AaredError } from "./errors.js";
import { parseChartGraphic } from "./chart.js";
import type {
  BoundsModel,
  ChartModel,
  PictureModel,
  ShapeFillModel,
  ShapeModel,
  SlideModel,
  TableCellModel,
  TableModel,
  TextBodyModel,
  TextParagraphModel,
  TextRunModel,
} from "./model.js";
import { fallbackGeometry, geometryForPreset } from "./drawingml-shape.js";

/** OOXML DrawingML coordinates use English Metric Units. */
const EMUS_PER_INCH = 914400;

function emuToInches(emu: number): number {
  return emu / EMUS_PER_INCH;
}

/** Cheap deck index: presentation metadata only, no slide XML inflated yet. */
export interface PptxDeck {
  size: { width: number; height: number };
  slideParts: string[];
  theme: DrawingTheme;
  tableStyles: Map<string, PptxTableStyle>;
  /** Embedded-font parts (`ppt/fonts/*.fntdata`) for future resolver delivery. */
  embeddedFonts: string[];
}

interface PptxTableStyle {
  whole?: string;
  firstRow?: string;
  band1H?: string;
  band2H?: string;
}

function numberAttr(element: XmlElement, name: string, context: string): number | undefined {
  const value = attr(element, name);
  if (value === undefined) return undefined;
  return numberValue(value, name, context);
}

function numberValue(value: string, name: string, context: string): number {
  const number = Number(value);
  if (!Number.isFinite(number)) throw new Error(`aared/core: invalid ${name}="${value}" on ${context}`);
  return number;
}

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

function resolveGradientFill(gradient: XmlElement, theme: DrawingTheme): ShapeFillModel {
  if (firstChild(gradient, "path") !== undefined) return { kind: "unsupported", source: "gradient-path" };
  const linear = firstChild(gradient, "lin");
  if (linear === undefined) return { kind: "unsupported", source: "gradient-linear" };
  const stopList = firstChild(gradient, "gsLst");
  if (stopList === undefined) return { kind: "unsupported", source: "gradient-linear" };
  const stops = [] as NonNullable<Extract<ShapeFillModel, { kind: "linear-gradient" }>["stops"]>;
  for (const stop of childElements(stopList, "gs")) {
    const position = numberAttr(stop, "pos", "a:gs");
    const color = resolveColor(stop, theme);
    if (position === undefined || position < 0 || position > 100000 || color === undefined) {
      return { kind: "unsupported", source: "gradient-linear" };
    }
    stops.push({ position: position / 100000, color });
  }
  if (stops.length === 0) return { kind: "unsupported", source: "gradient-linear" };
  const drawingMlAngle = numberAttr(linear, "ang", "a:lin") ?? 0;
  return {
    kind: "linear-gradient",
    angle: ((drawingMlAngle / 60000) % 360 + 360) % 360,
    scaled: onOff(linear, "scaled") ?? true,
    rotatesWithShape: onOff(gradient, "rotWithShape") ?? true,
    stops,
  };
}

function resolveShapeFill(spPr: XmlElement | undefined, style: XmlElement | undefined, theme: DrawingTheme): ShapeFillModel | undefined {
  if (spPr !== undefined) {
    if (firstChild(spPr, "noFill") !== undefined) return undefined;
    const gradient = firstChild(spPr, "gradFill");
    if (gradient !== undefined) return resolveGradientFill(gradient, theme);
    if (firstChild(spPr, "blipFill") !== undefined) return { kind: "unsupported", source: "image" };
    if (firstChild(spPr, "pattFill") !== undefined) return { kind: "unsupported", source: "pattern" };
    const color = resolveSolidFill(spPr, theme);
    if (color !== undefined) return { kind: "solid", color };
  }
  const color = resolveColor(style === undefined ? undefined : firstChild(style, "fillRef"), theme);
  return color === undefined ? undefined : { kind: "solid", color };
}

async function readTableStyles(pkg: OpcPackage, theme: DrawingTheme, signal?: AbortSignal): Promise<Map<string, PptxTableStyle>> {
  throwIfAborted(signal);
  const part = pkg.partNames().find((name) => pkg.contentTypeOf(name)?.endsWith("tableStyles+xml"));
  if (part === undefined) return new Map();
  const root = await rootForPart(pkg, part, signal);
  throwIfAborted(signal);
  const styles = new Map<string, PptxTableStyle>();
  for (const style of childElements(root, "tblStyle")) {
    throwIfAborted(signal);
    const id = attr(style, "styleId");
    if (id === undefined) continue;
    const fill = (region: string): string | undefined => {
      const regionElement = firstChild(style, region);
      const tcStyle = regionElement === undefined ? undefined : firstChild(regionElement, "tcStyle");
      const fillElement = tcStyle === undefined ? undefined : firstChild(tcStyle, "fill");
      return resolveSolidFill(fillElement, theme);
    };
    const resolved: PptxTableStyle = {};
    const whole = fill("wholeTbl");
    const firstRow = fill("firstRow");
    const band1H = fill("band1H");
    const band2H = fill("band2H");
    if (whole !== undefined) resolved.whole = whole;
    if (firstRow !== undefined) resolved.firstRow = firstRow;
    if (band1H !== undefined) resolved.band1H = band1H;
    if (band2H !== undefined) resolved.band2H = band2H;
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

function parseRun(run: XmlElement, theme: DrawingTheme, paragraphProperties: XmlElement[], paragraphText: string): TextRunModel {
  const properties = runProperties(run, paragraphProperties);
  const sizeElement = properties.find((property) => attr(property, "sz") !== undefined);
  const size = sizeElement === undefined ? undefined : numberAttr(sizeElement, "sz", "a:rPr");
  const solidFill = firstElement(properties, "solidFill");
  const latin = firstElement(properties, "latin");
  const cs = firstElement(properties, "cs");
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
  if (latinTypeface !== undefined) out.latinTypeface = latinTypeface;
  if (complexScriptTypeface !== undefined) out.complexScriptTypeface = complexScriptTypeface;
  const color = resolveColor(solidFill, theme);
  if (color !== undefined) out.color = color;
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

interface TextBodyInheritance {
  /** Text bodies in decreasing precedence after the slide's own body. */
  bodies?: XmlElement[];
  /** The applicable p:txStyles entry from the slide master. */
  masterStyle?: XmlElement;
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
  const autoNumCounters = new Map<number, number>();
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
      const runs = childElements(paragraph, "r");
      const paragraphText = runs.map((run) => textContent(firstChild(run, "t") ?? run)).join("");
      const out: TextParagraphModel = { runs: runs.map((run) => parseRun(run, theme, properties, paragraphText)) };
      const defaultRunProperties = properties
        .map((property) => firstChild(property, "defRPr"))
        .filter((property): property is XmlElement => property !== undefined);
      const defaultSize = defaultRunProperties.find((property) => attr(property, "sz") !== undefined);
      if (defaultSize !== undefined) out.fontSizePt = numberAttr(defaultSize, "sz", "a:defRPr")! / 100;
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
        || firstChild(property, "buChar") !== undefined || firstChild(property, "buAutoNum") !== undefined);
      const hasBuNone = bulletProperties !== undefined && firstChild(bulletProperties, "buNone") !== undefined;
      const autoNumElement = hasBuNone ? undefined : firstElement(properties, "buAutoNum");
      const buChar = bulletProperties === undefined || hasBuNone || autoNumElement !== undefined
        ? undefined
        : firstElement(properties, "buChar");
      if (autoNumElement !== undefined) {
        // Numbered lists resolve to their display string in the model so the
        // DOM renderer paints them through the same bullet path as buChar.
        // Counters are per indentation level; deeper levels reset when a
        // shallower numbered paragraph follows.
        const paragraphLevel = level ?? 0;
        const startAtRaw = attr(autoNumElement, "startAt");
        const startAt = startAtRaw === undefined ? undefined : Number(startAtRaw);
        let count: number;
        if (startAt !== undefined && Number.isFinite(startAt)) {
          count = Math.trunc(startAt);
          autoNumCounters.set(paragraphLevel, count);
        } else {
          count = (autoNumCounters.get(paragraphLevel) ?? 0) + 1;
          autoNumCounters.set(paragraphLevel, count);
        }
        for (const key of [...autoNumCounters.keys()]) {
          if (key > paragraphLevel) autoNumCounters.delete(key);
        }
        const buFontForNum = firstElement(properties, "buFont");
        const numTypeface = buFontForNum === undefined ? undefined : attr(buFontForNum, "typeface");
        out.bullet = {
          char: formatAutoNum(attr(autoNumElement, "type") ?? "arabicPeriod", count),
          ...(numTypeface === undefined ? {} : { typeface: numTypeface }),
        };
      } else {
        const buFont = buChar === undefined ? undefined : firstElement(properties, "buFont");
        const char = buChar === undefined ? undefined : attr(buChar, "char");
        const typeface = buFont === undefined ? undefined : attr(buFont, "typeface");
        if (char !== undefined || typeface !== undefined) {
          out.bullet = { ...(char === undefined ? {} : { char }), ...(typeface === undefined ? {} : { typeface }) };
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
      ...(insets.left === undefined ? {} : { left: emuToInches(numberValue(insets.left, "lIns", "a:bodyPr")) }),
      ...(insets.right === undefined ? {} : { right: emuToInches(numberValue(insets.right, "rIns", "a:bodyPr")) }),
      ...(insets.top === undefined ? {} : { top: emuToInches(numberValue(insets.top, "tIns", "a:bodyPr")) }),
      ...(insets.bottom === undefined ? {} : { bottom: emuToInches(numberValue(insets.bottom, "bIns", "a:bodyPr")) }),
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

async function rootForPart(pkg: OpcPackage, part: string, signal?: AbortSignal): Promise<XmlElement> {
  throwIfAborted(signal);
  const xml = await pkg.readText(part, signal);
  // A completed inflate/read may have been cancelled while it was pending.
  throwIfAborted(signal);
  const root = parseXml(xml);
  throwIfAborted(signal);
  return root;
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
  const layoutPart = layoutRel.target;
  const layoutRoot = await rootForPart(pkg, layoutPart, signal);
  throwIfAborted(signal);
  const layoutShape = matchingPlaceholder(layoutRoot, placeholder);
  const layoutBounds = layoutShape === undefined ? undefined : shapeBounds(layoutShape);
  const layoutRelationships = await pkg.relationships(layoutPart, signal);
  throwIfAborted(signal);
  const masterRel = layoutRelationships.find((relationship) => relationship.type.endsWith("/slideMaster"));
  if (masterRel === undefined) {
    return {
      ...(layoutBounds === undefined ? {} : { bounds: layoutBounds }),
      text: {
        bodies: layoutShape === undefined ? [] : [firstChild(layoutShape, "txBody")].filter((value): value is XmlElement => value !== undefined),
      },
    };
  }
  const masterRoot = await rootForPart(pkg, masterRel.target, signal);
  throwIfAborted(signal);
  const masterShape = matchingPlaceholder(masterRoot, placeholder);
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
    },
  };
}

async function parseShape(pkg: OpcPackage, slidePart: string, shape: XmlElement, theme: DrawingTheme, signal?: AbortSignal): Promise<ShapeModel> {
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
  const out: ShapeModel = {
    kind: "shape",
    ...shapeIdentity(shape),
    ...(bounds === undefined ? {} : { bounds }),
    ...(text === undefined ? {} : { text: parseTextBody(text, theme, inherited?.text) }),
  };
  const size = localGeometrySize(xfrm);
  const geometry = geometryForPreset(preset, size, parseAdjustments(prstGeom));
  out.geometry = geometry ?? {
    ...fallbackGeometry(size, spPr !== undefined && firstChild(spPr, "custGeom") !== undefined ? "custom" : preset === undefined ? "missing" : "unsupported"),
    ...(preset === undefined ? {} : { preset }),
  };
  const fill = resolveShapeFill(spPr, style, theme);
  if (fill !== undefined) out.fill = fill;
  const directLineHasNoFill = line !== undefined && firstChild(line, "noFill") !== undefined;
  const styleLine = style === undefined ? undefined : firstChild(style, "lnRef");
  const strokeColor = directLineHasNoFill ? undefined : resolveSolidFill(line, theme) ?? (line === undefined ? resolveColor(styleLine, theme) : undefined);
  if (strokeColor !== undefined) {
    const strokeWidth = line === undefined ? undefined : numberAttr(line, "w", "a:ln");
    out.stroke = { color: strokeColor, width: emuToInches(strokeWidth ?? 12700) };
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

function guessImageContentType(part: string): string {
  const dot = part.lastIndexOf(".");
  const ext = dot === -1 ? "" : part.slice(dot + 1).toLowerCase();
  switch (ext) {
    case "png":
      return "image/png";
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "tif":
    case "tiff":
      return "image/tiff";
    case "svg":
      return "image/svg+xml";
    case "emf":
      return "image/x-emf";
    case "wmf":
      return "image/x-wmf";
    default:
      return "application/octet-stream";
  }
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
  const contentType = pkg.contentTypeOf(mediaId) ?? guessImageContentType(mediaId);
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
    root = await rootForPart(pkg, notesRel.target, signal);
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
    chartRoot = await rootForPart(pkg, rel.target, signal);
  } catch {
    return out;
  }
  throwIfAborted(signal);
  const chartGraphic = parseChartGraphic(chartRoot, theme);
  throwIfAborted(signal);
  if (chartGraphic !== undefined) out.graphic = chartGraphic;
  return out;
}

function parseCellMargins(tcPr: XmlElement | undefined): TableCellModel["margins"] {
  if (tcPr === undefined) return undefined;
  const left = numberAttr(tcPr, "marL", "a:tcPr");
  const right = numberAttr(tcPr, "marR", "a:tcPr");
  const top = numberAttr(tcPr, "marT", "a:tcPr");
  const bottom = numberAttr(tcPr, "marB", "a:tcPr");
  if (left === undefined && right === undefined && top === undefined && bottom === undefined) return undefined;
  return {
    ...(left === undefined ? {} : { left: emuToInches(left) }),
    ...(right === undefined ? {} : { right: emuToInches(right) }),
    ...(top === undefined ? {} : { top: emuToInches(top) }),
    ...(bottom === undefined ? {} : { bottom: emuToInches(bottom) }),
  };
}

function parseCellBorder(edge: XmlElement | undefined, theme: DrawingTheme): { color: string; width: number } | undefined {
  if (edge === undefined || firstChild(edge, "noFill") !== undefined) return undefined;
  const color = resolveSolidFill(edge, theme);
  if (color === undefined) return undefined;
  return { color, width: emuToInches(numberAttr(edge, "w", "a:ln") ?? 12700) };
}

function parseCellBorders(tcPr: XmlElement | undefined, theme: DrawingTheme): TableCellModel["borders"] {  if (tcPr === undefined) return undefined;
  const left = parseCellBorder(firstChild(tcPr, "lnL"), theme);
  const right = parseCellBorder(firstChild(tcPr, "lnR"), theme);
  const top = parseCellBorder(firstChild(tcPr, "lnT"), theme);
  const bottom = parseCellBorder(firstChild(tcPr, "lnB"), theme);
  if (left === undefined && right === undefined && top === undefined && bottom === undefined) return undefined;
  return {
    ...(left === undefined ? {} : { left }),
    ...(right === undefined ? {} : { right }),
    ...(top === undefined ? {} : { top }),
    ...(bottom === undefined ? {} : { bottom }),
  };
}

export function parseTable(frame: XmlElement, theme: DrawingTheme, styles: Map<string, PptxTableStyle>): TableModel | undefined {
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
  const firstRow = tblPr === undefined ? false : onOff(tblPr, "firstRow") ?? false;
  const bandRow = tblPr === undefined ? false : onOff(tblPr, "bandRow") ?? false;
  const rowElements = childElements(table, "tr");
  const columns = grid === undefined ? [] : childElements(grid, "gridCol").map((column) => emuToInches(numberAttr(column, "w", "a:gridCol") ?? 0));
  const bounds = parseBounds(firstChild(frame, "xfrm"));
  const xfrm = firstChild(frame, "xfrm");
  const directFills: Array<Array<string | undefined>> = [];
  const rows = rowElements.map((row) => {
    const height = numberAttr(row, "h", "a:tr");
    const cells: TableCellModel[] = [];
    const fills: Array<string | undefined> = [];
    for (const cell of childElements(row, "tc")) {
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
      const cellOut: TableCellModel = { text: parseTextBody(body, theme) };
      const anchor = verticalAlign(tcPr === undefined ? undefined : attr(tcPr, "anchor"));
      if (anchor !== undefined) cellOut.verticalAlign = anchor;
      if (gridSpan !== undefined && gridSpan > 1) cellOut.columnSpan = gridSpan;
      if (rowSpan !== undefined && rowSpan > 1) cellOut.rowSpan = rowSpan;
      const margins = parseCellMargins(tcPr);
      if (margins !== undefined) cellOut.margins = margins;
      const borders = parseCellBorders(tcPr, theme);
      if (borders !== undefined) cellOut.borders = borders;
      cells.push(cellOut);
      // Direct cell fills win over table-style bands, matching PowerPoint.
      fills.push(
        resolveSolidFill(tcPr, theme) ?? resolveColor(tcPr === undefined ? undefined : firstChild(tcPr, "fillRef"), theme),
      );
    }
    directFills.push(fills);
    return {
      ...(height === undefined ? {} : { height: emuToInches(height) }),
      cells,
    };
  });
  const cellFills = rows.map((row, rowIndex) => {
    let styleFill = style?.whole;
    if (firstRow && rowIndex === 0) styleFill = style?.firstRow ?? styleFill;
    else if (bandRow && rowIndex > 0) styleFill = (rowIndex - 1) % 2 === 0 ? style?.band1H ?? styleFill : style?.band2H ?? styleFill;
    return row.cells.map((_, cellIndex) => directFills[rowIndex]?.[cellIndex] ?? styleFill);
  });
  const out: TableModel = {
    kind: "table",
    id,
    name: cNvPr === undefined ? "" : attr(cNvPr, "name") ?? "",
    ...(bounds === undefined ? {} : { bounds }),
    direction: tblPr === undefined || onOff(tblPr, "rtl") !== true ? "ltr" : "rtl",
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
  const root = await rootForPart(pkg, presentationPart, signal);
  throwIfAborted(signal);
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
 * Index a PPTX without opening a single slide part. This is the performance
 * boundary: a 200-slide deck should be as cheap to open as a two-slide deck;
 * `parsePptxSlide()` inflates only the unit currently entering the viewport.
 */
export async function indexPptx(pkg: OpcPackage, signal?: AbortSignal): Promise<PptxDeck> {
  throwIfAborted(signal);
  const packageRelationships = await pkg.packageRelationships(signal);
  throwIfAborted(signal);
  const office = packageRelationships.find((relationship) => relationship.type.endsWith("/officeDocument"));
  if (office === undefined || office.targetMode === "External") throw new Error("aared/core: PPTX has no internal officeDocument relationship");
  if (pkg.contentTypeOf(office.target) !== "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml") {
    throw new Error("aared/core: package is not a PPTX presentation");
  }
  const presentation = await rootForPart(pkg, office.target, signal);
  throwIfAborted(signal);
  const sldSz = firstChild(presentation, "sldSz");
  const width = sldSz === undefined ? undefined : numberAttr(sldSz, "cx", "p:sldSz");
  const height = sldSz === undefined ? undefined : numberAttr(sldSz, "cy", "p:sldSz");
  if (width === undefined || height === undefined) throw new Error("aared/core: PPTX is missing p:sldSz");
  const theme = await readDrawingTheme(pkg, signal);
  throwIfAborted(signal);
  const parts = await slideParts(pkg, office.target, signal);
  throwIfAborted(signal);
  const tableStyles = await readTableStyles(pkg, theme, signal);
  throwIfAborted(signal);
  return { size: { width: emuToInches(width), height: emuToInches(height) }, slideParts: parts, theme, tableStyles, embeddedFonts: listEmbeddedFonts(pkg) };
}

/** Inflate and parse exactly one slide previously listed by `indexPptx()`. */
export async function parsePptxSlide(pkg: OpcPackage, deck: PptxDeck, index: number, signal?: AbortSignal): Promise<SlideModel> {
  throwIfAborted(signal);
  const part = deck.slideParts[index];
  if (part === undefined) throw new AaredError("not-found", `aared/core: slide index ${index} is outside this ${deck.slideParts.length}-slide deck`);
  const root = await rootForPart(pkg, part, signal);
  throwIfAborted(signal);
  const cSld = firstChild(root, "cSld");
  const tree = cSld === undefined ? undefined : firstChild(cSld, "spTree");
  if (tree === undefined) throw new Error(`aared/core: slide "${part}" is missing p:spTree`);
  const relationships = await pkg.relationships(part, signal);
  throwIfAborted(signal);
  const elements: SlideModel["elements"] = [];
  const pushFallbackPicture = (container: XmlElement): void => {
    const fallback = firstChild(container, "Fallback");
    const pic = fallback === undefined ? undefined : firstChild(fallback, "pic");
    if (pic === undefined) return;
    const picture = parsePicture(pic, relationships, pkg);
    if (picture !== undefined) elements.push(picture);
  };
  for (const child of tree.children) {
    throwIfAborted(signal);
    if (child.local === "sp") {
      elements.push(await parseShape(pkg, part, child, deck.theme, signal));
      throwIfAborted(signal);
      continue;
    }
    if (child.local === "pic") {
      const picture = parsePicture(child, relationships, pkg);
      if (picture !== undefined) elements.push(picture);
      continue;
    }
    if (child.local === "AlternateContent") {
      // Ink and other modern content surface their raster fallback here.
      pushFallbackPicture(child);
      continue;
    }
    if (child.local === "grpSp") {
      // Grouped shapes keep document order. Group transforms are intentionally
      // not applied yet; nested fallbacks still render at their authored bounds
      // rather than disappearing.
      for (const nested of child.children) {
        throwIfAborted(signal);
        if (nested.local === "pic") {
          const picture = parsePicture(nested, relationships, pkg);
          if (picture !== undefined) elements.push(picture);
        } else if (nested.local === "AlternateContent") {
          pushFallbackPicture(nested);
        }
      }
      continue;
    }
    if (child.local === "graphicFrame") {
      throwIfAborted(signal);
      const table = parseTable(child, deck.theme, deck.tableStyles);
      if (table !== undefined) {
        elements.push(table);
        continue;
      }
      const chart = await parseChartFrame(pkg, child, relationships, deck.theme, signal);
      throwIfAborted(signal);
      if (chart !== undefined) elements.push(chart);
    }
  }
  throwIfAborted(signal);
  const notes = await parseSlideNotes(pkg, part, signal);
  throwIfAborted(signal);
  return { part, size: deck.size, elements, ...(notes === undefined ? {} : { notes }) };
}
