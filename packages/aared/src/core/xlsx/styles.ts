/**
 * SpreadsheetML styles reader: the exact gap SheetJS Community Edition
 * leaves (`cell.s` arrives empty there). Resolves fills, fonts, borders,
 * number formats, and alignment from `xl/styles.xml` into plain,
 * structured-cloneable cell styles, plus differential (conditional-format)
 * styles from `dxfs`.
 *
 * Color notes: `rgb` is author intent; `theme` indexes the DrawingML theme
 * palette (shared with PPTX via `core/theme.ts`); `indexed` falls back to
 * the file's `indexedColors` or the legacy 56-color table. `tint` blends in
 * sRGB toward white (positive) or black (negative) — a documented preview
 * approximation, not Office's exact transform.
 */

import { attr, childElements, firstChild } from "../xml.js";
import type { XmlElement } from "../xml.js";
import type { OpcPackage } from "../opc.js";
import { mixColors, parseHexDigits, toHexColor } from "../color.js";
import { readDrawingTheme } from "../theme.js";
import type { DrawingTheme } from "../theme.js";
import { damagedDocument, throwIfAborted } from "../errors.js";
import { formatCodeForId } from "./numberfmt.js";
import type { XlsxCellStyleModel } from "../model.js";

/** Theme color order for SpreadsheetML `theme="n"` indexes. */
const THEME_ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

/** Legacy indexed palette (ECMA-376 system colors 0-63) for `indexed="n"`. */
const LEGACY_INDEXED_COLORS = [
  "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
  "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
  "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080",
  "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF",
  "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF",
  "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",
  "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696",
  "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
];

/** sRGB lerp toward white (positive tint) or black (negative tint). */
function applyTint(color: string, tint: number): string {
  if (tint === 0) return color;
  return mixColors(color, tint > 0 ? "#FFFFFF" : "#000000", Math.abs(tint));
}

function numberAttribute(element: XmlElement, name: string): number | undefined {
  const value = attr(element, name);
  if (value === undefined) return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function onOff(element: XmlElement, name: string): boolean | undefined {
  const value = attr(element, name);
  if (value === undefined) return undefined;
  if (value === "1" || value === "true") return true;
  if (value === "0" || value === "false") return false;
  return undefined;
}

/** Resolve a SpreadsheetML `<color>`/`fgColor`/`bgColor` element to CSS. */
export function resolveStyleColor(
  element: XmlElement | undefined,
  theme: DrawingTheme,
  indexedColors: string[] | undefined,
): string | undefined {
  if (element === undefined) return undefined;
  const rgb = attr(element, "rgb");
  if (rgb !== undefined) {
    const channels = parseHexDigits(rgb.length === 8 ? rgb.slice(2) : rgb);
    if (channels === undefined) return undefined;
    if (rgb.length === 8 && !rgb.toUpperCase().startsWith("FF")) {
      const alpha = Number.parseInt(rgb.slice(0, 2), 16) / 255;
      const [red, green, blue] = channels;
      return `rgba(${red}, ${green}, ${blue}, ${Math.round(alpha * 100) / 100})`;
    }
    return toHexColor(channels);
  }
  const themeIndex = numberAttribute(element, "theme");
  if (themeIndex !== undefined) {
    const scheme = THEME_ORDER[themeIndex];
    const base = scheme === undefined ? undefined : theme.colors[scheme];
    if (base === undefined) return undefined;
    return applyTint(base, numberAttribute(element, "tint") ?? 0);
  }
  const indexed = numberAttribute(element, "indexed");
  if (indexed !== undefined) {
    const override = indexedColors?.[indexed];
    if (override !== undefined) return override;
    return LEGACY_INDEXED_COLORS[indexed];
  }
  return undefined;
}

interface ResolvedFont {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  size?: number;
  color?: string;
  name?: string;
}

function resolveFont(font: XmlElement, theme: DrawingTheme, indexedColors: string[] | undefined): ResolvedFont {
  const out: ResolvedFont = {};
  if (firstChild(font, "b") !== undefined && onOff(firstChild(font, "b")!, "val") !== false) out.bold = true;
  if (firstChild(font, "i") !== undefined && onOff(firstChild(font, "i")!, "val") !== false) out.italic = true;
  const underline = firstChild(font, "u");
  if (underline !== undefined && attr(underline, "val") !== "none") out.underline = true;
  const size = firstChild(font, "sz");
  const points = size === undefined ? undefined : numberAttribute(size, "val");
  if (points !== undefined) out.size = points;
  const color = resolveStyleColor(firstChild(font, "color"), theme, indexedColors);
  if (color !== undefined) out.color = color;
  const name = firstChild(font, "name") ?? firstChild(font, "rFont");
  const family = name === undefined ? undefined : attr(name, "val");
  if (family !== undefined && family !== "") out.name = family;
  return out;
}

function resolveFill(fill: XmlElement, theme: DrawingTheme, indexedColors: string[] | undefined): string | undefined {
  const pattern = firstChild(fill, "patternFill");
  if (pattern === undefined) return undefined;
  const type = attr(pattern, "patternType");
  if (type === "none" || type === undefined) {
    // Differential fills sometimes carry only a `bgColor` with no pattern
    // type; honor it as a solid intent rather than dropping the rule.
    const loose = resolveStyleColor(firstChild(pattern, "bgColor"), theme, indexedColors);
    return type === "none" ? undefined : loose;
  }
  if (type !== "solid") return undefined;
  return resolveStyleColor(firstChild(pattern, "fgColor"), theme, indexedColors)
    ?? resolveStyleColor(firstChild(pattern, "bgColor"), theme, indexedColors);
}

export interface ResolvedBorder {
  left?: { style?: string; color?: string };
  right?: { style?: string; color?: string };
  top?: { style?: string; color?: string };
  bottom?: { style?: string; color?: string };
}

function resolveBorderSide(side: XmlElement | undefined, theme: DrawingTheme, indexedColors: string[] | undefined): ResolvedBorder["left"] {
  if (side === undefined) return undefined;
  const style = attr(side, "style");
  const color = resolveStyleColor(firstChild(side, "color"), theme, indexedColors);
  if ((style === undefined || style === "none") && color === undefined) return undefined;
  return {
    ...(style === undefined || style === "none" ? {} : { style }),
    ...(color === undefined ? {} : { color }),
  };
}

function resolveBorder(border: XmlElement, theme: DrawingTheme, indexedColors: string[] | undefined): ResolvedBorder | undefined {
  const left = resolveBorderSide(firstChild(border, "left"), theme, indexedColors);
  const right = resolveBorderSide(firstChild(border, "right"), theme, indexedColors);
  const top = resolveBorderSide(firstChild(border, "top"), theme, indexedColors);
  const bottom = resolveBorderSide(firstChild(border, "bottom"), theme, indexedColors);
  const out: ResolvedBorder = {
    ...(left === undefined ? {} : { left }),
    ...(right === undefined ? {} : { right }),
    ...(top === undefined ? {} : { top }),
    ...(bottom === undefined ? {} : { bottom }),
  };
  return Object.keys(out).length === 0 ? undefined : out;
}

interface ResolvedAlignment {
  horizontal?: XlsxCellStyleModel["align"];
  vertical?: XlsxCellStyleModel["verticalAlign"];
  wrapText?: boolean;
  indent?: number;
  readingOrder?: 0 | 1 | 2;
}

function resolveAlignment(alignment: XmlElement | undefined): ResolvedAlignment {
  if (alignment === undefined) return {};
  const out: ResolvedAlignment = {};
  switch (attr(alignment, "horizontal")) {
    case "left": out.horizontal = "left"; break;
    case "right": out.horizontal = "right"; break;
    case "center": case "centerContinuous": out.horizontal = "center"; break;
    case "justify": case "distributed": out.horizontal = "justify"; break;
    case "fill": out.horizontal = "left"; break;
    case undefined: case "general": break;
    default: break;
  }
  switch (attr(alignment, "vertical")) {
    case "top": out.vertical = "top"; break;
    case "center": out.vertical = "middle"; break;
    case "bottom": out.vertical = "bottom"; break;
    case undefined: case "justify": case "distributed": break;
    default: break;
  }
  if (onOff(alignment, "wrapText") === true) out.wrapText = true;
  const indent = numberAttribute(alignment, "indent");
  if (indent !== undefined && indent > 0) out.indent = indent;
  const readingOrder = numberAttribute(alignment, "readingOrder");
  if (readingOrder === 1 || readingOrder === 2) out.readingOrder = readingOrder;
  return out;
}

export interface DifferentialStyle {
  font?: { bold?: boolean; color?: string; italic?: boolean };
  fill?: string;
}

/** Shared by every unstyled cell; never mutated. */
const EMPTY_STYLE: XlsxCellStyleModel = Object.freeze({});

/** Resolved workbook styles: direct cell formats plus conditional differentials. */
export interface XlsxStyles {
  theme: DrawingTheme;
  /** Number-format code for a numFmtId (custom records win over builtins). */
  numberFormatCode(id: number): string;
  /**
   * Resolved direct style for a cellXfs index; unknown indexes yield {}.
   * The object is shared by every cell using that index (one allocation per
   * style, not per cell; structured clone keeps the sharing), so treat it
   * as read-only.
   */
  cellStyle(xfId: number): XlsxCellStyleModel;
  /** Resolved differential style for a dxfs index; unknown indexes yield {}. */
  differential(dxfId: number): DifferentialStyle;
  count: number;
}

function differentialFor(dxf: XmlElement, theme: DrawingTheme, indexedColors: string[] | undefined): DifferentialStyle {
  const out: DifferentialStyle = {};
  const font = firstChild(dxf, "font");
  if (font !== undefined) {
    const resolved = resolveFont(font, theme, indexedColors);
    if (resolved.bold !== undefined || resolved.color !== undefined || resolved.italic !== undefined) {
      out.font = {
        ...(resolved.bold === undefined ? {} : { bold: resolved.bold }),
        ...(resolved.color === undefined ? {} : { color: resolved.color }),
        ...(resolved.italic === undefined ? {} : { italic: resolved.italic }),
      };
    }
  }
  const fill = firstChild(dxf, "fill");
  if (fill !== undefined) {
    const color = resolveFill(fill, theme, indexedColors);
    if (color !== undefined) out.fill = color;
  }
  return out;
}

function cellStyleFor(
  xf: XmlElement,
  fonts: ResolvedFont[],
  fills: Array<string | undefined>,
  borders: Array<ResolvedBorder | undefined>,
  numberFormats: ReadonlyMap<number, string>,
): XlsxCellStyleModel {
  const out: XlsxCellStyleModel = {};
  const fontId = numberAttribute(xf, "fontId") ?? 0;
  const fillId = numberAttribute(xf, "fillId") ?? 0;
  const borderId = numberAttribute(xf, "borderId") ?? 0;
  const numFmtId = numberAttribute(xf, "numFmtId") ?? 0;
  const font = fonts[fontId];
  if (font?.bold !== undefined) out.bold = font.bold;
  if (font?.italic !== undefined) out.italic = font.italic;
  if (font?.underline !== undefined) out.underline = font.underline;
  if (font?.size !== undefined) out.fontSizePt = font.size;
  if (font?.color !== undefined) out.color = font.color;
  if (font?.name !== undefined) out.fontFamily = font.name;
  const fill = fills[fillId];
  if (fill !== undefined) out.fill = fill;
  const border = borders[borderId];
  if (border !== undefined) {
    out.border = {
      ...(border.left === undefined ? {} : { left: border.left }),
      ...(border.right === undefined ? {} : { right: border.right }),
      ...(border.top === undefined ? {} : { top: border.top }),
      ...(border.bottom === undefined ? {} : { bottom: border.bottom }),
    };
  }
  const code = formatCodeForId(numFmtId, numberFormats);
  if (code !== "General") out.numberFormat = code;
  const alignment = resolveAlignment(firstChild(xf, "alignment"));
  if (alignment.horizontal !== undefined) out.align = alignment.horizontal;
  if (alignment.vertical !== undefined) out.verticalAlign = alignment.vertical;
  if (alignment.wrapText !== undefined) out.wrapText = alignment.wrapText;
  if (alignment.indent !== undefined) out.indent = alignment.indent;
  if (alignment.readingOrder !== undefined) out.readingOrder = alignment.readingOrder;
  return out;
}

/** Read `xl/styles.xml` (via the workbook relationship, else the conventional part). */
export async function readXlsxStyles(
  pkg: OpcPackage,
  workbookPart: string,
  signal?: AbortSignal,
  /** The workbook's already-parsed theme; read from the package when omitted. */
  knownTheme?: DrawingTheme,
): Promise<XlsxStyles> {
  throwIfAborted(signal);
  const theme = knownTheme ?? await readDrawingTheme(pkg, signal);
  throwIfAborted(signal);
  const relationships = await pkg.relationships(workbookPart, signal);
  throwIfAborted(signal);
  const stylesPart = relationships.find((relationship) => relationship.type.endsWith("/styles"))?.target
    ?? (pkg.hasPart("xl/styles.xml") ? "xl/styles.xml" : undefined);
  const empty: XlsxStyles = {
    theme,
    numberFormatCode: (id) => formatCodeForId(id, new Map()),
    cellStyle: () => EMPTY_STYLE,
    differential: () => ({}),
    count: 0,
  };
  const root = stylesPart === undefined ? undefined : await pkg.readXmlIfPresent(stylesPart, signal);
  if (root === undefined) return empty;
  if (root.local !== "styleSheet") throw damagedDocument("xl/styles.xml root must be <styleSheet>");

  const numberFormats = new Map<number, string>();
  const numFmts = firstChild(root, "numFmts");
  if (numFmts !== undefined) {
    for (const numFmt of childElements(numFmts, "numFmt")) {
      throwIfAborted(signal);
      const id = numberAttribute(numFmt, "numFmtId");
      const code = attr(numFmt, "formatCode");
      if (id !== undefined && code !== undefined) numberFormats.set(id, code);
    }
  }

  let indexedColors: string[] | undefined;
  const colors = firstChild(root, "colors");
  const indexed = colors === undefined ? undefined : firstChild(colors, "indexedColors");
  if (indexed !== undefined) {
    // An unreadable entry keeps its slot (indexes are positional) as black.
    indexedColors = childElements(indexed, "rgbColor").map((entry) => {
      const rgb = attr(entry, "rgb") ?? "";
      return toHexColor(parseHexDigits(rgb.length === 8 ? rgb.slice(2) : rgb) ?? [0, 0, 0]);
    });
  }

  const fontsRoot = firstChild(root, "fonts");
  const fonts = fontsRoot === undefined ? [] : childElements(fontsRoot, "font").map((font) => resolveFont(font, theme, indexedColors));
  const fillsRoot = firstChild(root, "fills");
  const fills = fillsRoot === undefined ? [] : childElements(fillsRoot, "fill").map((fill) => resolveFill(fill, theme, indexedColors));
  const bordersRoot = firstChild(root, "borders");
  const borders = bordersRoot === undefined ? [] : childElements(bordersRoot, "border").map((border) => resolveBorder(border, theme, indexedColors));
  const cellXfs = firstChild(root, "cellXfs");
  const xfElements = cellXfs === undefined ? [] : childElements(cellXfs, "xf");
  const resolved = xfElements.map((xf) => cellStyleFor(xf, fonts, fills, borders, numberFormats));
  const dxfs = firstChild(root, "dxfs");
  const differentials = dxfs === undefined ? [] : childElements(dxfs, "dxf").map((dxf) => differentialFor(dxf, theme, indexedColors));
  throwIfAborted(signal);
  return {
    theme,
    numberFormatCode: (id) => formatCodeForId(id, numberFormats),
    cellStyle: (xfId) => resolved[xfId] ?? EMPTY_STYLE,
    differential: (dxfId) => {
      const found = differentials[dxfId];
      if (found === undefined) return {};
      return {
        ...(found.font === undefined ? {} : { font: { ...found.font } }),
        ...(found.fill === undefined ? {} : { fill: found.fill }),
      };
    },
    count: resolved.length,
  };
}
