/** Shared DrawingML theme reader for PPTX, XLSX, and DOCX drawing content. */

import { attr, childElements, firstChild, ownTree } from "./xml.js";
import type { XmlElement } from "./xml.js";
import type { OpcPackage } from "./opc.js";
import { clampChannel, linearToSrgb, parseHexColor, srgbToLinear, toHexColor } from "./color.js";
import { throwIfAborted } from "./errors.js";
import { usesArabic } from "./typography.js";

export interface DrawingTheme {
  colors: Record<string, string>;
  /**
   * The theme's background fill styles (`a:bgFillStyleLst`), in order. A
   * `p:bgRef idx="1001+"` picks one and substitutes its own colour for the
   * style's `phClr` placeholder.
   */
  backgroundFills?: XmlElement[];
  /**
   * The theme's line style widths (`a:lnStyleLst/a:ln/@w`), in EMU, in order:
   * what `a:lnRef idx="1|2|3"` draws at. Absent widths are left out, so read
   * with a fallback.
   */
  lineWidths?: Array<number | undefined>;
  fonts: {
    major: { latin: string; complexScript: string; scripts: Record<string, string> };
    minor: { latin: string; complexScript: string; scripts: Record<string, string> };
  };
}

function readColor(container: XmlElement): string | undefined {
  const srgb = firstChild(container, "srgbClr");
  const system = firstChild(container, "sysClr");
  const value = srgb === undefined ? undefined : attr(srgb, "val");
  if (value !== undefined) return `#${value}`;
  const last = system === undefined ? undefined : attr(system, "lastClr");
  return last === undefined ? undefined : `#${last}`;
}

function readFontFamily(font: XmlElement | undefined): { latin: string; complexScript: string; scripts: Record<string, string> } {
  const latin = font === undefined ? undefined : firstChild(font, "latin");
  const cs = font === undefined ? undefined : firstChild(font, "cs");
  const scripts: Record<string, string> = {};
  if (font !== undefined) {
    for (const script of childElements(font, "font")) {
      const name = attr(script, "script");
      const typeface = attr(script, "typeface");
      if (name !== undefined && typeface !== undefined && typeface !== "") scripts[name] = typeface;
    }
  }
  return {
    latin: latin === undefined ? "" : attr(latin, "typeface") ?? "",
    complexScript: cs === undefined ? "" : attr(cs, "typeface") ?? "",
    scripts,
  };
}

/**
 * Read the first declared DrawingML theme from an OOXML package. Word and Excel have one theme;
 * PowerPoint reads each master's own with `readThemePart`.
 */
export async function readDrawingTheme(pkg: OpcPackage, signal?: AbortSignal): Promise<DrawingTheme> {
  throwIfAborted(signal);
  const part = pkg.partNames().find((name) => pkg.contentTypeOf(name) === "application/vnd.openxmlformats-officedocument.theme+xml");
  if (part === undefined) return { colors: {}, fonts: { major: { latin: "", complexScript: "", scripts: {} }, minor: { latin: "", complexScript: "", scripts: {} } } };
  return readThemePart(pkg, part, signal);
}

/** Read one DrawingML theme part by name. */
export async function readThemePart(pkg: OpcPackage, part: string, signal?: AbortSignal): Promise<DrawingTheme> {
  throwIfAborted(signal);
  const root = await pkg.readXml(part, signal);
  const elements = firstChild(root, "themeElements");
  const scheme = elements === undefined ? undefined : firstChild(elements, "clrScheme");
  const fontScheme = elements === undefined ? undefined : firstChild(elements, "fontScheme");
  const colors: Record<string, string> = {};
  if (scheme !== undefined) {
    for (const color of scheme.children) {
      const value = readColor(color);
      if (value !== undefined) colors[color.local] = value;
    }
  }
  const formats = elements === undefined ? undefined : firstChild(elements, "fmtScheme");
  const backgroundStyles = formats === undefined ? undefined : firstChild(formats, "bgFillStyleLst");
  const lineStyles = formats === undefined ? undefined : firstChild(formats, "lnStyleLst");
  const lineWidths = lineStyles === undefined ? [] : lineStyles.children.map((line) => {
    const width = Number(attr(line, "w"));
    return Number.isFinite(width) ? width : undefined;
  });
  return {
    colors,
    ...(lineWidths.some((width) => width !== undefined) ? { lineWidths } : {}),
    // Kept for the session: owned copies, so the theme part's text can be collected.
    ...(backgroundStyles === undefined ? {} : { backgroundFills: backgroundStyles.children.map(ownTree) }),
    fonts: {
      major: readFontFamily(fontScheme === undefined ? undefined : firstChild(fontScheme, "majorFont")),
      minor: readFontFamily(fontScheme === undefined ? undefined : firstChild(fontScheme, "minorFont")),
    },
  };
}

/**
 * A theme font slot as a browser family: `latin` for the Latin slot, `cs` for
 * the complex-script slot. Themes routinely leave `cs` empty (Word's and
 * PowerPoint's defaults do), and Arabic text then takes the theme's `Arab`
 * script font. Ground truth: native Word sets a `w:rtl` run whose theme `cs`
 * is empty in that entry (a probe with `Arab` = Tahoma painted Tahoma), and a
 * literal `w:cs` typeface still wins over it.
 */
export function themeFontFamily(font: DrawingTheme["fonts"]["major"], slot: "latin" | "cs", arabic: boolean): string | undefined {
  if (slot === "latin") return font.latin === "" ? undefined : font.latin;
  if (font.complexScript !== "") return font.complexScript;
  return arabic ? font.scripts["Arab"] : undefined;
}

/**
 * `bg1`/`bg2`/`tx1`/`tx2` are colour-map names that a `p:clrMap` points at
 * theme slots (`lt1`, `dk1`, ...); dark-master decks swap them. PPTX passes
 * the slide's map. Excel and Word have no colour map: their DrawingML (charts)
 * always uses the default map, Office's light one, which `clrMap` undefined
 * gives.
 */
export function themeForColorMap(theme: DrawingTheme, clrMap: XmlElement | undefined): DrawingTheme {
  const slot = (name: string, fallback: string): string | undefined => theme.colors[(clrMap === undefined ? undefined : attr(clrMap, name)) ?? fallback];
  const mapped: Record<string, string> = {};
  for (const [name, fallback] of [["bg1", "lt1"], ["bg2", "lt2"], ["tx1", "dk1"], ["tx2", "dk2"]] as const) {
    const color = slot(name, fallback);
    if (color !== undefined) mapped[name] = color;
  }
  return { ...theme, colors: { ...theme.colors, ...mapped } };
}

/** Resolve a theme token or literal DrawingML typeface to a browser family. */
export function resolveThemeTypeface(typeface: string | undefined, text: string, theme: DrawingTheme): string | undefined {
  if (typeface === undefined || typeface === "") return undefined;
  const match = /^\+m([jn])-(lt|cs)$/.exec(typeface);
  if (match === null) return typeface;
  return themeFontFamily(match[1] === "j" ? theme.fonts.major : theme.fonts.minor, match[2] === "lt" ? "latin" : "cs", usesArabic(text));
}

function rgbToHsl(channels: number[]): [number, number, number] {
  const [red, green, blue] = channels.map((channel) => channel / 255);
  const maximum = Math.max(red!, green!, blue!);
  const minimum = Math.min(red!, green!, blue!);
  const lightness = (maximum + minimum) / 2;
  if (maximum === minimum) return [0, 0, lightness];
  const delta = maximum - minimum;
  const saturation = lightness > 0.5 ? delta / (2 - maximum - minimum) : delta / (maximum + minimum);
  const hue = maximum === red ? ((green! - blue!) / delta + (green! < blue! ? 6 : 0)) / 6
    : maximum === green ? ((blue! - red!) / delta + 2) / 6
      : ((red! - green!) / delta + 4) / 6;
  return [hue, saturation, lightness];
}

function hslToRgb(hue: number, saturation: number, lightness: number): number[] {
  if (saturation === 0) return [lightness * 255, lightness * 255, lightness * 255];
  const chroma = saturation * Math.min(lightness, 1 - lightness);
  const hueChannel = (channel: number): number => {
    const value = (channel + hue * 12) % 12;
    return (lightness - chroma * Math.max(-1, Math.min(value - 3, Math.min(9 - value, 1)))) * 255;
  };
  return [hueChannel(0), hueChannel(8), hueChannel(4)];
}

/**
 * A SpreadsheetML `tint` (-1..1) on a colour: Excel moves its lightness in
 * HLS space, scaling it down for a negative tint (`L * (1 + tint)`) and toward
 * white for a positive one (`L * (1 - tint) + tint`), hue and saturation kept
 * (ECMA-376 18.3.1.15). Not the DrawingML `a:tint` above, which blends in
 * linear light. Excel's "Darker 25%" of accent 5B9BD5 is 2E75B6; blending the
 * channels toward black gave 4474A0.
 */
export function applySpreadsheetTint(color: string, tint: number): string {
  const channels = tint === 0 ? undefined : parseHexColor(color);
  if (channels === undefined) return color;
  const [hue, saturation, lightness] = rgbToHsl(channels);
  const shifted = tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint;
  return toHexColor(hslToRgb(hue, saturation, Math.min(Math.max(shifted, 0), 1)));
}

const WORD_THEME_SLOTS: Readonly<Record<string, string>> = {
  dark1: "dk1", light1: "lt1", dark2: "dk2", light2: "lt2",
  text1: "dk1", background1: "lt1", text2: "dk2", background2: "lt2",
  hyperlink: "hlink", followedHyperlink: "folHlink",
};

/**
 * Word theme ink: shade/tint are hex bytes retaining that fraction of HSL
 * lightness (shade), or of the distance from white (tint). Native Word's
 * TextColor.TintAndShade probe writes BF/99 for -25%/+40% and paints
 * 156082 as 0F4761/45B0E1. Its resulting channels truncate, rather than round.
 * The cached `val` is a fallback, not authority over a present theme slot.
 */
export function resolveWordRunColor(element: XmlElement, theme: DrawingTheme): string | undefined {
  const token = attr(element, "themeColor");
  const source = token === undefined ? undefined : theme.colors[WORD_THEME_SLOTS[token] ?? token];
  const channels = source === undefined ? undefined : parseHexColor(source);
  if (channels === undefined) {
    const cached = attr(element, "val");
    const parsed = cached === undefined ? undefined : parseHexColor(cached);
    return parsed === undefined ? undefined : toHexColor(parsed);
  }
  const byte = (name: string): number | undefined => {
    const raw = attr(element, name);
    return raw !== undefined && /^[0-9a-f]{2}$/i.test(raw) ? Number.parseInt(raw, 16) / 255 : undefined;
  };
  const shade = byte("themeShade"), tint = byte("themeTint");
  if (shade === undefined && tint === undefined) return source;
  const shifted = withLightness(channels, (lightness) => tint !== undefined
    ? lightness * tint + (1 - tint)
    : lightness * shade!);
  return toHexColor(shifted.map((channel) => Math.floor(channel + 1e-9)));
}

/** `channels` with its HSL lightness replaced by `next(lightness)`, kept inside 0..1. */
function withLightness(channels: number[], next: (lightness: number) => number): number[] {
  const [hue, saturation, lightness] = rgbToHsl(channels);
  return hslToRgb(hue, saturation, Math.min(Math.max(next(lightness), 0), 1));
}

/**
 * A colour with DrawingML's `a:lumMod` (a multiplier, 1 = 100%) and then
 * `a:lumOff` (an offset) applied: the shades Office derives for series and
 * gridlines. The same arithmetic `resolveDrawingColor` applies to the modifiers
 * of an authored colour. `undefined` when `hex` is not a colour.
 */
export function shadeColor(hex: string, lumMod: number, lumOff: number): string | undefined {
  const channels = parseHexColor(hex);
  if (channels === undefined) return undefined;
  return toHexColor(withLightness(withLightness(channels, (l) => l * lumMod), (l) => l + lumOff));
}

function tintChannels(channels: number[], sourceRatio: number): number[] {
  return channels.map((channel) => linearToSrgb(srgbToLinear(channel) * sourceRatio + (1 - sourceRatio)));
}

/**
 * ECMA-376 ST_PresetColorVal RGB table (20.1.10.48), not CSS named colors.
 * Token aliases follow the Open XML SDK's PresetColorValues vocabulary.
 */
const PRESET_COLORS: Readonly<Record<string, string>> = {
  aliceBlue: "F0F8FF", antiqueWhite: "FAEBD7", aqua: "00FFFF", aquamarine: "7FFFD4",
  azure: "F0FFFF", beige: "F5F5DC", bisque: "FFE4C4", black: "000000",
  blanchedAlmond: "FFEBCD", blue: "0000FF", blueViolet: "8A2BE2", brown: "A52A2A",
  burlyWood: "DEB887", cadetBlue: "5F9EA0", chartreuse: "7FFF00", chocolate: "D2691E",
  coral: "FF7F50", cornflowerBlue: "6495ED", cornsilk: "FFF8DC", crimson: "DC143C",
  cyan: "00FFFF", deepPink: "FF1493", deepSkyBlue: "00BFFF", dimGray: "696969",
  dkBlue: "00008B", dkCyan: "008B8B", dkGoldenrod: "B8860B", dkGray: "A9A9A9",
  dkGreen: "006400", dkKhaki: "BDB76B", dkMagenta: "8B008B", dkOliveGreen: "556B2F",
  dkOrange: "FF8C00", dkOrchid: "9932CC", dkRed: "8B0000", dkSalmon: "E9967A",
  dkSeaGreen: "8FBC8B", dkSlateBlue: "483D8B", dkSlateGray: "2F4F4F", dkTurquoise: "00CED1",
  dkViolet: "9400D3", dodgerBlue: "1E90FF", firebrick: "B22222", floralWhite: "FFFAF0",
  forestGreen: "228B22", fuchsia: "FF00FF", gainsboro: "DCDCDC", ghostWhite: "F8F8FF",
  gold: "FFD700", goldenrod: "DAA520", gray: "808080", green: "008000",
  greenYellow: "ADFF2F", honeydew: "F0FFF0", hotPink: "FF69B4", indianRed: "CD5C5C",
  indigo: "4B0082", ivory: "FFFFF0", khaki: "F0E68C", lavender: "E6E6FA",
  lavenderBlush: "FFF0F5", lawnGreen: "7CFC00", lemonChiffon: "FFFACD", lime: "00FF00",
  limeGreen: "32CD32", linen: "FAF0E6", ltBlue: "ADD8E6", ltCoral: "F08080",
  ltCyan: "E0FFFF", ltGoldenrodYellow: "FAFA78", ltGray: "D3D3D3", ltGreen: "90EE90",
  ltPink: "FFB6C1", ltSalmon: "FFA07A", ltSeaGreen: "20B2AA", ltSkyBlue: "87CEFA",
  ltSlateGray: "778899", ltSteelBlue: "B0C4DE", ltYellow: "FFFFE0", magenta: "FF00FF",
  maroon: "800000", medAquamarine: "66CDAA", medBlue: "0000CD", medOrchid: "BA55D3",
  medPurple: "9370DB", medSeaGreen: "3CB371", medSlateBlue: "7B68EE", medSpringGreen: "00FA9A",
  medTurquoise: "48D1CC", medVioletRed: "C71585", midnightBlue: "191970", mintCream: "F5FFFA",
  mistyRose: "FFE4E1", moccasin: "FFE4B5", navajoWhite: "FFDEAD", navy: "000080",
  oldLace: "FDF5E6", olive: "808000", oliveDrab: "6B8E23", orange: "FFA500",
  orangeRed: "FF4500", orchid: "DA70D6", paleGoldenrod: "EEE8AA", paleGreen: "98FB98",
  paleTurquoise: "AFEEEE", paleVioletRed: "DB7093", papayaWhip: "FFEFD5", peachPuff: "FFDAB9",
  peru: "CD853F", pink: "FFC0CB", plum: "DDA0DD", powderBlue: "B0E0E6",
  purple: "800080", red: "FF0000", rosyBrown: "BC8F8F", royalBlue: "4169E1",
  saddleBrown: "8B4513", salmon: "FA8072", sandyBrown: "F4A460", seaGreen: "2E8B57",
  seaShell: "FFF5EE", sienna: "A0522D", silver: "C0C0C0", skyBlue: "87CEEB",
  slateBlue: "6A5ACD", slateGray: "708090", snow: "FFFAFA", springGreen: "00FF7F",
  steelBlue: "4682B4", tan: "D2B48C", teal: "008080", thistle: "D8BFD8",
  tomato: "FF6347", turquoise: "40E0D0", violet: "EE82EE", wheat: "F5DEB3",
  white: "FFFFFF", whiteSmoke: "F5F5F5", yellow: "FFFF00", yellowGreen: "9ACD32",
};

function presetColor(name: string | undefined): string | undefined {
  if (name === undefined) return undefined;
  // Office 2010's additional ST_PresetColorVal tokens spell out dark/light/
  // medium and accept Grey variants (Open XML SDK PresetColorValues). These
  // exact aliases still go through the closed ECMA table; CSS-only names do not.
  const canonical = name.replace(/^dark(?=[A-Z])/, "dk").replace(/^light(?=[A-Z])/, "lt")
    .replace(/^medium(?=[A-Z])/, "med").replace(/Grey$/, "Gray").replace(/^grey$/, "gray");
  return Object.hasOwn(PRESET_COLORS, canonical) ? `#${PRESET_COLORS[canonical]}` : undefined;
}

/**
 * Resolve a DrawingML solid-color container (`a:solidFill`, `a:gs`, `a:fillRef`,
 * …) to a browser CSS color. Handles `srgbClr`/`schemeClr`/`sysClr`/`prstClr` plus
 * `tint` (linear-light, measured in PowerPoint's table styles), `shade`,
 * `lumMod`/`lumOff`, and `alpha`. Returns undefined when the source has no
 * usable color. Shared by shape, table, and chart parsing so all three agree.
 */
export function resolveDrawingColor(element: XmlElement | undefined, theme: DrawingTheme): string | undefined {
  if (element === undefined) return undefined;
  const color = firstChild(element, "srgbClr") ?? firstChild(element, "schemeClr") ?? firstChild(element, "sysClr") ?? firstChild(element, "prstClr");
  if (color === undefined) return undefined;
  const raw = attr(color, "val");
  const system = attr(color, "lastClr");
  const source = color.local === "srgbClr" ? raw : color.local === "schemeClr" ? (raw === undefined ? undefined : theme.colors[raw])
    : color.local === "prstClr" ? presetColor(raw) : system;
  const parsed = source === undefined ? undefined : parseHexColor(source);
  if (parsed === undefined) return undefined;
  let channels: number[] = parsed;
  let alpha = 1;
  for (const transform of color.children) {
    const rawValue = attr(transform, "val");
    if (rawValue === undefined) continue;
    const numeric = Number(rawValue);
    // A modifier that is not a number is skipped: one odd value does not cost the colour.
    if (!Number.isFinite(numeric)) continue;
    const ratio = Math.min(Math.max(numeric / 100000, 0), 1);
    switch (transform.local) {
      case "tint":
        // Inspecting table-style colour variants in PowerPoint (the
        // table-style-colors probe) established that Office resolves
        // `a:tint` as a linear-light blend where `val` retains that proportion
        // of the source color, not a gamma-encoded RGB interpolation.
        channels = tintChannels(channels, ratio);
        break;
      case "shade":
        channels = channels.map((channel) => channel * ratio);
        break;
      case "lumMod":
        // A multiplier, not a fraction: 110000 (110%) brightens.
        channels = withLightness(channels, (l) => l * Math.max(numeric / 100000, 0));
        break;
      case "satMod": {
        // Ground truth: PowerPoint's own render of a theme background style
        // (`tint 95000` + `satMod 170000` on a dark scheme colour).
        const [hue, saturation, lightness] = rgbToHsl(channels);
        channels = hslToRgb(hue, Math.min(saturation * Math.max(numeric / 100000, 0), 1), lightness);
        break;
      }
      case "lumOff":
        // Signed: a negative offset darkens.
        channels = withLightness(channels, (l) => l + numeric / 100000);
        break;
      case "alpha":
        alpha *= ratio;
        break;
    }
  }
  return alpha === 1 ? toHexColor(channels) : `rgb(${channels.map(clampChannel).join(" ")} / ${alpha})`;
}
