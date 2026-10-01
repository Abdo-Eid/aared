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

/** Read the first declared DrawingML theme from an OOXML package. */
export async function readDrawingTheme(pkg: OpcPackage, signal?: AbortSignal): Promise<DrawingTheme> {
  throwIfAborted(signal);
  const part = pkg.partNames().find((name) => pkg.contentTypeOf(name) === "application/vnd.openxmlformats-officedocument.theme+xml");
  if (part === undefined) return { colors: {}, fonts: { major: { latin: "", complexScript: "", scripts: {} }, minor: { latin: "", complexScript: "", scripts: {} } } };
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

function tintChannels(channels: number[], sourceRatio: number): number[] {
  return channels.map((channel) => linearToSrgb(srgbToLinear(channel) * sourceRatio + (1 - sourceRatio)));
}

/**
 * Resolve a DrawingML solid-color container (`a:solidFill`, `a:gs`, `a:fillRef`,
 * …) to a browser CSS color. Handles `srgbClr`/`schemeClr`/`sysClr` plus
 * `tint` (linear-light, per tools/probe/table-style-colors), `shade`,
 * `lumMod`/`lumOff`, and `alpha`. Returns undefined when the source has no
 * usable color. Shared by shape, table, and chart parsing so all three agree.
 */
export function resolveDrawingColor(element: XmlElement | undefined, theme: DrawingTheme): string | undefined {
  if (element === undefined) return undefined;
  const color = firstChild(element, "srgbClr") ?? firstChild(element, "schemeClr") ?? firstChild(element, "sysClr");
  if (color === undefined) return undefined;
  const raw = attr(color, "val");
  const system = attr(color, "lastClr");
  const source = color.local === "srgbClr" ? raw : color.local === "schemeClr" ? (raw === undefined ? undefined : theme.colors[raw]) : system;
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
        // Ground truth: tools/probe/table-style-colors. PowerPoint resolves
        // `a:tint` as a linear-light blend where `val` retains that proportion
        // of the source color, not a gamma-encoded RGB interpolation.
        channels = tintChannels(channels, ratio);
        break;
      case "shade":
        channels = channels.map((channel) => channel * ratio);
        break;
      case "lumMod": {
        // A multiplier, not a fraction: 110000 (110%) brightens.
        const [hue, saturation, lightness] = rgbToHsl(channels);
        channels = hslToRgb(hue, saturation, Math.min(lightness * Math.max(numeric / 100000, 0), 1));
        break;
      }
      case "satMod": {
        // Ground truth: PowerPoint's own render of a theme background style
        // (`tint 95000` + `satMod 170000` on a dark scheme colour).
        const [hue, saturation, lightness] = rgbToHsl(channels);
        channels = hslToRgb(hue, Math.min(saturation * Math.max(numeric / 100000, 0), 1), lightness);
        break;
      }
      case "lumOff": {
        const [hue, saturation, lightness] = rgbToHsl(channels);
        channels = hslToRgb(hue, saturation, Math.min(lightness + ratio, 1));
        break;
      }
      case "alpha":
        alpha *= ratio;
        break;
    }
  }
  return alpha === 1 ? toHexColor(channels) : `rgb(${channels.map(clampChannel).join(" ")} / ${alpha})`;
}
