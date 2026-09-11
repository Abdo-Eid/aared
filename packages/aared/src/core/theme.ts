/** Shared DrawingML theme reader for PPTX, XLSX, and DOCX drawing content. */

import { attr, childElements, firstChild, parseXml } from "./xml.js";
import type { XmlElement } from "./xml.js";
import type { OpcPackage } from "./opc.js";
import { throwIfAborted } from "./zip.js";

export interface DrawingTheme {
  colors: Record<string, string>;
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
  const xml = await pkg.readText(part, signal);
  throwIfAborted(signal);
  const root = parseXml(xml);
  throwIfAborted(signal);
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
  return {
    colors,
    fonts: {
      major: readFontFamily(fontScheme === undefined ? undefined : firstChild(fontScheme, "majorFont")),
      minor: readFontFamily(fontScheme === undefined ? undefined : firstChild(fontScheme, "minorFont")),
    },
  };
}

/** Whether text needs the theme's Arabic complex-script slot. */
export function usesArabic(text: string): boolean {
  return /[\u0600-\u06ff]/u.test(text);
}

/** Resolve a theme token or literal DrawingML typeface to a browser family. */
export function resolveThemeTypeface(typeface: string | undefined, text: string, theme: DrawingTheme): string | undefined {
  if (typeface === undefined || typeface === "") return undefined;
  const match = /^\+m([jn])-(lt|cs)$/.exec(typeface);
  if (match === null) return typeface;
  const family = match[1] === "j" ? theme.fonts.major : theme.fonts.minor;
  if (match[2] === "lt") return family.latin || undefined;
  if (family.complexScript !== "") return family.complexScript;
  return usesArabic(text) ? family.scripts["Arab"] : undefined;
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

function sRgbToLinear(channel: number): number {
  const value = channel / 255;
  return value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
}

function linearToSRgb(value: number): number {
  const clamped = Math.min(Math.max(value, 0), 1);
  return Math.round((clamped <= 0.0031308 ? clamped * 12.92 : 1.055 * clamped ** (1 / 2.4) - 0.055) * 255);
}

function tintChannels(channels: number[], sourceRatio: number): number[] {
  return channels.map((channel) => linearToSRgb(sRgbToLinear(channel) * sourceRatio + (1 - sourceRatio)));
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
  if (source === undefined || !/^#[0-9a-fA-F]{6}$/.test(source.startsWith("#") ? source : `#${source}`)) return undefined;
  let channels = [1, 3, 5].map((offset) => Number.parseInt((source.startsWith("#") ? source : `#${source}`).slice(offset, offset + 2), 16));
  let alpha = 1;
  for (const transform of color.children) {
    const rawValue = attr(transform, "val");
    if (rawValue === undefined) continue;
    const numeric = Number(rawValue);
    if (!Number.isFinite(numeric)) throw new Error(`aared/core: invalid val="${rawValue}" on a:${transform.local}`);
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
        const [hue, saturation, lightness] = rgbToHsl(channels);
        channels = hslToRgb(hue, saturation, lightness * ratio);
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
  const hex = `#${channels.map((channel) => Math.round(Math.min(Math.max(channel, 0), 255)).toString(16).padStart(2, "0")).join("").toUpperCase()}`;
  return alpha === 1 ? hex : `rgb(${channels.map((channel) => Math.round(Math.min(Math.max(channel, 0), 255))).join(" ")} / ${alpha})`;
}
