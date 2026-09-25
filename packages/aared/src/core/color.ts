/**
 * Colour arithmetic shared by every parser: hex parsing and formatting, sRGB
 * mixing, and linear-light conversion. Pure numbers and strings, no DOM.
 * Format-specific colour *semantics* (DrawingML `tint`, SpreadsheetML `tint`,
 * PowerPoint's automatic text colour) stay with their parsers; only the
 * arithmetic underneath lives here.
 */

/** Red, green and blue, each 0-255. */
export type Channels = [number, number, number];

const HEX_DIGITS = /^[0-9a-fA-F]{6}$/;

/** Channels of six bare hex digits (`RRGGBB`); `undefined` for anything else. */
export function parseHexDigits(digits: string): Channels | undefined {
  if (!HEX_DIGITS.test(digits)) return undefined;
  return [
    Number.parseInt(digits.slice(0, 2), 16),
    Number.parseInt(digits.slice(2, 4), 16),
    Number.parseInt(digits.slice(4, 6), 16),
  ];
}

/** Channels of `#RRGGBB` or `RRGGBB`; `undefined` for anything else (no alpha, no short forms). */
export function parseHexColor(value: string): Channels | undefined {
  return parseHexDigits(value.startsWith("#") ? value.slice(1) : value);
}

/**
 * Reads this module's own resolved-color output back into RGBA channels —
 * `resolveDrawingColor`'s `#RRGGBB` (opaque) or `rgb(r g b / a)` (has alpha)
 * — never a general CSS color parser.
 */
export function parseResolvedColor(value: string): { r: number; g: number; b: number; a: number } | undefined {
  const hex = parseHexColor(value);
  if (hex !== undefined) return { r: hex[0], g: hex[1], b: hex[2], a: 1 };
  const match = /^rgb\((\d+(?:\.\d+)?) (\d+(?:\.\d+)?) (\d+(?:\.\d+)?) \/ (\d+(?:\.\d+)?)\)$/.exec(value);
  if (match === null) return undefined;
  return { r: Number(match[1]), g: Number(match[2]), b: Number(match[3]), a: Number(match[4]) };
}

/** One channel as an integer, clamped to 0-255. */
export function clampChannel(channel: number): number {
  return Math.round(Math.min(Math.max(channel, 0), 255));
}

/** Upper-case `#RRGGBB` from channels; each is rounded and clamped. */
export function toHexColor(channels: readonly number[]): string {
  return `#${channels.map((channel) => clampChannel(channel).toString(16).padStart(2, "0")).join("").toUpperCase()}`;
}

/** The sRGB colour `ratio` (0-1) of the way from `from` to `to`, both `#RRGGBB`; `from` when either does not parse. */
export function mixColors(from: string, to: string, ratio: number): string {
  const start = parseHexColor(from);
  const end = parseHexColor(to);
  if (start === undefined || end === undefined) return from;
  const clamped = Math.min(Math.max(ratio, 0), 1);
  return toHexColor(start.map((channel, index) => channel + (end[index]! - channel) * clamped));
}

/** An 8-bit sRGB channel as linear light, 0-1. */
export function srgbToLinear(channel: number): number {
  const value = channel / 255;
  return value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
}

/** Linear light (clamped to 0-1) as an 8-bit sRGB channel. */
export function linearToSrgb(value: number): number {
  const clamped = Math.min(Math.max(value, 0), 1);
  return Math.round((clamped <= 0.0031308 ? clamped * 12.92 : 1.055 * clamped ** (1 / 2.4) - 0.055) * 255);
}

function mixResolved(from: string, to: string, ratio: number, channel: (a: number, b: number, ratio: number) => number): string {
  const start = parseResolvedColor(from);
  const end = parseResolvedColor(to);
  if (start === undefined || end === undefined) return from;
  const clamped = Math.min(Math.max(ratio, 0), 1);
  const channels = [channel(start.r, end.r, clamped), channel(start.g, end.g, clamped), channel(start.b, end.b, clamped)];
  const alpha = start.a + (end.a - start.a) * clamped;
  return alpha >= 1 ? toHexColor(channels) : `rgb(${channels.join(" ")} / ${Math.round(alpha * 1000) / 1000})`;
}

/**
 * The colour `ratio` (0-1) of the way from `from` to `to`, mixed in linear
 * light; alpha mixes linearly. Both are resolved colours (`#RRGGBB` or
 * `rgb(r g b / a)`); `from` comes back unchanged when either does not parse.
 */
export function mixResolvedColorsLinear(from: string, to: string, ratio: number): string {
  return mixResolved(from, to, ratio, (a, b, t) => linearToSrgb(srgbToLinear(a) + (srgbToLinear(b) - srgbToLinear(a)) * t));
}

/** As `mixResolvedColorsLinear`, but mixing the sRGB channel values directly. */
export function mixResolvedColorsSrgb(from: string, to: string, ratio: number): string {
  return mixResolved(from, to, ratio, (a, b, t) => clampChannel(a + (b - a) * t));
}

/** Relative luminance (Rec. 709 weights on linear light), 0-1. */
export function relativeLuminance(channels: readonly number[]): number {
  const [red, green, blue] = channels.map(srgbToLinear);
  return 0.2126 * red! + 0.7152 * green! + 0.0722 * blue!;
}

/** WCAG contrast ratio (1 to 21) between two `#RRGGBB` colours; `undefined` when either does not parse. */
export function contrastRatio(first: string, second: string): number | undefined {
  const a = parseHexColor(first);
  const b = parseHexColor(second);
  if (a === undefined || b === undefined) return undefined;
  const [lighter, darker] = [relativeLuminance(a), relativeLuminance(b)].sort((x, y) => y - x);
  return (lighter! + 0.05) / (darker! + 0.05);
}
