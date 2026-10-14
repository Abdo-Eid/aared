/**
 * Headless text measurement for chart layout. Office lays a chart out from
 * the font's advance widths, so the engine does too: the layout comes out the
 * same whatever fonts the reader's device has, and stays testable without a
 * browser. Tables for the faces charts are actually set in ship as two base-36
 * digits per code point (see metrics-data.ts); anything else falls to the
 * nearest metric class. Kerning is ignored: Office does not kern chart text.
 *
 * Arabic is measured per letter and joining form (isolated, initial, medial,
 * final), in the run's complex-script face: a joined letter is about half the
 * width of the isolated shape, so one flat width per letter laid Arabic
 * labels out a third too wide or narrow depending on the word. The page can
 * also hand in a per-face scale (`setFontScales`) measured from the font it
 * actually paints, for readers without Office's faces.
 * Runs on the page (the layout half), so the family strings it caches by are
 * copies from the unit model, not slices of a part held by the Worker.
 */

import { ARABIC_LETTERS, ARABIC_MARKS, ARABIC_NEUTRALS, ARABIC_TABLES, FACE_TABLES } from "./metrics-data.js";
import type { TextLook } from "./spec.js";

/** Advance of an Arabic-script character the tables do not cover (a presentation form typed directly, a rare letter), em. */
const ARABIC_FALLBACK = 0.45;

interface Face { ascent: number; descent: number; regular: Float32Array; bold: Float32Array; scale: number }

const faces = new Map<string, Face>();

function decode(table: string): Float32Array {
  const out = new Float32Array(table.length / 2);
  for (let i = 0; i < out.length; i++) out[i] = parseInt(table.slice(i * 2, i * 2 + 2), 36) / 1000;
  return out;
}

/** A serif face by name ("Microsoft Sans Serif" and "sans-serif" are not). `f` is lower case. */
function isSerif(f: string): boolean {
  return !f.includes("sans") && (f.includes("times") || f.includes("serif") || f.includes("cambria") || f.includes("georgia") || f.includes("garamond") || f.includes("caladea"));
}

/** Which table a family measures with, and a width scale for near relatives. */
function classify(family: string): [string, number] {
  const f = family.toLowerCase();
  if (FACE_TABLES[f] !== undefined) return [f, 1];
  if (f.includes("aptos")) return ["arial", 0.97];
  if (f.includes("segoe") || f.includes("tahoma") || f.includes("verdana")) return ["arial", f.includes("verdana") ? 1.12 : 1.02];
  if (f.includes("calibri") || f.includes("carlito") || f.includes("corbel") || f.includes("candara")) return ["calibri", 1];
  if (isSerif(f)) return ["times new roman", f.includes("cambria") || f.includes("georgia") ? 1.08 : 1];
  if (f.includes("courier") || f.includes("mono") || f.includes("consol")) return ["courier", 1];
  return ["arial", 1];
}

function faceFor(family: string): Face {
  let face = faces.get(family);
  if (face !== undefined) return face;
  const [key, scale] = classify(family);
  if (key === "courier") {
    const mono = new Float32Array(191).fill(0.6);
    face = { ascent: 0.833, descent: 0.3, regular: mono, bold: mono, scale: 1 };
  } else {
    const table = FACE_TABLES[key]!;
    face = { ascent: table.ascent / 1000, descent: table.descent / 1000, regular: decode(table.regular), bold: decode(table.bold), scale };
  }
  if (faces.size < 64) faces.set(family, face);
  return face;
}

function advance(code: number, widths: Float32Array): number {
  if (code >= 0x20 && code <= 0x7e) return widths[code - 0x20]!;
  if (code >= 0xa0 && code <= 0xff) return widths[code - 0xa0 + 95]!;
  if (code === 0x2013) return 0.5;
  if (code === 0x2014) return 1;
  if (code === 0x2026) return 1;
  if (code >= 0x2e80) return 1;
  if (code === 0x200f || code === 0x200e || code === 0x200b || code === 0x200c || code === 0x200d) return 0;
  return 0.55;
}

// ---------------------------------------------------------------- Arabic

/** Joining types: none (a letter that joins nothing), right-joining, dual-joining, transparent (a mark). Zero is "not in the tables". */
const J_NONE = 1;
const J_RIGHT = 2;
const J_DUAL = 3;
const J_MARK = 4;

const ALEF_FOR_LAM = [0x0622, 0x0623, 0x0625, 0x0627];
const LAM = 0x0644;
const ZWJ = 0x200d;

interface ArabicShapes {
  /** Joining type per code point of U+0600-06FF. */
  kind: Uint8Array;
  /** Index of the letter's first width in a face table, per code point of U+0600-06FF. */
  at: Int16Array;
  /** Where lam-alef, the digits and the spaces and punctuation start in a face table. */
  lamAlef: number;
  digits: number;
  neutral: number;
  /** Index among the tabulated spaces and punctuation, per ASCII code; -1 for any other. */
  neutralAt: Int8Array;
}

let shapes: ArabicShapes | undefined;

function arabicShapes(): ArabicShapes {
  if (shapes !== undefined) return shapes;
  const kind = new Uint8Array(0x700);
  const at = new Int16Array(0x700);
  let next = 0;
  for (let i = 0; i < ARABIC_LETTERS.length; i += 3) {
    const code = parseInt(ARABIC_LETTERS.slice(i, i + 2), 36);
    const type = ARABIC_LETTERS[i + 2];
    kind[code] = type === "d" ? J_DUAL : type === "r" ? J_RIGHT : J_NONE;
    at[code] = next;
    next += type === "d" ? 4 : type === "r" ? 2 : 1;
  }
  for (let i = 0; i < ARABIC_MARKS.length; i += 2) {
    for (let code = ARABIC_MARKS[i]!; code <= ARABIC_MARKS[i + 1]!; code++) kind[code - 0x0600] = J_MARK;
  }
  const neutralAt = new Int8Array(128).fill(-1);
  let index = 0;
  for (let i = 0; i < ARABIC_NEUTRALS.length; i += 2) {
    for (let code = ARABIC_NEUTRALS[i]!; code <= ARABIC_NEUTRALS[i + 1]!; code++) neutralAt[code] = index++;
  }
  shapes = { kind, at, lamAlef: next, digits: next + 8, neutral: next + 10, neutralAt };
  return shapes;
}

interface ArabicFace { widths: Float32Array; scale: number }

const arabicFaces = new Map<string, ArabicFace>();

/** Which Arabic table a complex-script face measures with. Times New Roman's Arabic glyphs are Arial's. */
function arabicFaceFor(family: string): ArabicFace {
  let face = arabicFaces.get(family);
  if (face !== undefined) return face;
  const f = family.toLowerCase();
  const key = f.includes("tahoma") || f.includes("verdana") ? "tahoma" : isSerif(f) ? "times new roman" : "arial";
  const table = ARABIC_TABLES[key] ?? ARABIC_TABLES["arial"]!;
  face = { widths: decode(table), scale: f.includes("verdana") ? 1.1 : f.includes("segoe") ? 1.02 : 1 };
  if (arabicFaces.size < 64) arabicFaces.set(family, face);
  return face;
}

/** Whether `code` is an Arabic-script character (the blocks Office draws with the complex-script face). */
function isArabicScript(code: number): boolean {
  return (code >= 0x0600 && code <= 0x06ff) || (code >= 0x0750 && code <= 0x077f) || (code >= 0x08a0 && code <= 0x08ff)
    || (code >= 0xfb50 && code <= 0xfdff) || (code >= 0xfe70 && code <= 0xfeff);
}

/** Whether `text` holds any Arabic-script character. */
export function hasArabicScript(text: string): boolean {
  for (let i = 0; i < text.length; i++) if (isArabicScript(text.charCodeAt(i))) return true;
  return false;
}

/**
 * The script of one character in a line that carries Arabic: 2 Arabic, 1 Latin (an ASCII
 * letter or digit, a Latin-1 or Latin Extended letter, which keep the Latin face even inside
 * Arabic), 0 neutral (a space, punctuation, a zero-width character).
 */
function scriptOf(code: number): 0 | 1 | 2 {
  if (isArabicScript(code)) return 2;
  if ((code >= 0x30 && code <= 0x39) || (code >= 0x41 && code <= 0x5a) || (code >= 0x61 && code <= 0x7a)
    || (code >= 0xc0 && code <= 0x24f && code !== 0xd7 && code !== 0xf7)) return 1;
  return 0;
}

/** Whether the first strong (Arabic or Latin) character of `text` is Arabic: what a leading space or bracket belongs to. */
function startsArabic(text: string): boolean {
  for (let i = 0; i < text.length; i++) {
    const script = scriptOf(text.charCodeAt(i));
    if (script !== 0) return script === 2;
  }
  return false;
}

/**
 * `text` cut into runs by the face that sets each: Arabic letters in the complex-script
 * face, Latin letters and digits in the Latin one. A neutral character (a space, a bracket)
 * goes with the strong character before it, so a space between two Arabic words is the
 * Arabic face's; leading ones go with what follows. `textWidth` measures by the same rule.
 */
export function scriptRuns(text: string): Array<{ text: string; arabic: boolean }> {
  const runs: Array<{ text: string; arabic: boolean }> = [];
  let arabic = startsArabic(text);
  let from = 0;
  for (let i = 0; i < text.length; i++) {
    const script = scriptOf(text.charCodeAt(i));
    if (script === 0 || (script === 2) === arabic) continue;
    if (i > from) runs.push({ text: text.slice(from, i), arabic });
    from = i;
    arabic = script === 2;
  }
  runs.push({ text: text.slice(from), arabic });
  return runs;
}

/**
 * Whether the first strong character of `text` is right to left, which is the
 * base direction the Unicode bidi algorithm gives the line. Digits and
 * punctuation are neutral and skipped; text with no strong letter is left to
 * right.
 */
export function isRtlText(text: string): boolean {
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c === 0x200f) return true;
    if (c === 0x200e) return false;
    if ((c >= 0x0590 && c <= 0x08ff) || (c >= 0xfb1d && c <= 0xfdff) || (c >= 0xfe70 && c <= 0xfefc)) {
      // Arabic-Indic digits, their extended forms and the Arabic comma are weak, not strong.
      if ((c >= 0x0660 && c <= 0x066c) || (c >= 0x06f0 && c <= 0x06f9) || c === 0x060c) continue;
      if (c >= 0x0591 && c <= 0x05bd) continue;
      return true;
    }
    if ((c >= 0x41 && c <= 0x5a) || (c >= 0x61 && c <= 0x7a) || c === 0xaa || c === 0xb5 || c === 0xba || (c >= 0xc0 && c <= 0x24f && c !== 0xd7 && c !== 0xf7)) return false;
    if ((c >= 0x0370 && c < 0x0590) || (c >= 0x0900 && c < 0x2000) || (c >= 0x3040 && c < 0xd800)) return false;
  }
  return false;
}

/**
 * The face Arabic text in `look` is set in: the run's complex-script face, or
 * Office's own fallback when nothing names one (Arial, Times New Roman beside
 * a serif face).
 */
export function arabicFamily(look: TextLook): string {
  if (look.csFamily !== undefined && look.csFamily !== "") return look.csFamily;
  return classify(look.family)[0] === "times new roman" ? "Times New Roman" : "Arial";
}

/** The key a font scale for a face's Arabic glyphs goes under (the same face name can scale differently per script). */
export function arabicScaleKey(family: string): string {
  return `${family}|cs`;
}

let fontScales: Readonly<Record<string, number>> | undefined;

/**
 * Scale every width by what the reader's own fonts measured against the
 * tables: a family key (`look.family`, or `arabicScaleKey(face)` for Arabic
 * glyphs) maps to painted width over table width. `undefined` clears it. The
 * layout sets it around one `drawChart` call, never longer.
 */
export function setFontScales(scales: Readonly<Record<string, number>> | undefined): void {
  fontScales = scales;
}

function jKind(text: string, index: number, shape: ArabicShapes): number {
  const code = text.charCodeAt(index);
  if (code === ZWJ) return J_DUAL;
  if (code >= 0x0600 && code <= 0x06ff) {
    const kind = shape.kind[code - 0x0600]!;
    return kind === 0 ? J_NONE : kind;
  }
  return J_NONE;
}

/** Skip joining-transparent marks and zero-width non-joiners: the index of the nearest real neighbour, or -1 / length. */
function neighbour(text: string, from: number, step: 1 | -1, shape: ArabicShapes): number {
  let i = from + step;
  while (i >= 0 && i < text.length && (jKind(text, i, shape) === J_MARK)) i += step;
  return i;
}

/**
 * Width in em of `text` split into [Latin part, Arabic part]: Arabic letters by
 * joining form in `arabic`'s face, the rest through the Latin face's `widths`
 * (a space or punctuation mark inside an Arabic run, in the Arabic face's own).
 * A non-Arabic character (a space, a digit, a Latin letter) ends a join; marks
 * take no advance and let one pass through; lam followed by alef is one glyph.
 */
function mixedEm(text: string, widths: Float32Array, arabic: Float32Array): [number, number] {
  const shape = arabicShapes();
  let latin = 0;
  let joined = 0;
  const n = text.length;
  // A neutral character is measured in the face of the run it belongs to (see scriptRuns).
  let inArabic = startsArabic(text);
  for (let i = 0; i < n; i++) {
    const code = text.charCodeAt(i);
    const script = scriptOf(code);
    if (script !== 2) {
      if (script === 1) inArabic = false;
      const slot = inArabic && code < 128 ? shape.neutralAt[code]! : -1;
      if (slot >= 0) joined += arabic[shape.neutral + slot]!;
      else latin += advance(code, widths);
      continue;
    }
    inArabic = true;
    if (code > 0x06ff) {
      joined += ARABIC_FALLBACK;
      continue;
    }
    if ((code >= 0x0660 && code <= 0x0669)) { joined += arabic[shape.digits]!; continue; }
    if (code >= 0x06f0 && code <= 0x06f9) { joined += arabic[shape.digits + 1]!; continue; }
    const kind = shape.kind[code - 0x0600]!;
    if (kind === J_MARK) continue;
    if (kind === 0) { joined += ARABIC_FALLBACK; continue; }
    const prev = neighbour(text, i, -1, shape);
    const next = neighbour(text, i, 1, shape);
    const joinsPrev = prev >= 0 && jKind(text, prev, shape) === J_DUAL;
    const base = shape.at[code - 0x0600]!;
    if (code === LAM && next < n) {
      const alef = ALEF_FOR_LAM.indexOf(text.charCodeAt(next));
      if (alef >= 0) {
        joined += arabic[shape.lamAlef + alef * 2 + (joinsPrev ? 1 : 0)]!;
        i = next;
        continue;
      }
    }
    const nextKind = next < n ? jKind(text, next, shape) : J_NONE;
    const joinsNext = nextKind === J_DUAL || nextKind === J_RIGHT;
    if (kind === J_DUAL) joined += arabic[base + (joinsPrev ? (joinsNext ? 2 : 3) : joinsNext ? 1 : 0)]!;
    else if (kind === J_RIGHT) joined += arabic[base + (joinsPrev ? 1 : 0)]!;
    else joined += arabic[base]!;
  }
  return [latin, joined];
}

/**
 * Width of one line of text in points: Latin characters in `look.family`,
 * Arabic ones in its complex-script face. `joined_width` in
 * tools/chart-oracle/gen_metrics.py is the same rule over the font files.
 */
export function textWidth(text: string, look: TextLook): number {
  const face = faceFor(look.family);
  const widths = look.bold ? face.bold : face.regular;
  const latinScale = fontScales === undefined ? 1 : fontScales[look.family] ?? 1;
  let em = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (isArabicScript(code)) {
      const cs = arabicFamily(look);
      const arabic = arabicFaceFor(cs);
      const csScale = fontScales === undefined ? 1 : fontScales[arabicScaleKey(cs)] ?? 1;
      const [latin, joined] = mixedEm(text, widths, arabic.widths);
      return (latin * face.scale * latinScale + joined * arabic.scale * csScale) * look.size;
    }
    em += advance(code, widths);
  }
  return em * look.size * face.scale * latinScale;
}

/** Distance between baselines of consecutive lines, in points. */
export function lineHeight(look: TextLook): number {
  const face = faceFor(look.family);
  const natural = (face.ascent + face.descent) * look.size;
  // A stored `normAutofit/@lnSpcReduction` removes this share of the leading
  // without shrinking the glyphs (Excel 16.0 build 20430, probe
  // `chart-txpr-autofit`: 25.39pt -> 20.30pt with 20000).
  return look.lineSpacingReduction === undefined ? natural : natural * (1 - look.lineSpacingReduction);
}

/** Baseline offset below the top of a line box, in points. */
export function ascent(look: TextLook): number {
  return faceFor(look.family).ascent * look.size;
}

/**
 * Greedy word wrap at spaces into lines no wider than `maxWidth`. A single
 * word wider than the limit keeps its own line (Office does not break inside
 * words for chart labels); the caller decides what to do with the overflow.
 */
export function wrapText(text: string, look: TextLook, maxWidth: number): string[] {
  const out: string[] = [];
  for (const paragraph of text.split("\n")) {
    const words = paragraph.split(" ");
    let line = "";
    for (const word of words) {
      const candidate = line === "" ? word : `${line} ${word}`;
      if (line !== "" && textWidth(candidate, look) > maxWidth) {
        out.push(line);
        line = word;
      } else {
        line = candidate;
      }
    }
    out.push(line);
  }
  return out;
}

/**
 * `text` cut to fit `maxWidth` with "..." (ASCII, so shipped scripts stay
 * Latin-1); empty when not even the dots fit.
 */
export function fitText(text: string, look: TextLook, maxWidth: number): string {
  if (textWidth(text, look) <= maxWidth) return text;
  const dots = "...";
  let out = "";
  for (const ch of text) {
    if (textWidth(out + ch + dots, look) > maxWidth) break;
    out += ch;
  }
  if (out === "") return textWidth(dots, look) <= maxWidth ? dots : "";
  return out + dots;
}

/** Word-wrap to `maxWidth`, at most `maxLines` lines; what does not fit is cut with "...". */
export function wrapFit(text: string, look: TextLook, maxWidth: number, maxLines: number): string[] {
  let lines = wrapText(text, look, maxWidth);
  if (lines.length > maxLines) lines = [...lines.slice(0, maxLines - 1), lines.slice(maxLines - 1).join(" ")];
  return lines.map((l) => fitText(l, look, maxWidth));
}

/** Size of a block of lines. */
export function blockSize(lines: readonly string[], look: TextLook): { width: number; height: number } {
  let width = 0;
  for (const line of lines) width = Math.max(width, textWidth(line, look));
  return { width, height: lines.length * lineHeight(look) };
}
