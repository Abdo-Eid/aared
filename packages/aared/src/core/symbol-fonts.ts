/**
 * Unicode stand-ins for the symbol-font characters Office offers as bullets
 * and list markers. Wingdings and Symbol keep their glyphs at letter codes
 * (Wingdings' square is "q", its check mark a u with diaeresis), in either
 * the plain form or the private-use form Word writes (U+F000 + code). Only
 * machines with Office have those fonts: anywhere else the marker paints as
 * the letter, or as an empty box. A stand-in paints the same shape in any
 * font that has it. Each entry was matched by eye against the real glyph
 * (Notes §52); a character with no entry keeps its font and code.
 *
 * Code points, never literals: shipped JS stays Latin-1 (see `chars.ts`,
 * which holds the one fixed character this needs).
 */

import { TEXT_PRESENTATION } from "./chars.js";

const WINGDINGS: ReadonlyMap<number, number> = new Map([
  [0x6c, 0x25cf], // filled circle
  [0x6e, 0x25a0], // filled square
  [0x6f, 0x25a1], // outline square
  [0x70, 0x25a1], // heavy outline square
  [0x71, 0x2751], // shadowed square
  [0x75, 0x25c6], // filled diamond
  [0x76, 0x2756], // four-diamond flower
  [0x77, 0x2b25], // small diamond
  [0x9f, 0x2022], // small bullet
  [0xa1, 0x25cb], // outline circle
  [0xa7, 0x25aa], // small square
  [0xa8, 0x25fb], // outline square, medium
  [0xd8, 0x27a2], // arrowhead
  [0xe0, 0x2192], // plain arrow
  [0xe8, 0x2794], // heavy arrow
  [0xf0, 0x21e8], // outline arrow
  [0xfb, 0x2717], // ballot x
  [0xfc, 0x2713], // check mark
  [0xfd, 0x2612], // box with x
  [0xfe, 0x2611], // box with check
]);

const SYMBOL: ReadonlyMap<number, number> = new Map([
  [0xa7, 0x2663], // club
  [0xa8, 0x2666], // diamond
  [0xae, 0x2192], // arrow
  [0xb7, 0x2022], // bullet: Word's default first-level bullet
  [0xde, 0x21d2], // double arrow
]);

const TABLES: ReadonlyMap<string, ReadonlyMap<number, number>> = new Map([
  ["wingdings", WINGDINGS],
  ["symbol", SYMBOL],
]);

/** Stand-ins a phone may draw as colour emoji: followed by the text-style selector (U+FE0E). */
const TEXT_STYLE: ReadonlySet<number> = new Set([0x2611, 0x25aa, 0x25fb, 0x2663, 0x2666]);

/**
 * `text` set in `typeface` as Unicode any font can draw, or `undefined` when
 * the typeface is not Wingdings or Symbol or a character has no stand-in (the
 * caller then keeps the text and its font as they are).
 */
export function symbolFontStandIn(typeface: string | undefined, text: string): string | undefined {
  if (typeface === undefined || text === "") return undefined;
  const table = TABLES.get(typeface.trim().toLowerCase());
  if (table === undefined) return undefined;
  let out = "";
  for (const char of text) {
    let code = char.codePointAt(0)!;
    if (code >= 0xf000 && code <= 0xf0ff) code -= 0xf000;
    const standIn = table.get(code);
    if (standIn === undefined) return undefined;
    out += String.fromCodePoint(standIn);
    if (TEXT_STYLE.has(standIn)) out += TEXT_PRESENTATION;
  }
  return out;
}
