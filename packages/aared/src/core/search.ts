/**
 * Shared `search()` and lookup helpers for the three format entrypoints.
 * Internal: not part of the public surface. `search()` matching is a
 * substring test over `searchKey` (case and Arabic spelling variants folded
 * away); each format owns what text it searches and how a hit is shaped. A hit
 * reports the original text, never the key, so the key may change length.
 * Lookups by label (`findByTitle`) are exact-then-case-folded equality, never
 * substring, and never Arabic-folded.
 */

import { ARABIC_ALEF, ARABIC_YEH } from "./chars.js";

/** Longest snippet a search match reports. */
const SNIPPET_LENGTH = 120;

/**
 * The case fold. Title lookups compare through it; searches go on to
 * `searchKey`, which starts from it.
 */
export function foldText(text: string): string {
  return text.toLowerCase();
}

/**
 * The one fold every `search()` compares through: needles and haystacks both
 * pass it, so changing the rule here changes both sides together. After the
 * case fold it makes the spellings Arabic readers treat as one word equal:
 *
 * - tashkeel (U+064B-065F, superscript alef U+0670) and tatweel (U+0640) are
 *   dropped;
 * - alef with madda, hamza above or below, and alef wasla become a bare alef;
 * - alef maqsura becomes ya;
 * - Arabic-Indic and extended (Persian) digits become 0-9, and the Arabic
 *   decimal and thousands separators become `.` and `,`;
 * - Arabic presentation forms (what a PDF paste leaves behind) take their
 *   base letters first, through NFKC.
 *
 * Left alone on purpose, because each merge changes a word's meaning or is a
 * habit of some writers only: ta marbuta against ha, and the hamza seats waw
 * and ya (and a bare hamza). Text without any Arabic-script code unit takes
 * the plain case fold with no further work, so a large sheet of Latin text
 * indexes as fast as before.
 */
export function searchKey(text: string): string {
  let arabic = false;
  let presentation = false;
  for (let at = 0; at < text.length; at++) {
    const code = text.charCodeAt(at);
    if (code < 0x0622) continue;
    if (code <= 0x06f9) arabic = true;
    else if ((code >= 0xfb50 && code <= 0xfdff) || (code >= 0xfe70 && code <= 0xfeff)) presentation = true;
  }
  if (!arabic && !presentation) return foldText(text);

  const base = foldText(presentation ? text.normalize("NFKC") : text);
  let key = "";
  for (let at = 0; at < base.length; at++) {
    const code = base.charCodeAt(at);
    if (code < 0x0622 || code > 0x06f9) key += base[at];
    else if (code === 0x0640 || (code >= 0x064b && code <= 0x065f) || code === 0x0670) continue;
    else if (code === 0x0622 || code === 0x0623 || code === 0x0625 || code === 0x0671) key += ARABIC_ALEF;
    else if (code === 0x0649) key += ARABIC_YEH;
    else if (code >= 0x0660 && code <= 0x0669) key += String.fromCharCode(code - 0x0660 + 48);
    else if (code >= 0x06f0 && code <= 0x06f9) key += String.fromCharCode(code - 0x06f0 + 48);
    else if (code === 0x066b) key += ".";
    else if (code === 0x066c) key += ",";
    else key += base[at];
  }
  return key;
}

/**
 * The one rule every lookup by label follows (`sheets.get("Q1")`,
 * `headings.get("Scope")`, `goToHeading("Scope")`): the item whose title is
 * exactly `title`, else the first whose title matches ignoring case. Never a
 * substring — partial matching belongs to `search()`. The caller raises its
 * own `not-found` when nothing matches.
 */
export function findByTitle<T>(items: readonly T[], title: string, titleOf: (item: T) => string): T | undefined {
  const exact = items.find((item) => titleOf(item) === title);
  if (exact !== undefined) return exact;
  const folded = foldText(title);
  return items.find((item) => foldText(titleOf(item)) === folded);
}

/**
 * The search key for `query`, or `null` when it matches nothing: an empty
 * query, or one made only of marks the key drops (a lone tashkeel).
 */
export function foldQuery(query: string): string | null {
  const needle = searchKey(query);
  return needle === "" ? null : needle;
}

/** A block's plain text: its runs joined in order. */
export function runsText(runs: readonly { text: string }[]): string {
  return runs.map((run) => run.text).join("");
}

/** Trim a matching block to the reported snippet length. */
export function clipSnippet(text: string): string {
  return text.slice(0, SNIPPET_LENGTH);
}
