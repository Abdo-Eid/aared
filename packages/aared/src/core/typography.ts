/**
 * Script and line-box rules every format shares. DOM-free, so the Worker
 * (font-slot choice, pagination estimates) and the renderers agree on them.
 */

/**
 * Smallest line-height multiple painted or estimated. Below it Arabic glyph
 * ink (tall marks above and below the line) clips at the line box.
 */
export const MINIMUM_LINE_HEIGHT = 1.1;

/** Office faces whose closest generic family is serif or monospace; everything else falls back to sans-serif. */
const SERIF_FACES = new Set([
  "cambria", "times new roman", "times", "georgia", "garamond", "book antiqua", "palatino linotype", "palatino",
  "constantia", "century", "century schoolbook", "bookman old style", "baskerville old face", "bodoni mt",
  "sylfaen", "traditional arabic", "simplified arabic", "amiri", "noto serif", "liberation serif",
]);
const MONOSPACE_FACES = new Set(["consolas", "courier new", "courier", "lucida console", "cascadia code", "cascadia mono", "liberation mono"]);

/**
 * A CSS `font-family` value for an Office face: the face, then the generic
 * family nearest it. A bare family has no fallback of its own, so a face the
 * reader lacks drops to the browser default (usually a serif) — Arial table
 * text painting in Times, say. PowerPoint, Word and Excel substitute a
 * similar face instead.
 */
export function cssFontFamily(family: string): string {
  const key = family.trim().toLowerCase();
  const generic = SERIF_FACES.has(key) ? "serif" : MONOSPACE_FACES.has(key) ? "monospace" : "sans-serif";
  return `"${family.replace(/["\\]/g, "")}", ${generic}`;
}

const TAB_ONLY = /^\t+$/;

/**
 * The size that sets a paragraph's line box. Word and PowerPoint both size a
 * line by the runs on it: the paragraph's default only shows through on a
 * run that states no size (and on an empty paragraph), so 9pt runs under a
 * 10.5pt style make a 9pt line. Ground truth: a native Word export's table
 * rows of 9pt text measure 38.7px, not the 42.9px of a 10.5pt line; in
 * PowerPoint, `post-production-review.pptx` slide 2's 13pt badge numbers
 * under an 18pt deck default sit centred only on a 13pt line. Shared by the
 * DOCX paginator and both renderers so painted boxes and estimates agree.
 */
export function lineFontSizePt(runs: ReadonlyArray<{ text: string; fontSizePt?: number }>, paragraphPt: number | undefined): number | undefined {
  let size: number | undefined;
  for (const run of runs) {
    // A run holding only tabs sets no line size: Word ignores an unsized tab
    // between 8pt footer runs, which measure as an 8pt line, not a 10.5pt one.
    if (run.text === "" || TAB_ONLY.test(run.text)) continue;
    const runSize = run.fontSizePt ?? paragraphPt;
    if (runSize !== undefined && (size === undefined || runSize > size)) size = runSize;
  }
  return size ?? paragraphPt;
}

/** Arabic, Arabic Supplement, Arabic Extended-A, and both presentation-form blocks. */
const ARABIC_CLASS = "[\\u0600-\\u06ff\\u0750-\\u077f\\u08a0-\\u08ff\\ufb50-\\ufdff\\ufe70-\\ufeff]";
const ARABIC_ANY = new RegExp(ARABIC_CLASS, "u");
const ARABIC_ALL = new RegExp(ARABIC_CLASS, "gu");

/** Whether text carries any Arabic-script character, so it needs the complex-script font slot. */
export function usesArabic(text: string): boolean {
  return ARABIC_ANY.test(text);
}

/** Share of the characters in `text` that are Arabic script, from 0 to 1. */
export function arabicShare(text: string): number {
  return text.length === 0 ? 0 : (text.match(ARABIC_ALL)?.length ?? 0) / text.length;
}
