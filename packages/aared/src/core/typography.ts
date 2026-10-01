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
 * Office faces a browser's own generic keyword commonly resolves to by
 * itself on a real machine (Arial for `sans-serif`, Times New Roman for
 * `serif`, Courier New for `monospace` — this is why those three names were
 * picked as the historical archetypes the generics stand in for). Text set
 * in one of these paints correctly whether the exact face is truly installed
 * or the platform's default coincidentally matches it, so
 * `detectFontFidelity` (`font-fidelity.ts`) never calls one of these
 * "unavailable" just because it measured the same as its own generic —
 * that comparison is genuinely inconclusive here, unlike for any other face.
 */
const COMMONLY_THE_GENERIC_DEFAULT = new Set(["arial", "helvetica", "times new roman", "times", "courier new", "courier"]);

/** Whether `family` is one a browser's own generic keyword commonly already resolves to (see above). */
export function resolvesLikeGenericDefault(family: string): boolean {
  return COMMONLY_THE_GENERIC_DEFAULT.has(family.trim().toLowerCase());
}

/**
 * Suffix a delivered font substitute registers its `FontFace` under. A
 * `FontResolver` (`aared-fonts`'s `defaultResolver`, or a caller's own) never
 * registers a substitute under the requested name itself — that would shadow
 * a reader who actually has the requested font installed, and would apply a
 * calibrated substitute's size/metric overrides to that real, unrelated
 * font. Registering under the alias instead lets `cssFontFamily`'s stack
 * fall through to it only when nothing local answers to the requested name.
 * The one home for this rule: both `cssFontFamily` (below) and every
 * `FontResolver` that delivers something call `fontAlias` on the same
 * requested family, so the two sides always agree on the name.
 */
export function fontAlias(family: string): string {
  return `${family} (aared)`;
}

/**
 * The generic CSS family nearest an Office face — the same classification
 * `cssFontFamily` stacks onto every requested family, exposed so
 * `detectFontFidelity` (`font-fidelity.ts`) can build an apples-to-apples
 * control of the same generic category (comparing a serif face's width
 * against the `sans-serif` default would read as "missing" even when it
 * isn't).
 */
export function genericFontFamily(family: string): "serif" | "monospace" | "sans-serif" {
  const key = family.trim().toLowerCase();
  return SERIF_FACES.has(key) ? "serif" : MONOSPACE_FACES.has(key) ? "monospace" : "sans-serif";
}

/**
 * A CSS `font-family` value for an Office face: the face, its delivered
 * substitute's alias (see `fontAlias`), then the generic family nearest it.
 * A bare family has no fallback of its own, so a face the reader lacks (and
 * for which no resolver delivered a substitute) drops to the browser default
 * (usually a serif) — Arial table text painting in Times, say.
 */
export function cssFontFamily(family: string): string {
  const clean = family.replace(/["\\]/g, "");
  return `"${clean}", "${fontAlias(clean)}", ${genericFontFamily(family)}`;
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

const RTL_LETTER = new RegExp(`[\\u0590-\\u05ff${ARABIC_CLASS.slice(1, -1)}]`, "u");

/**
 * Whether a run is a left-to-right island in a right-to-left paragraph, the
 * way Word and PowerPoint order it: the run is not right-to-left in Office's
 * terms (`runRtl`) and carries no Hebrew or Arabic letter. See
 * `TextRunModel.ltrRun`.
 */
export function isLtrRunInRtl(text: string, paragraphRtl: boolean, runRtl: boolean): boolean {
  return paragraphRtl && !runRtl && text.length > 0 && !RTL_LETTER.test(text);
}

/** A maximal stretch of ASCII punctuation or symbols (no spaces, letters or digits). */
const ASCII_PUNCTUATION = /[!-/:-@[-`{-~]+/g;
const ARABIC_AT_START = new RegExp(`^ *${ARABIC_CLASS}`, "u");
const ARABIC_AT_END = new RegExp(`${ARABIC_CLASS} *$`, "u");

/**
 * `text` cut around each stretch of ASCII punctuation that sits between two
 * Arabic words, or undefined when none does. PowerPoint orders that
 * punctuation left to right, and so the words on either side, in a
 * left-to-right paragraph's run tagged with a left-to-right language (probe
 * `en-us-arabic-run`: every ASCII punctuation character, not spaces, so
 * `W1 + W2` and `W1 (W2) W3` keep their written order while `W1 W2 W3` still
 * reads right to left). The caller flags the `punctuation` pieces as
 * left-to-right runs. Text with no Arabic letter exits before any slicing.
 */
export function arabicPunctuationPieces(text: string): Array<{ text: string; punctuation: boolean }> | undefined {
  if (!ARABIC_ANY.test(text)) return undefined;
  const pieces: Array<{ text: string; punctuation: boolean }> = [];
  let taken = 0;
  for (const match of text.matchAll(ASCII_PUNCTUATION)) {
    const start = match.index;
    const end = start + match[0].length;
    if (!ARABIC_AT_END.test(text.slice(taken, start)) || !ARABIC_AT_START.test(text.slice(end))) continue;
    if (start > taken) pieces.push({ text: text.slice(taken, start), punctuation: false });
    pieces.push({ text: match[0], punctuation: true });
    taken = end;
  }
  if (pieces.length === 0) return undefined;
  if (taken < text.length) pieces.push({ text: text.slice(taken), punctuation: false });
  return pieces;
}

/** Share of the characters in `text` that are Arabic script, from 0 to 1. */
export function arabicShare(text: string): number {
  return text.length === 0 ? 0 : (text.match(ARABIC_ALL)?.length ?? 0) / text.length;
}
