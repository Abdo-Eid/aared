/**
 * Characters above U+00FF that aared emits at runtime.
 *
 * V8 keeps a script's source one byte per character only while every
 * character fits in Latin-1; a single wider one makes the whole script
 * (the Worker's is ~128 KB) two bytes per character, in every Worker. So
 * shipped string, template and regex literals stay Latin-1, and the few
 * wider characters live here, spelled so no build step can turn them back
 * into raw characters: a `\u` escape in a literal doesn't survive (Vite's
 * esbuild prints with `charset: "utf8"`), and `String.fromCharCode(…)` is
 * constant-folded by rolldown. `JSON.parse` of an ASCII escape is evaluated
 * only at runtime. `test/shipped-latin1.test.ts` and
 * `tools/check-consumer-build.ts` enforce it.
 */

const chars = (...escapes: string[]): string => JSON.parse(`"${escapes.map((escape) => `\\u${escape}`).join("")}"`) as string;

/** U+2022, the default bullet (PPTX empty and picture-fallback bullets, DOCX fallback list level). */
export const BULLET = /* @__PURE__ */ chars("2022");
/** U+2011, Word's `w:noBreakHyphen`. */
export const NON_BREAKING_HYPHEN = /* @__PURE__ */ chars("2011");
/** U+200E, the left-to-right mark after a DOCX tab spacer in an LTR paragraph. */
export const LEFT_TO_RIGHT_MARK = /* @__PURE__ */ chars("200e");
/** U+200F, the right-to-left mark after a DOCX tab spacer in an RTL paragraph. */
export const RIGHT_TO_LEFT_MARK = /* @__PURE__ */ chars("200f");
/** U+21A9, the footnote back-link glyph. */
export const RETURN_ARROW = /* @__PURE__ */ chars("21a9");
/** U+FE0E, the text-style selector after a symbol-font stand-in a phone may draw as colour emoji. */
export const TEXT_PRESENTATION = /* @__PURE__ */ chars("fe0e");
/** U+0627, bare alef: what `core/search.ts` folds every alef and hamza-on-alef form to. */
export const ARABIC_ALEF = /* @__PURE__ */ chars("0627");
/** U+064A, ya: what `core/search.ts` folds alef maqsura to. */
export const ARABIC_YEH = /* @__PURE__ */ chars("064a");
/**
 * An Arabic sentence ("sales by quarter") that takes every joining form, for the chart engine's font probe
 * (measuring how wide the reader's complex-script face paints it) and as the sample that makes a theme's
 * `cs` slot resolve to its Arabic font.
 */
export const ARABIC_PROBE = /* @__PURE__ */ chars(
  "0627", "0644", "0645", "0628", "064a", "0639", "0627", "062a", "0020",
  "062d", "0633", "0628", "0020",
  "0627", "0644", "0631", "0628", "0639",
);
