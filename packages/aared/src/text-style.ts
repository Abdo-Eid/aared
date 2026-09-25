/**
 * The text painting PPTX and DOCX share: a run's character formatting and a
 * paragraph's box, as CSS. Renderer-side. What only one format has (PPTX's
 * outline, gradient and shadows; DOCX's shading, scripts and borders) stays
 * with that format, applied on top.
 */

import type { TextAlign, TextDirection, TextRunModel } from "./core/model.js";
import { MINIMUM_LINE_HEIGHT, cssFontFamily } from "./core/typography.js";
import { inchesToCss } from "./core/units.js";

/** A point value with float noise trimmed, so `0.30000000000000004pt` never reaches a style. */
export function roundPt(value: number): number {
  return Math.round(value * 1000) / 1000;
}

export function pt(value: number): string {
  return `${roundPt(value)}pt`;
}

/**
 * Paint a run's character formatting onto its span: weight, style, size,
 * colour, family, underline and strikethrough, caps, letter-spacing and
 * highlight. Outline-only text (`noFill`) keeps its colour unset: the caller
 * decides how to paint it.
 */
export function applyRunBase(span: HTMLElement, run: TextRunModel): void {
  if (run.bold === true) span.style.fontWeight = "700";
  if (run.italic === true) span.style.fontStyle = "italic";
  if (run.fontSizePt !== undefined) span.style.fontSize = `${run.fontSizePt}pt`;
  if (run.color !== undefined && run.noFill !== true) span.style.color = run.color;
  if (run.fontFamily !== undefined) span.style.fontFamily = cssFontFamily(run.fontFamily);
  const decorations: string[] = [];
  if (run.underline === true) decorations.push("underline");
  // A double strike paints the same single CSS line; documented, not doubled.
  if (run.strikethrough === true) decorations.push("line-through");
  if (decorations.length > 0) {
    span.style.textDecoration = decorations.join(" ");
    // Browsers skip an underline around descenders and diacritics by default
    // (`text-decoration-skip-ink: auto`), tuned for Latin text. Arabic joins,
    // dots and tashkeel cross the line constantly, so PowerPoint and Word
    // both draw one continuous line; skip-ink instead breaks it into pieces.
    if (run.underline === true) span.style.textDecorationSkipInk = "none";
  }
  if (run.caps === "all") span.style.textTransform = "uppercase";
  else if (run.caps === "small") span.style.fontVariantCaps = "small-caps";
  if (run.letterSpacingPt !== undefined) span.style.letterSpacing = pt(run.letterSpacingPt);
  if (run.highlight !== undefined) span.style.backgroundColor = run.highlight;
}

/** The paragraph fields both formats' models carry. */
export interface ParagraphBox {
  direction?: TextDirection;
  align?: TextAlign;
  /** Authored line height as a multiple of the font size. */
  lineSpacing?: number;
  /** Explicit line height in points. */
  lineSpacingPt?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  /** Indents in inches. */
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndent?: number;
}

/** A paragraph's physical alignment: authored, else the side its direction starts on. */
export function paragraphAlign(box: Pick<ParagraphBox, "align" | "direction">): TextAlign {
  return box.align ?? (box.direction === "rtl" ? "right" : "left");
}

export interface ParagraphBoxOptions {
  /**
   * The CSS line-height multiple that an authored spacing of 1 stands for
   * (default 1). PowerPoint's 100% is 1.2em and Word's "single" about
   * 1.15em; each format passes its own, and the authored multiple scales it.
   */
  singleLine?: number;
  /** Paint the first-line indent. A bulleted PPTX paragraph gets its hanging separation from the bullet grid instead. */
  firstLineIndent?: boolean;
}

/**
 * Paint a paragraph's box: direction, alignment, line height, spacing and
 * indents. The authored line spacing stays on the element as data; only the
 * painted line box is floored at `MINIMUM_LINE_HEIGHT`, because CSS would crop
 * an Arabic glyph's ink in anything tighter.
 */
export function applyParagraphBox(line: HTMLElement, box: ParagraphBox, options: ParagraphBoxOptions = {}): void {
  const { singleLine = 1, firstLineIndent = true } = options;
  if (box.direction !== undefined) {
    line.style.direction = box.direction;
    line.dataset.aaredDirection = box.direction;
  }
  const align = paragraphAlign(box);
  line.style.textAlign = align;
  line.dataset.aaredAlign = align;
  if (box.lineSpacing !== undefined) {
    const multiple = singleLine === 1 ? box.lineSpacing : Number((box.lineSpacing * singleLine).toFixed(3));
    line.style.lineHeight = String(Math.max(multiple, MINIMUM_LINE_HEIGHT));
    line.dataset.aaredLineSpacing = String(box.lineSpacing);
  }
  if (box.lineSpacingPt !== undefined) line.style.lineHeight = `${box.lineSpacingPt}pt`;
  if (box.spaceBeforePt !== undefined) line.style.marginTop = `${box.spaceBeforePt}pt`;
  if (box.spaceAfterPt !== undefined) line.style.marginBottom = `${box.spaceAfterPt}pt`;
  if (box.leftIndent !== undefined) line.style.marginLeft = inchesToCss(box.leftIndent);
  if (box.rightIndent !== undefined) line.style.marginRight = inchesToCss(box.rightIndent);
  if (firstLineIndent && box.firstLineIndent !== undefined) line.style.textIndent = inchesToCss(box.firstLineIndent);
}
