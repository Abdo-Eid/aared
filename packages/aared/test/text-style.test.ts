import { describe, expect, test } from "bun:test";
import type { TextRunModel } from "../src/core/model.js";
import type { ParagraphBox, ParagraphBoxOptions } from "../src/text-style.js";
import { applyParagraphBox, applyRunBase, paragraphAlign, pt, roundPt } from "../src/text-style.js";

// A stand-in element records the CSS and data attributes written to it; the
// real-DOM output of both formats is covered by their visual suites.
interface FakeElement {
  style: Record<string, string>;
  dataset: Record<string, string>;
}

const element = (): FakeElement => ({ style: {}, dataset: {} });

function applyRun(span: FakeElement, run: TextRunModel): void {
  applyRunBase(span as unknown as HTMLElement, run);
}

function applyBox(line: FakeElement, box: ParagraphBox, options?: ParagraphBoxOptions): void {
  applyParagraphBox(line as unknown as HTMLElement, box, options);
}

describe("point values", () => {
  test("float noise is trimmed to thousandths and never reaches a style", () => {
    expect(roundPt(0.1 + 0.2)).toBe(0.3);
    expect(pt(0.1 + 0.2)).toBe("0.3pt");
    expect(pt(1.5)).toBe("1.5pt");
    expect(pt(6.96)).toBe("6.96pt");
  });
});

describe("run formatting", () => {
  test("writes weight, style, size, colour, family, caps, spacing and highlight", () => {
    const span = element();
    applyRun(span, {
      text: "x", bold: true, italic: true, fontSizePt: 10.5, color: "#1B2029", fontFamily: "Arial",
      caps: "all", letterSpacingPt: 1.5, highlight: "#FFFF00",
    });
    // The face, then its nearest generic family, so a missing Arial never drops to the browser's serif default.
    expect(span.style).toEqual({
      fontWeight: "700", fontStyle: "italic", fontSize: "10.5pt", color: "#1B2029", fontFamily: "\"Arial\", sans-serif",
      textTransform: "uppercase", letterSpacing: "1.5pt", backgroundColor: "#FFFF00",
    });
  });

  test("underline and strikethrough share one text-decoration, and small caps is its own variant", () => {
    const both = element();
    applyRun(both, { text: "x", underline: true, strikethrough: true });
    expect(both.style.textDecoration).toBe("underline line-through");
    const struck = element();
    applyRun(struck, { text: "x", strikethrough: true, caps: "small" });
    expect(struck.style).toEqual({ textDecoration: "line-through", fontVariantCaps: "small-caps" });
  });

  test("outline-only text keeps its colour unset for the caller to paint", () => {
    const span = element();
    applyRun(span, { text: "x", color: "#FF0000", noFill: true });
    expect(span.style.color).toBeUndefined();
  });

  test("a run that authors nothing writes nothing", () => {
    const span = element();
    applyRun(span, { text: "plain" });
    expect(span.style).toEqual({});
  });

  test("character spacing is trimmed like every other point value", () => {
    const span = element();
    applyRun(span, { text: "x", letterSpacingPt: 0.1 + 0.2 });
    expect(span.style.letterSpacing).toBe("0.3pt");
  });
});

describe("paragraph box", () => {
  test("an unset alignment starts on the side of the direction", () => {
    expect(paragraphAlign({})).toBe("left");
    expect(paragraphAlign({ direction: "ltr" })).toBe("left");
    expect(paragraphAlign({ direction: "rtl" })).toBe("right");
    expect(paragraphAlign({ direction: "rtl", align: "center" })).toBe("center");
  });

  test("writes direction, alignment, spacing and indents, and marks direction and alignment as data", () => {
    const line = element();
    applyBox(line, {
      direction: "rtl", spaceBeforePt: 6, spaceAfterPt: 8, leftIndent: 0.25, rightIndent: 0.5, firstLineIndent: -0.25,
    });
    expect(line.style).toEqual({
      direction: "rtl", textAlign: "right", marginTop: "6pt", marginBottom: "8pt",
      marginLeft: "24px", marginRight: "48px", textIndent: "-24px",
    });
    expect(line.dataset).toEqual({ aaredDirection: "rtl", aaredAlign: "right" });
  });

  test("a paragraph with no direction leaves the direction to its context", () => {
    const line = element();
    applyBox(line, {});
    expect(line.style.direction).toBeUndefined();
    expect(line.dataset.aaredDirection).toBeUndefined();
    expect(line.style.textAlign).toBe("left");
  });

  test("DrawingML spacing is a multiple of the font size; Word's scales its 1.15em single line", () => {
    const drawing = element();
    applyBox(drawing, { lineSpacing: 1.2 });
    expect(drawing.style.lineHeight).toBe("1.2");
    const word = element();
    applyBox(word, { lineSpacing: 1.25 }, { singleLine: 1.15 });
    // 1.25 x 1.15 = 1.4375, painted to three decimals as the DOCX renderer always did.
    expect(word.style.lineHeight).toBe("1.438");
    // The authored value stays on the element as data, whatever the painted box is.
    expect(drawing.dataset.aaredLineSpacing).toBe("1.2");
    expect(word.dataset.aaredLineSpacing).toBe("1.25");
  });

  test("the painted line box is floored so Arabic ink is not cropped, and the data keeps the authored value", () => {
    const tight = element();
    applyBox(tight, { lineSpacing: 0.8 });
    expect(tight.style.lineHeight).toBe("1.1");
    expect(tight.dataset.aaredLineSpacing).toBe("0.8");
    const tightWord = element();
    applyBox(tightWord, { lineSpacing: 0.9 }, { singleLine: 1.15 });
    expect(tightWord.style.lineHeight).toBe("1.1");
  });

  test("an explicit point line height wins over a multiple", () => {
    const line = element();
    applyBox(line, { lineSpacing: 1.5, lineSpacingPt: 14 });
    expect(line.style.lineHeight).toBe("14pt");
  });

  test("a bulleted PPTX paragraph skips the first-line indent that the bullet grid supplies", () => {
    const bulleted = element();
    applyBox(bulleted, { firstLineIndent: 0.5, leftIndent: 0.5 }, { firstLineIndent: false });
    expect(bulleted.style.textIndent).toBeUndefined();
    expect(bulleted.style.marginLeft).toBe("48px");
  });
});
