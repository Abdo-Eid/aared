import { describe, expect, test } from "bun:test";
import type { DocxImageAnchor, DocxImageNode } from "../src/core/model.js";
import { horizontalOffsetPx, planFloat, polygonSpan } from "../src/docx/float-layout.js";
import type { AnchorFrame } from "../src/docx/float-layout.js";

/** A US Letter body at 96 dpi: a 6.5in column on an 8.5in page with one-inch margins. */
const LETTER: AnchorFrame = {
  columnPx: 624,
  columnLeftPx: 96,
  page: { widthPx: 816, heightPx: 1056, marginTopPx: 96, originXPx: 0, originYPx: 0 },
};
const FLOW: AnchorFrame = { columnPx: 624, columnLeftPx: 0 };

const NONE = { top: 0, bottom: 0, left: 0, right: 0 };

function picture(anchor: Partial<DocxImageAnchor>, size = 1.5): DocxImageNode {
  return {
    type: "image",
    mediaId: "word/media/x.png",
    contentType: "image/png",
    width: size,
    height: size,
    anchor: {
      wrap: "square",
      behind: false,
      wrapText: "bothSides",
      horizontal: { from: "column", align: "left" },
      vertical: { from: "paragraph", offset: 0 },
      distance: NONE,
      ...anchor,
    },
  };
}

describe("anchored picture placement", () => {
  test("a picture on a column edge wraps on that side, with text on the larger one", () => {
    expect(planFloat(picture({ horizontal: { from: "column", align: "left" } }), FLOW)).toEqual({ kind: "wrap", side: "left", edgePx: 0, topPx: 0 });
    expect(planFloat(picture({ horizontal: { from: "column", align: "right" } }), FLOW)).toEqual({ kind: "wrap", side: "right", edgePx: 0, topPx: 0 });
  });

  test("`wrapText` left or right forces the side text uses", () => {
    // Text only on the left of the picture: the picture rides the right edge.
    expect(planFloat(picture({ wrapText: "left", horizontal: { from: "column", align: "left" } }), FLOW)).toMatchObject({ kind: "wrap", side: "right" });
    expect(planFloat(picture({ wrapText: "right", horizontal: { from: "column", align: "right" } }), FLOW)).toMatchObject({ kind: "wrap", side: "left" });
  });

  test("an offset picture keeps its offset and wraps on its larger side (Word's `largest`)", () => {
    // 110pt from the column's left edge, as the probe's case 5.
    const plan = planFloat(picture({ wrapText: "largest", horizontal: { from: "column", offset: 110 / 72 } }), FLOW);
    expect(plan).toMatchObject({ kind: "wrap", side: "left" });
    if (plan.kind === "wrap") expect(plan.edgePx).toBeCloseTo((110 / 72) * 96, 1);
  });

  test("bothSides on a picture with room on both sides is a split: text down both sides, laid out line by line", () => {
    const plan = planFloat(picture({ horizontal: { from: "column", align: "center" } }), FLOW);
    expect(plan).toEqual({ kind: "split", leftPx: (624 - 144) / 2, topPx: 0 });
    // No room for text on one side (the picture against the column edge) is an ordinary one-sided float.
    expect(planFloat(picture({ horizontal: { from: "column", align: "left" }, distance: { top: 0, bottom: 0, left: 0.125, right: 0.125 } }), FLOW)).toMatchObject({ kind: "wrap", side: "left" });
  });

  test("top and bottom is always a band, and the vertical offset is kept", () => {
    const plan = planFloat(picture({ wrap: "topAndBottom", horizontal: { from: "column", align: "center" }, vertical: { from: "paragraph", offset: 30 / 72 } }), FLOW);
    expect(plan).toMatchObject({ kind: "band", topPx: 40 });
  });

  test("a picture with no wrap is an overlay over or under the text", () => {
    const behind = planFloat(picture({ wrap: "none", behind: true, horizontal: { from: "column", align: "center" }, vertical: { from: "paragraph", offset: 6 / 72 } }), FLOW);
    expect(behind).toEqual({ kind: "overlay", to: "paragraph", leftPx: (624 - 144) / 2, topPx: 8, behind: true });
    expect(planFloat(picture({ wrap: "none", behind: false }), FLOW)).toMatchObject({ kind: "overlay", behind: false });
  });

  test("an overlay measured from the paragraph takes the paragraph's indent back out", () => {
    const plan = planFloat(picture({ wrap: "none", horizontal: { from: "column", offset: 1 } }), FLOW, 24);
    expect(plan).toMatchObject({ kind: "overlay", to: "paragraph", leftPx: 96 - 24 });
  });

  // Ground truth: Word's own render of the probe `page.docx` (docx-float-wrap):
  // a 280pt picture centred on the page in both directions, behind the text,
  // painted with its centre at (408, 528) on the 816x1056 page.
  test("a picture centred on the page is placed from the page, not the column", () => {
    const watermark = picture({ wrap: "none", behind: true, horizontal: { from: "page", align: "center" }, vertical: { from: "page", align: "center" } }, 280 / 72);
    const plan = planFloat(watermark, LETTER);
    if (plan.kind !== "overlay") throw new Error("expected an overlay");
    expect(plan.to).toBe("page");
    const size = (280 / 72) * 96;
    expect(plan.leftPx + size / 2).toBeCloseTo(408, 1);
    expect(plan.topPx + size / 2).toBeCloseTo(528, 1);
  });

  // The probe's header logo: 30pt from the page's left and 20pt from its top,
  // inside a header positioned at the left margin and the header distance.
  test("a page-relative picture in a positioned header is translated by where the header sits", () => {
    const header: AnchorFrame = { ...LETTER, page: { ...LETTER.page!, originXPx: 96, originYPx: 48 } };
    const logo = picture({ wrap: "none", behind: true, horizontal: { from: "page", offset: 30 / 72 }, vertical: { from: "page", offset: 20 / 72 } }, 60 / 72);
    const plan = planFloat(logo, header);
    if (plan.kind !== "overlay") throw new Error("expected an overlay");
    // On the page the logo's left edge is at 40px and its top at 26.67px.
    expect(plan.leftPx + 96).toBeCloseTo(40, 1);
    expect(plan.topPx + 48).toBeCloseTo((20 / 72) * 96, 1);
  });

  test("in flow mode there is no page, so a page-relative picture is placed against the paragraph", () => {
    const plan = planFloat(picture({ wrap: "none", horizontal: { from: "page", offset: 1 }, vertical: { from: "page", offset: 1 } }), FLOW);
    expect(plan).toMatchObject({ kind: "overlay", to: "paragraph", leftPx: 96 });
  });

  test("the extent of a polygon over a band is what a line beside it keeps clear", () => {
    // A diamond: 0..100 wide and tall, widest (0..100) at y = 50.
    const diamond: Array<[number, number]> = [[50, 0], [100, 50], [50, 100], [0, 50]];
    expect(polygonSpan(diamond, 45, 55)).toEqual([0, 100]);
    // Near the top only the tip is there: the band 0..10 spans x from 40 to 60.
    const top = polygonSpan(diamond, 0, 10)!;
    expect(top[0]).toBeCloseTo(40, 6);
    expect(top[1]).toBeCloseTo(60, 6);
    // Above and below the polygon there is nothing to keep clear.
    expect(polygonSpan(diamond, -30, -10)).toBeUndefined();
    expect(polygonSpan(diamond, 110, 130)).toBeUndefined();
    // A flat edge inside the band counts, one outside does not.
    expect(polygonSpan([[10, 20], [90, 20], [50, 80]], 15, 25)).toEqual([10, 90]);
  });

  test("horizontal alignment inside and outside stand for the near and far edge", () => {
    expect(horizontalOffsetPx(picture({ horizontal: { from: "column", align: "inside" } }), FLOW)).toBe(0);
    expect(horizontalOffsetPx(picture({ horizontal: { from: "column", align: "outside" } }), FLOW)).toBe(624 - 144);
  });
});
