/**
 * Where an anchored picture goes, worked out from its `wp:anchor` and the
 * geometry it sits in. Pure numbers in, a plan out, so the placement rules are
 * testable without a browser; `docx/index.ts` turns the plan into CSS.
 *
 * A CSS float wraps text on one side only: a line box is one interval between
 * the left and right floats. Word's `bothSides` for a picture with room on
 * each side puts text down both, so that case is a `split` plan, which
 * `split-wrap.ts` lays out itself, line by line, with browser text.
 */

import type { DocxDrawingNode, DocxImageAnchor } from "../core/model.js";
import { PX_PER_INCH } from "../core/units.js";

/** The box an anchored picture is placed in. Lengths in CSS pixels. */
export interface AnchorFrame {
  /** Text column width, and its left edge measured from the page's left edge. */
  columnPx: number;
  columnLeftPx: number;
  /**
   * Layout mode only. The page, the top margin, and where the nearest
   * positioned ancestor of the paragraph (the page, or a header) has its
   * top-left corner on the page.
   */
  page?: { widthPx: number; heightPx: number; marginTopPx: number; originXPx: number; originYPx: number };
}

export type FloatPlan =
  /** A CSS float on `side`; `edgePx` is the picture's distance from that column edge, `topPx` from the paragraph top. */
  | { kind: "wrap"; side: "left" | "right"; edgePx: number; topPx: number }
  /** A full-width band: text above `topPx` and below the picture, none beside it. `leftPx` is from the column's left edge. */
  | { kind: "band"; leftPx: number; topPx: number }
  /**
   * Text on both sides of the picture, line by line, as Word's `bothSides` does
   * with room on each side. CSS cannot wrap a line around both sides of one
   * obstacle, so the renderer splits each line beside the picture itself
   * (`split-wrap.ts`). `leftPx` is from the column's left edge.
   */
  | { kind: "split"; leftPx: number; topPx: number }
  /** Out of flow, over or under the text. `leftPx`/`topPx` are from the paragraph or, for `to: "page"`, from its positioned container. */
  | { kind: "overlay"; to: "paragraph" | "page"; leftPx: number; topPx: number; behind: boolean };

/** `relativeFrom` values that measure from the page or its margins rather than from the paragraph. */
const PAGE_EDGES = new Set(["page", "margin", "topMargin", "bottomMargin", "insideMargin", "outsideMargin"]);

const px = (inches: number | undefined): number => (inches ?? 0) * PX_PER_INCH;

/** The picture's left edge, in pixels from the text column's left edge. */
export function horizontalOffsetPx(image: DocxDrawingNode, frame: AnchorFrame): number {
  const { horizontal } = image.anchor!;
  const width = px(image.width);
  const fromPage = horizontal.from === "page" && frame.page !== undefined;
  const spanLeft = fromPage ? -frame.columnLeftPx : 0;
  const span = fromPage ? frame.page!.widthPx : frame.columnPx;
  switch (horizontal.align) {
    case "left":
    case "inside":
      return spanLeft;
    case "center":
      return spanLeft + (span - width) / 2;
    case "right":
    case "outside":
      return spanLeft + span - width;
    default:
      return spanLeft + px(horizontal.offset);
  }
}

/** The picture's top edge on the page, when its vertical position is measured from the page or a margin. */
function pageTopPx(image: DocxDrawingNode, frame: AnchorFrame): number {
  const { vertical } = image.anchor!;
  const page = frame.page!;
  const base = vertical.from === "page" ? 0 : page.marginTopPx;
  const height = px(image.height);
  switch (vertical.align) {
    case "center":
      return (page.heightPx - height) / 2;
    case "bottom":
      return page.heightPx - height;
    case "top":
    case "inside":
    case "outside":
      return base;
    default:
      return base + px(vertical.offset);
  }
}

/**
 * Plan one anchored picture. `paragraphIndentPx` is the anchor paragraph's left
 * indent, which an overlay measured from the paragraph has to take back out.
 */
export function planFloat(image: DocxDrawingNode, frame: AnchorFrame, paragraphIndentPx = 0): FloatPlan {
  const anchor: DocxImageAnchor = image.anchor!;
  const width = px(image.width);
  const x = horizontalOffsetPx(image, frame);
  // A picture text wraps around sits in the paragraph's flow, so its top is
  // measured from the paragraph. An offset from the page or a margin, or a
  // vertical alignment, would need the paragraph's place on the page, which
  // the flow does not know yet: those start at the paragraph top instead of
  // that many inches below it (a documented limit).
  const fromParagraph = anchor.vertical.from === "paragraph" || anchor.vertical.from === "line";
  const paragraphTop = fromParagraph && anchor.vertical.align === undefined ? Math.max(0, px(anchor.vertical.offset)) : 0;

  if (anchor.wrap === "none") {
    const toPage = frame.page !== undefined && PAGE_EDGES.has(anchor.vertical.from);
    if (toPage) {
      return {
        kind: "overlay",
        to: "page",
        leftPx: frame.columnLeftPx + x - frame.page!.originXPx,
        topPx: pageTopPx(image, frame) - frame.page!.originYPx,
        behind: anchor.behind,
      };
    }
    return { kind: "overlay", to: "paragraph", leftPx: x - paragraphIndentPx, topPx: anchor.vertical.offset === undefined ? 0 : px(anchor.vertical.offset), behind: anchor.behind };
  }

  if (anchor.wrap === "topAndBottom") return { kind: "band", leftPx: clampLeft(x, width, frame), topPx: paragraphTop };

  // square, tight and through: text on one side. Room on each side of the picture:
  const roomLeft = Math.max(0, x);
  const roomRight = Math.max(0, frame.columnPx - x - width);
  const largerIsRight = roomRight >= roomLeft;
  let side: "left" | "right";
  switch (anchor.wrapText) {
    // Text may only use the left of the picture: the picture rides the right edge.
    case "left":
      side = "right";
      break;
    case "right":
      side = "left";
      break;
    case "largest":
      side = largerIsRight ? "left" : "right";
      break;
    default:
      // bothSides with room on each side: text runs down both, split line by
      // line. A picture against a column edge has one side only, which a CSS
      // float does natively (and keeps wrapping the paragraphs after it).
      if (roomLeft > anchor.distance.left * PX_PER_INCH && roomRight > anchor.distance.right * PX_PER_INCH) {
        return { kind: "split", leftPx: x, topPx: paragraphTop };
      }
      side = largerIsRight ? "left" : "right";
  }
  return { kind: "wrap", side, edgePx: side === "left" ? roomLeft : roomRight, topPx: paragraphTop };
}

/**
 * The horizontal extent of a polygon over the band `y0..y1`: the smallest and
 * largest x any part of its outline reaches there, or undefined when the
 * polygon does not reach the band. What a line beside a tight or through wrap
 * has to keep clear (the hole is the extent, not the polygon's whole box).
 */
export function polygonSpan(points: ReadonlyArray<readonly [number, number]>, y0: number, y1: number): [number, number] | undefined {
  let min = Number.POSITIVE_INFINITY;
  let max = Number.NEGATIVE_INFINITY;
  for (let index = 0; index < points.length; index++) {
    const [ax, ay] = points[index]!;
    const [bx, by] = points[(index + 1) % points.length]!;
    // The part of this edge inside the band, by clipping its parameter range.
    let from = 0;
    let to = 1;
    if (ay === by) {
      if (ay < y0 || ay > y1) continue;
    } else {
      const tA = (y0 - ay) / (by - ay);
      const tB = (y1 - ay) / (by - ay);
      from = Math.max(from, Math.min(tA, tB));
      to = Math.min(to, Math.max(tA, tB));
      if (from > to) continue;
    }
    for (const t of [from, to]) {
      const x = ax + (bx - ax) * t;
      if (x < min) min = x;
      if (x > max) max = x;
    }
  }
  return min > max ? undefined : [min, max];
}

function clampLeft(x: number, width: number, frame: AnchorFrame): number {
  return Math.min(Math.max(0, x), Math.max(0, frame.columnPx - width));
}
