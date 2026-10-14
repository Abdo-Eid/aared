/**
 * WordArt outline text (`a:ln` on a run): PowerPoint traces the outline of
 * the run's whole *merged* glyph silhouette — Arabic glyphs that join or
 * overlap read as one continuous shape, with no seam where they touch.
 * `-webkit-text-stroke` instead strokes every glyph's own outline
 * separately, which leaves a visible line wherever two glyphs' outlines
 * cross (ground truth: `pptx-coverage.pptx` slide 6, hollow "تعبئة" —
 * PowerPoint draws one clean outline, a per-glyph stroke shows a seam at
 * every join).
 *
 * An SVG filter fixes it: filters operate on the element's whole rendered
 * (rasterized) alpha, so touching glyphs already read as one shape before
 * the outline is traced — unlike `-webkit-text-stroke`, which strokes each
 * glyph's vector outline before they are composited together.
 *
 * The ring is `dilate(alpha) XOR alpha` (an outside-only band, `dilate`'s
 * own edge back to the original edge), never `dilate(alpha) - erode(alpha)`
 * (a centered band). A centered band needs `feMorphology operator="erode"`,
 * and Chromium's `erode` silently no-ops whenever the filtered element sits
 * under an ancestor CSS `transform: scale()` — which this renderer's own
 * zoom/fit always applies — collapsing the ring into the full dilated shape
 * (a solid blob at any real radius, confirmed in isolation against a plain
 * filled glyph at three ancestor scales, including none). `dilate` alone
 * has no such problem at any of those scales, so the outline stays outside
 * the glyph's own edge rather than straddling it — a visible, deliberate
 * difference from a centered stroke, not a bug.
 */

import { svgElement } from "../svg-element.js";
import { ptToPx } from "../core/units.js";
import { resourceId } from "../resource-id.js";

/**
 * Outline `span`'s merged glyph silhouette in `color`, `widthPt` outside the
 * glyph edge. The `<filter>` goes inside the span itself: a `url(#id)` reference
 * resolves in the tree scope of the element it styles, so a filter kept on
 * `document.body` is not found from inside `<aared-viewer>`'s shadow root and the
 * run would paint as solid text. Each span owns its definition and name, even
 * when another render or another aared bundle uses the same outline width.
 */
export function hollowText(span: HTMLElement, widthPt: number, color: string): void {
  const id = resourceId("hollow");
  const filter = svgElement("filter", {
    id,
    x: "-20%", y: "-20%", width: "140%", height: "140%",
    "color-interpolation-filters": "sRGB",
  });
  filter.append(
    svgElement("feMorphology", { in: "SourceAlpha", operator: "dilate", radius: Math.max(ptToPx(widthPt), 0.01), result: "dilated" }),
    svgElement("feComposite", { in: "dilated", in2: "SourceAlpha", operator: "xor", result: "ring" }),
    svgElement("feFlood", { "flood-color": color, "flood-opacity": "1", result: "color" }),
    svgElement("feComposite", { in: "color", in2: "ring", operator: "in" }),
  );
  const defs = svgElement("defs");
  defs.append(filter);
  const host = svgElement("svg", { "aria-hidden": "true" });
  host.style.position = "absolute";
  host.style.width = "0";
  host.style.height = "0";
  host.style.overflow = "hidden";
  host.append(defs);
  span.append(host);
  span.style.filter = `url(#${id})`;
}
