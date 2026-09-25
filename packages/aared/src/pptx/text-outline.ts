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

const SVG_NS = "http://www.w3.org/2000/svg";
const DEFS_HOST_ID = "aared-hollow-text-defs";

/** Points to CSS pixels (96 DPI). */
function ptToPx(pt: number): number {
  return (pt * 4) / 3;
}

function svgElement<K extends keyof SVGElementTagNameMap>(name: K, attributes: Record<string, string>): SVGElementTagNameMap[K] {
  const element = document.createElementNS(SVG_NS, name);
  for (const [key, value] of Object.entries(attributes)) element.setAttribute(key, value);
  return element;
}

/** One hidden `<svg><defs>` shared by every hollow-text filter on the page, created lazily. */
function defsRoot(): SVGDefsElement {
  const existing = document.getElementById(DEFS_HOST_ID);
  if (existing instanceof SVGSVGElement) return existing.querySelector("defs")!;
  const host = svgElement("svg", { id: DEFS_HOST_ID, "aria-hidden": "true" });
  host.style.position = "absolute";
  host.style.width = "0";
  host.style.height = "0";
  host.style.overflow = "hidden";
  const defs = svgElement("defs", {});
  host.append(defs);
  document.body.append(host);
  return defs;
}

/** One `<filter>` per (colour, width) pair, reused across every run that needs it. */
const filters = new Map<string, string>();

/** CSS `filter` value that outlines an element's merged glyph silhouette in `color`, `widthPt` outside the glyph edge. */
export function hollowTextFilter(widthPt: number, color: string): string {
  const key = `${widthPt}|${color}`;
  let id = filters.get(key);
  if (id === undefined) {
    id = `aared-hollow-${filters.size}`;
    filters.set(key, id);
    const radius = String(Math.max(ptToPx(widthPt), 0.01));
    const filter = svgElement("filter", {
      id,
      x: "-20%", y: "-20%", width: "140%", height: "140%",
      "color-interpolation-filters": "sRGB",
    });
    filter.append(
      svgElement("feMorphology", { in: "SourceAlpha", operator: "dilate", radius, result: "dilated" }),
      svgElement("feComposite", { in: "dilated", in2: "SourceAlpha", operator: "xor", result: "ring" }),
      svgElement("feFlood", { "flood-color": color, "flood-opacity": "1", result: "color" }),
      svgElement("feComposite", { in: "color", in2: "ring", operator: "in" }),
    );
    defsRoot().append(filter);
  }
  return `url(#${id})`;
}
