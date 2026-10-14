/** Browser-native filters over DOM/SVG content, including a whole group's silhouette. */
import type { DrawingEffectsModel, TransformModel } from "./core/model.js";
import { parseResolvedColor, toHexColor } from "./core/color.js";
import { PX_PER_INCH, PT_PER_INCH } from "./core/units.js";
import { svgElement } from "./svg-element.js";
import { resourceId } from "./resource-id.js";

const PX_PER_PT = PX_PER_INCH / PT_PER_INCH;

function shadowOffset(shadow: NonNullable<DrawingEffectsModel["outerShadow"]>, transform: TransformModel | undefined): [number, number] {
  let dx = shadow.dxPt * PX_PER_PT, dy = shadow.dyPt * PX_PER_PT;
  if (!shadow.rotatesWithShape) {
    const angle = -(transform?.rotation ?? 0) * Math.PI / 180;
    [dx, dy] = [dx * Math.cos(angle) - dy * Math.sin(angle), dx * Math.sin(angle) + dy * Math.cos(angle)];
    if (transform?.flipHorizontal) dx = -dx;
    if (transform?.flipVertical) dy = -dy;
  }
  return [dx, dy];
}

export function paintDrawingEffects(host: HTMLElement, effects: DrawingEffectsModel | undefined, definitions: HTMLElement, transform: TransformModel | undefined, cleanups: Array<() => void>): void {
  if (effects === undefined) return;
  if (effects.unsupported !== undefined) host.dataset.aaredEffectFallback = effects.unsupported.join(" ");
  if (effects.reflection !== undefined) {
    const r = effects.reflection;
    if (CSS.supports("-webkit-box-reflect", "below")) {
      // The mask is applied before reflection, so it runs upward: the
      // source's bottom becomes the reflection's near edge.
      host.style.setProperty("-webkit-box-reflect", `below ${r.distancePt * PX_PER_PT}px linear-gradient(to top, rgb(0 0 0 / ${r.startOpacity}) ${r.startPosition * 100}%, rgb(0 0 0 / ${r.endOpacity}) ${r.endPosition * 100}%, transparent ${r.endPosition * 100}%)`);
    } else {
      // Firefox has no box-reflect. Mirror browser DOM with a standard CSS
      // mask, excluded from selection, links, accessibility and search.
      const reflection = document.createElement("div");
      reflection.dataset.aaredElement = "reflection";
      reflection.setAttribute("aria-hidden", "true");
      reflection.inert = true;
      reflection.style.cssText = host.style.cssText;
      // Word's paint surfaces are in-flow/relative, unlike PPTX's absolute
      // wrappers. The decorative mirror must never become layout content.
      reflection.style.position = "absolute";
      reflection.style.left = host.style.left || "0";
      reflection.style.top = host.style.top || "0";
      reflection.style.width = host.style.width || "100%";
      reflection.style.height = host.style.height || "100%";
      reflection.style.filter = "";
      reflection.style.overflow = "visible";
      reflection.style.pointerEvents = "none";
      reflection.style.userSelect = "none";
      const layer = document.createElement("div");
      layer.style.position = "absolute";
      layer.style.top = `calc(100% + ${r.distancePt * PX_PER_PT}px)`;
      layer.style.width = "100%"; layer.style.height = "100%";
      layer.style.maskImage = `linear-gradient(to bottom, rgb(0 0 0 / ${r.startOpacity}) ${r.startPosition * 100}%, rgb(0 0 0 / ${r.endOpacity}) ${r.endPosition * 100}%, transparent ${r.endPosition * 100}%)`;
      reflection.append(layer); definitions.append(reflection);
      const refresh = (): void => {
        // Read layout-space lengths, never transformed bounding rectangles.
        // This also handles Word's auto-height surfaces and padded inline boxes.
        const style = getComputedStyle(host);
        reflection.style.left = host.style.left || `${host.offsetLeft}px`;
        reflection.style.top = host.style.top || `${host.offsetTop}px`;
        reflection.style.width = style.width;
        reflection.style.height = style.height;
        const copy = host.cloneNode(true) as HTMLElement;
        // References keep pointing at the original paint/filter definitions
        // in this root. Never duplicate their IDs in a reflected subtree.
        for (const element of [copy, ...copy.querySelectorAll("*")]) {
          element.removeAttribute("id");
          for (const name of element.getAttributeNames()) if (name.startsWith("data-aared-")) element.removeAttribute(name);
        }
        copy.style.left = "0"; copy.style.top = "0";
        copy.style.width = "100%"; copy.style.height = "100%";
        copy.style.transformOrigin = "center"; copy.style.transform = "scaleY(-1)";
        layer.replaceChildren(copy);
      };
      // Native picture loads and run updates continue to reach the mirror.
      const observer = new MutationObserver(refresh);
      observer.observe(host, { childList: true, subtree: true, attributes: true, characterData: true });
      const resize = new ResizeObserver(refresh);
      resize.observe(host);
      cleanups.push(() => { observer.disconnect(); resize.disconnect(); reflection.remove(); });
      queueMicrotask(refresh);
    }
  }
  if (effects.outerShadow === undefined && effects.innerShadow === undefined && effects.glow === undefined && effects.softEdgePt === undefined) return;
  if (effects.outerShadow !== undefined && effects.innerShadow === undefined && effects.glow === undefined && effects.softEdgePt === undefined) {
    // The browser already provides a silhouette shadow. Use its native
    // primitive for the common case, without a custom SVG filter graph.
    const s = effects.outerShadow, [dx, dy] = shadowOffset(s, transform);
    host.style.filter = `drop-shadow(${dx}px ${dy}px ${s.blurPt * PX_PER_PT / 2}px ${s.color})`;
    return;
  }
  const id = resourceId("effect");
  const svg = svgElement("svg", { width: 0, height: 0, "aria-hidden": "true" });
  svg.dataset.aaredElement = "effect-definitions";
  svg.style.position = "absolute";
  svg.style.pointerEvents = "none";
  const defs = svgElement("defs");
  const filter = svgElement("filter", { id, x: "-100%", y: "-100%", width: "300%", height: "300%", "color-interpolation-filters": "sRGB" });
  defs.append(filter); svg.append(defs); definitions.append(svg);
  const merge = (below: string, above: string, result: string): void => {
    const node = svgElement("feMerge", { result });
    node.append(svgElement("feMergeNode", { in: below }), svgElement("feMergeNode", { in: above }));
    filter.append(node);
  };
  const flood = (color: string, result: string): void => {
    const c = parseResolvedColor(color);
    filter.append(svgElement("feFlood", { "flood-color": c === undefined ? color : toHexColor([c.r, c.g, c.b]),
      "flood-opacity": c?.a ?? 1, result }));
  };
  let source = "SourceGraphic";
  if (effects.softEdgePt !== undefined && effects.softEdgePt > 0) {
    const radius = effects.softEdgePt * PX_PER_PT / 2;
    // Morphology was the costly path (8 soft-edge groups: ~58ms zoom
    // frames). A blurred alpha threshold feathers inward without erosion.
    filter.append(svgElement("feGaussianBlur", { in: "SourceAlpha", stdDeviation: radius, result: "edge-blur" }));
    const transfer = svgElement("feComponentTransfer", { in: "edge-blur", result: "edge-mask" });
    transfer.append(svgElement("feFuncA", { type: "linear", slope: 4, intercept: -3 }));
    filter.append(transfer, svgElement("feComposite", { in: source, in2: "edge-mask", operator: "in", result: "soft-source" }));
    source = "soft-source";
  }
  if (effects.glow !== undefined) {
    const radius = effects.glow.radiusPt * PX_PER_PT;
    // Strengthen blurred alpha rather than dilating it: the same visible
    // halo at normal use, without the morphology path's measured cost. Word
    // holds a glow near full strength for several points out and fades to
    // nothing by the authored radius (measured against Word 16.0 on the
    // native DOCX page), so blur wide, gain it, then bend the falloff with
    // a table into that plateau.
    filter.append(svgElement("feGaussianBlur", { in: "SourceAlpha", stdDeviation: radius * 0.6, result: "glow-blur" }));
    const gain = svgElement("feComponentTransfer", { in: "glow-blur", result: "glow-gain" });
    gain.append(svgElement("feFuncA", { type: "linear", slope: 1.25, intercept: 0 }));
    const transfer = svgElement("feComponentTransfer", { in: "glow-gain", result: "glow-mask" });
    transfer.append(svgElement("feFuncA", { type: "table", tableValues: "0 0.05 0.5 0.82 0.92 0.95 0.95 0.95 0.95 0.95 0.95 0.95 0.95 0.95 0.95 0.95" }));
    filter.append(gain, transfer);
    flood(effects.glow.color, "glow-color");
    filter.append(svgElement("feComposite", { in: "glow-color", in2: "glow-mask", operator: "in", result: "glow" }));
    merge("glow", source, "with-glow"); source = "with-glow";
  }
  if (effects.outerShadow !== undefined) {
    const s = effects.outerShadow, [dx, dy] = shadowOffset(s, transform);
    filter.append(svgElement("feGaussianBlur", { in: "SourceAlpha", stdDeviation: s.blurPt * PX_PER_PT / 2, result: "shadow-blur" }),
      svgElement("feOffset", { in: "shadow-blur", dx, dy, result: "shadow-mask" }));
    flood(s.color, "shadow-color");
    filter.append(svgElement("feComposite", { in: "shadow-color", in2: "shadow-mask", operator: "in", result: "shadow" }));
    merge("shadow", source, "with-shadow"); source = "with-shadow";
  }
  if (effects.innerShadow !== undefined) {
    const s = effects.innerShadow, [dx, dy] = shadowOffset(s, transform);
    // Office's positive inner-shadow offset darkens bottom/right; subtract
    // the oppositely shifted alpha to expose those interior edges.
    filter.append(svgElement("feOffset", { in: "SourceAlpha", dx: -dx, dy: -dy, result: "inner-offset" }),
      svgElement("feGaussianBlur", { in: "inner-offset", stdDeviation: s.blurPt * PX_PER_PT / 2, result: "inner-blur" }),
      svgElement("feComposite", { in: "SourceAlpha", in2: "inner-blur", operator: "out", result: "inner-mask" }));
    flood(s.color, "inner-color");
    filter.append(svgElement("feComposite", { in: "inner-color", in2: "inner-mask", operator: "in", result: "inner-shadow" }));
    merge(source, "inner-shadow", "with-inner-shadow");
  }
  host.style.filter = `url(#${id})`;
}
