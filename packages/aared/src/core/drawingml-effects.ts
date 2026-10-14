/** DrawingML's 2D effect list; pure resolved values, no DOM or retained XML. */
import { attr, firstChild, numberAttr, onOff } from "./xml.js";
import type { XmlElement } from "./xml.js";
import { resolveDrawingColor } from "./theme.js";
import type { DrawingTheme } from "./theme.js";
import type { DrawingEffectsModel } from "./model.js";
import { emuToPt } from "./units.js";

export function readDrawingEffects(properties: XmlElement | undefined, theme: DrawingTheme): DrawingEffectsModel | undefined {
  if (properties === undefined) return undefined;
  const list = firstChild(properties, "effectLst");
  const dag = firstChild(properties, "effectDag");
  if (list === undefined && dag === undefined) return undefined;
  const out: DrawingEffectsModel = {};
  const unsupported = (source: string): void => { (out.unsupported ??= []).push(source); };
  // A DAG can reference/reorder effects; flattening it changes composition.
  if (dag !== undefined && dag.children.length > 0) unsupported("effectDag");
  for (const effect of list?.children ?? []) {
    const point = (name: string, fallback = 0): number => emuToPt(numberAttr(effect, name, `a:${effect.local}`) ?? fallback);
    const ratio = (name: string, fallback: number): number => Math.min(1, Math.max(0, (numberAttr(effect, name, `a:${effect.local}`) ?? fallback) / 100000));
    switch (effect.local) {
      case "outerShdw":
      case "innerShdw": {
        const color = resolveDrawingColor(effect, theme);
        if (color === undefined) { unsupported(effect.local); break; }
        // Native PowerPoint offsets 8pt/6pt write dist=127000, dir=2212190.
        const angle = (numberAttr(effect, "dir", `a:${effect.local}`) ?? 0) / 60000 * Math.PI / 180;
        const shadow = { color, blurPt: point("blurRad"), dxPt: point("dist") * Math.cos(angle),
          dyPt: point("dist") * Math.sin(angle), rotatesWithShape: onOff(effect, "rotWithShape") ?? true };
        // Perspective shadows need a transformed silhouette, not an offset.
        if (["sx", "sy"].some((name) => (numberAttr(effect, name, `a:${effect.local}`) ?? 100000) !== 100000)
          || ["kx", "ky"].some((name) => (numberAttr(effect, name, `a:${effect.local}`) ?? 0) !== 0)) unsupported("shadow-transform");
        else if (effect.local === "outerShdw") out.outerShadow = shadow;
        else out.innerShadow = shadow;
        break;
      }
      case "glow": {
        const color = resolveDrawingColor(effect, theme);
        if (color === undefined) unsupported("glow");
        else out.glow = { color, radiusPt: point("rad") };
        break;
      }
      case "softEdge": out.softEdgePt = point("rad"); break;
      case "reflection": {
        const sy = numberAttr(effect, "sy", "a:reflection") ?? 100000;
        if (sy !== -100000 || (numberAttr(effect, "sx", "a:reflection") ?? 100000) !== 100000
          || (numberAttr(effect, "dir", "a:reflection") ?? 5400000) !== 5400000
          || ["kx", "ky"].some((name) => (numberAttr(effect, name, "a:reflection") ?? 0) !== 0)) {
          unsupported("reflection-transform"); break;
        }
        out.reflection = { distancePt: point("dist"), startOpacity: ratio("stA", 100000), endOpacity: ratio("endA", 0),
          startPosition: ratio("stPos", 0), endPosition: ratio("endPos", 100000) };
        break;
      }
      default: unsupported(effect.local);
    }
  }
  return Object.keys(out).length === 0 ? undefined : out;
}
