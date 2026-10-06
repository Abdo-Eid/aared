/**
 * Where text lands: the box a text primitive occupies, estimated from the same
 * metrics the layout uses. The layout uses it to keep text inside the frame;
 * tests use it, with `pathBox` for paths, to prove nothing is drawn outside it.
 */

import { ascent, lineHeight, textWidth } from "./metrics.js";
import type { ChartTextPrimitive } from "./scene.js";

export interface Bounds { x0: number; y0: number; x1: number; y1: number }

export function textBox(p: ChartTextPrimitive): Bounds {
  const look = { size: p.fontSize, bold: p.fontWeight === "bold", italic: p.fontStyle === "italic", color: "", family: p.fontFamily };
  const w = textWidth(p.text, look);
  const lh = lineHeight(look);
  const asc = ascent(look);
  const left = p.textAnchor === "start" ? 0 : p.textAnchor === "middle" ? -w / 2 : -w;
  const corners: Array<[number, number]> = [[left, -asc], [left + w, -asc], [left + w, lh - asc], [left, lh - asc]];
  const rad = ((p.rotate ?? 0) * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const b: Bounds = { x0: Infinity, y0: Infinity, x1: -Infinity, y1: -Infinity };
  for (const [cx, cy] of corners) {
    const x = p.x + cx * cos - cy * sin;
    const y = p.y + cx * sin + cy * cos;
    b.x0 = Math.min(b.x0, x); b.x1 = Math.max(b.x1, x);
    b.y0 = Math.min(b.y0, y); b.y1 = Math.max(b.y1, y);
  }
  return b;
}
