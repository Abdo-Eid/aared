/**
 * Primitive emitter. Shapes that share one look and are drawn back-to-back
 * merge into a single path, so a series of a thousand columns is one
 * primitive rather than a thousand: the painted DOM stays small whatever the
 * point count, and paint order (which Office defines per series) is kept.
 *
 * Primitives are what the page paints (`chart-svg.ts`); they never leave the
 * page and are not part of the unit model. Coordinates are points, in a view
 * box equal to the chart frame.
 */

import { arabicFamily, hasArabicScript, isRtlText, scriptRuns } from "./metrics.js";
import type { Fill, Line, Look, TextLook } from "./spec.js";

export interface ChartStroke {
  color: string;
  width: number;
  /** Dash and gap lengths in points. */
  dash?: number[];
  cap?: "round" | "square";
  join?: "round" | "bevel";
}

/** A rectangle in the chart frame's points. */
export interface ChartBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface ChartPathPrimitive {
  kind: "path";
  /** SVG path commands only, never SVG markup. */
  path: string;
  fill?: Fill;
  /**
   * The bounding box of `path`. Set when `fill` is not a plain color: a
   * gradient, pattern or picture fill is laid out against the box of the shape
   * it paints, in the frame's coordinates.
   */
  box?: ChartBox;
  stroke?: ChartStroke;
}

export interface ChartTextPrimitive {
  kind: "text";
  /** Anchor point; `y` is the alphabetic baseline. */
  x: number;
  y: number;
  text: string;
  fontSize: number;
  fontFamily: string;
  /**
   * Set on a line that mixes Arabic with Latin letters or digits: the line cut by the face that
   * sets each part (Office sets Arabic in its complex-script face, the rest in `fontFamily`), in
   * reading order of the text. The runs' texts join back to `text`. A line of Arabic alone has no
   * runs; its complex-script face is `fontFamily`.
   */
  runs?: Array<{ text: string; fontFamily: string }>;
  fill: string;
  fontWeight?: "bold";
  fontStyle?: "italic";
  /** Which side of (x, y) the text sits on, on the page: always left-to-right terms, whatever `direction` says. */
  textAnchor: "start" | "middle" | "end";
  /** Set when the text's first strong character is right to left: the painter gives the element that base direction. */
  direction?: "rtl";
  /** Rotation about (x, y), degrees clockwise. */
  rotate?: number;
}

export type ChartPrimitive = ChartPathPrimitive | ChartTextPrimitive;

const r = (v: number): string => {
  const n = Math.round(v * 100) / 100;
  return Object.is(n, -0) ? "0" : String(n);
};

const round = (v: number): number => Math.round(v * 100) / 100;

function stroke(line: Line): ChartStroke {
  const out: ChartStroke = { color: line.color, width: round(line.width) };
  if (line.dash !== undefined) out.dash = line.dash.map(round);
  if (line.cap === "round" || line.cap === "square") out.cap = line.cap;
  if (line.join === "round" || line.join === "bevel") out.join = line.join;
  return out;
}

function lookKey(fill: Fill | undefined, line: Line | undefined): string {
  return JSON.stringify([fill ?? null, line ?? null]);
}

/**
 * The bounding box of path data this engine emits: absolute or relative
 * M L H V C Z, and circular arcs (A with equal radii, whose extremes are found
 * from the arc's centre; control points bound a curve, which is exact enough
 * to place a fill). Anything else is not emitted here.
 */
export function pathBox(d: string): ChartBox {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  const add = (x: number, y: number): void => {
    if (x < minX) minX = x;
    if (x > maxX) maxX = x;
    if (y < minY) minY = y;
    if (y > maxY) maxY = y;
  };
  const tokens = d.match(/[A-Za-z]|-?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/g) ?? [];
  let x = 0;
  let y = 0;
  let startX = 0;
  let startY = 0;
  let command = "";
  let i = 0;
  const next = (): number => Number(tokens[i++]);
  while (i < tokens.length) {
    if (/[A-Za-z]/.test(tokens[i]!)) command = tokens[i++]!;
    else if (command === "") break;
    const relative = command >= "a";
    const dx = relative ? x : 0;
    const dy = relative ? y : 0;
    switch (command.toUpperCase()) {
      case "M":
      case "L":
        x = next() + dx;
        y = next() + dy;
        add(x, y);
        if (command.toUpperCase() === "M") {
          startX = x;
          startY = y;
          // Further pairs after a move are implicit lines.
          command = relative ? "l" : "L";
        }
        break;
      case "H":
        x = next() + dx;
        add(x, y);
        break;
      case "V":
        y = next() + dy;
        add(x, y);
        break;
      case "C": {
        const cx1 = next() + dx;
        const cy1 = next() + dy;
        const cx2 = next() + dx;
        const cy2 = next() + dy;
        x = next() + dx;
        y = next() + dy;
        add(cx1, cy1);
        add(cx2, cy2);
        add(x, y);
        break;
      }
      case "A": {
        const radius = next();
        next();
        next();
        const large = next() !== 0;
        const sweep = next() !== 0;
        const x1 = next() + dx;
        const y1 = next() + dy;
        arcExtremes(x, y, x1, y1, radius, large, sweep, add);
        x = x1;
        y = y1;
        add(x, y);
        break;
      }
      case "Z":
        x = startX;
        y = startY;
        break;
      default:
        i = tokens.length;
    }
  }
  if (minX === Infinity) return { x: 0, y: 0, width: 0, height: 0 };
  return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
}

/** The points of a circular arc that sit on the circle's top, bottom, left or right, for `pathBox`. */
function arcExtremes(
  x0: number, y0: number, x1: number, y1: number, radius: number, large: boolean, sweep: boolean,
  add: (x: number, y: number) => void,
): void {
  const halfX = (x0 - x1) / 2;
  const halfY = (y0 - y1) / 2;
  const chord = halfX * halfX + halfY * halfY;
  if (chord === 0 || radius <= 0) return;
  // A chord longer than the diameter scales the radius up to fit, as SVG does.
  const r = Math.max(radius, Math.sqrt(chord));
  const coefficient = (large === sweep ? -1 : 1) * Math.sqrt(Math.max(0, (r * r - chord) / chord));
  const cx = coefficient * halfY + (x0 + x1) / 2;
  const cy = -coefficient * halfX + (y0 + y1) / 2;
  const start = Math.atan2(y0 - cy, x0 - cx);
  let delta = Math.atan2(y1 - cy, x1 - cx) - start;
  if (!sweep && delta > 0) delta -= 2 * Math.PI;
  else if (sweep && delta < 0) delta += 2 * Math.PI;
  for (let quarter = 0; quarter < 4; quarter++) {
    const angle = (quarter * Math.PI) / 2;
    // The quarter turn's offset from the start, in the arc's own direction.
    let offset = sweep ? angle - start : start - angle;
    offset = ((offset % (2 * Math.PI)) + 2 * Math.PI) % (2 * Math.PI);
    if (offset <= Math.abs(delta)) add(cx + r * Math.cos(angle), cy + r * Math.sin(angle));
  }
}

export class Scene {
  readonly primitives: ChartPrimitive[] = [];
  private openKey = "";
  private openLook: Look | undefined;
  private open: ChartPathPrimitive | undefined;

  /** Append path commands under `look`, merging with the previous shape when the look matches. */
  path(d: string, look: Look, mergeable = true, box?: ChartBox): void {
    if (look.fill === undefined && look.line === undefined) return;
    if (d === "") return;
    // Same look object as the open run: no need to compare by value.
    if (look !== this.openLook) {
      const key = lookKey(look.fill, look.line);
      if (key !== this.openKey) this.open = undefined;
      this.openKey = key;
      this.openLook = look;
    }
    // A gradient spans its own shape's box, so gradient shapes never merge.
    if (mergeable && this.open !== undefined && typeof look.fill !== "object") {
      this.open.path += ` ${d}`;
      return;
    }
    const primitive: ChartPathPrimitive = { kind: "path", path: d };
    if (look.fill !== undefined) primitive.fill = look.fill;
    if (typeof look.fill === "object") {
      const own = box ?? pathBox(d);
      primitive.box = { x: round(own.x), y: round(own.y), width: round(own.width), height: round(own.height) };
    }
    if (look.line !== undefined) primitive.stroke = stroke(look.line);
    this.primitives.push(primitive);
    this.open = primitive;
  }

  rect(x: number, y: number, w: number, h: number, look: Look, mergeable = true): void {
    if (w < 0) { x += w; w = -w; }
    if (h < 0) { y += h; h = -h; }
    this.path(`M${r(x)} ${r(y)}h${r(w)}v${r(h)}h${r(-w)}Z`, look, mergeable, { x, y, width: w, height: h });
  }

  /** A stroked polyline; `points` flat as x0,y0,x1,y1... */
  polyline(points: readonly number[], line: Line | undefined, closed = false): void {
    if (line === undefined || points.length < 4) return;
    let d = `M${r(points[0]!)} ${r(points[1]!)}`;
    for (let i = 2; i < points.length; i += 2) d += `L${r(points[i]!)} ${r(points[i + 1]!)}`;
    if (closed) d += "Z";
    this.path(d, { line });
  }

  segment(x1: number, y1: number, x2: number, y2: number, line: Line | undefined): void {
    if (line === undefined) return;
    this.path(`M${r(x1)} ${r(y1)}L${r(x2)} ${r(y2)}`, { line });
  }

  /**
   * One line of text; `y` is the baseline. `rtl` is its paragraph's direction when the line is one of several (a label's
   * "58%" under an Arabic category runs right to left, as "%58", with the rest of its label);
   * absent, the line's own first strong character decides.
   */
  text(x: number, y: number, text: string, look: TextLook, anchor: "start" | "middle" | "end", rotate = 0, rtl?: boolean): void {
    this.breakRun();
    if (text === "") return;
    const primitive: ChartTextPrimitive = {
      kind: "text", x: round(x), y: round(y), text,
      fontSize: round(look.size), fontFamily: look.family, fill: look.color, textAnchor: anchor,
    };
    if (hasArabicScript(text)) {
      const cs = arabicFamily(look);
      const runs = scriptRuns(text);
      if (runs.every((run) => run.arabic)) primitive.fontFamily = cs;
      else if (cs !== look.family && runs.some((run) => run.arabic)) {
        primitive.runs = runs.map((run) => ({ text: run.text, fontFamily: run.arabic ? cs : look.family }));
      }
    }
    if (rtl ?? isRtlText(text)) primitive.direction = "rtl";
    if (look.bold) primitive.fontWeight = "bold";
    if (look.italic) primitive.fontStyle = "italic";
    if (rotate !== 0) primitive.rotate = round(rotate);
    this.primitives.push(primitive);
  }

  /** End the current merge run (e.g. between series, so z-order follows series order). */
  breakRun(): void {
    this.open = undefined;
    this.openKey = "";
    this.openLook = undefined;
  }
}

export { r as fmt };
