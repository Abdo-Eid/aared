/**
 * Evaluates a geometry program (see `drawingml-program.ts`) to SVG path data
 * in the shape's local EMU space: the ECMA-376 §20.1.9 guide language with
 * its 17 formula operators and built-in guides, and the six path verbs.
 *
 * Loaded lazily on the main thread by the session, only when a slide carries
 * geometry the worker's fast path did not resolve. Decks made of rectangles,
 * text boxes and the other hand-coded presets never download it.
 */

export interface GeometryPaths {
  /** Subpaths both filled and stroked (the normal case). */
  path: string;
  /** Subpaths filled but not stroked (`stroke="false"`), painted beneath `path`. */
  fillOnly?: string;
  /** Subpaths stroked but not filled (`fill="none"`), painted above `path`. */
  strokeOnly?: string;
}

/** One DrawingML angle unit (1/60000 degree) in radians. */
const ANGLE = Math.PI / 10800000;
const FULL_TURN = 21600000;

function builtinGuide(name: string, w: number, h: number): number | undefined {
  switch (name) {
    case "l":
    case "t":
      return 0;
    case "r":
    case "w":
      return w;
    case "b":
    case "h":
      return h;
    case "hc":
      return w / 2;
    case "vc":
      return h / 2;
    case "ss":
      return Math.min(w, h);
    case "ls":
      return Math.max(w, h);
  }
  // wd2…wd32, hd2…hd10, ssd2…ssd32: the base divided by the suffix.
  const divided = /^(wd|hd|ssd)(\d+)$/.exec(name);
  if (divided !== null) return (divided[1] === "wd" ? w : divided[1] === "hd" ? h : Math.min(w, h)) / Number(divided[2]);
  // cd2, cd4, cd8, 3cd4, 3cd8, 5cd8, 7cd8: fractions of a full turn.
  const turn = /^(\d*)cd(\d+)$/.exec(name);
  if (turn !== null) return (Number(turn[1] || 1) * FULL_TURN) / Number(turn[2]);
  return undefined;
}

function formula(op: string, x: () => number, y: () => number, z: () => number): number {
  switch (op) {
    case "val": return x();
    case "*/": return (x() * y()) / z();
    case "+-": return x() + y() - z();
    case "+/": return (x() + y()) / z();
    case "?:": return x() > 0 ? y() : z();
    case "abs": return Math.abs(x());
    case "at2": return Math.atan2(y(), x()) / ANGLE;
    case "cat2": return x() * Math.cos(Math.atan2(z(), y()));
    case "cos": return x() * Math.cos(y() * ANGLE);
    case "max": return Math.max(x(), y());
    case "min": return Math.min(x(), y());
    case "mod": return Math.hypot(x(), y(), z());
    case "pin": {
      const low = x(), value = y(), high = z();
      return value < low ? low : value > high ? high : value;
    }
    case "sat2": return x() * Math.sin(Math.atan2(z(), y()));
    case "sin": return x() * Math.sin(y() * ANGLE);
    case "sqrt": return Math.sqrt(x());
    case "tan": return x() * Math.tan(y() * ANGLE);
  }
  throw new Error(`unknown guide operator ${op}`);
}

/** Two decimals of an EMU is far below a device pixel; keeps path data short. */
function fmt(value: number): string {
  if (!Number.isFinite(value)) throw new Error("non-finite coordinate");
  const rounded = Math.round(value * 100) / 100;
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

/**
 * The point of an ellipse centred at the origin at DrawingML angle `angle`.
 * DrawingML arc angles are visual (the direction of the ray from the
 * centre), not the ellipse parameter; the two agree only on the axes.
 */
function ellipsePoint(wR: number, hR: number, angle: number): { x: number; y: number; t: number } {
  const theta = angle * ANGLE;
  const t = Math.atan2(wR * Math.sin(theta), hR * Math.cos(theta));
  return { x: wR * Math.cos(t), y: hR * Math.sin(t), t };
}

/**
 * Evaluate a geometry program for a `width` × `height` shape (EMU). Authored
 * `adjustments` override the program's `avLst` defaults by name. A
 * `symmetry` is a symmetry of the square as catalog aliases spell it: any of
 * `t` (transpose, swap the axes), `x` (mirror left-right), `y` (mirror
 * top-bottom), applied in that order (`leftArrow` is `x` of `rightArrow`); the
 * program is evaluated in its own frame (axes swapped for `t`) and its output
 * mapped onto this shape. Returns `undefined` for
 * anything malformed or non-finite, so callers keep their bounds fallback
 * rather than painting a wrong outline.
 */
export function evaluateGeometry(
  program: string,
  width: number,
  height: number,
  adjustments: Readonly<Record<string, number>> = {},
  symmetry = "",
): GeometryPaths | undefined {
  try {
    return run(program, width, height, adjustments, symmetry);
  } catch {
    return undefined;
  }
}

function run(program: string, width: number, height: number, adjustments: Readonly<Record<string, number>>, symmetry: string): GeometryPaths | undefined {
  const transpose = symmetry.includes("t");
  const mirrorX = symmetry.includes("x");
  const mirrorY = symmetry.includes("y");
  // The program runs in its own frame; only emitted coordinates are mapped.
  const w = transpose ? height : width;
  const h = transpose ? width : height;
  const point = (px: number, py: number): string => {
    let x = transpose ? py : px;
    let y = transpose ? px : py;
    if (mirrorX) x = width - x;
    if (mirrorY) y = height - y;
    return `${fmt(x)} ${fmt(y)}`;
  };
  // Each reflection reverses orientation, so arc sweep directions flip.
  const reversed = (Number(transpose) + Number(mirrorX) + Number(mirrorY)) % 2 === 1;
  const guides = new Map<string, number>();
  const value = (token: string | undefined): number => {
    if (token === undefined) throw new Error("missing argument");
    const guide = guides.get(token);
    if (guide !== undefined) return guide;
    const literal = Number(token);
    if (token !== "" && Number.isFinite(literal)) return literal;
    const builtin = builtinGuide(token, w, h);
    if (builtin === undefined) throw new Error(`unknown guide ${token}`);
    return builtin;
  };
  const buckets = { both: [] as string[], fillOnly: [] as string[], strokeOnly: [] as string[] };

  for (const statement of program.split(";")) {
    const parts = statement.split(" ");
    const kind = parts[0];
    if (kind === "a") {
      const name = parts[1]!;
      guides.set(name, adjustments[name] ?? value(parts[2]));
    } else if (kind === "g") {
      const [, name, op, ...args] = parts;
      guides.set(name!, formula(op!, () => value(args[0]), () => value(args[1]), () => value(args[2])));
    } else if (kind === "p") {
      const [, pw, ph, flags, ...commands] = parts;
      const fills = !flags!.includes("f");
      const strokes = !flags!.includes("s");
      if (!fills && !strokes) continue;
      const sx = pw === "-" ? 1 : w / Number(pw);
      const sy = ph === "-" ? 1 : h / Number(ph);
      const out: string[] = [];
      let x = 0, y = 0, startX = 0, startY = 0;
      let i = 0;
      const X = (): number => value(commands[i++]) * sx;
      const Y = (): number => value(commands[i++]) * sy;
      while (i < commands.length) {
        const verb = commands[i++];
        if (verb === "M" || verb === "L") {
          x = X(); y = Y();
          if (verb === "M") { startX = x; startY = y; }
          out.push(`${verb} ${point(x, y)}`);
        } else if (verb === "Q") {
          const x1 = X(), y1 = Y();
          x = X(); y = Y();
          out.push(`Q ${point(x1, y1)} ${point(x, y)}`);
        } else if (verb === "C") {
          const x1 = X(), y1 = Y(), x2 = X(), y2 = Y();
          x = X(); y = Y();
          out.push(`C ${point(x1, y1)} ${point(x2, y2)} ${point(x, y)}`);
        } else if (verb === "A") {
          const wR = X(), hR = Y();
          const start = value(commands[i++]), sweep = value(commands[i++]);
          [x, y] = arc(out, x, y, wR, hR, start, sweep, point, transpose, reversed);
        } else if (verb === "Z") {
          out.push("Z");
          x = startX; y = startY;
        } else {
          throw new Error(`unknown path verb ${verb}`);
        }
      }
      if (out.length > 0) (fills && strokes ? buckets.both : fills ? buckets.fillOnly : buckets.strokeOnly).push(out.join(" "));
    } else {
      throw new Error(`unknown statement ${kind}`);
    }
  }
  if (buckets.both.length + buckets.fillOnly.length + buckets.strokeOnly.length === 0) return undefined;
  return {
    path: buckets.both.join(" "),
    ...(buckets.fillOnly.length === 0 ? {} : { fillOnly: buckets.fillOnly.join(" ") }),
    ...(buckets.strokeOnly.length === 0 ? {} : { strokeOnly: buckets.strokeOnly.join(" ") }),
  };
}

/**
 * Append an `arcTo` from the current point (program frame); returns the new
 * current point. `point` maps program-frame coordinates onto the shape.
 */
function arc(
  out: string[],
  x: number,
  y: number,
  wR: number,
  hR: number,
  start: number,
  sweep: number,
  point: (x: number, y: number) => string,
  transpose: boolean,
  reversed: boolean,
): [number, number] {
  if (sweep === 0) return [x, y];
  // SVG cannot draw a closed ellipse in one arc (start == end paints nothing).
  if (Math.abs(sweep) >= FULL_TURN) {
    const [mx, my] = arc(out, x, y, wR, hR, start, sweep / 2, point, transpose, reversed);
    return arc(out, mx, my, wR, hR, start + sweep / 2, sweep / 2, point, transpose, reversed);
  }
  const from = ellipsePoint(wR, hR, start);
  const to = ellipsePoint(wR, hR, start + sweep);
  const endX = x - from.x + to.x;
  const endY = y - from.y + to.y;
  if (wR === 0 || hR === 0) {
    out.push(`L ${point(endX, endY)}`);
    return [endX, endY];
  }
  // The large-arc flag depends on the parametric sweep, which can cross 180°
  // while the visual sweep does not (and vice versa) on a stretched ellipse.
  let parametric = to.t - from.t;
  const target = sweep * ANGLE;
  parametric += 2 * Math.PI * Math.round((target - parametric) / (2 * Math.PI));
  const large = Math.abs(parametric) > Math.PI ? 1 : 0;
  const [rx, ry] = transpose ? [hR, wR] : [wR, hR];
  out.push(`A ${fmt(Math.abs(rx))} ${fmt(Math.abs(ry))} 0 ${large} ${(sweep > 0) !== reversed ? 1 : 0} ${point(endX, endY)}`);
  return [endX, endY];
}
