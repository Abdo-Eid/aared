import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { geometryForPreset } from "../src/core/drawingml-shape.js";
import { encodeGeometry } from "../src/core/drawingml-program.js";
import { evaluateGeometry } from "../src/core/drawingml-engine.js";
import { parseXml } from "../src/core/xml.js";

const W = 539640;
const H = 251640;
const seed = parseXml(readFileSync(new URL("../scripts/presetShapeDefinitions.xml", import.meta.url), "utf8"));

function program(preset: string): string {
  const definition = seed.children.find((child) => child.local === preset);
  const encoded = definition === undefined ? undefined : encodeGeometry(definition);
  if (encoded === undefined) throw new Error(`seed lacks ${preset}`);
  return encoded;
}

/** Vertex list of a polygon path, whatever mix of M/L/H/V it is written in. */
function vertices(d: string): Array<[number, number]> {
  const tokens = d.split(" ");
  const out: Array<[number, number]> = [];
  let x = 0, y = 0;
  for (let i = 0; i < tokens.length;) {
    const verb = tokens[i++];
    if (verb === "M" || verb === "L") { x = Number(tokens[i++]); y = Number(tokens[i++]); }
    else if (verb === "H") x = Number(tokens[i++]);
    else if (verb === "V") y = Number(tokens[i++]);
    else if (verb === "Z") continue;
    else throw new Error(`not a polygon: ${verb}`);
    out.push([x, y]);
  }
  return out;
}

/** Same polygon, same winding; the start vertex may differ (ECMA's diamond starts left, ours top). */
function expectSameVertices(actual: string, expected: string): void {
  const a = vertices(actual), e = vertices(expected);
  expect(a).toHaveLength(e.length);
  // 1 EMU tolerance: ECMA's own sin-based guides (hexagon's dy1) land a
  // fraction of an EMU off exact — some 10,000x below a device pixel, while
  // a wrong formula basis misses by thousands.
  const near = (p: [number, number], q: [number, number]): boolean => Math.abs(p[0] - q[0]) < 1 && Math.abs(p[1] - q[1]) < 1;
  const offset = e.findIndex((vertex) => near(vertex, a[0]!));
  expect(offset).toBeGreaterThan(-1);
  a.forEach((vertex, index) => expect(near(vertex, e[(index + offset) % e.length]!)).toBe(true));
}

describe("geometry engine", () => {
  test("reproduces the hand-coded fast path from the ECMA definitions", () => {
    const cases: Array<[string, Record<string, number>]> = [
      ["rect", {}], ["triangle", {}], ["triangle", { adj: 120000 }], ["rtTriangle", {}], ["diamond", {}],
      ["parallelogram", {}], ["parallelogram", { adj: 200000 }], ["hexagon", {}], ["octagon", {}], ["star5", {}], ["line", {}],
    ];
    for (const [preset, adjustments] of cases) {
      for (const [width, height] of [[W, H], [H, W]] as const) {
        const native = geometryForPreset(preset, { width, height }, adjustments)!.path;
        expectSameVertices(evaluateGeometry(program(preset), width, height, adjustments)!.path, native);
      }
    }
  });

  test("draws roundRect corners as quarter arcs ending on the edges", () => {
    const radius = (H * 16667) / 100000;
    const path = evaluateGeometry(program("roundRect"), W, H)!.path;
    expect(path.startsWith(`M 0 ${Math.round(radius * 100) / 100} A`)).toBe(true);
    expect(path).toContain(`0 0 1 ${Math.round(radius * 100) / 100} 0 L`);
    expect(path.match(/ A /g)).toHaveLength(4);
  });

  test("authored adjustments override avLst defaults by name", () => {
    const narrow = evaluateGeometry(program("rightArrow"), W, H, { adj1: 20000 })!.path;
    const wide = evaluateGeometry(program("rightArrow"), W, H)!.path;
    expect(narrow).not.toBe(wide);
    expect(narrow).toContain(`M 0 ${H / 2 - (H * 20000) / 200000}`);
  });

  test("splits DrawingML paths into fill-only, outlined and stroke-only layers", () => {
    const can = evaluateGeometry(program("can"), 400000, 600000)!;
    expect(can.path).toBe("");
    expect(can.fillOnly!.match(/M /g)).toHaveLength(2);
    expect(can.strokeOnly!.match(/M /g)).toHaveLength(1);
    const frame = evaluateGeometry(program("frame"), W, H)!;
    expect(frame.fillOnly).toBeUndefined();
    expect(frame.path.match(/M /g)).toHaveLength(2);
  });

  test("scales custGeom paths from their own w/h space", () => {
    const custom = parseXml(
      `<custGeom><avLst/><gdLst/><pathLst><path w="100" h="50"><moveTo><pt x="0" y="0"/></moveTo><lnTo><pt x="100" y="25"/></lnTo><quadBezTo><pt x="50" y="50"/><pt x="0" y="50"/></quadBezTo><close/></path></pathLst></custGeom>`,
    );
    const encoded = encodeGeometry(custom)!;
    expect(encoded).toBe("p 100 50 - M 0 0 L 100 25 Q 50 50 0 50 Z");
    expect(evaluateGeometry(encoded, 1000, 200)!.path).toBe("M 0 0 L 1000 100 Q 500 200 0 200 Z");
  });

  test("encodes fill='none' and stroke='false' paths as layer flags", () => {
    const custom = parseXml(
      `<custGeom><pathLst><path fill="none"><moveTo><pt x="l" y="t"/></moveTo><lnTo><pt x="r" y="b"/></lnTo></path><path stroke="false" fill="darken"><moveTo><pt x="l" y="b"/></moveTo><lnTo><pt x="r" y="t"/></lnTo><close/></path></pathLst></custGeom>`,
    );
    expect(encodeGeometry(custom)).toBe("p - - f M l t L r b;p - - s M l b L r t Z");
  });

  test("resolves the built-in guides, including the parametric families", () => {
    const probe = (guide: string): number => Number(evaluateGeometry(`p - - - M ${guide} 0`, 3200, 1600)!.path.split(" ")[1]);
    expect(probe("wd32")).toBe(100);
    expect(probe("hd4")).toBe(400);
    expect(probe("ssd16")).toBe(100);
    expect(probe("ls")).toBe(3200);
    expect(probe("3cd4")).toBe(16200000);
    expect(probe("7cd8")).toBe(18900000);
  });

  test("implements every ECMA formula operator", () => {
    const run = (formula: string): number => Number(evaluateGeometry(`g v ${formula};p - - - M v 0`, 100, 100)!.path.split(" ")[1]);
    expect(run("*/ 6 4 3")).toBe(8);
    expect(run("+- 6 4 3")).toBe(7);
    expect(run("+/ 6 4 5")).toBe(2);
    expect(run("?: 1 7 9")).toBe(7);
    expect(run("?: 0 7 9")).toBe(9);
    expect(run("abs -5")).toBe(5);
    expect(run("at2 0 10")).toBe(5400000);
    expect(run("cat2 10 1 0")).toBe(10);
    expect(run("cos 10 5400000")).toBe(0);
    expect(run("max 3 9")).toBe(9);
    expect(run("min 3 9")).toBe(3);
    expect(run("mod 3 4 12")).toBe(13);
    expect(run("pin 0 150 100")).toBe(100);
    expect(run("pin 0 -5 100")).toBe(0);
    expect(run("sat2 10 0 1")).toBe(10);
    expect(run("sin 10 5400000")).toBe(10);
    expect(run("sqrt 81")).toBe(9);
    expect(run("tan 10 2700000")).toBe(10);
    expect(run("val 42")).toBe(42);
  });

  test("arcs set the large-arc flag from the sweep and split full turns", () => {
    const quarter = evaluateGeometry("p - - - M 100 50 A 50 50 0 5400000", 100, 100)!.path;
    expect(quarter).toBe("M 100 50 A 50 50 0 0 1 50 100");
    const threeQuarters = evaluateGeometry("p - - - M 100 50 A 50 50 0 16200000", 100, 100)!.path;
    expect(threeQuarters).toBe("M 100 50 A 50 50 0 1 1 50 0");
    const counter = evaluateGeometry("p - - - M 100 50 A 50 50 0 -5400000", 100, 100)!.path;
    expect(counter).toBe("M 100 50 A 50 50 0 0 0 50 0");
    // A 360° arc starts and ends on the same point, which SVG would drop.
    const circle = evaluateGeometry("p - - - M 100 50 A 50 50 0 21600000", 100, 100)!.path;
    expect(circle).toBe("M 100 50 A 50 50 0 0 1 0 50 A 50 50 0 0 1 100 50");
  });

  test("maps output through symmetries, flipping arc sweeps on reflection", () => {
    const quarter = "p - - - M 100 50 A 50 50 0 5400000";
    // Mirrored: left extreme down to the bottom, now counter-clockwise.
    expect(evaluateGeometry(quarter, 100, 100, {}, "x")!.path).toBe("M 0 50 A 50 50 0 0 0 50 100");
    expect(evaluateGeometry(quarter, 100, 100, {}, "xy")!.path).toBe("M 0 50 A 50 50 0 0 1 50 0");
    // Transposed: evaluated in the program's own 200×100 frame, then axes
    // swapped onto the 100×200 shape — radii swap and the sweep reverses.
    const stretched = "p - - - M w vc A wd2 hd2 0 5400000";
    expect(evaluateGeometry(stretched, 200, 100)!.path).toBe("M 200 50 A 100 50 0 0 1 100 100");
    expect(evaluateGeometry(stretched, 100, 200, {}, "t")!.path).toBe("M 50 200 A 50 100 0 0 0 100 100");
    // A symmetric shape is its own image: frame's corners are unchanged.
    const vertexSet = (d: string): string => d.split(" ").filter((token) => !/[A-Z]/.test(token)).join(",").split(",").reduce<string[]>((pairs, value, index, all) => (index % 2 === 0 ? [...pairs, `${value} ${all[index + 1]}`] : pairs), []).sort().join("|");
    expect(vertexSet(evaluateGeometry(program("frame"), W, H, {}, "xy")!.path)).toBe(vertexSet(evaluateGeometry(program("frame"), W, H)!.path));
  });

  test("uses visual, not parametric, arc angles on stretched ellipses", () => {
    // 45° visually on a 200×50 ellipse lands where the ray at 45° meets it.
    const d = evaluateGeometry("p - - - M 0 0 A 100 25 0 2700000", 1, 1)!.path.split(" ");
    const r = (100 * 25) / Math.hypot(25 * Math.SQRT1_2, 100 * Math.SQRT1_2);
    expect(Number(d[9])).toBeCloseTo(r * Math.SQRT1_2 - 100, 1);
    expect(Number(d[10])).toBeCloseTo(r * Math.SQRT1_2, 1);
  });

  test("returns undefined for anything it cannot draw faithfully", () => {
    expect(evaluateGeometry("g v nope 1 2;p - - - M v 0", 10, 10)).toBeUndefined();
    expect(evaluateGeometry("p - - - M missing 0", 10, 10)).toBeUndefined();
    expect(evaluateGeometry("p - - - M 0", 10, 10)).toBeUndefined();
    expect(evaluateGeometry("p - - - X 0 0", 10, 10)).toBeUndefined();
    expect(evaluateGeometry("g v */ 1 1 0;p - - - M v 0", 10, 10)).toBeUndefined();
    expect(evaluateGeometry("p - - fs M 0 0 L 1 1", 10, 10)).toBeUndefined();
    expect(encodeGeometry(parseXml(`<custGeom><pathLst><path><spline/></path></pathLst></custGeom>`))).toBeUndefined();
    expect(encodeGeometry(parseXml(`<custGeom><pathLst><path/></pathLst></custGeom>`))).toBeUndefined();
  });
});
