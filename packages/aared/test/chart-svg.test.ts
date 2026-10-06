import { describe, expect, test } from "bun:test";
import { appendChartPrimitive } from "../src/chart-svg.js";
import type { ChartPaintContext } from "../src/chart-svg.js";
import type { ChartPathPrimitive, ChartPrimitive, ChartTextPrimitive } from "../src/core/chart/scene.js";
import { cssFontFamily, fontAlias } from "../src/core/typography.js";

// `appendChartPrimitive` needs a DOM. The unit suite runs headless, so this
// stand-in records the elements it is asked to create and their attributes;
// the real-DOM output of PPTX and XLSX charts is covered by their visual suites.
interface FakeNode {
  tag: string;
  attributes: Record<string, string>;
  textContent: string;
  children: FakeNode[];
  setAttribute(name: string, value: string): void;
  append(node: FakeNode): void;
}

function fakeNode(tag: string): FakeNode {
  return {
    tag,
    attributes: {},
    textContent: "",
    children: [],
    setAttribute(name, value) {
      this.attributes[name] = value;
    },
    append(node) {
      this.children.push(node);
    },
  };
}

/** Paint one primitive into a fake SVG; returns the SVG's children (gradient defs first, when any). */
function paint(primitive: ChartPrimitive, context?: ChartPaintContext): FakeNode[] {
  const original = (globalThis as { document?: unknown }).document;
  (globalThis as { document?: unknown }).document = { createElementNS: (_namespace: string, tag: string) => fakeNode(tag) };
  try {
    const svg = fakeNode("svg");
    const host = Object.assign(svg, {
      dataset: {} as Record<string, string>,
      querySelector: (tag: string) => svg.children.find((child) => child.tag === tag) ?? null,
      prepend: (node: FakeNode) => svg.children.unshift(node),
    });
    appendChartPrimitive(host as unknown as SVGSVGElement, primitive, context);
    return svg.children;
  } finally {
    (globalThis as { document?: unknown }).document = original;
  }
}

const text = (fields: Partial<ChartTextPrimitive> = {}): ChartTextPrimitive => ({
  kind: "text", x: 0, y: 0, text: "t", fontSize: 10, fontFamily: "Arial", fill: "#000000", textAnchor: "start", ...fields,
});

describe("chart primitives", () => {
  test("a path is filled and stroked from its own fields, dashes and joins included", () => {
    const [node] = paint({ kind: "path", path: "M0 0h1v1h-1Z", fill: "#2E7D74", stroke: { color: "#000", width: 2, dash: [4, 3], cap: "round", join: "bevel" } });
    expect(node!.tag).toBe("path");
    expect(node!.attributes).toMatchObject({
      d: "M0 0h1v1h-1Z", fill: "#2E7D74", stroke: "#000", "stroke-width": "2", "stroke-dasharray": "4 3", "stroke-linecap": "round", "stroke-linejoin": "bevel",
    });
  });

  test("an unfilled path stays unfilled instead of taking SVG's black default", () => {
    const [node] = paint({ kind: "path", path: "M0 0L1 1", stroke: { color: "#123", width: 4 } });
    expect(node!.attributes).toMatchObject({ d: "M0 0L1 1", fill: "none", stroke: "#123", "stroke-width": "4" });
    expect(node!.attributes["stroke-dasharray"]).toBeUndefined();
  });

  const stops = [{ position: 0, color: "#ffffff" }, { position: 1, color: "#000000" }];
  const gradient: ChartPathPrimitive = {
    kind: "path", path: "M10 20h30v40h-30Z", box: { x: 10, y: 20, width: 30, height: 40 },
    fill: { kind: "linear-gradient", angle: 90, scaled: false, rotatesWithShape: true, stops },
  };

  test("a gradient fill goes through the shared fill painter, in the box of the shape it paints", () => {
    const calls: Array<{ id: string; box: Record<string, number>; urls: unknown }> = [];
    const fills = {
      fillPaint: (_svg: unknown, id: string, _fill: unknown, box: Record<string, number>, options: { urls: unknown }) => {
        calls.push({ id, box, urls: options.urls });
        return `url(#${id})`;
      },
      fillFallbackReason: () => "never",
    } as unknown as ChartPaintContext["fills"];
    const [node] = paint(gradient, { fills });
    expect(calls).toHaveLength(1);
    // A point is 0.75 of a CSS pixel's worth of user units; the box starts where the shape does.
    expect(calls[0]!.box).toMatchObject({ width: 30, height: 40, unitsPerPxX: 0.75, unitsPerPxY: 0.75, slideX: 0, slideY: 0, originX: 10, originY: 20 });
    expect(node!.attributes.fill).toBe(`url(#${calls[0]!.id})`);
    expect(node!.attributes["data-aared-fill-fallback"]).toBeUndefined();
  });

  test("a fill the painter cannot paint stays unfilled and says why, never a substitute color", () => {
    const fills = { fillPaint: () => undefined, fillFallbackReason: () => "pattern" } as unknown as ChartPaintContext["fills"];
    const [node] = paint({ ...gradient, fill: { kind: "unsupported", source: "group" } }, { fills });
    expect(node!.attributes.fill).toBe("none");
    expect(node!.attributes["data-aared-fill-fallback"]).toBe("pattern");
    // Without the painter at all (its file did not arrive) the shape is marked too.
    const [bare] = paint(gradient);
    expect(bare!.attributes).toMatchObject({ fill: "none", "data-aared-fill-fallback": "painter-unavailable" });
  });

  test("text carries its font, anchor, rotation and content", () => {
    const [node] = paint(text({ x: 5, y: 6, text: "الاسم", fontSize: 20, fontWeight: "bold", fontStyle: "italic", textAnchor: "end", fill: "#5B6470", rotate: -45 }));
    expect(node!.textContent).toBe("الاسم");
    expect(node!.attributes).toMatchObject({
      x: "5", y: "6", "font-size": "20", "font-weight": "bold", "font-style": "italic", "text-anchor": "end", fill: "#5B6470", transform: "rotate(-45 5 6)",
    });
  });

  test("text paints its family through the same stack as every run: the family, its delivered substitute, a generic", () => {
    // A bare family the reader lacks (Aptos, the default theme's) fell to the SVG default serif.
    expect(paint(text({ fontFamily: "Aptos" }))[0]!.attributes["font-family"]).toBe(`"Aptos", "${fontAlias("Aptos")}", sans-serif`);
    expect(paint(text({ fontFamily: "Arial" }))[0]!.attributes["font-family"]).toBe(cssFontFamily("Arial"));
  });

  test("optional fields that are absent leave no attribute behind", () => {
    const [node] = paint(text());
    for (const name of ["font-weight", "font-style", "transform", "stroke", "direction"]) expect(node!.attributes[name]).toBeUndefined();
  });

  test("right-to-left text sets direction on its own element and keeps the engine's side of the anchor", () => {
    // Under direction=rtl, `start` is the right edge: the painter swaps start and end so the text sits
    // where the layout (which anchors in left-to-right terms) put it. Middle is the same both ways.
    const anchor = (textAnchor: "start" | "middle" | "end", direction?: "rtl"): string | undefined =>
      paint(text({ textAnchor, ...(direction === undefined ? {} : { direction }) }))[0]!.attributes["text-anchor"];
    expect(paint(text({ direction: "rtl" }))[0]!.attributes.direction).toBe("rtl");
    expect(anchor("start", "rtl")).toBe("end");
    expect(anchor("end", "rtl")).toBe("start");
    expect(anchor("middle", "rtl")).toBe("middle");
    expect(anchor("start")).toBe("start");
    expect(anchor("end")).toBe("end");
  });

  test("a line of Arabic mixed with Latin paints each run in its own face, as tspans of one text element", () => {
    // A font list would not do: Calibri and Arial both have Arabic glyphs, so the Arabic would take the Latin face.
    const [node] = paint(text({ fontFamily: "Calibri", text: "الربع 1", direction: "rtl", runs: [{ text: "الربع ", fontFamily: "Tahoma" }, { text: "1", fontFamily: "Calibri" }] }));
    expect(node!.tag).toBe("text");
    expect(node!.textContent).toBe("");
    expect(node!.children.map((part) => [part.tag, part.textContent, part.attributes["font-family"]])).toEqual([
      ["tspan", "الربع ", cssFontFamily("Tahoma")],
      ["tspan", "1", cssFontFamily("Calibri")],
    ]);
    expect(node!.attributes["font-family"]).toBe(cssFontFamily("Calibri"));
    expect(node!.attributes.direction).toBe("rtl");
  });
});
