import { describe, expect, test } from "bun:test";
import { appendChartPrimitive } from "../src/chart-svg.js";
import type { ChartGraphicPrimitive } from "../src/core/model.js";

// `appendChartPrimitive` needs a DOM. The unit suite runs headless, so this
// stand-in records the elements it is asked to create and their attributes;
// the real-DOM output of PPTX and XLSX charts is covered by their visual suites.
interface FakeNode {
  tag: string;
  attributes: Record<string, string>;
  textContent: string;
  setAttribute(name: string, value: string): void;
}

function paint(primitive: ChartGraphicPrimitive): FakeNode {
  const created: FakeNode[] = [];
  const original = (globalThis as { document?: unknown }).document;
  (globalThis as { document?: unknown }).document = {
    createElementNS: (_namespace: string, tag: string): FakeNode => {
      const node: FakeNode = {
        tag,
        attributes: {},
        textContent: "",
        setAttribute(name, value) {
          this.attributes[name] = value;
        },
      };
      created.push(node);
      return node;
    },
  };
  try {
    const appended: FakeNode[] = [];
    appendChartPrimitive({ append: (node: FakeNode) => appended.push(node) } as unknown as SVGSVGElement, primitive);
    expect(appended).toEqual(created);
    return appended[0]!;
  } finally {
    (globalThis as { document?: unknown }).document = original;
  }
}

describe("chart primitives", () => {
  test("a rect is placed and styled from its own fields", () => {
    const node = paint({ kind: "rect", x: 1, y: 2, width: 3, height: 4, fill: "#2E7D74", opacity: 0.5, stroke: { color: "#000", width: 2 } });
    expect(node.tag).toBe("rect");
    expect(node.attributes).toMatchObject({ x: "1", y: "2", width: "3", height: "4", fill: "#2E7D74", stroke: "#000", "stroke-width": "2", opacity: "0.5" });
  });

  test("an unfilled rect or path stays unfilled instead of taking SVG's black default", () => {
    expect(paint({ kind: "rect", x: 0, y: 0, width: 1, height: 1 }).attributes.fill).toBe("none");
    const path = paint({ kind: "path", path: "M 0 0 L 1 1", stroke: { color: "#123", width: 4 } });
    expect(path.attributes).toMatchObject({ d: "M 0 0 L 1 1", fill: "none", stroke: "#123", "stroke-width": "4" });
  });

  test("a line has a stroke and no fill, and follows currentColor when unstyled", () => {
    const styled = paint({ kind: "line", x1: 0, y1: 1, x2: 2, y2: 3, stroke: { color: "#888", width: 2 }, fill: "#ff0000" });
    expect(styled.attributes).toMatchObject({ x1: "0", y1: "1", x2: "2", y2: "3", stroke: "#888", "stroke-width": "2" });
    // A line is never filled: SpreadsheetML used to turn a stray fill into a 1px stroke, which no parser output relies on.
    expect(styled.attributes.fill).toBeUndefined();
    expect(paint({ kind: "line", x1: 0, y1: 0, x2: 1, y2: 1 }).attributes.stroke).toBe("currentColor");
  });

  test("text carries its font, anchor and content, and is black unless filled", () => {
    const node = paint({
      kind: "text", x: 5, y: 6, text: "الاسم", fontSize: 20, fontFamily: "Arial", fontWeight: "bold", textAnchor: "end", fill: "#5B6470",
    });
    expect(node.textContent).toBe("الاسم");
    expect(node.attributes).toMatchObject({ x: "5", y: "6", "font-size": "20", "font-family": "Arial", "font-weight": "bold", "text-anchor": "end", fill: "#5B6470" });
    expect(paint({ kind: "text", x: 0, y: 0, text: "t" }).attributes.fill).toBe("#000000");
  });

  test("optional fields that are absent leave no attribute behind", () => {
    const node = paint({ kind: "text", x: 0, y: 0, text: "t" });
    for (const name of ["font-size", "font-family", "font-weight", "text-anchor", "opacity", "stroke", "stroke-width"]) {
      expect(node.attributes[name]).toBeUndefined();
    }
  });
});
