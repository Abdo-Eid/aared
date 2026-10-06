import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { layoutChart } from "../src/core/chart/draw.js";
import type { ChartTextPrimitive } from "../src/core/chart/scene.js";
import type { ChartModel } from "../src/core/model.js";
import { withPart } from "./mutate.js";

const PRODUCTION = `${import.meta.dir}/../../../fixtures/post-production-review.pptx`;

async function productionChart(): Promise<ChartModel> {
  const pkg = await OpcPackage.open(await Bun.file(PRODUCTION).arrayBuffer());
  const deck = await indexPptx(pkg);
  const slide = await parsePptxSlide(pkg, deck, 3);
  const chart = slide.elements.find((element): element is ChartModel => element.kind === "chart");
  if (chart === undefined) throw new Error("fixture structure changed: expected a chart on slide 4");
  return chart;
}

describe("PPTX charts the preview does not draw", () => {
  test("say why, as XLSX does: the reader's reason travels on the model", async () => {
    // The same deck with its column chart turned into a surface chart, a type the preview does not draw.
    const pkg = await withPart("post-production-review.pptx", "ppt/charts/chart1.xml", (xml) => xml.replaceAll("c:barChart>", "c:surfaceChart>"));
    const slide = await parsePptxSlide(pkg, await indexPptx(pkg), 3);
    const chart = slide.elements.find((element): element is ChartModel => element.kind === "chart")!;
    expect(chart.spec).toBeUndefined();
    expect(chart.unsupportedReason).toBe("chart type surfaceChart is not previewed");
  });
});

describe("PPTX charts", () => {
  test("the Worker sends the chart resolved, as plain data, never markup", async () => {
    const chart = await productionChart();
    expect(chart).toMatchObject({ id: "4", name: "Chart 0", sourceChartId: "ppt/charts/chart1.xml", bounds: { x: 0.6, y: 1.45 } });
    const spec = chart.spec!;
    const bars = spec.groups[0]!;
    expect(bars).toMatchObject({ kind: "bar", direction: "col", grouping: "clustered" });
    expect(bars.series.map((s) => [s.name, s.values, s.look.fill])).toEqual([
      ["Logged", [39, 27, 8], "#2E7D74"],
      ["Planned", [42, 40, 38], "#C9D6D3"],
    ]);
    expect(bars.series[0]!.categories).toEqual(["Ep 04", "Ep 05", "Ep 06"]);
    // Survives structured clone and JSON alike: no Map, Set or class instances.
    expect(JSON.parse(JSON.stringify(spec))).toEqual(spec);
    expect(JSON.stringify(spec)).not.toContain("<");
  });

  test("laid out at its frame it reads as PowerPoint draws it: axis 0-45 by 5, every value labelled", async () => {
    const chart = await productionChart();
    const drawing = layoutChart(chart.spec!, chart.bounds);
    // The view box is the frame in points.
    expect(drawing.width).toBeCloseTo(chart.bounds!.width * 72, 2);
    expect(drawing.height).toBeCloseTo(chart.bounds!.height * 72, 2);
    const texts = drawing.primitives.filter((p): p is ChartTextPrimitive => p.kind === "text").map((p) => p.text);
    expect(texts).toEqual(expect.arrayContaining(["0", "5", "10", "15", "20", "25", "30", "35", "40", "45"]));
    expect(texts).not.toContain("50");
    expect(texts).toEqual(expect.arrayContaining(["39", "27", "8", "42", "40", "38", "Ep 04", "Ep 05", "Ep 06", "Logged", "Planned"]));
  });
});
