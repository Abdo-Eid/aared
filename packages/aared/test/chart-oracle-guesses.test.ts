import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { layoutChart } from "../src/core/chart/draw.js";
import type { ChartDrawing } from "../src/core/chart/draw.js";
import type { ChartPathPrimitive, ChartTextPrimitive } from "../src/core/chart/scene.js";
import type { ChartSpecModel } from "../src/core/chart/spec.js";
import type { ChartModel } from "../src/core/model.js";

// The rules the chart engine once assumed, each now checked against PowerPoint on the chart oracle's
// `g-*` cases (tools/chart-oracle: cases.json, chart-oracle.pptx, and PowerPoint's PDF of it).
// The numbers asserted are what PowerPoint drew. DECISION_LOG §74.

const ORACLE = `${import.meta.dir}/../../../tools/chart-oracle`;
const cases = (await Bun.file(`${ORACLE}/cases.json`).json()) as { cases: Array<{ id: string }> };
const pkg = await OpcPackage.open(await Bun.file(`${ORACLE}/chart-oracle.pptx`).arrayBuffer());
const deck = await indexPptx(pkg);

async function oracle(id: string): Promise<{ spec: ChartSpecModel; drawing: ChartDrawing }> {
  const index = cases.cases.findIndex((c) => c.id === id);
  if (index < 0) throw new Error(`oracle corpus changed: no case ${id}`);
  const slide = await parsePptxSlide(pkg, deck, index);
  const chart = slide.elements.find((e): e is ChartModel => e.kind === "chart");
  if (chart?.spec === undefined) throw new Error(`oracle corpus changed: case ${id} has no chart`);
  return { spec: chart.spec, drawing: layoutChart(chart.spec, chart.bounds) };
}

const texts = (g: ChartDrawing) => g.primitives.filter((p): p is ChartTextPrimitive => p.kind === "text");
const paths = (g: ChartDrawing) => g.primitives.filter((p): p is ChartPathPrimitive => p.kind === "path");

describe("unstyled charts take Office's defaults", () => {
  test("axes: black lines, tick marks crossing the axis, minor ticks on the value axis (g-err-stddev)", async () => {
    const { spec } = await oracle("g-err-stddev");
    for (const axis of spec.axes) {
      expect(axis.majorTick).toBe("cross");
      expect(axis.minorTick).toBe("cross");
      expect(axis.line?.color).toBe("#000000");
    }
  });

  test("error bars are thin and black, centred on the mean for standard deviations (g-err-stddev)", async () => {
    const { spec } = await oracle("g-err-stddev");
    expect(spec.groups[0]!.series[0]!.errBars![0]!.line).toEqual({ color: "#000000", width: 0.75 });
  });

  test("a trendline is thin, solid and black (g-trend-linear-bare)", async () => {
    const { spec } = await oracle("g-trend-linear-bare");
    expect(spec.groups[0]!.series[0]!.trendlines![0]!.line).toEqual({ color: "#000000", width: 0.75 });
  });

  test("a stock chart's series take automatic markers; down bars are dark grey, up bars white (g-stock-bare)", async () => {
    const { spec } = await oracle("g-stock-bare");
    const group = spec.groups[0]!;
    expect(group.series.every((s) => s.marker.symbol !== "none")).toBe(true);
    expect(group.upDownBars!.down.fill).toBe("#404040");
    expect(group.upDownBars!.up.fill).toBe("#FFFFFF");
    expect(group.hiLowLines).toEqual({ color: "#000000", width: 0.75 });
  });
});

describe("value axes", () => {
  test("reach the error bars' ends: 9 +- 10% runs the axis to 12, as PowerPoint does (g-err-fixed-pct)", async () => {
    const { drawing } = await oracle("g-err-fixed-pct");
    expect(texts(drawing).map((t) => t.text)).toContain("12");
  });
});

describe("pie labels", () => {
  test("a label dragged off the pie gets a black leader line; one still touching it gets none; slices keep their colours (g-pie-leaders)", async () => {
    const { drawing } = await oracle("g-pie-leaders");
    const fills = new Set(paths(drawing).filter((p) => p.path.includes("A") && typeof p.fill === "string").map((p) => p.fill));
    expect(fills.size).toBe(4);
    // 8.2, 1.4 and 1.2 sit clear of the pie; 3.2 still overlaps its slice.
    const leaders = paths(drawing).filter((p) => p.stroke?.color === "#000000" && p.fill === undefined && /^M[^M]*L[^M]*L[^M]*$/.test(p.path));
    expect(leaders).toHaveLength(3);
  });

  test("a label wraps at a fifth of the plot; one too wide for its slice goes outside (g-pie-bestfit-long)", async () => {
    const { drawing } = await oracle("g-pie-bestfit-long");
    const lines = texts(drawing).map((t) => t.text);
    expect(lines).toContain("Southern Europe");
    expect(lines).toContain("North Africa and");
    expect(lines).toContain("Northern Europe and");
  });

  test("inside its slice an Arabic name wraps to the slice's width, a shorter one stays whole (ar-pie-labels)", async () => {
    const { drawing } = await oracle("ar-pie-labels");
    const lines = texts(drawing).map((t) => t.text);
    expect(lines).toContain("آسيا والمحيط");
    expect(lines).toContain("أمريكا الشمالية");
  });
});

describe("date axes", () => {
  test("label every day while they can, turned 45 degrees when flat ones would crowd (g-date-days)", async () => {
    const { drawing } = await oracle("g-date-days");
    const days = texts(drawing).filter((t) => /^1\/\d+\/2024$/.test(t.text));
    expect(days).toHaveLength(10);
    expect(days.every((t) => t.rotate === -45)).toBe(true);
  });

  test("with no date format show the serial numbers, as PowerPoint does (g-date-general)", async () => {
    const { drawing } = await oracle("g-date-general");
    expect(texts(drawing).map((t) => t.text)).toContain("45292");
  });
});

describe("trendline labels", () => {
  test("General is Office's rounding, and the block's lines are centred, its right edge at the line's end (g-trend-poly-eq)", async () => {
    const { drawing } = await oracle("g-trend-poly-eq");
    const eq = texts(drawing).find((t) => t.text.startsWith("y ="))!;
    const r2 = texts(drawing).find((t) => t.text.startsWith("R"))!;
    expect(eq.text).toBe("y = 1.3214x² - 9.25x + 17");
    expect(r2.text).toBe("R² = 0.9978");
    expect(eq.textAnchor).toBe("middle");
    expect(r2.x).toBeCloseTo(eq.x, 6);
  });
});
