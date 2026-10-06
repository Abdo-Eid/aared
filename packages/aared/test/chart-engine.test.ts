import { describe, expect, test } from "bun:test";
import { parseXml } from "../src/core/xml.js";
import { readChartModel } from "../src/core/chart.js";
import { readChart } from "../src/core/chart/read.js";
import { drawChart } from "../src/core/chart/engine.js";
import { chartFontProbes, layoutChart } from "../src/core/chart/draw.js";
import type { ChartDrawing } from "../src/core/chart/draw.js";
import type { ChartPathPrimitive, ChartTextPrimitive } from "../src/core/chart/scene.js";
import { niceStep, valueScale } from "../src/core/chart/scale.js";
import { textBox } from "../src/core/chart/bounds.js";
import { pathBox } from "../src/core/chart/scene.js";
import type { Bounds } from "../src/core/chart/bounds.js";
import { errorAmounts, fitTrend, movingAverage } from "../src/core/chart/trend.js";
import { scriptRuns, textWidth, wrapText } from "../src/core/chart/metrics.js";
import { themeForColorMap } from "../src/core/theme.js";
import { ARABIC_PROBE } from "../src/core/chars.js";

// Office's default theme under the default colour map, as XLSX charts get it.
const theme = themeForColorMap({
  colors: { dk1: "#000000", lt1: "#FFFFFF", dk2: "#44546A", lt2: "#E7E6E6", accent1: "#4472C4", accent2: "#ED7D31", accent3: "#A5A5A5", accent4: "#FFC000", accent5: "#5B9BD5", accent6: "#70AD47" },
  fonts: { major: { latin: "Calibri Light", complexScript: "", scripts: {} }, minor: { latin: "Calibri", complexScript: "", scripts: {} } },
}, undefined);

const NS = 'xmlns:c="c" xmlns:a="a" xmlns:c15="c15"';
const cache = (kind: "str" | "num", values: Array<string | number | null>, fmt = "General"): string =>
  kind === "str"
    ? `<c:strRef><c:strCache><c:ptCount val="${values.length}"/>${values.map((v, i) => (v === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join("")}</c:strCache></c:strRef>`
    : `<c:numRef><c:numCache><c:formatCode>${fmt}</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => (v === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join("")}</c:numCache></c:numRef>`;

function ser(i: number, name: string, cats: string[], vals: Array<number | null>, extra = ""): string {
  return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${cache("str", [name])}</c:tx>${extra}<c:cat>${cache("str", cats)}</c:cat><c:val>${cache("num", vals)}</c:val></c:ser>`;
}

const axes = (cat = "", val = "<c:majorGridlines/>") =>
  `<c:catAx><c:axId val="1"/><c:axPos val="b"/>${cat}<c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:axPos val="l"/>${val}<c:crossAx val="1"/></c:valAx>`;

function chart(plot: string, rest = ""): string {
  return `<c:chartSpace ${NS}><c:chart>${rest}<c:plotArea><c:layout/>${plot}</c:plotArea></c:chart></c:chartSpace>`;
}

/** The whole route: read as the Worker does, lay out as the page does. */
function draw(xml: string, frame = { width: 8, height: 4.5 }): ChartDrawing {
  const read = readChartModel(parseXml(xml), theme);
  if (!("spec" in read)) throw new Error(read.unsupportedReason);
  return layoutChart(read.spec, frame);
}

const texts = (g: ChartDrawing) => g.primitives.filter((p): p is ChartTextPrimitive => p.kind === "text");
const paths = (g: ChartDrawing) => g.primitives.filter((p): p is ChartPathPrimitive => p.kind === "path");

/** Rect subpaths (`M x y h w v h h -w Z`) of every path filled with `color`. */
function rectsFilled(g: ChartDrawing, color: string): Array<{ x: number; y: number; w: number; h: number }> {
  const out: Array<{ x: number; y: number; w: number; h: number }> = [];
  for (const p of paths(g)) {
    if (p.fill !== color) continue;
    for (const m of p.path.matchAll(/M(-?[\d.]+) (-?[\d.]+)h(-?[\d.]+)v(-?[\d.]+)h-?[\d.]+Z/g)) {
      out.push({ x: Number(m[1]), y: Number(m[2]), w: Number(m[3]), h: Number(m[4]) });
    }
  }
  return out;
}

const Q = ["Q1", "Q2", "Q3", "Q4"];

describe("value scaling follows Office's automatic rules", () => {
  test("1-2-5 steps", () => {
    expect([0.07, 0.15, 0.3, 0.9, 1, 1.2, 3, 7, 12, 4500].map(niceStep)).toEqual([0.1, 0.2, 0.5, 1, 1, 2, 5, 10, 20, 5000]);
  });
  test("zero is kept and 5% headroom rounds up to the next unit (PowerPoint's default chart: 0-6 by 1)", () => {
    expect(valueScale(undefined, 1.8, 5, 6)).toMatchObject({ min: 0, max: 6, major: 1 });
  });
  test("data in the top sixth of its magnitude floats the minimum half its span below (Excel: 97-100 reads 95.5-100.5 by 0.5)", () => {
    expect(valueScale(undefined, 97, 100, 10)).toMatchObject({ min: 95.5, max: 100.5, major: 0.5 });
  });
  test("never more than ten intervals when room allows (PowerPoint: max 19 reads 0-20 by 2, max 5 reads 0-6 by 1)", () => {
    expect(valueScale(undefined, 3, 19, 10)).toMatchObject({ min: 0, max: 20, major: 2 });
    expect(valueScale(undefined, 1.8, 5, 10)).toMatchObject({ min: 0, max: 6, major: 1 });
  });
  test("radar axes end on the data (PowerPoint: max 5 reads 0-5, not 0-6)", () => {
    expect(valueScale(undefined, 0, 5, 6, false, false)).toMatchObject({ min: 0, max: 5, major: 1 });
    expect(valueScale(undefined, 0, 5, 6)).toMatchObject({ max: 6 });
  });
  test("negative-only data pins the maximum at zero", () => {
    expect(valueScale(undefined, -15, -2, 6)).toMatchObject({ max: 0 });
  });
  test("fixed bounds and unit win", () => {
    const axis = { min: 0, max: 10, majorUnit: 2.5 } as never;
    expect(valueScale(axis, 1, 5, 20)).toMatchObject({ min: 0, max: 10, major: 2.5 });
  });
  test("percent stacking spans 0-100%", () => {
    expect(valueScale(undefined, 0, 1, 5, true)).toMatchObject({ min: 0, max: 1 });
  });
  test("log axes span whole decades", () => {
    expect(valueScale({ logBase: 10 } as never, 3, 12000, 10)).toMatchObject({ min: 1, max: 100000, log: 10 });
  });
});

describe("text metrics", () => {
  test("Calibri advance widths: 'Q1' at 10pt", () => {
    // Calibri: Q = 0.653em, 1 = 0.507em.
    expect(textWidth("Q1", { size: 10, bold: false, italic: false, color: "#000", family: "Calibri" })).toBeCloseTo(11.6, 0);
  });
  test("bold is wider, unknown families fall to a metric class", () => {
    const look = { size: 12, bold: false, italic: false, color: "#000", family: "Arial" };
    expect(textWidth("Revenue", { ...look, bold: true })).toBeGreaterThan(textWidth("Revenue", look));
    expect(textWidth("Revenue", { ...look, family: "Helvetica Neue" })).toBeCloseTo(textWidth("Revenue", look), 5);
  });
  test("wrapping breaks at spaces only", () => {
    const look = { size: 10, bold: false, italic: false, color: "#000", family: "Calibri" };
    expect(wrapText("Europe, Middle East and Africa", look, 75)).toEqual(["Europe, Middle", "East and Africa"]);
    expect(wrapText("Supercalifragilistic", look, 20)).toEqual(["Supercalifragilistic"]);
  });
});

describe("reading resolves Office defaults", () => {
  test("series without spPr take the theme accents in order", () => {
    const spec = readChart(parseXml(chart(`<c:barChart><c:barDir val="col"/>${ser(0, "A", Q, [1, 2, 3, 4])}${ser(1, "B", Q, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`)), theme)!;
    expect(spec.groups[0]!.series.map((s) => s.look.fill)).toEqual(["#4472C4", "#ED7D31"]);
    expect(spec.groups[0]!.gapWidth).toBe(150);
    expect(spec.text).toMatchObject({ size: 10, family: "Calibri" });
  });
  test("series seven onward darken (Office's colour variations)", () => {
    const series = Array.from({ length: 7 }, (_, i) => ser(i, `S${i}`, Q, [1, 2, 3, 4])).join("");
    const spec = readChart(parseXml(chart(`<c:barChart>${series}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`)), theme)!;
    const seventh = spec.groups[0]!.series[6]!.look.fill as string;
    expect(seventh).not.toBe("#4472C4");
    expect(parseInt(seventh.slice(1, 3), 16)).toBeLessThan(0x44);
  });
  test("txPr on the chart space cascades to every element; element txPr overrides", () => {
    const xml = `<c:chartSpace ${NS}><c:chart><c:plotArea>` +
      `<c:barChart>${ser(0, "A", Q, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      `<c:catAx><c:axId val="1"/><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr></a:p></c:txPr><c:crossAx val="2"/></c:catAx>` +
      `<c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart>` +
      `<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1200"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr></c:chartSpace>`;
    const spec = readChart(parseXml(xml), theme)!;
    const axis = (id: number) => spec.axes.find((a) => a.id === id)!;
    expect(axis(1).text).toMatchObject({ size: 14, bold: true, color: "#595959" });
    expect(axis(2).text).toMatchObject({ size: 12, bold: false, color: "#595959" });
  });
  test("the model is plain data: it survives JSON as it is", () => {
    const xml = chart(`<c:pieChart><c:varyColors val="1"/>${ser(0, "A", Q, [1, 2, 3, 4])}</c:pieChart>`).replace("</c:plotArea>", '</c:plotArea><c:legend><c:legendPos val="r"/><c:legendEntry><c:idx val="2"/><c:delete val="1"/></c:legendEntry></c:legend>');
    const spec = readChart(parseXml(xml), theme)!;
    expect(JSON.parse(JSON.stringify(spec))).toEqual(spec);
    expect(spec.legend!.deleted).toEqual([2]);
    expect(Object.keys(spec.groups[0]!.series[0]!.points)).toEqual(["0", "1", "2", "3"]);
  });
  test("tx1 text resolves through the colour map Excel always uses", () => {
    const raw = { ...theme, colors: { dk1: "#112233", lt1: "#FFFFFF", accent1: "#4472C4" } };
    const xml = chart(`<c:barChart>${ser(0, "A", Q, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`)
      .replace("</c:chartSpace>", '<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr></c:chartSpace>');
    expect(readChart(parseXml(xml), themeForColorMap(raw, undefined))!.text.color).toBe("#112233");
  });
  test("stacked bars ignore the authored overlap", () => {
    const spec = readChart(parseXml(chart(`<c:barChart><c:grouping val="stacked"/>${ser(0, "A", Q, [1, 2, 3, 4])}<c:overlap val="-20"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`)), theme)!;
    expect(spec.groups[0]!.overlap).toBe(100);
  });
});

describe("drawing", () => {
  test("the view box is the frame in points", () => {
    const g = draw(chart(`<c:barChart>${ser(0, "A", Q, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`), { width: 8, height: 4.5 });
    expect({ width: g.width, height: g.height }).toEqual({ width: 576, height: 324 });
    expect(JSON.stringify(g)).not.toContain("<");
  });

  test("clustered bars: width = slot / (n − overlap·(n−1) + gap), side by side, centred in the slot", () => {
    const g = draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, "A", Q, [4, 3, 2, 1])}${ser(1, "B", Q, [1, 2, 3, 4])}<c:gapWidth val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`));
    const a = rectsFilled(g, "#4472C4");
    const b = rectsFilled(g, "#ED7D31");
    expect(a).toHaveLength(4);
    expect(b).toHaveLength(4);
    const slot = a[1]!.x - a[0]!.x;
    expect(a[0]!.w).toBeCloseTo(slot / 3, 1);
    expect(b[0]!.x).toBeCloseTo(a[0]!.x + a[0]!.w, 1);
    // Taller value, taller bar; all share one baseline.
    expect(a[0]!.h).toBeGreaterThan(a[3]!.h);
    expect(a[0]!.y + a[0]!.h).toBeCloseTo(a[3]!.y + a[3]!.h, 1);
  });

  test("a series of 1,000 columns is one primitive", () => {
    const cats = Array.from({ length: 1000 }, (_, i) => `C${i}`);
    const g = draw(chart(`<c:barChart>${ser(0, "A", cats, cats.map((_, i) => i % 17))}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`));
    expect(paths(g).filter((p) => p.fill === "#4472C4")).toHaveLength(1);
    expect(rectsFilled(g, "#4472C4").length).toBeGreaterThan(900);
  });

  test("value axis labels come from the scale and the series format", () => {
    const g = draw(chart(`<c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:cat>${cache("str", Q)}</c:cat><c:val>${cache("num", [125000, 180500, 97250, 210000], "#,##0")}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`));
    const labels = texts(g).map((t) => t.text);
    expect(labels).toEqual(expect.arrayContaining(["0", "50,000", "100,000", "150,000", "200,000", "250,000"]));
  });

  test("crowded category labels stand upright, all of them while they fit (PowerPoint: 30 items, 12pt, 8in)", () => {
    const cats = Array.from({ length: 30 }, (_, i) => `Item ${i + 1}`);
    const twelve = '<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1200"/></a:pPr></a:p></c:txPr></c:chartSpace>';
    const g = draw(chart(`<c:barChart>${ser(0, "A", cats, cats.map((_, i) => i))}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`).replace("</c:chartSpace>", twelve));
    const catLabels = texts(g).filter((t) => t.text.startsWith("Item"));
    expect(catLabels).toHaveLength(30);
    expect(catLabels.every((t) => t.rotate === -90)).toBe(true);
  });

  test("upright labels thin out only when even upright they overlap", () => {
    const cats = Array.from({ length: 60 }, (_, i) => `Item ${i + 1}`);
    const g = draw(chart(`<c:barChart>${ser(0, "A", cats, cats.map((_, i) => i))}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`));
    const catLabels = texts(g).filter((t) => t.text.startsWith("Item"));
    expect(catLabels.length).toBeLessThan(60);
    expect(catLabels.every((t) => t.rotate === -90)).toBe(true);
  });

  test("pie percentages add up to 100: the rounding goes to the largest slice (PowerPoint: 58/23/10/9)", () => {
    const xml = chart(`<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:dLbls><c:showVal val="0"/><c:showCatName val="0"/><c:showPercent val="1"/></c:dLbls><c:cat>${cache("str", ["a", "b", "c", "d"])}</c:cat><c:val>${cache("num", [8.2, 3.2, 1.4, 1.2])}</c:val></c:ser></c:pieChart>`);
    expect(texts(draw(xml)).map((t) => t.text)).toEqual(expect.arrayContaining(["58%", "23%", "10%", "9%"]));
  });

  test("pie labels put category and percent on separate lines unless a separator is set", () => {
    const xml = chart(`<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:dLbls><c:showVal val="0"/><c:showCatName val="1"/><c:showPercent val="1"/></c:dLbls><c:cat>${cache("str", ["Web", "TV"])}</c:cat><c:val>${cache("num", [1, 3])}</c:val></c:ser></c:pieChart>`);
    expect(texts(draw(xml)).map((t) => t.text)).toEqual(expect.arrayContaining(["Web", "25%", "TV", "75%"]));
  });

  test("long category labels wrap before they rotate", () => {
    const cats = ["North America region", "Europe, Middle East and Africa", "Asia Pacific and Japan", "Latin America"];
    const g = draw(chart(`<c:barChart>${ser(0, "A", cats, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`), { width: 5, height: 3 });
    const t = texts(g).map((x) => x.text);
    expect(t).toContain("Latin America");
    expect(t.some((x) => x.startsWith("Europe,") && x !== cats[1])).toBe(true);
    expect(texts(g).every((x) => x.rotate === undefined)).toBe(true);
  });

  test("a reversed category axis moves the value axis to the right", () => {
    const g = draw(chart(`<c:barChart>${ser(0, "A", Q, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes('<c:scaling><c:orientation val="maxMin"/></c:scaling>')}`));
    const zero = texts(g).find((t) => t.text === "0")!;
    const q1 = texts(g).find((t) => t.text === "Q1")!;
    const q4 = texts(g).find((t) => t.text === "Q4")!;
    expect(q1.x).toBeGreaterThan(q4.x);
    expect(zero.x).toBeGreaterThan(q1.x);
    expect(zero.textAnchor).toBe("start");
  });

  test("negative values: bars hang from zero, category labels sit at the axis line", () => {
    const g = draw(chart(`<c:barChart>${ser(0, "A", Q, [5, -5, 5, -5])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`));
    const bars = rectsFilled(g, "#4472C4");
    expect(bars[0]!.y + bars[0]!.h).toBeCloseTo(bars[1]!.y, 1);
    const q1 = texts(g).find((t) => t.text === "Q1")!;
    expect(q1.y).toBeGreaterThan(bars[1]!.y);
    expect(q1.y).toBeLessThan(bars[1]!.y + bars[1]!.h);
  });

  test("line: gaps break the path, markers sit on the points", () => {
    const g = draw(chart(`<c:lineChart><c:grouping val="standard"/>${ser(0, "A", ["a", "b", "c", "d", "e"], [1, 2, null, 4, 5], '<c:marker><c:symbol val="square"/><c:size val="6"/></c:marker>')}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${axes()}`, '<c:dispBlanksAs val="gap"/>'));
    const line = paths(g).find((p) => p.stroke?.width === 2.25)!;
    expect(line.path.match(/M/g)).toHaveLength(2);
    const markers = paths(g).find((p) => p.fill === "#4472C4" && p.stroke?.width === 0.75)!;
    expect(markers.path.match(/Z/g)).toHaveLength(4);
  });

  test("pie: slices sweep clockwise from 12 o'clock, percentages label the slices", () => {
    const xml = chart(`<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/></c:dLbls><c:cat>${cache("str", ["Web", "TV"])}</c:cat><c:val>${cache("num", [1, 3])}</c:val></c:ser><c:firstSliceAng val="0"/></c:pieChart>`);
    const g = draw(xml);
    const slices = paths(g).filter((p) => p.path.includes("A"));
    expect(slices.map((p) => p.fill)).toEqual(["#4472C4", "#ED7D31"]);
    // First slice starts straight up: its first edge point has the centre's x.
    const m = slices[0]!.path.match(/^M([\d.]+) ([\d.]+)L([\d.]+) ([\d.]+)/)!;
    expect(Number(m[3])).toBeCloseTo(Number(m[1]), 1);
    expect(Number(m[4])).toBeLessThan(Number(m[2]));
    expect(texts(g).map((t) => t.text)).toEqual(expect.arrayContaining(["25%", "75%"]));
  });

  test("title, legend and axis titles reserve their bands", () => {
    const xml = chart(`<c:barChart>${ser(0, "A", Q, [1, 2, 3, 4])}${ser(1, "B", Q, [2, 3, 4, 5])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`,
      '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>')
      .replace("</c:plotArea>", '</c:plotArea><c:legend><c:legendPos val="b"/></c:legend>');
    const g = draw(xml);
    const title = texts(g).find((t) => t.text === "Sales")!;
    expect(title.fontSize).toBe(18);
    expect(title.fontWeight).toBe("bold");
    const legendA = texts(g).find((t) => t.text === "A")!;
    const q1 = texts(g).find((t) => t.text === "Q1")!;
    expect(legendA.y).toBeGreaterThan(q1.y);
    const bars = rectsFilled(g, "#4472C4");
    expect(Math.min(...bars.map((b) => b.y))).toBeGreaterThan(title.y);
  });

  test("scatter uses numeric x values on a value axis", () => {
    const xml = chart(`<c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/><c:xVal>${cache("num", [0, 10, 20, 35])}</c:xVal><c:yVal>${cache("num", [1, 4, 9, 16])}</c:yVal></c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart>` +
      `<c:valAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>`);
    const g = draw(xml);
    const t = texts(g).map((x) => x.text);
    expect(t).toEqual(expect.arrayContaining(["0", "10", "20", "30", "40"]));
    // The x axis sits on one baseline; Office keeps its labels well apart.
    const xLabels = texts(g).filter((x) => x.y === texts(g).find((l) => l.text === "40")!.y);
    expect(xLabels.length).toBeLessThanOrEqual(10);
  });

  test("every supported group kind draws, and other kinds explain themselves", () => {
    const kinds = ["areaChart", "radarChart", "doughnutChart", "bubbleChart"];
    for (const kind of kinds) {
      const body = kind === "bubbleChart"
        ? `<c:ser><c:idx val="0"/><c:xVal>${cache("num", [1, 2])}</c:xVal><c:yVal>${cache("num", [3, 4])}</c:yVal><c:bubbleSize>${cache("num", [1, 2])}</c:bubbleSize></c:ser>`
        : ser(0, "A", Q, [1, 2, 3, 4]);
      const ax = kind === "doughnutChart" ? "" : '<c:axId val="1"/><c:axId val="2"/>';
      const axesXml = kind === "doughnutChart" ? "" : kind === "bubbleChart" ? `<c:valAx><c:axId val="1"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>` : axes();
      const g = draw(chart(`<c:${kind}>${body}${ax}</c:${kind}>${axesXml}`));
      expect(paths(g).length).toBeGreaterThan(0);
    }
    const surface = parseXml(chart(`<c:surfaceChart>${ser(0, "A", Q, [1, 2, 3, 4])}</c:surfaceChart>`));
    expect(readChartModel(surface, theme)).toEqual({ unsupportedReason: "chart type surfaceChart is not previewed" });
  });

  test("stock and of-pie charts draw (their geometry is pinned in chart-types.test.ts); an unknown type still shows its placeholder, even beside a drawn plot", () => {
    const stock = draw(chart(`<c:stockChart>${ser(0, "H", Q, [4, 5, 6, 7])}${ser(1, "L", Q, [1, 2, 3, 4])}${ser(2, "C", Q, [2, 3, 5, 6])}<c:hiLowLines/><c:axId val="1"/><c:axId val="2"/></c:stockChart>${axes()}`));
    expect(paths(stock).length).toBeGreaterThan(1);
    const ofPie = draw(chart(`<c:ofPieChart><c:ofPieType val="bar"/><c:varyColors val="1"/>${ser(0, "A", Q, [4, 3, 2, 1])}</c:ofPieChart>`));
    expect(paths(ofPie).length).toBeGreaterThan(1);
    const combo = parseXml(chart(`<c:barChart>${ser(0, "A", Q, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:surfaceChart>${ser(1, "B", Q, [1, 2, 3, 4])}</c:surfaceChart>${axes()}`));
    expect(readChartModel(combo, theme)).toEqual({ unsupportedReason: "chart type surfaceChart is not previewed" });
  });

  test("a part that claims millions of points is refused at once, not allocated", () => {
    // 600 bytes claiming nine million points froze the reader; the layout now runs on the page.
    const claim = `<c:val><c:numRef><c:numCache><c:ptCount val="9000000"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="8999999"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>`;
    const xml = parseXml(chart(`<c:barChart><c:ser><c:idx val="0"/>${claim}</c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`));
    const t0 = performance.now();
    const read = readChartModel(xml, theme);
    // The refusal itself is the check; the bound only guards against allocating the claim (seconds, not milliseconds).
    expect(performance.now() - t0).toBeLessThan(5000);
    expect(read).toEqual({ unsupportedReason: "this chart has a series of more than 100,000 points, more than the preview draws" });
    // Many series each under the bound still meet the chart-wide one.
    const wide = Array.from({ length: 3 }, (_, i) => `<c:ser><c:idx val="${i}"/><c:val><c:numRef><c:numCache><c:ptCount val="99000"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>`).join("");
    const many = readChartModel(parseXml(chart(`<c:lineChart>${wide}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${axes()}`)), theme);
    expect(many).toEqual({ unsupportedReason: "this chart has more than 250,000 points in all, more than the preview draws" });
    // The bound is met as each series claims its points, not after all of them are allocated.
    const thousands = Array.from({ length: 5000 }, (_, i) => `<c:ser><c:idx val="${i}"/><c:val><c:numRef><c:numCache><c:ptCount val="100000"/></c:numCache></c:numRef></c:val></c:ser>`).join("");
    const started = performance.now();
    const hostile = readChartModel(parseXml(chart(`<c:lineChart>${thousands}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${axes()}`)), theme);
    expect(performance.now() - started).toBeLessThan(5000);
    expect(hostile).toEqual({ unsupportedReason: "this chart has more than 250,000 points in all, more than the preview draws" });
  });

  test("numeric x values and bubble sizes count against the budget too", () => {
    const big = (tag: string): string => `<c:${tag}><c:numRef><c:numCache><c:ptCount val="99000"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:${tag}>`;
    const small = (tag: string): string => `<c:${tag}><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:${tag}>`;
    const tooMany = { unsupportedReason: "this chart has more than 250,000 points in all, more than the preview draws" };
    const series = (parts: (i: number) => string): string => Array.from({ length: 3 }, (_, i) => `<c:ser><c:idx val="${i}"/>${parts(i)}</c:ser>`).join("");
    // A scatter whose y values are short but whose numeric x values are not: the x arrays are allocated, so they are claimed.
    const scatter = series(() => `${big("xVal")}${small("yVal")}`);
    expect(readChartModel(parseXml(chart(`<c:scatterChart><c:scatterStyle val="lineMarker"/>${scatter}<c:axId val="1"/><c:axId val="2"/></c:scatterChart>${axes()}`)), theme)).toEqual(tooMany);
    // A bubble chart whose sizes are the long array.
    const bubble = series(() => `${small("xVal")}${small("yVal")}${big("bubbleSize")}`);
    expect(readChartModel(parseXml(chart(`<c:bubbleChart>${bubble}<c:axId val="1"/><c:axId val="2"/></c:bubbleChart>${axes()}`)), theme)).toEqual(tooMany);
    // Under the budget, the same shapes still read.
    const one = readChartModel(parseXml(chart(`<c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/>${big("xVal")}${small("yVal")}</c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart>${axes()}`)), theme);
    expect(one).not.toHaveProperty("unsupportedReason");
  });

  test("rotation past a quarter turn reads as automatic", () => {
    const auto = '<c:txPr><a:bodyPr rot="-60000000"/><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:txPr>';
    const spec = readChart(parseXml(chart(`<c:barChart>${ser(0, "A", Q, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes(auto)}`)), theme)!;
    expect(spec.axes.find((a) => a.id === 1)!.text.rotation).toBeUndefined();
  });

  test("Arabic text measures as joined letters: PowerPoint drew 'al-mabi'at' (sales) 31.4pt wide at 12pt in Arial", () => {
    const look = { size: 12, bold: false, italic: false, color: "#000", family: "Arial" };
    expect(textWidth("المبيعات", look)).toBeCloseTo(31.4, 0);
    expect(Math.abs(textWidth("المبيعات", look) - 31.4)).toBeLessThan(0.5);
  });
});

describe("Arabic text", () => {
  const look = { size: 100, bold: false, italic: false, color: "#000", family: "Arial" };
  const em = (text: string, over: Partial<typeof look & { csFamily: string }> = {}): number => textWidth(text, { ...look, ...over }) / 100;

  test("a letter is as wide as its joining form: isolated, initial, medial, final (Arial's own shaped glyphs)", () => {
    // Beh: 0.713 isolated or final, 0.244 initial or medial. The third letter of a word is medial, so
    // adding one beh to "bb" adds a medial beh; a flat per-letter width could not tell these apart.
    expect(em("ب")).toBeCloseTo(0.713, 2);
    expect(em("بب")).toBeCloseTo(0.244 + 0.713, 2);
    expect(em("ببب")).toBeCloseTo(0.244 + 0.244 + 0.713, 2);
    // A right-joining letter (alef) joins what is before it and breaks what is after it.
    expect(em("اب")).toBeCloseTo(em("ا") + em("ب"), 3);
    expect(em("ا ب")).toBeCloseTo(em("ا") + em(" ") + em("ب"), 3);
    // Words of the same letter count differ in width, which the flat 0.333 em never showed.
    expect(em("المبيعات")).toBeCloseTo(2.633, 2);
    expect(em("التكاليف")).toBeCloseTo(2.521, 2);
    expect(em("الربع")).not.toBeCloseTo(5 * 0.333, 1);
  });

  test("lam followed by alef is one glyph, and harakat take no advance and let a join pass", () => {
    expect(em("لا")).toBeCloseTo(0.544, 2);
    expect(em("كَتَبَ")).toBeCloseTo(em("كتب"), 5);
    expect(em("كتب")).toBeCloseTo(1.352, 2);
  });

  test("Arabic-Indic digits have their own width, and a Latin character ends a join", () => {
    expect(em("٢٠")).toBeCloseTo(2 * 0.526, 2);
    // The Latin run measures in the Latin face, and the Arabic letters around it do not join across it.
    expect(em("با1با")).toBeCloseTo(2 * em("با") + textWidth("1", look) / 100, 3);
  });

  test("the complex-script face measures Arabic, the Latin face the rest", () => {
    // Tahoma's Arabic is wider than Arial's; Latin text is untouched by csFamily.
    expect(em("المبيعات", { csFamily: "Tahoma" })).toBeGreaterThan(em("المبيعات", { csFamily: "Arial" }) * 1.2);
    expect(em("Revenue", { csFamily: "Tahoma" })).toBeCloseTo(em("Revenue"), 6);
    // Times New Roman's Arabic glyphs are Arial's; with nothing named, Arabic takes Arial beside a sans face.
    expect(em("المبيعات", { csFamily: "Times New Roman" })).toBeCloseTo(em("المبيعات"), 6);
    expect(em("المبيعات", { family: "Calibri" })).toBeCloseTo(em("المبيعات", { family: "Arial" }), 6);
    const mixed = em("Q1 المبيعات", { family: "Calibri", csFamily: "Tahoma" });
    expect(mixed).toBeCloseTo(textWidth("Q1 ", { ...look, family: "Calibri" }) / 100 + em("المبيعات", { csFamily: "Tahoma" }), 3);
  });

  test("a:cs on the text properties, and the theme's cs slot, become the spec's csFamily", () => {
    const withCs = (cs: string): string => `<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1200"><a:latin typeface="+mn-lt"/><a:cs typeface="${cs}"/></a:defRPr></a:pPr></a:p></c:txPr>`;
    const spec = (cs: string, th = theme) => readChart(parseXml(chart(`<c:barChart>${ser(0, "A", Q, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes(withCs(cs))}`)), th)!;
    expect(spec("Tahoma").axes.find((a) => a.id === 1)!.text.csFamily).toBe("Tahoma");
    const arabTheme = themeForColorMap({
      colors: { dk1: "#000000", lt1: "#FFFFFF", accent1: "#4472C4" },
      fonts: { major: { latin: "Calibri Light", complexScript: "", scripts: {} }, minor: { latin: "Calibri", complexScript: "", scripts: { Arab: "Arial" } } },
    }, undefined);
    // `+mn-cs` is Office's own default; with the theme's cs empty it is the Arabic script font.
    expect(spec("+mn-cs", arabTheme).axes.find((a) => a.id === 1)!.text.csFamily).toBe("Arial");
    expect(spec("+mn-cs", arabTheme).axes.find((a) => a.id === 2)!.text.csFamily).toBe("Arial");
    // A theme with no Arabic font names none: the engine's fallback applies, and nothing is invented.
    expect(spec("+mn-cs").axes.find((a) => a.id === 1)!.text.csFamily).toBeUndefined();
  });

  const arabicCats = ["الربع الأول", "الربع الثاني", "الربع الثالث", "الربع الرابع"];
  const arabicChart = (cs = "Tahoma", extra = ""): string => chart(`<c:barChart>${ser(0, "المبيعات", arabicCats, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes(`<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr><a:cs typeface="${cs}"/></a:defRPr></a:pPr></a:p></c:txPr>`)}`, extra)
    .replace("</c:plotArea>", '</c:plotArea><c:legend><c:legendPos val="r"/></c:legend>');

  test("Arabic text paints in the complex-script face, a mixed line in the Latin face with the Arabic one behind it", () => {
    const g = draw(arabicChart());
    const category = texts(g).find((t) => t.text === arabicCats[0])!;
    expect(category.fontFamily).toBe("Tahoma");
    // The legend keeps the chart's own text (no cs named): Office's Arabic fallback, Arial.
    expect(texts(g).find((t) => t.text === "المبيعات")!.fontFamily).toBe("Arial");
    const mixed = texts(draw(chart(`<c:barChart>${ser(0, "A", ["الربع 1", "Q1 الربع"], [1, 2])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes('<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr><a:cs typeface="Tahoma"/></a:defRPr></a:pPr></a:p></c:txPr>')}`)));
    // Calibri, Arial and Tahoma all carry Arabic glyphs, so a font list alone would paint the Arabic in the
    // Latin face; the line is cut into runs, each in the face that sets it. A space between Arabic words
    // is the Arabic face's, the neutral after a Latin run the Latin one's.
    expect(mixed.find((t) => t.text === "الربع 1")!.runs).toEqual([{ text: "الربع ", fontFamily: "Tahoma" }, { text: "1", fontFamily: "Calibri" }]);
    expect(mixed.find((t) => t.text === "Q1 الربع")!.runs).toEqual([{ text: "Q1 ", fontFamily: "Calibri" }, { text: "الربع", fontFamily: "Tahoma" }]);
    expect(mixed.find((t) => t.text === "الربع 1")!.fontFamily).toBe("Calibri");
    // Latin text, and pure Arabic, have no runs.
    expect(texts(g).find((t) => t.text === "0")!.fontFamily).toBe("Calibri");
    expect(texts(g).find((t) => t.text === "0")!.runs).toBeUndefined();
    expect(category.runs).toBeUndefined();
  });

  test("the measure follows the same runs: a space inside Arabic is the Arabic face's", () => {
    // Tahoma's space is 0.313 em, Arial's 0.278: the Arabic face's wins between Arabic words, the Latin face's after Latin.
    const tahoma = { csFamily: "Tahoma" };
    expect(em("ا ب", tahoma)).toBeCloseTo(em("ا", tahoma) + 0.312 + em("ب", tahoma), 2);
    expect(em("Q ب", tahoma)).toBeCloseTo(textWidth("Q ", look) / 100 + em("ب", tahoma), 3);
    expect(scriptRuns("الربع 1").map((r) => r.text)).toEqual(["الربع ", "1"]);
    expect(scriptRuns(" الربع").map((r) => r.arabic)).toEqual([true]);
    expect(scriptRuns("Revenue").map((r) => r.arabic)).toEqual([false]);
  });

  test("text whose first strong character is right to left gets direction rtl, anything else none", () => {
    const g = draw(chart(`<c:barChart>${ser(0, "A", ["الربع 1", "Q1 الربع", "2024", "12 الربع"], [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`));
    const dir = (text: string) => texts(g).find((t) => t.text === text)!.direction;
    expect(dir("الربع 1")).toBe("rtl");
    expect(dir("12 الربع")).toBe("rtl");
    expect(dir("Q1 الربع")).toBeUndefined();
    expect(dir("2024")).toBeUndefined();
    expect(dir("0")).toBeUndefined();
  });

  test("the layout keeps its anchors in left-to-right terms whatever the text direction", () => {
    // Value axis labels sit left of the axis, anchored at their end; Arabic or not, the engine's x is the same.
    const arabicValues = draw(arabicChart());
    const latin = draw(chart(`<c:barChart>${ser(0, "A", Q, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`));
    const zero = (g: ChartDrawing) => texts(g).find((t) => t.text === "0")!;
    expect(zero(arabicValues).textAnchor).toBe(zero(latin).textAnchor);
  });

  test("a font scale from the page widens what the layout measured", () => {
    const xml = chart(`<c:barChart>${ser(0, "A", Q, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`);
    const read = readChartModel(parseXml(xml), theme);
    if (!("spec" in read)) throw new Error(read.unsupportedReason);
    const firstBar = (g: ChartDrawing) => Math.min(...rectsFilled(g, "#4472C4").map((r) => r.x));
    const plain = layoutChart(read.spec, { width: 8, height: 4.5 });
    const wide = layoutChart(read.spec, { width: 8, height: 4.5 }, { Calibri: 1.5 });
    // Wider value labels take more of the left edge, so the plot (and its first bar) starts further right.
    expect(firstBar(wide)).toBeGreaterThan(firstBar(plain) + 3);
    // The scale lasts one layout: the next one without it is the plain one again.
    expect(firstBar(layoutChart(read.spec, { width: 8, height: 4.5 }))).toBe(firstBar(plain));
    // Arabic glyphs scale under their own key, apart from the Latin of the same face.
    const arabic = readChartModel(parseXml(arabicChart("Arial")), theme);
    if (!("spec" in arabic)) throw new Error(arabic.unsupportedReason);
    const legend = (g: ChartDrawing) => texts(g).find((t) => t.text === "المبيعات")!.x;
    const base = layoutChart(arabic.spec, { width: 8, height: 4.5 });
    expect(legend(layoutChart(arabic.spec, { width: 8, height: 4.5 }, { "Arial|cs": 1.6 }))).toBeLessThan(legend(base));
    expect(legend(layoutChart(arabic.spec, { width: 8, height: 4.5 }, { Arial: 1.6 }))).toBe(legend(base));
  });

  test("the probes name each face the chart is set in, Arabic ones only for a chart with Arabic text", () => {
    const latin = readChart(parseXml(chart(`<c:barChart>${ser(0, "A", Q, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`)), theme)!;
    expect(chartFontProbes(latin, ARABIC_PROBE).map((p) => p.key)).toEqual(["Calibri"]);
    const arabic = readChart(parseXml(arabicChart()), theme)!;
    const probes = chartFontProbes(arabic, ARABIC_PROBE);
    expect(probes.map((p) => p.key).sort()).toEqual(["Arial|cs", "Calibri", "Tahoma|cs"]);
    for (const probe of probes) expect(probe.expected).toBeGreaterThan(0);
  });
});

describe("drawing, continued", () => {
  test("drawing a 5,000-point line is deterministic (and nowhere near slow)", () => {
    const cats = Array.from({ length: 5000 }, (_, i) => `C${i}`);
    const xml = parseXml(chart(`<c:lineChart>${ser(0, "A", cats, cats.map((_, i) => Math.sin(i / 50)))}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${axes()}`));
    const spec = readChart(xml, theme)!;
    const first = JSON.stringify(drawChart(spec, 576, 324));
    const t0 = performance.now();
    const second = JSON.stringify(drawChart(spec, 576, 324));
    expect(performance.now() - t0).toBeLessThan(5000);
    expect(second).toBe(first);
  });

  test("a label's lines are one paragraph: under an Arabic category the percentage runs right to left", () => {
    const cats = ["أوروبا", "Europe"];
    const dLbls = '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/></c:dLbls>';
    const g = draw(chart(`<c:pieChart><c:varyColors val="1"/>${ser(0, "S", cats, [3, 1], dLbls)}</c:pieChart>`));
    const percent = (share: string) => texts(g).find((t) => t.text === share)!;
    // "75%" belongs to the Arabic label: right to left, shown as "%75", as PowerPoint draws it.
    expect(percent("75%").direction).toBe("rtl");
    // "25%" belongs to a Latin label and stays left to right.
    expect(percent("25%").direction).toBeUndefined();
  });

  test("a line longer than the plot is wide keeps every peak and dip, in a few points per column", () => {
    const n = 50_000;
    const cats = Array.from({ length: n }, (_, i) => `C${i}`);
    // A gentle wave with one spike up and one down: the spikes must survive thinning.
    const values = cats.map((_, i) => (i === 31_337 ? 100 : i === 12_345 ? -100 : Math.sin(i / 400)));
    const spec = readChart(parseXml(chart(`<c:lineChart><c:grouping val="standard"/>${ser(0, "A", cats, values, '<c:marker><c:symbol val="none"/></c:marker>')}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${axes('<c:tickLblSkip val="5000"/>')}`)), theme)!;
    const drawn = drawChart(spec, 576, 324);
    const line = (drawn.primitives as ChartPathPrimitive[]).find((p) => p.kind === "path" && p.stroke !== undefined && p.path.split("L").length > 100)!;
    const ys = [...line.path.matchAll(/[ML](-?[\d.]+) (-?[\d.]+)/g)].map((m) => Number(m[2]));
    // At most four points per half-point column of the plot, far fewer than the data.
    expect(ys.length).toBeLessThan(n / 4);
    // The spikes are still the line's extremes: exactly where a three-point line with the same
    // range (so the same value axis) puts 100 and -100, unthinned.
    const short = readChart(parseXml(chart(`<c:lineChart><c:grouping val="standard"/>${ser(0, "A", ["a", "b", "c"], [100, -100, 0], '<c:marker><c:symbol val="none"/></c:marker>')}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${axes('<c:tickLblSkip val="5000"/>')}`)), theme)!;
    const reference = (drawChart(short, 576, 324).primitives as ChartPathPrimitive[]).find((p) => p.kind === "path" && p.stroke !== undefined && /^M[^M]*L[^M]*L/.test(p.path) && p.path.split("L").length === 3)!;
    const refYs = [...reference.path.matchAll(/[ML](-?[\d.]+) (-?[\d.]+)/g)].map((m) => Number(m[2]));
    expect(Math.min(...ys)).toBeCloseTo(Math.min(...refYs), 1);
    expect(Math.max(...ys)).toBeCloseTo(Math.max(...refYs), 1);
  });
});

// ------------------------------------------------------------------ content Office shows

const rich = (...paragraphs: string[]): string => `<c:tx><c:rich><a:bodyPr/>${paragraphs.map((p) => `<a:p>${p}</a:p>`).join("")}</c:rich></c:tx>`;
const run = (text: string): string => `<a:r><a:t>${text}</a:t></a:r>`;
const field = (type: string, cached: string): string => `<a:fld type="${type}"><a:t>${cached}</a:t></a:fld>`;
const flags = (val: number, cat = 0, ser = 0, pct = 0): string =>
  `<c:showLegendKey val="0"/><c:showVal val="${val}"/><c:showCatName val="${cat}"/><c:showSerName val="${ser}"/><c:showPercent val="${pct}"/><c:showBubbleSize val="0"/>`;

/** A four-point column chart whose series carries `dLbls` (and optional series-level `extra`). Values avoid the axis' own tick labels. */
function labelled(dLbls: string, extra = ""): string {
  return chart(`<c:barChart>${ser(0, "Sales", Q, [11, 23, 37, 41], `${dLbls}${extra}`)}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`);
}

describe("data labels show what the author typed", () => {
  test("a point with typed text draws it, the others keep their values", () => {
    const g = draw(labelled(`<c:dLbls><c:dLbl><c:idx val="3"/>${rich(run("Record high"))}${flags(1)}</c:dLbl>${flags(1)}</c:dLbls>`));
    const t = texts(g).map((x) => x.text);
    expect(t).toContain("Record high");
    expect(t).not.toContain("41");
    expect(t).toEqual(expect.arrayContaining(["11", "23", "37"]));
  });

  test("fields are filled from the point; paragraphs become lines; a field with no source keeps its cached text", () => {
    const g = draw(labelled(`<c:dLbls><c:dLbl><c:idx val="1"/>${rich(field("SERIESNAME", "[SERIES NAME]") + run(": ") + field("CATEGORYNAME", "[CATEGORY NAME]"), field("VALUE", "[VALUE]") + run(" / ") + field("PERCENTAGE", "[PERCENTAGE]"))}${flags(1)}</c:dLbl>${flags(1)}</c:dLbls>`));
    const t = texts(g).map((x) => x.text);
    expect(t).toContain("Sales: Q2");
    expect(t).toContain("23 / [PERCENTAGE]");
  });

  test("typed text follows the label's own font", () => {
    const xml = labelled(`<c:dLbls><c:dLbl><c:idx val="0"/><c:tx><c:rich><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1600" b="1"/></a:pPr><a:r><a:t>Big</a:t></a:r></a:p></c:rich></c:tx>${flags(1)}</c:dLbl>${flags(1)}</c:dLbls>`);
    const big = texts(draw(xml)).find((x) => x.text === "Big")!;
    expect(big).toMatchObject({ fontSize: 16, fontWeight: "bold" });
  });

  test("value from cells: the range text leads the label, and a CELLRANGE field takes its point's text", () => {
    const range = `<c:extLst><c:ext><c15:datalabelsRange><c15:f>Sheet1!$C$2:$C$5</c15:f><c15:dlblRangeCache><c:ptCount val="4"/>${["Alpha", "Beta", "Gamma", "Delta"].map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}</c15:dlblRangeCache></c15:datalabelsRange></c:ext></c:extLst>`;
    const show = (val: number, sep = "") => `<c:dLbls>${sep}${flags(val)}<c:extLst><c:ext><c15:showDataLabelsRange val="1"/></c:ext></c:extLst></c:dLbls>`;
    const only = texts(draw(labelled(show(0), range))).map((x) => x.text);
    expect(only).toEqual(expect.arrayContaining(["Alpha", "Beta", "Gamma", "Delta"]));
    expect(only).not.toContain("11");
    const both = texts(draw(labelled(show(1, "<c:separator> / </c:separator>"), range))).map((x) => x.text);
    expect(both).toEqual(expect.arrayContaining(["Alpha / 11", "Delta / 41"]));
    const typed = labelled(`<c:dLbls><c:dLbl><c:idx val="2"/>${rich(field("CELLRANGE", "[CELLRANGE]") + run("!"))}${flags(0)}</c:dLbl>${flags(0)}</c:dLbls>`, range);
    expect(texts(draw(typed)).map((x) => x.text)).toContain("Gamma!");
  });

  test("a deleted label draws nothing, its neighbours stay", () => {
    const t = texts(draw(labelled(`<c:dLbls><c:dLbl><c:idx val="1"/><c:delete val="1"/></c:dLbl>${flags(1)}</c:dLbls>`))).map((x) => x.text);
    expect(t).not.toContain("23");
    expect(t).toEqual(expect.arrayContaining(["11", "37", "41"]));
  });

  test("a manual position moves the label by that fraction of the chart", () => {
    const plain = texts(draw(labelled(`<c:dLbls>${flags(1)}</c:dLbls>`))).find((x) => x.text === "11")!;
    const moved = texts(draw(labelled(`<c:dLbls><c:dLbl><c:idx val="0"/><c:layout><c:manualLayout><c:x val="0.1"/><c:y val="-0.05"/></c:manualLayout></c:layout>${flags(1)}</c:dLbl>${flags(1)}</c:dLbls>`))).find((x) => x.text === "11")!;
    expect(moved.x - plain.x).toBeCloseTo(0.1 * 576, 1);
    expect(moved.y - plain.y).toBeCloseTo(-0.05 * 324, 1);
  });

  test("series name, category and value join with the file's separator, in Office's order", () => {
    const t = texts(draw(labelled(`<c:dLbls>${flags(1, 1, 1)}<c:separator> | </c:separator></c:dLbls>`))).map((x) => x.text);
    expect(t).toEqual(expect.arrayContaining(["Sales | Q1 | 11", "Sales | Q4 | 41"]));
  });

  test("one point can show a different combination than its series", () => {
    const t = texts(draw(labelled(`<c:dLbls><c:dLbl><c:idx val="2"/>${flags(0, 1)}</c:dLbl>${flags(1)}</c:dLbls>`))).map((x) => x.text);
    expect(t).toContain("Q3");
    expect(t).not.toContain("37");
  });

  test("on an XY chart the category name is the x value", () => {
    const xml = chart(`<c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/><c:dLbls>${flags(1, 1)}<c:separator>;</c:separator></c:dLbls><c:xVal>${cache("num", [5, 10])}</c:xVal><c:yVal>${cache("num", [1, 4])}</c:yVal></c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart>` +
      `<c:valAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>`);
    expect(texts(draw(xml)).map((x) => x.text)).toEqual(expect.arrayContaining(["5;1", "10;4"]));
  });

  test("the model keeps the typed text as plain data, and bounds it", () => {
    const spec = readChart(parseXml(labelled(`<c:dLbls><c:dLbl><c:idx val="3"/>${rich(run("Record high") + field("VALUE", "[VALUE]"))}${flags(1)}</c:dLbl>${flags(1)}</c:dLbls>`)), theme)!;
    expect(spec.groups[0]!.series[0]!.points[3]!.labels!.custom).toEqual(["Record high", { field: "VALUE", cached: "[VALUE]" }]);
    expect(JSON.parse(JSON.stringify(spec))).toEqual(spec);
    const long = Array.from({ length: 200 }, () => run("x")).join("");
    const bounded = readChart(parseXml(labelled(`<c:dLbls><c:dLbl><c:idx val="0"/>${rich(long)}${flags(1)}</c:dLbl>${flags(1)}</c:dLbls>`)), theme)!;
    expect(bounded.groups[0]!.series[0]!.points[0]!.labels!.custom!.length).toBe(64);
  });
});

describe("grouped categories draw their outer levels", () => {
  const inner = ["Q1", "Q2", "Q3", "Q1", "Q2", "Q3"];
  const level = (names: Array<[number, string]>): string => `<c:lvl>${names.map(([i, v]) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}</c:lvl>`;
  const multi = (...outer: string[]): string =>
    `<c:multiLvlStrRef><c:multiLvlStrCache><c:ptCount val="6"/>${level(inner.map((v, i) => [i, v]))}${outer.join("")}</c:multiLvlStrCache></c:multiLvlStrRef>`;
  const grouped = (cat: string, axisExtra = "", dir = "col"): string => {
    const s = `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${cache("str", ["Units"])}</c:tx><c:cat>${cat}</c:cat><c:val>${cache("num", [1, 2, 3, 4, 5, 6])}</c:val></c:ser>`;
    return chart(`<c:barChart><c:barDir val="${dir}"/>${s}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes(axisExtra)}`);
  };
  const years = level([[0, "2023"], [3, "2024"]]);

  /** Vertical strokes `M x y0 V y1` of every path. */
  const verticals = (g: ChartDrawing): Array<{ x: number; y0: number; y1: number }> =>
    paths(g).flatMap((p) => [...p.path.matchAll(/M(-?[\d.]+) (-?[\d.]+)V(-?[\d.]+)/g)].map((m) => ({ x: Number(m[1]), y0: Number(m[2]), y1: Number(m[3]) })));

  test("the reader keeps every level: the inner one as the categories, the others as group starts", () => {
    const spec = readChart(parseXml(grouped(multi(years, level([[0, "FY"]])))), theme)!;
    const s = spec.groups[0]!.series[0]!;
    expect(s.categories).toEqual(inner);
    expect(s.categoryGroups).toEqual([[{ start: 0, text: "2023" }, { start: 3, text: "2024" }], [{ start: 0, text: "FY" }]]);
    expect(spec.axes.find((a) => a.id === 1)!.multiLevel).toBe(true);
    expect(JSON.parse(JSON.stringify(spec))).toEqual(spec);
  });

  test("levels past the bound are dropped, not allocated", () => {
    const many = Array.from({ length: 30 }, (_, i) => level([[0, `L${i}`]]));
    const s = readChart(parseXml(grouped(multi(...many))), theme)!.groups[0]!.series[0]!;
    expect(s.categoryGroups!.length).toBe(7);
  });

  test("each outer label is centred under its categories, below the inner labels", () => {
    const g = draw(grouped(multi(years)));
    const t = texts(g);
    const q = t.filter((x) => x.text.startsWith("Q"));
    const y2023 = t.find((x) => x.text === "2023")!;
    const y2024 = t.find((x) => x.text === "2024")!;
    expect(y2023.x).toBeCloseTo((q[0]!.x + q[2]!.x) / 2, 1);
    expect(y2024.x).toBeCloseTo((q[3]!.x + q[5]!.x) / 2, 1);
    expect(y2023.y).toBeGreaterThan(q[0]!.y);
    expect(y2023.textAnchor).toBe("middle");
  });

  test("a tick runs between groups, through the label rows, and at both ends", () => {
    const g = draw(grouped(multi(years)));
    const t = texts(g);
    const q = t.filter((x) => x.text.startsWith("Q"));
    const slot = q[1]!.x - q[0]!.x;
    const innerBottom = q[0]!.y;
    const ticks = verticals(g).filter((v) => v.y1 > innerBottom + 5);
    const xs = ticks.map((v) => v.x);
    // Between Q3 and the second Q1, and half a slot outside the first and last categories.
    expect(xs.some((x) => Math.abs(x - (q[2]!.x + slot / 2)) < 0.1)).toBe(true);
    expect(xs.some((x) => Math.abs(x - (q[0]!.x - slot / 2)) < 0.1)).toBe(true);
    expect(xs.some((x) => Math.abs(x - (q[5]!.x + slot / 2)) < 0.1)).toBe(true);
    // The tick starts at the axis line, above the labels.
    expect(Math.min(...ticks.map((v) => v.y0))).toBeLessThan(innerBottom);
  });

  test("an outer level of an outer level sits further out, and its ticks reach further", () => {
    const g = draw(grouped(multi(years, level([[0, "FY"]]))));
    const t = texts(g);
    expect(t.find((x) => x.text === "FY")!.y).toBeGreaterThan(t.find((x) => x.text === "2023")!.y);
    const ys = verticals(g).map((v) => v.y1);
    expect(Math.max(...ys)).toBeGreaterThan(t.find((x) => x.text === "2023")!.y);
  });

  test("noMultiLvlLbl hides the outer levels and gives their room back", () => {
    const withRows = draw(grouped(multi(years)));
    const without = draw(grouped(multi(years), '<c:noMultiLvlLbl val="1"/>'));
    expect(texts(without).some((x) => x.text === "2023")).toBe(false);
    const bottom = (g: ChartDrawing): number => Math.max(...rectsFilled(g, "#4472C4").map((r) => r.y + r.h));
    expect(bottom(without)).toBeGreaterThan(bottom(withRows));
  });

  test("beside a horizontal bar chart the outer labels turn upright to the left of the inner ones", () => {
    const g = draw(grouped(multi(years), "", "bar"));
    const t = texts(g);
    const label = t.find((x) => x.text === "2023")!;
    expect(label.rotate).toBe(-90);
    expect(label.x).toBeLessThan(t.find((x) => x.text === "Q1")!.x);
  });
});

describe("date axes place points by date", () => {
  const dateSer = (dates: number[], values: number[], fmt = "d-mmm", extra = ""): string =>
    `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${cache("str", ["Units"])}</c:tx>${extra}<c:cat>${cache("num", dates, fmt)}</c:cat><c:val>${cache("num", values)}</c:val></c:ser>`;
  const dateAxes = (axisExtra = "", tag = "c:dateAx"): string =>
    `<${tag}><c:axId val="1"/><c:axPos val="b"/>${axisExtra}<c:crossAx val="2"/></${tag}><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>`;
  const dated = (dates: number[], axisExtra = "", tag = "c:dateAx", group = "barChart", fmt = "d-mmm"): string =>
    chart(`<c:${group}>${dateSer(dates, dates.map((_, i) => i + 1), fmt)}<c:axId val="1"/><c:axId val="2"/></c:${group}>${dateAxes(axisExtra, tag)}`);
  const centers = (g: ChartDrawing): number[] => rectsFilled(g, "#4472C4").map((r) => r.x + r.w / 2);

  // 45000 is 2023-03-15.
  test("columns 0, 1, 31 and 32 days apart land by distance, not evenly", () => {
    const g = draw(dated([45000, 45001, 45031, 45032], '<c:baseTimeUnit val="days"/>'));
    const [a, b, c, d] = centers(g) as [number, number, number, number];
    expect((b - a) / (d - a)).toBeCloseTo(1 / 32, 3);
    expect((c - a) / (d - a)).toBeCloseTo(31 / 32, 3);
    // A column is one base unit wide (gap 150% by default): a day slot over 2.5.
    const widths = rectsFilled(g, "#4472C4").map((r) => r.w);
    expect(widths[0]).toBeCloseTo((d - a) / 32 / 2.5, 2);
  });

  test("the base unit follows the closest pair of dates when the file names none", () => {
    const days = centers(draw(dated([45000, 45001, 45031])));
    expect((days[2]! - days[0]!) / (days[1]! - days[0]!)).toBeCloseTo(31, 1);
    // A month or more apart: months are the slots (March to July is five of them).
    const auto = centers(draw(dated([45000, 45046, 45108])));
    expect((auto[1]! - auto[0]!) / (auto[2]! - auto[0]!)).toBeCloseTo(1 / 4, 1);
    // Named months: two dates in March share a column position.
    const months = centers(draw(dated([45000, 45010, 45046, 45108], '<c:baseTimeUnit val="months"/>')));
    expect(months[0]).toBeCloseTo(months[1]!, 5);
    expect((months[3]! - months[0]!) / (months[2]! - months[0]!)).toBeCloseTo(4, 1);
  });

  test("a catAx on automatic whose data are dates becomes a date axis; a text axis stays by index", () => {
    const auto = readChart(parseXml(dated([45000, 45001, 45031], '<c:auto val="1"/>', "c:catAx")), theme)!;
    expect(auto.axes.find((a) => a.id === 1)!.kind).toBe("date");
    const text = readChart(parseXml(dated([45000, 45001, 45031], '<c:auto val="0"/>', "c:catAx")), theme)!;
    expect(text.axes.find((a) => a.id === 1)!.kind).toBe("cat");
    const evenly = centers(draw(dated([45000, 45001, 45031], '<c:auto val="0"/>', "c:catAx")));
    expect(evenly[1]! - evenly[0]!).toBeCloseTo(evenly[2]! - evenly[1]!, 1);
    const numbers = readChart(parseXml(dated([1, 2, 3], '<c:auto val="1"/>', "c:catAx", "barChart", "General")), theme)!;
    expect(numbers.axes.find((a) => a.id === 1)!.kind).toBe("cat");
  });

  test("lines go by date too, and a point outside the axis range is not drawn", () => {
    const g = draw(dated([45000, 45001, 45031, 45032], '<c:baseTimeUnit val="days"/>', "c:dateAx", "lineChart"));
    const line = paths(g).find((p) => p.stroke?.width === 2.25)!;
    const xs = [...line.path.matchAll(/[ML](-?[\d.]+) /g)].map((m) => Number(m[1]));
    expect(xs).toHaveLength(4);
    // A line chart's points sit mid-slot: 33 slots over the plot.
    expect((xs[1]! - xs[0]!) / (xs[3]! - xs[0]!)).toBeCloseTo(1 / 32, 3);
    const clipped = draw(dated([45000, 45001, 45031, 45032], '<c:scaling><c:max val="45001"/></c:scaling><c:baseTimeUnit val="days"/>'));
    expect(centers(clipped)).toHaveLength(2);
  });

  test("major ticks step by the file's unit: calendar months from the first date, labelled by the axis format", () => {
    const g = draw(dated([45000, 45046, 45092], '<c:baseTimeUnit val="days"/><c:majorUnit val="1"/><c:majorTimeUnit val="months"/>'));
    const t = texts(g);
    const labels = ["15-Mar", "15-Apr", "15-May", "15-Jun"].map((s) => t.find((x) => x.text === s)!);
    expect(labels.every((l) => l !== undefined)).toBe(true);
    // Slots 0, 31, 61 and 92 of 93: their spacing is proportional to the days between.
    const dx = labels.map((l) => l.x - labels[0]!.x);
    expect(dx[1]! / dx[3]!).toBeCloseTo(31 / 92, 3);
    expect(dx[2]! / dx[3]!).toBeCloseTo(61 / 92, 3);
  });

  test("without a major unit the labels thin out to fit the axis, never overlapping", () => {
    const dates = Array.from({ length: 200 }, (_, i) => 45000 + i);
    const t = texts(draw(dated(dates))).filter((x) => /^\d+-\w{3}$/.test(x.text));
    expect(t.length).toBeGreaterThan(2);
    expect(t.length).toBeLessThan(40);
    // As Office does (the oracle's g-date-* cases): turned 45 degrees before stepping up a unit, so
    // neighbours clear each other along the slant, a line height apart, rather than side by side.
    expect(t.every((x) => x.rotate === -45)).toBe(true);
    const lh = 10 * 1.22;
    const xs = t.map((x) => x.x).sort((a, b) => a - b);
    for (let i = 1; i < xs.length; i++) expect(xs[i]! - xs[i - 1]!).toBeGreaterThanOrEqual(lh / Math.SQRT1_2 - 0.5);
  });

  test("a date axis labels every day it can: turned 45 degrees when flat labels would crowd (PowerPoint, g-date-days)", () => {
    const days = Array.from({ length: 10 }, (_, i) => 45292 + i);
    // Room for every day flat in a wide chart; a narrow one turns them rather than skipping days.
    const wide = texts(draw(dated(days))).filter((x) => /^\d+-\w{3}$/.test(x.text));
    expect(wide).toHaveLength(10);
    expect(wide.every((x) => x.rotate === undefined)).toBe(true);
    const narrow = texts(draw(dated(days), { width: 3.2, height: 2.4 })).filter((x) => /^\d+-\w{3}$/.test(x.text));
    expect(narrow).toHaveLength(10);
    expect(narrow.every((x) => x.rotate === -45)).toBe(true);
  });

  test("years between dates give a year axis, labelled by the axis' own format when it has one", () => {
    const g = draw(dated([36526, 40179, 43831], '<c:numFmt formatCode="yyyy" sourceLinked="0"/>'));
    const t = texts(g).map((x) => x.text);
    expect(t).toEqual(expect.arrayContaining(["2000", "2010", "2020"]));
    expect(centers(g)).toHaveLength(3);
  });

  test("the workbook's 1904 flag moves the dates", () => {
    const xml = dated([1, 2], '<c:baseTimeUnit val="days"/>', "c:dateAx", "barChart", "m/d/yyyy").replace("<c:chart>", '<c:date1904 val="1"/><c:chart>');
    expect(texts(draw(xml)).map((x) => x.text)).toEqual(expect.arrayContaining(["1/2/1904", "1/3/1904"]));
    const plain = draw(dated([61, 62], '<c:baseTimeUnit val="days"/>', "c:dateAx", "barChart", "m/d/yyyy"));
    expect(texts(plain).map((x) => x.text)).toEqual(expect.arrayContaining(["3/1/1900", "3/2/1900"]));
  });

  test("a date axis over text categories falls back to the index", () => {
    const xml = chart(`<c:barChart>${ser(0, "A", Q, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${dateAxes()}`);
    expect(texts(draw(xml)).map((x) => x.text)).toEqual(expect.arrayContaining(["Q1", "Q4"]));
  });

  test("a span of decades in days moves to a coarser unit instead of laying out a hundred thousand slots", () => {
    const t0 = performance.now();
    const g = draw(dated([20000, 60000], '<c:baseTimeUnit val="days"/>'));
    expect(performance.now() - t0).toBeLessThan(5000);
    expect(centers(g)).toHaveLength(2);
  });
});

// ---------------------------------------------------------------- spills and overlays

const EPS = 0.5;

function textBoxes(g: ChartDrawing): Array<{ t: ChartTextPrimitive; b: Bounds }> {
  return texts(g).map((t) => ({ t, b: textBox(t) }));
}

function overlapPairs(g: ChartDrawing): number {
  const boxes = textBoxes(g);
  let n = 0;
  for (let i = 0; i < boxes.length; i++) {
    for (let j = i + 1; j < boxes.length; j++) {
      const a = boxes[i]!.b, c = boxes[j]!.b;
      if (a.x0 < c.x1 - EPS && c.x0 < a.x1 - EPS && a.y0 < c.y1 - EPS && c.y0 < a.y1 - EPS) n++;
    }
  }
  return n;
}

/** Every text box and path of a drawing, as offenders against the frame (empty when all inside). */
function outsideFrame(g: ChartDrawing): string[] {
  const out: string[] = [];
  const inside = (b: Bounds): boolean => b.x0 >= -EPS && b.y0 >= -EPS && b.x1 <= g.width + EPS && b.y1 <= g.height + EPS;
  for (const { t, b } of textBoxes(g)) if (!inside(b)) out.push(`text "${t.text}" ${JSON.stringify(b)}`);
  for (const p of paths(g)) {
    const box = pathBox(p.path);
    const b: Bounds = { x0: box.x, y0: box.y, x1: box.x + box.width, y1: box.y + box.height };
    if (!inside(b)) out.push(`path ${p.path.slice(0, 60)} ${JSON.stringify(b)}`);
  }
  return out;
}

const PIE_VALUES = [30, 18, 10, 8, 7, 6, 6, 5, 5, 5];
function pieXml(labels: string[], values = PIE_VALUES, extra = '<c:showLeaderLines val="1"/>', pos = ""): string {
  return chart(`<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:dLbls>${pos}<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/>${extra}</c:dLbls><c:cat>${cache("str", labels)}</c:cat><c:val>${cache("num", values)}</c:val></c:ser></c:pieChart>`);
}
const PIE_NAMES = PIE_VALUES.map((_, i) => `Category ${i + 1}`);

/** The case that overlapped before: ten slices, five of them small, labelled with name and percent. */
const BUSY_PIE = [40, 20, 10, 6, 5, 4, 4, 4, 4, 3];
const LEADER = '<c:showLeaderLines val="1"/><c:leaderLines><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></a:ln></c:spPr></c:leaderLines>';

/** Subpath coordinates of a stroked-only path: [[x0, y0, x1, y1], ...] for each `M a b L c d`. */
function segments(p: ChartPathPrimitive): Array<[number, number, number, number]> {
  return [...p.path.matchAll(/M(-?[\d.]+) (-?[\d.]+)L(-?[\d.]+) (-?[\d.]+)/g)].map((m) => [Number(m[1]), Number(m[2]), Number(m[3]), Number(m[4])]);
}

describe("pie labels", () => {
  test("ten slices: labels spread so none overlaps, where the plain placement overlapped (2 pairs before, 0 after)", () => {
    const g = draw(pieXml(BUSY_PIE.map((_, i) => `Category ${i + 1}`), BUSY_PIE, "", '<c:dLblPos val="outEnd"/>'), { width: 5, height: 3.5 });
    expect(overlapPairs(g)).toBe(0);
    expect(outsideFrame(g)).toEqual([]);
    expect(texts(g).filter((t) => t.text.startsWith("Category"))).toHaveLength(10);
  });

  test("a label that left its slice gets a leader line in the authored style, from the rim to the label", () => {
    const g = draw(pieXml(BUSY_PIE.map((_, i) => `Category ${i + 1}`), BUSY_PIE, LEADER, '<c:dLblPos val="outEnd"/>'), { width: 5, height: 3.5 });
    const leaders = paths(g).filter((p) => p.stroke?.color === "#FF00FF");
    expect(leaders.length).toBeGreaterThan(0);
    expect(leaders[0]!.stroke!.width).toBe(1);
    const pie = paths(g).filter((p) => p.path.includes("A"));
    // A leader starts on the pie's rim: its distance from the centre is the pie's radius.
    const m = pie[0]!.path.match(/^M([\d.]+) ([\d.]+)L([\d.]+) ([\d.]+)A([\d.]+)/)!;
    const [cx, cy, r] = [Number(m[1]), Number(m[2]), Number(m[5])];
    for (const [x0, y0] of leaders.flatMap(segments)) expect(Math.hypot(x0 - cx, y0 - cy)).toBeCloseTo(r, 0);
    expect(overlapPairs(g)).toBe(0);
  });

  test("no showLeaderLines, no leader lines", () => {
    const g = draw(pieXml(BUSY_PIE.map((_, i) => `Category ${i + 1}`), BUSY_PIE, "", '<c:dLblPos val="outEnd"/>'), { width: 5, height: 3.5 });
    expect(paths(g).filter((p) => p.stroke !== undefined && !p.path.includes("A") && !p.path.includes("h"))).toHaveLength(0);
  });

  test("a label too big for its slice moves outside, and wide labels shrink the pie instead of leaving the chart", () => {
    const names = BUSY_PIE.map((_, i) => `Category with a long name ${i + 1}`);
    // 4.6in tall: labels wrap at a fifth of the plot (PowerPoint's width), three lines each, and a
    // shorter chart holds a column of them only overlapping, as PowerPoint's does (Known limits).
    const g = draw(pieXml(names, BUSY_PIE, LEADER, '<c:dLblPos val="ctr"/>'), { width: 6, height: 4.6 });
    expect(outsideFrame(g)).toEqual([]);
    expect(overlapPairs(g)).toBe(0);
    // Centred labels this wide cannot sit in their slices: the small slices' labels left them, joined by leader lines.
    expect(paths(g).some((p) => p.stroke?.color === "#FF00FF")).toBe(true);
    // Even the 3-line wrap of a long label stays within its share of the plot.
    const widest = Math.max(...texts(g).filter((t) => t.text.startsWith("Category")).map((t) => textBox(t).x1 - textBox(t).x0));
    expect(widest).toBeLessThan(g.width * 0.3);
    // The same chart in a frame too small for any pie shrinks the pie, never the labels' place in the frame.
    expect(outsideFrame(draw(pieXml(names, BUSY_PIE, LEADER, '<c:dLblPos val="ctr"/>'), { width: 3, height: 3 }))).toEqual([]);
  });
});

describe("legend", () => {
  const many = (n: number, name = (i: number) => `Series ${i + 1}`): string => Array.from({ length: n }, (_, i) => ser(i, name(i), Q, [1, 2, 3, 4])).join("");
  const barChart = (n: number, legendPos: string, name?: (i: number) => string, rest = ""): string =>
    chart(`<c:barChart>${many(n, name)}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`, rest).replace("</c:plotArea>", `</c:plotArea><c:legend><c:legendPos val="${legendPos}"/></c:legend>`);

  test("20 series: a side legend keeps its width to a third of the chart and drops what does not fit, the plot keeps its room", () => {
    const g = draw(barChart(20, "r"));
    expect(outsideFrame(g)).toEqual([]);
    const names = texts(g).filter((t) => t.text.startsWith("Series"));
    expect(names.length).toBeGreaterThan(5);
    expect(names.length).toBeLessThan(20);
    for (const t of names) expect(t.x).toBeGreaterThan(g.width * (2 / 3) - 10);
    const bars = rectsFilled(g, "#4472C4");
    expect(Math.max(...bars.map((b) => b.x + b.w))).toBeLessThan(Math.min(...names.map((t) => t.x)));
  });

  test("20 series along the bottom wrap into rows no taller than a third of the chart", () => {
    const g = draw(barChart(20, "b"));
    expect(outsideFrame(g)).toEqual([]);
    const names = texts(g).filter((t) => t.text.startsWith("Series"));
    const rows = new Set(names.map((t) => t.y));
    expect(rows.size).toBeGreaterThan(1);
    expect(Math.min(...names.map((t) => t.y))).toBeGreaterThan(g.height * (2 / 3) - 10);
  });

  test("long names wrap in a side legend and are cut in a bottom one", () => {
    const long = (i: number): string => `Quarterly revenue from the region number ${i + 1} excluding returns`;
    const side = draw(barChart(3, "r", long));
    expect(outsideFrame(side)).toEqual([]);
    expect(texts(side).filter((t) => t.text.includes("Quarterly")).length).toBe(3);
    expect(texts(side).some((t) => t.text.includes("excluding") || t.text.includes("returns"))).toBe(true);
    const bottom = draw(barChart(2, "b", (i) => `Name ${i + 1} ${"x".repeat(400)}`));
    expect(outsideFrame(bottom)).toEqual([]);
    expect(texts(bottom).some((t) => t.text.endsWith("..."))).toBe(true);
  });

  test("a manual legend box past the frame is pulled back in", () => {
    const xml = barChart(3, "r").replace("<c:legendPos val=\"r\"/>", '<c:legendPos val="r"/><c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.9"/><c:y val="0.95"/><c:w val="0.5"/><c:h val="0.4"/></c:manualLayout></c:layout>');
    expect(outsideFrame(draw(xml))).toEqual([]);
  });
});

describe("nothing is drawn outside the frame", () => {
  const stress: Record<string, string> = {
    "20 series": chart(`<c:lineChart>${Array.from({ length: 20 }, (_, i) => ser(i, `Series ${i + 1}`, Q, [i, i + 1, i + 2, i + 3])).join("")}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${axes()}`)
      .replace("</c:plotArea>", '</c:plotArea><c:legend><c:legendPos val="r"/></c:legend>'),
    "long series names": chart(`<c:barChart>${Array.from({ length: 4 }, (_, i) => ser(i, `A series whose name keeps going ${"and going ".repeat(8)}${i}`, Q, [1, 2, 3, 4])).join("")}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`)
      .replace("</c:plotArea>", '</c:plotArea><c:legend><c:legendPos val="t"/></c:legend>'),
    "busy pie with long labels": pieXml(BUSY_PIE.map((_, i) => `Category with a rather long name ${i + 1}`), BUSY_PIE, LEADER, '<c:dLblPos val="outEnd"/>'),
    "long title": chart(`<c:barChart>${ser(0, "A", Q, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`,
      `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${"A title that runs on ".repeat(12)}</a:t></a:r></a:p></c:rich></c:tx></c:title>`),
    "one unbreakable title word": chart(`<c:barChart>${ser(0, "A", Q, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`,
      `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${"W".repeat(120)}</a:t></a:r></a:p></c:rich></c:tx></c:title>`),
    "labels on the tallest bars": chart(`<c:barChart>${ser(0, "A", Q, [10, 10, 10, 10], '<c:dLbls><c:dLblPos val="outEnd"/><c:showVal val="1"/></c:dLbls>')}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes("", '<c:scaling><c:max val="10"/><c:min val="0"/></c:scaling>')}`),
    "manual plot beyond the frame": chart(`<c:barChart>${ser(0, "A", Q, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`)
      .replace("<c:layout/>", '<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.5"/><c:y val="0.5"/><c:w val="0.9"/><c:h val="0.9"/></c:manualLayout></c:layout>'),
  };
  for (const [name, xml] of Object.entries(stress)) {
    test(name, () => {
      expect(outsideFrame(draw(xml, { width: 5, height: 3 }))).toEqual([]);
    });
  }
});

describe("trendlines", () => {
  const scatter = (trend: string, ys = [3, 5, 7, 9, 11], err = ""): string =>
    chart(`<c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:ln w="9525"><a:noFill/></a:ln></c:spPr>${trend}${err}<c:xVal>${cache("num", [1, 2, 3, 4, 5])}</c:xVal><c:yVal>${cache("num", ys)}</c:yVal></c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart>` +
      `<c:valAx><c:axId val="1"/><c:scaling><c:max val="10"/><c:min val="0"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:scaling><c:max val="20"/><c:min val="0"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>`);
  const TREND_LINE = '<c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></a:ln></c:spPr>';
  const route = (xml: string) => {
    const read = readChartModel(parseXml(xml), theme);
    if (!("spec" in read)) throw new Error(read.unsupportedReason);
    const drawn = drawChart(read.spec, 576, 324);
    const px = (x: number) => drawn.plot.x + (x / 10) * drawn.plot.w;
    const py = (y: number) => drawn.plot.y + drawn.plot.h - (y / 20) * drawn.plot.h;
    return { spec: read.spec, primitives: drawn.primitives, px, py, line: drawn.primitives.filter((p): p is ChartPathPrimitive => p.kind === "path" && p.stroke?.color === "#00AA00") };
  };

  test("linear fit: a straight path through the fitted line at both ends, extended forward and backward", () => {
    const r = route(scatter(`<c:trendline>${TREND_LINE}<c:trendlineType val="linear"/><c:forward val="2"/><c:backward val="0.5"/></c:trendline>`));
    expect(r.spec.groups[0]!.series[0]!.trendlines![0]).toMatchObject({ type: "linear", forward: 2, backward: 0.5 });
    expect(r.line).toHaveLength(1);
    const [x0, y0, x1, y1] = segments(r.line[0]!)[0]!;
    // y = 2x + 1 from x = 0.5 to x = 7.
    expect(x0).toBeCloseTo(r.px(0.5), 1);
    expect(y0).toBeCloseTo(r.py(2), 1);
    expect(x1).toBeCloseTo(r.px(7), 1);
    expect(y1).toBeCloseTo(r.py(15), 1);
  });

  test("a trendline is clipped to the plot where it would leave it", () => {
    const r = route(scatter(`<c:trendline>${TREND_LINE}<c:trendlineType val="linear"/><c:forward val="9"/></c:trendline>`));
    const [, , x1, y1] = segments(r.line[0]!)[0]!;
    // y = 2x + 1 reaches the axis maximum (20) at x = 9.5, before x = 14.
    expect(y1).toBeCloseTo(r.py(20), 1);
    expect(x1).toBeCloseTo(r.px(9.5), 1);
  });

  test("equation and R-squared labels", () => {
    const r = route(scatter(`<c:trendline>${TREND_LINE}<c:trendlineType val="linear"/><c:dispRSqr val="1"/><c:dispEq val="1"/></c:trendline>`));
    const t = r.primitives.filter((p): p is ChartTextPrimitive => p.kind === "text").map((p) => p.text);
    expect(t).toContain("y = 2x + 1");
    expect(t.some((x) => x.startsWith("R") && x.endsWith("= 1.0000"))).toBe(true);
  });

  test("the fits recover their curves", () => {
    const xs = [1, 2, 3, 4, 5, 6];
    const base = { order: 2, period: 2, forward: 0, backward: 0, dispEq: false, dispRSqr: false, line: { color: "#000", width: 1 }, labelText: { size: 10, bold: false, italic: false, color: "#000", family: "Calibri" } };
    const quad = fitTrend({ ...base, type: "poly" }, xs, xs.map((x) => 3 + 2 * x + 0.5 * x * x))!;
    expect(quad.coefficients.map((c) => Math.round(c * 1e6) / 1e6)).toEqual([3, 2, 0.5]);
    const exp = fitTrend({ ...base, type: "exp" }, xs, xs.map((x) => 2 * Math.exp(0.3 * x)))!;
    expect(exp.at(7)).toBeCloseTo(2 * Math.exp(2.1), 6);
    expect(exp.r2).toBeCloseTo(1, 9);
    const pow = fitTrend({ ...base, type: "power" }, xs, xs.map((x) => 4 * x ** 1.5))!;
    expect(pow.at(9)).toBeCloseTo(4 * 27, 6);
    const log = fitTrend({ ...base, type: "log" }, xs, xs.map((x) => 1 + 3 * Math.log(x)))!;
    expect(log.at(10)).toBeCloseTo(1 + 3 * Math.log(10), 6);
    expect(movingAverage(xs, [2, 4, 6, 8, 10, 12], 3)).toEqual([[3, 4], [4, 6], [5, 8], [6, 10]]);
  });

  test("a trendline appears in the legend after the series, named as Office names it", () => {
    const xml = scatter(`<c:trendline><c:trendlineType val="linear"/></c:trendline>`).replace("</c:plotArea>", '</c:plotArea><c:legend><c:legendPos val="b"/></c:legend>');
    expect(texts(draw(xml)).map((t) => t.text)).toContain("Linear (Series1)");
  });
});

describe("error bars", () => {
  const bars = (err: string, vals = [4, 3, 2, 1]): string =>
    chart(`<c:barChart><c:barDir val="col"/>${ser(0, "A", Q, vals, err)}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes("", '<c:scaling><c:max val="10"/><c:min val="0"/></c:scaling><c:majorGridlines/>')}`);
  const BAR_LN = '<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln></c:spPr>';
  const route = (xml: string) => {
    const read = readChartModel(parseXml(xml), theme);
    if (!("spec" in read)) throw new Error(read.unsupportedReason);
    const drawn = drawChart(read.spec, 576, 324);
    const bar = drawn.primitives.find((p): p is ChartPathPrimitive => p.kind === "path" && p.stroke?.color === "#123456");
    const seg = bar === undefined ? [] : segments(bar);
    return { plot: drawn.plot, vertical: seg.filter(([x0, , x1]) => Math.abs(x0 - x1) < 0.01), caps: seg.filter(([, y0, , y1]) => Math.abs(y0 - y1) < 0.01) };
  };

  test("fixed value: every bar spans value +/- 1, with end caps", () => {
    const r = route(bars(`<c:errBars><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:noEndCap val="0"/><c:val val="1"/>${BAR_LN}</c:errBars>`));
    expect(r.vertical).toHaveLength(4);
    for (const [, y0, , y1] of r.vertical) expect(Math.abs(y1 - y0)).toBeCloseTo((2 / 10) * r.plot.h, 1);
    expect(r.caps).toHaveLength(8);
    // The first bar (4) is centred on 4: its middle sits at value 4.
    const [, y0, , y1] = r.vertical[0]!;
    expect((y0 + y1) / 2).toBeCloseTo(r.plot.y + r.plot.h - (4 / 10) * r.plot.h, 1);
  });

  test("percentage, plus only, no caps: length = value x 50%, one side", () => {
    const r = route(bars(`<c:errBars><c:errBarType val="plus"/><c:errValType val="percentage"/><c:noEndCap val="1"/><c:val val="50"/>${BAR_LN}</c:errBars>`));
    expect(r.caps).toHaveLength(0);
    const lengths = r.vertical.map(([, y0, , y1]) => Math.abs(y1 - y0));
    [4, 3, 2, 1].forEach((v, i) => expect(lengths[i]).toBeCloseTo((v * 0.5 / 10) * r.plot.h, 1));
  });

  test("custom amounts come from the plus and minus caches", () => {
    const r = route(bars(`<c:errBars><c:errBarType val="both"/><c:errValType val="cust"/><c:noEndCap val="1"/><c:plus>${cache("num", [1, 1, 1, 1])}</c:plus><c:minus>${cache("num", [0, 0, 2, 0])}</c:minus>${BAR_LN}</c:errBars>`));
    const lengths = r.vertical.map(([, y0, , y1]) => Math.abs(y1 - y0));
    expect(lengths[0]).toBeCloseTo((1 / 10) * r.plot.h, 1);
    expect(lengths[2]).toBeCloseTo((3 / 10) * r.plot.h, 1);
  });

  test("a bar cut off at the plot edge loses its cap on that end", () => {
    const r = route(bars(`<c:errBars><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="5"/>${BAR_LN}</c:errBars>`, [8, 8, 8, 8]));
    // Value 8 + 5 passes the axis maximum (10): the top is cut, the bottom (3) keeps its cap.
    expect(r.caps).toHaveLength(4);
    for (const [, y0, , y1] of r.vertical) expect(Math.min(y0, y1)).toBeGreaterThanOrEqual(r.plot.y - 0.01);
  });


  test("standard deviation bars are centred on the series mean", () => {
    const eb = { dir: "y" as const, type: "both" as const, valType: "stdDev" as const, value: 1, noEndCap: false, line: { color: "#000", width: 1 } };
    const a = errorAmounts(eb, [2, 4, 6, 8])!;
    expect(a[0]!.centre).toBe(5);
    expect(a[0]!.plus).toBeCloseTo(Math.sqrt(20 / 3), 9);
  });
});
