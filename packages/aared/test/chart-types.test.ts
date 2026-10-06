import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { readWorkbook } from "../src/core/xlsx/workbook.js";
import { parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { parseXml } from "../src/core/xml.js";
import { readChartModel } from "../src/core/chart.js";
import { readChart } from "../src/core/chart/read.js";
import { layoutChart } from "../src/core/chart/draw.js";
import type { ChartDrawing } from "../src/core/chart/draw.js";
import type { ChartPathPrimitive, ChartTextPrimitive } from "../src/core/chart/scene.js";
import type { ChartSpecModel } from "../src/core/chart/spec.js";
import { themeForColorMap } from "../src/core/theme.js";

// Stock and of-pie charts as Excel wrote them: tools/chart-oracle/types-oracle.xlsx
// (made by gen_types.py) holds one chart per sheet, over eight days of prices or
// eight shares. The numbers asserted below are the ones Excel's own export of
// those charts shows (measured once, see the chart-oracle README).

const ORACLE = `${import.meta.dir}/../../../tools/chart-oracle/types-oracle.xlsx`;

interface Loaded { spec: ChartSpecModel; drawing: ChartDrawing }

async function load(sheetName: string): Promise<Loaded> {
  const pkg = await OpcPackage.open(await Bun.file(ORACLE).arrayBuffer());
  const workbook = await readWorkbook(pkg);
  const index = workbook.sheets.findIndex((sheet) => sheet.name === sheetName);
  if (index < 0) throw new Error(`oracle structure changed: no sheet ${sheetName}`);
  const sheet = await parseXlsxSheet(pkg, workbook, index);
  const chart = sheet.drawings[0];
  if (chart?.kind !== "chart" || chart.spec === undefined) throw new Error(`oracle structure changed: sheet ${sheetName} lost its chart (${chart?.kind === "chart" ? chart.unsupportedReason : "none"})`);
  return { spec: chart.spec, drawing: layoutChart(chart.spec, chart.bounds) };
}

const paths = (g: ChartDrawing) => g.primitives.filter((p): p is ChartPathPrimitive => p.kind === "path");
const texts = (g: ChartDrawing) => g.primitives.filter((p): p is ChartTextPrimitive => p.kind === "text");

/** Every `M x y h w v h h -w Z` subpath of the paths filled with `color` and stroked with `stroke`. */
function rects(g: ChartDrawing, fill: string, stroke?: string): Array<{ x: number; y: number; w: number; h: number }> {
  const out: Array<{ x: number; y: number; w: number; h: number }> = [];
  for (const p of paths(g)) {
    if (p.fill !== fill || (stroke !== undefined && p.stroke?.color !== stroke)) continue;
    for (const m of p.path.matchAll(/M(-?[\d.]+) (-?[\d.]+)h(-?[\d.]+)v(-?[\d.]+)h-?[\d.]+Z/g)) {
      out.push({ x: Number(m[1]), y: Number(m[2]), w: Number(m[3]), h: Number(m[4]) });
    }
  }
  return out;
}

/** Every straight segment `M x0 y0 L x1 y1` drawn in `color`, in paint order. */
function segments(g: ChartDrawing, color: string): Array<[number, number, number, number]> {
  const out: Array<[number, number, number, number]> = [];
  for (const p of paths(g)) {
    if (p.stroke?.color !== color) continue;
    for (const m of p.path.matchAll(/M(-?[\d.]+) (-?[\d.]+)L(-?[\d.]+) (-?[\d.]+)(?=$|\s|M)/g)) out.push([Number(m[1]), Number(m[2]), Number(m[3]), Number(m[4])]);
  }
  return out;
}

/** The y of each horizontal gridline, top to bottom. */
function gridYs(g: ChartDrawing, color: string): number[] {
  const ys: number[] = [];
  for (const p of paths(g)) {
    if (p.stroke?.color !== color) continue;
    for (const m of p.path.matchAll(/M(-?[\d.]+) (-?[\d.]+)H(-?[\d.]+)/g)) ys.push(Number(m[2]));
  }
  return ys.sort((a, b) => a - b);
}

/** A mapping from value to y, from the value axis' own gridlines (`ticks` values, the lowest at the bottom). */
function valueToY(g: ChartDrawing, spec: ChartSpecModel, axisId: number, lo: number, hi: number): (v: number) => number {
  const axis = spec.axes.find((a) => a.id === axisId)!;
  const ys = gridYs(g, axis.majorGrid!.color);
  const top = ys[0]!;
  const bottom = ys[ys.length - 1]!;
  return (v) => bottom - ((v - lo) / (hi - lo)) * (bottom - top);
}

// Excel's days: open, high, low, close per category.
const OHLC = [
  [20.0, 24.0, 18.5, 23.0], [23.0, 25.5, 21.0, 21.5], [21.5, 22.5, 19.0, 22.0], [22.0, 26.0, 21.5, 25.0],
  [25.0, 27.0, 22.5, 23.0], [23.0, 24.0, 20.0, 20.5], [20.5, 23.5, 20.0, 23.0], [23.0, 28.0, 22.0, 27.0],
] as const;

describe("stock charts", () => {
  test("open-high-low-close: one line per day from its high to its low, on the value axis' own scale", async () => {
    const { spec, drawing } = await load("OHLC");
    const group = spec.groups[0]!;
    expect(group.kind).toBe("stock");
    expect(group.series.map((s) => s.name)).toEqual(["Open", "High", "Low", "Close"]);
    // Series lines are noFill in Excel's files: only the lines between series show.
    expect(group.series.every((s) => s.look.line === undefined)).toBe(true);
    const y = valueToY(drawing, spec, group.axisIds![1], 0, 30);
    const line = paths(drawing).find((p) => p.stroke?.color === group.hiLowLines!.color && /V/.test(p.path))!;
    const spans = [...line.path.matchAll(/M(-?[\d.]+) (-?[\d.]+)V(-?[\d.]+)/g)].map((m) => ({ x: Number(m[1]), top: Number(m[2]), bottom: Number(m[3]) }));
    expect(spans).toHaveLength(8);
    OHLC.forEach(([, high, low], i) => {
      expect(spans[i]!.top).toBeCloseTo(y(high), 0);
      expect(spans[i]!.bottom).toBeCloseTo(y(low), 0);
    });
    // One per category slot, evenly spaced.
    const step = spans[1]!.x - spans[0]!.x;
    spans.forEach((s, i) => expect(s.x - spans[0]!.x).toBeCloseTo(i * step, 1));
  });

  test("up bars where the close is above the open, down bars where it is below, between open and close", async () => {
    const { spec, drawing } = await load("OHLC");
    const bars = spec.groups[0]!.upDownBars!;
    expect(bars.gapWidth).toBe(150);
    // Excel's defaults, resolved: white up bars and dark down bars, both with a grey outline.
    expect(bars.up.fill).toBe("#FFFFFF");
    expect(bars.down.fill).not.toBe(bars.up.fill);
    const y = valueToY(drawing, spec, spec.groups[0]!.axisIds![1], 0, 30);
    const up = rects(drawing, bars.up.fill as string, bars.up.line!.color);
    const down = rects(drawing, bars.down.fill as string, bars.down.line!.color);
    const upDays = OHLC.flatMap(([o, , , c], i) => (c >= o ? [i] : []));
    const downDays = OHLC.flatMap(([o, , , c], i) => (c < o ? [i] : []));
    expect(up).toHaveLength(upDays.length);
    expect(down).toHaveLength(downDays.length);
    for (const [bar, days] of [[up, upDays], [down, downDays]] as const) {
      bar.forEach((r, k) => {
        const [open, , , close] = OHLC[days[k]!]!;
        expect(r.y).toBeCloseTo(y(Math.max(open, close)), 0);
        expect(r.y + r.h).toBeCloseTo(y(Math.min(open, close)), 0);
      });
    }
    // Slot / (1 + gap): 1/2.5 of the category slot (Excel's export: a 30 px bar in a 73.6 px slot).
    const slot = (up[1]!.x + up[1]!.w / 2 - (up[0]!.x + up[0]!.w / 2)) / (upDays[1]! - upDays[0]!);
    expect(up[0]!.w / slot).toBeCloseTo(0.4, 2);
  });

  test("paint order: the high-low line, then the up-down bars over it", async () => {
    const { spec, drawing } = await load("OHLC");
    const group = spec.groups[0]!;
    const at = (match: (p: ChartPathPrimitive) => boolean): number => drawing.primitives.findIndex((p) => p.kind === "path" && match(p));
    const hiLow = at((p) => p.stroke?.color === group.hiLowLines!.color && /V/.test(p.path));
    const up = at((p) => p.fill === group.upDownBars!.up.fill && p.stroke?.color === group.upDownBars!.up.line!.color);
    const down = at((p) => p.fill === group.upDownBars!.down.fill && p.stroke?.color === group.upDownBars!.down.line!.color);
    expect(hiLow).toBeGreaterThan(-1);
    expect(up).toBeGreaterThan(hiLow);
    expect(down).toBeGreaterThan(hiLow);
  });

  test("the close's marker is painted over the high-low line", async () => {
    const { spec, drawing } = await load("HLC");
    const group = spec.groups[0]!;
    // Three series: no up-down bars; the close carries a dot marker, the others none.
    expect(group.upDownBars).toBeUndefined();
    expect(group.series.map((s) => s.marker.symbol)).toEqual(["none", "none", "dot"]);
    const order = drawing.primitives.map((p) => (p.kind === "path" ? (p.stroke?.color === group.hiLowLines!.color ? "hilow" : p.fill === group.series[2]!.marker.look.fill ? "marker" : "") : ""));
    expect(order.indexOf("hilow")).toBeGreaterThan(-1);
    expect(order.indexOf("marker")).toBeGreaterThan(order.indexOf("hilow"));
    const dots = paths(drawing).find((p) => p.fill === group.series[2]!.marker.look.fill && p.path.includes("A"))!;
    expect(dots.path.match(/Z/g)).toHaveLength(8);
  });

  test("Arabic category labels sit under the days", async () => {
    const { drawing } = await load("HLC");
    expect(texts(drawing).map((t) => t.text)).toEqual(expect.arrayContaining(["الأحد", "الثلاثاء 2"]));
  });

  test("volume variants: the volume columns use the left axis, the prices a second axis on the right", async () => {
    for (const name of ["VHLC", "VOHLC"]) {
      const { spec, drawing } = await load(name);
      expect(spec.groups.map((g) => g.kind)).toEqual(["bar", "stock"]);
      const [volume, stock] = spec.groups;
      expect(volume!.axisIds![1]).not.toBe(stock!.axisIds![1]);
      const labels = texts(drawing);
      const left = labels.find((t) => t.text === "3000")!;
      const right = labels.find((t) => t.text === "30")!;
      expect(right.x).toBeGreaterThan(left.x);
      expect(right.textAnchor).toBe("start");
      // The columns are volume (1,200 of 3,000) and the hi-low lines are prices (18.5-24 of 30): both read off their own axis.
      const yVolume = valueToY(drawing, spec, volume!.axisIds![1], 0, 3000);
      const column = rects(drawing, volume!.series[0]!.look.fill as string)[0]!;
      expect(column.y).toBeCloseTo(yVolume(1200), 0);
      const yPrice = valueToY(drawing, spec, stock!.axisIds![1], 0, 30);
      const line = paths(drawing).find((p) => p.stroke?.color === stock!.hiLowLines!.color && /V/.test(p.path))!;
      const first = line.path.match(/M(-?[\d.]+) (-?[\d.]+)V(-?[\d.]+)/)!;
      expect(Number(first[2])).toBeCloseTo(yPrice(24), 0);
      expect(Number(first[3])).toBeCloseTo(yPrice(18.5), 0);
    }
  });

  test("the model is plain data", async () => {
    const { spec } = await load("VOHLC");
    expect(JSON.parse(JSON.stringify(spec))).toEqual(spec);
  });
});

// ------------------------------------------------------------------ of-pie

interface Sector { cx: number; cy: number; r: number; start: number; sweep: number; fill: unknown }

/** The pie sectors of a drawing in paint order, with angles in SVG degrees (0 at 3 o'clock, clockwise). */
function sectors(g: ChartDrawing): Sector[] {
  const out: Sector[] = [];
  for (const p of paths(g)) {
    const m = p.path.match(/^M(-?[\d.]+) (-?[\d.]+)L(-?[\d.]+) (-?[\d.]+)A(-?[\d.]+) [\d.]+ 0 [01] 1 (-?[\d.]+) (-?[\d.]+)Z$/);
    if (m === null) continue;
    const [cx, cy, x0, y0, r, x1, y1] = m.slice(1).map(Number) as [number, number, number, number, number, number, number];
    const a0 = (Math.atan2(y0 - cy, x0 - cx) * 180) / Math.PI;
    const a1 = (Math.atan2(y1 - cy, x1 - cx) * 180) / Math.PI;
    out.push({ cx, cy, r, start: a0, sweep: (((a1 - a0) % 360) + 360) % 360, fill: p.fill });
  }
  return out;
}

const circular = (a: number): number => ((((a + 180) % 360) + 360) % 360) - 180;

describe("pie of pie", () => {
  test("Excel's automatic split moves the last third of the points: 8 points leave 3", async () => {
    const { spec } = await load("PieOfPie");
    const group = spec.groups[0]!;
    expect(group.kind).toBe("ofPie");
    expect(group.ofPie!.secondary).toEqual([5, 6, 7]);
    expect(group.ofPie!.type).toBe("pie");
    expect(group.ofPie!.secondSize).toBe(75);
    expect(group.gapWidth).toBe(100);
    // The Other slice is the series' ninth styled point, as Excel writes it; the legend does not list it.
    expect(group.ofPie!.other.fill).toBe(group.series[0]!.points[8]!.look!.fill);
  });

  test("the Other slice is the share of the split points, centred on 3 o'clock", async () => {
    const { spec, drawing } = await load("PieOfPie");
    const all = sectors(drawing);
    expect(all).toHaveLength(9);
    const main = all.slice(0, 6);
    const other = main[5]!;
    expect(other.fill).toBe(spec.groups[0]!.ofPie!.other.fill);
    // 4 + 2 + 1 of 100.
    expect(other.sweep).toBeCloseTo(0.07 * 360, 0);
    expect(circular(other.start + other.sweep / 2)).toBeCloseTo(0, 0);
    // The first slice starts just below it, and the main slices add up to a whole turn.
    expect(main[0]!.start).toBeCloseTo(other.sweep / 2, 0);
    expect(main.reduce((sum, s) => sum + s.sweep, 0)).toBeCloseTo(360, 0);
    // The share of each ordinary slice: 38, 25, 14, 9, 7 of 100.
    expect(main.slice(0, 5).map((s) => Math.round(s.sweep / 3.6))).toEqual([38, 25, 14, 9, 7]);
  });

  test("the second pie holds exactly the split points, from the same start angle, 75% as wide", async () => {
    const { drawing } = await load("PieOfPie");
    const all = sectors(drawing);
    const main = all.slice(0, 6);
    const second = all.slice(6);
    expect(second).toHaveLength(3);
    // 4, 2, 1 of 7, in source order.
    expect(second.map((s) => Math.round((s.sweep / 360) * 7))).toEqual([4, 2, 1]);
    expect(second[0]!.start).toBeCloseTo(main[0]!.start, 0);
    const r1 = main[0]!.r;
    const r2 = second[0]!.r;
    expect(r2 / r1).toBeCloseTo(0.75, 2);
    // gapWidth 100: one main radius between the circles; both centred on one line.
    expect(second[0]!.cy).toBeCloseTo(main[0]!.cy, 1);
    expect(second[0]!.cx - main[0]!.cx - r1 - r2).toBeCloseTo(r1, 0);
  });

  test("the series lines run from the Other slice's rim corners tangent to the second pie", async () => {
    const { spec, drawing } = await load("PieOfPie");
    const all = sectors(drawing);
    const main = all.slice(0, 6)[0]!;
    const second = all[6]!;
    const half = (all[5]!.sweep / 2 * Math.PI) / 180;
    const color = spec.groups[0]!.serLines!.color;
    const lines = segments(drawing, color);
    expect(lines).toHaveLength(2);
    lines.forEach(([x0, y0, x1, y1], k) => {
      const sign = k === 0 ? -1 : 1;
      // Starts on the main rim at the Other slice's edge ...
      expect(x0).toBeCloseTo(main.cx + main.r * Math.cos(sign * half), 1);
      expect(y0).toBeCloseTo(main.cy + main.r * Math.sin(sign * half), 1);
      // ... and ends where it touches the second circle, square to its radius.
      expect(Math.hypot(x1 - second.cx, y1 - second.cy)).toBeCloseTo(second.r, 0);
      expect((x1 - x0) * (x1 - second.cx) + (y1 - y0) * (y1 - second.cy)).toBeCloseTo(0, -1);
    });
    // Excel's export measured the slope at -0.326 / +0.32.
    const upper = lines[0]!;
    expect((upper[3] - upper[1]) / (upper[2] - upper[0])).toBeCloseTo(-0.32, 1);
  });
});

describe("bar of pie", () => {
  test("the second plot is one bar of the split points, first point on top, as tall as 75% of the pie", async () => {
    const { spec, drawing } = await load("BarOfPie");
    const group = spec.groups[0]!;
    expect(group.ofPie!.type).toBe("bar");
    const all = sectors(drawing);
    expect(all).toHaveLength(6);
    const r1 = all[0]!.r;
    const fills = [5, 6, 7].map((i) => group.series[0]!.points[i]!.look!.fill as string);
    const stack = fills.map((f) => rects(drawing, f)[0]!);
    // Stacked top to bottom in source order: 4, 2, 1 of 7.
    expect(stack[0]!.y).toBeLessThan(stack[1]!.y);
    expect(stack[1]!.y).toBeLessThan(stack[2]!.y);
    expect(stack.map((r) => Math.round((r.h / (stack[0]!.h + stack[1]!.h + stack[2]!.h)) * 7))).toEqual([4, 2, 1]);
    const height = stack.reduce((sum, r) => sum + r.h, 0);
    expect(height / (2 * r1)).toBeCloseTo(0.75, 2);
    // Half as wide as it is tall; gapWidth 100 puts half a main radius before it.
    expect(stack[0]!.w / height).toBeCloseTo(0.5, 2);
    expect(stack[0]!.x - (all[0]!.cx + r1)).toBeCloseTo(r1 / 2, 0);
    // Centred on the pie's centre line.
    expect(stack[0]!.y + height / 2).toBeCloseTo(all[0]!.cy, 1);
  });

  test("the series lines join the Other slice's rim corners to the bar's top and bottom left corners", async () => {
    const { spec, drawing } = await load("BarOfPie");
    const all = sectors(drawing);
    const main = all[0]!;
    const half = (all[5]!.sweep / 2 * Math.PI) / 180;
    const top = rects(drawing, spec.groups[0]!.series[0]!.points[5]!.look!.fill as string)[0]!;
    const bottom = rects(drawing, spec.groups[0]!.series[0]!.points[7]!.look!.fill as string)[0]!;
    const color = spec.groups[0]!.serLines!.color;
    const pts = segments(drawing, color);
    expect(pts).toHaveLength(2);
    expect(pts[0]![0]).toBeCloseTo(main.cx + main.r * Math.cos(-half), 1);
    expect(pts[0]![1]).toBeCloseTo(main.cy + main.r * Math.sin(-half), 1);
    expect([pts[0]![2], pts[0]![3]]).toEqual([expect.closeTo(top.x, 1), expect.closeTo(top.y, 1)]);
    expect([pts[1]![2], pts[1]![3]]).toEqual([expect.closeTo(bottom.x, 1), expect.closeTo(bottom.y + bottom.h, 1)]);
  });
});

// ------------------------------------------------------------------ the split rules, on synthetic parts

const theme = themeForColorMap({
  colors: { dk1: "#000000", lt1: "#FFFFFF", accent1: "#4472C4", accent2: "#ED7D31", accent3: "#A5A5A5", accent4: "#FFC000", accent5: "#5B9BD5", accent6: "#70AD47" },
  fonts: { major: { latin: "Calibri Light", complexScript: "", scripts: {} }, minor: { latin: "Calibri", complexScript: "", scripts: {} } },
}, undefined);

function ofPie(values: number[], tail = ""): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
  const cats = values.map((_, i) => `<c:pt idx="${i}"><c:v>c${i}</c:v></c:pt>`).join("");
  return `<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/><c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/>` +
    `<c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:strCache><c:ptCount val="${values.length}"/>${cats}</c:strCache></c:strRef></c:cat>` +
    `<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numCache></c:numRef></c:val></c:ser>${tail}</c:ofPieChart></c:plotArea></c:chart></c:chartSpace>`;
}

function secondary(values: number[], tail = ""): number[] {
  return readChart(parseXml(ofPie(values, tail)), theme)!.groups[0]!.ofPie!.secondary;
}

describe("of-pie split rules", () => {
  const data = [38, 25, 14, 9, 7, 4, 2, 1];
  test("automatic: the last third, rounded up (Excel: 2-3 points leave 1, 4-6 leave 2, 7-9 leave 3, 10-12 leave 4)", () => {
    expect(secondary([5, 6]).length).toBe(1);
    expect(secondary([5, 6, 7]).length).toBe(1);
    expect(secondary([1, 2, 3, 4]).length).toBe(2);
    expect(secondary(data)).toEqual([5, 6, 7]);
    expect(secondary(Array.from({ length: 10 }, (_, i) => i + 1)).length).toBe(4);
  });
  test("by position, value, percentage and custom list", () => {
    expect(secondary(data, '<c:splitType val="pos"/><c:splitPos val="2"/>')).toEqual([6, 7]);
    expect(secondary(data, '<c:splitType val="val"/><c:splitPos val="5"/>')).toEqual([5, 6, 7]);
    expect(secondary(data, '<c:splitType val="percent"/><c:splitPos val="8"/>')).toEqual([4, 5, 6, 7]);
    expect(secondary(data, '<c:splitType val="cust"/><c:custSplit><c:secondPiePt val="0"/><c:secondPiePt val="1"/><c:secondPiePt val="5"/></c:custSplit>')).toEqual([0, 1, 5]);
  });
  test("a split by value may leave non-adjacent points; claims past the data are dropped", () => {
    expect(secondary([1, 50, 2, 40], '<c:splitType val="val"/><c:splitPos val="3"/>')).toEqual([0, 2]);
    expect(secondary(data, '<c:splitType val="cust"/><c:custSplit><c:secondPiePt val="3"/><c:secondPiePt val="3"/><c:secondPiePt val="99"/><c:secondPiePt val="-1"/></c:custSplit>')).toEqual([3]);
    expect(secondary(data, '<c:splitType val="pos"/><c:splitPos val="1000000000"/>').length).toBe(8);
  });
  test("with nothing to split off the chart draws its pie alone", () => {
    const read = readChartModel(parseXml(ofPie(data, '<c:splitType val="cust"/><c:custSplit/>')), theme);
    if (!("spec" in read)) throw new Error(read.unsupportedReason);
    const sweeps = sectors(layoutChart(read.spec, { width: 8, height: 4.5 })).map((s) => s.sweep);
    expect(sweeps).toHaveLength(8);
    expect(sweeps.reduce((a, b) => a + b, 0)).toBeCloseTo(360, 0);
  });
  test("the size is bounded when read; the gap, kept as written, is bounded when laid out", () => {
    const wild = readChart(parseXml(ofPie(data, '<c:gapWidth val="900000"/><c:secondPieSize val="99999"/>')), theme)!;
    expect(wild.groups[0]!.ofPie!.secondSize).toBe(200);
    expect(wild.groups[0]!.gapWidth).toBe(900000);
    // 900000 draws as 500, and the plots still fit the frame.
    const capped = readChart(parseXml(ofPie(data, '<c:gapWidth val="500"/><c:secondPieSize val="99999"/>')), theme)!;
    const drawing = layoutChart(wild, { width: 8, height: 4.5 });
    expect(JSON.stringify(drawing)).toBe(JSON.stringify(layoutChart(capped, { width: 8, height: 4.5 })));
    for (const s of sectors(drawing)) expect(s.cx + s.r).toBeLessThanOrEqual(drawing.width + 1);
  });
});
