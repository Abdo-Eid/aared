import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { readWorkbook } from "../src/core/xlsx/workbook.js";
import type { XlsxWorkbook } from "../src/core/xlsx/workbook.js";
import { parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { readChartModel } from "../src/core/chart.js";
import { MAX_SERIES_POINTS } from "../src/core/chart/read.js";
import { layoutChart } from "../src/core/chart/draw.js";
import type { ChartPrimitive } from "../src/core/chart/scene.js";
import type { ChartSpecModel } from "../src/core/chart/spec.js";
import { themeForColorMap } from "../src/core/theme.js";
import { parseXml } from "../src/core/xml.js";

// Office 2016 charts (waterfall, funnel, histogram, Pareto, box and whisker,
// treemap, sunburst) read through the XLSX parser from the workbook Excel
// itself wrote (`tools/chart-oracle/chartex-oracle.xlsx`, made by
// `gen_chartex.py`; Excel's own renders of each sheet are the layout oracle).
// Excel writes no data cache in a workbook's chartex part: the values come
// from the cells its hidden `_xlchart.*` names point at. The cached-`cx:lvl`
// form PowerPoint and Word embed is covered with synthetic parts at the end:
// pure-function edge cases, with no Office output to read it from.

const ORACLE = `${import.meta.dir}/../../../tools/chart-oracle/chartex-oracle.xlsx`;
/** The frame Excel drew every chart in (480 x 300 pt). */
const FRAME = { width: 480 / 72, height: 300 / 72 };

let opened: Promise<{ pkg: OpcPackage; workbook: XlsxWorkbook }> | undefined;
function oracle(): Promise<{ pkg: OpcPackage; workbook: XlsxWorkbook }> {
  opened ??= (async () => {
    const pkg = await OpcPackage.open(await Bun.file(ORACLE).arrayBuffer());
    return { pkg, workbook: await readWorkbook(pkg) };
  })();
  return opened;
}

async function chartOf(sheet: string): Promise<{ spec: ChartSpecModel; primitives: ChartPrimitive[] }> {
  const { pkg, workbook } = await oracle();
  const index = workbook.sheets.findIndex((s) => s.name === sheet);
  if (index < 0) throw new Error(`fixture structure changed: no sheet "${sheet}"`);
  const parsed = await parseXlsxSheet(pkg, workbook, index);
  const drawing = parsed.drawings[0];
  if (drawing?.kind !== "chart" || drawing.spec === undefined) {
    throw new Error(`"${sheet}" did not read as a chart: ${drawing?.kind === "chart" ? drawing.unsupportedReason : JSON.stringify(parsed.unsupported)}`);
  }
  return { spec: drawing.spec, primitives: layoutChart(drawing.spec, FRAME).primitives };
}

interface Rect { x: number; y: number; w: number; h: number; fill: string }

/** Every `M x y h w v h h -w Z` rectangle the scene emitted, in paint order. */
function rects(primitives: ChartPrimitive[]): Rect[] {
  const out: Rect[] = [];
  for (const p of primitives) {
    if (p.kind !== "path" || typeof p.fill !== "string") continue;
    for (const m of p.path.matchAll(/M(-?[\d.]+) (-?[\d.]+)h(-?[\d.]+)v(-?[\d.]+)h-?[\d.]+Z/g)) {
      out.push({ x: Number(m[1]), y: Number(m[2]), w: Number(m[3]), h: Number(m[4]), fill: p.fill });
    }
  }
  return out;
}

function texts(primitives: ChartPrimitive[]): string[] {
  return primitives.flatMap((p) => (p.kind === "text" ? [p.text] : []));
}

const near = (a: number, b: number, tolerance = 0.06): void => expect(Math.abs(a - b)).toBeLessThan(tolerance);

describe("chartex in XLSX: every Office 2016 chart type reads and draws", () => {
  for (const sheet of ["waterfall", "waterfall2", "funnel", "histogram", "histogram2", "pareto", "pareto2", "boxwhisker", "treemap", "treemap2", "sunburst", "sunburst2"]) {
    test(`${sheet} resolves its data from the workbook cells and draws`, async () => {
      const { spec, primitives } = await chartOf(sheet);
      expect(spec.groups.length).toBeGreaterThan(0);
      expect(primitives.length).toBeGreaterThan(5);
      // An empty title element shows Office's placeholder text.
      expect(texts(primitives)).toContain("Chart Title");
    });
  }
});

describe("waterfall", () => {
  test("bar tops follow the running totals, a total draws from zero, connectors join the levels", async () => {
    const { spec, primitives } = await chartOf("waterfall");
    const series = spec.groups[0]!.series[0]!;
    expect(series.spans).toEqual([[0, 100], [100, 160], [160, 120], [120, 105], [0, 105]]);
    const bars = rects(primitives).filter((r) => ["#156082", "#E97132", "#196B24"].includes(r.fill) && r.h > 8 && r.w < 100 && r.y > 30);
    const [start, up, down, down2, total] = bars;
    const unit = start!.h / 100;
    near(up!.h, 60 * unit);
    near(up!.y + up!.h, start!.y); // the rise starts where the first bar ends
    near(down!.y, up!.y); // the fall starts at the running level 160
    near(down!.h, 40 * unit);
    near(down2!.y, down!.y + down!.h);
    near(down2!.h, 15 * unit);
    near(total!.y + total!.h, start!.y + start!.h); // a total stands on the axis
    near(total!.h, 105 * unit);
    // Increase, decrease and total take the first three accents of the theme.
    expect(bars.map((b) => b.fill)).toEqual(["#156082", "#156082", "#E97132", "#E97132", "#196B24"]);
    const connectors = primitives.filter((p) => p.kind === "path" && p.stroke?.color === "#BFBFBF");
    expect(connectors.length).toBe(1);
    expect(connectors[0]!.kind === "path" && connectors[0]!.path.split("M").length - 1).toBe(4);
    expect(texts(primitives)).toEqual(expect.arrayContaining(["100", "60", "-40", "-15", "105", "البداية", "المبيعات", "الصافي", "Increase", "Decrease", "Total"]));
  });

  test("a total that is not the running sum draws its own value and the connector slopes to it", async () => {
    const { spec, primitives } = await chartOf("waterfall2");
    const series = spec.groups[0]!.series[0]!;
    // Observed in Excel: Total A (90) and Total B (120) stand on the axis at their own values; Up2 continues from 90.
    expect(series.spans).toEqual([[0, 100], [100, 160], [160, 120], [120, 105], [0, 90], [90, 100], [0, 120]]);
    expect(series.totals).toEqual([4, 6]);
    const sloped = primitives.flatMap((p) =>
      p.kind === "path" && p.stroke?.color === "#BFBFBF"
        ? [...p.path.matchAll(/M(-?[\d.]+) (-?[\d.]+)L(-?[\d.]+) (-?[\d.]+)/g)].filter((m) => Math.abs(Number(m[2]) - Number(m[4])) > 0.5)
        : [],
    );
    expect(sloped.length).toBe(2); // 105 -> 90 into Total A, 100 -> 120 into Total B
  });
});

describe("funnel", () => {
  test("each stage is centred and as wide as its value, the largest filling the plot", async () => {
    const { primitives } = await chartOf("funnel");
    const bars = rects(primitives).filter((r) => r.fill === "#156082" && r.w > 20);
    expect(bars.length).toBe(4);
    const centre = bars[0]!.x + bars[0]!.w / 2;
    [1000, 640, 380, 120].forEach((value, i) => {
      near(bars[i]!.w / bars[0]!.w, value / 1000, 0.001);
      near(bars[i]!.x + bars[i]!.w / 2, centre);
    });
    // The first stage is on top, its label inside the bar, the stage names on the left axis.
    expect(bars[0]!.y).toBeLessThan(bars[1]!.y);
    expect(texts(primitives)).toEqual(expect.arrayContaining(["1000", "640", "380", "120", "الزيارات", "الشراء"]));
  });
});

describe("histogram", () => {
  test("bins follow Scott's rule from the minimum, right-closed, and count the values", async () => {
    const { spec, primitives } = await chartOf("histogram");
    const series = spec.groups[0]!.series[0]!;
    // Observed in Excel: 30 values, width 9.9 (Scott's 9.87 rounded up to two digits), bins from the minimum.
    expect(series.categories).toEqual(["[1, 10.9]", "(10.9, 20.8]", "(20.8, 30.7]", "(30.7, 40.6]"]);
    expect(series.values).toEqual([19, 6, 4, 1]);
    expect(spec.groups[0]!.gapWidth).toBe(0);
    const bars = rects(primitives).filter((r) => r.fill === "#156082");
    near(bars[0]!.h / bars[3]!.h, 19, 0.01);
    near(bars[1]!.h / bars[3]!.h, 6, 0.01);
    expect(texts(primitives)).toEqual(expect.arrayContaining(["[1, 10.9]", "(30.7, 40.6]", "20"]));
  });

  test("a wide range rounds the width up to two significant digits", async () => {
    const { spec } = await chartOf("histogram2");
    const series = spec.groups[0]!.series[0]!;
    // Scott's width is 2686; Excel drew 2700.
    expect(series.categories).toEqual(["[1204, 3904]", "(3904, 6604]", "(6604, 9304]", "(9304, 12004]"]);
    expect(series.values).toEqual([16, 3, 0, 1]);
  });
});

describe("pareto", () => {
  test("bars largest first and a cumulative-percentage line on a second 0-100% axis", async () => {
    const { spec, primitives } = await chartOf("pareto");
    const [bars, line] = spec.groups;
    expect(bars!.kind).toBe("bar");
    expect(line!.kind).toBe("line");
    expect(bars!.series[0]!.categories).toEqual(["Scratch", "Dent", "Crack", "Stain", "Other"]);
    const shares = line!.series[0]!.values as number[];
    near(shares[0]!, 42 / 94, 1e-9);
    near(shares[4]!, 1, 1e-9);
    const secondary = spec.axes.find((a) => a.id === line!.axisIds![1])!;
    expect(secondary.position).toBe("r");
    expect(texts(primitives)).toEqual(expect.arrayContaining(["0%", "50%", "100%", "45"]));
  });

  test("categories repeated in the data add up and sort by size", async () => {
    const { spec } = await chartOf("pareto2");
    const bars = spec.groups[0]!.series[0]!;
    expect(bars.categories).toEqual(["Scratch", "Dent", "Crack", "Stain"]);
    expect(bars.values).toEqual([42, 30, 14, 9]);
  });
});

describe("box and whisker", () => {
  test("quartiles by the exclusive method, whiskers inside 1.5 IQR, the far value an outlier", async () => {
    const { spec, primitives } = await chartOf("boxwhisker");
    const series = spec.groups[0]!.series[0]!;
    expect(series.categories).toEqual(["Alpha", "بيتا"]);
    const [alpha, beta] = series.boxes!;
    // QUARTILE.EXC of 3 5 6 7 8 9 10 11 25: 5.5, 8, 10.5; Excel drew the same box.
    expect([alpha!.q1, alpha!.median, alpha!.q3, alpha!.low, alpha!.high, alpha!.outliers]).toEqual([5.5, 8, 10.5, 3, 11, [25]]);
    expect([beta!.q1, beta!.median, beta!.q3, beta!.low, beta!.high, beta!.outliers]).toEqual([12.5, 15, 18.5, 10, 20, []]);
    near(alpha!.mean, 9.3333, 1e-3);
    expect(spec.groups[0]!.box).toEqual({ meanMarker: true, meanLine: false, outliers: true, innerPoints: false });
    const boxes = rects(primitives).filter((r) => r.fill === "#156082" && r.w > 30);
    expect(boxes.length).toBe(2);
    // The box heights are the interquartile ranges (5 and 6) on one scale.
    near(boxes[1]!.h / boxes[0]!.h, 6 / 5, 0.01);
    expect(texts(primitives)).toEqual(expect.arrayContaining(["Alpha", "بيتا"]));
  });
});

describe("treemap", () => {
  test("rectangle areas are proportional to the values; each branch keeps its colour", async () => {
    const { spec, primitives } = await chartOf("treemap");
    const leaves = rects(primitives).filter((r) => ["#156082", "#E97132", "#196B24"].includes(r.fill) && r.w > 20 && r.h > 20);
    expect(leaves.length).toBe(6);
    const tree = spec.groups[0]!.tree!;
    expect(tree.map((n) => [n.name, n.value, n.color])).toEqual([["فاكهة", 80, "#156082"], ["خضار", 40, "#E97132"], ["ألبان", 20, "#196B24"]]);
    const byColor = (color: string): number => leaves.filter((r) => r.fill === color).reduce((s, r) => s + r.w * r.h, 0);
    const unit = byColor("#196B24") / 20;
    near(byColor("#156082") / unit, 80, 0.5);
    near(byColor("#E97132") / unit, 40, 0.5);
    // Leaves: 40 25 15 | 30 10 | 20, areas to scale (stroke widths aside).
    const areas = leaves.map((r) => (r.w * r.h) / unit).sort((a, b) => b - a);
    [40, 30, 25, 20, 15, 10].forEach((value, i) => near(areas[i]!, value, 0.6));
    expect(texts(primitives)).toEqual(expect.arrayContaining(["فاكهة", "تفاح", "موز", "برتقال", "حليب"]));
  });

  test("sorts every level by size, colours by first appearance, and lists the branches in data order", async () => {
    const { spec, primitives } = await chartOf("treemap2");
    const tree = spec.groups[0]!.tree!;
    expect(tree.map((n) => n.name)).toEqual(["C", "B", "A"]);
    expect(tree.map((n) => n.color)).toEqual(["#196B24", "#156082", "#E97132"]);
    expect(tree[1]!.children!.map((n) => n.name)).toEqual(["b2", "b1"]);
    expect(spec.groups[0]!.legendEntries!.map((e) => e.label)).toEqual(["B", "A", "C"]);
    // The biggest branch is the leftmost, full-height block.
    const biggest = rects(primitives).filter((r) => r.fill === "#196B24" && r.w > 20)[0]!;
    const others = rects(primitives).filter((r) => r.fill !== "#196B24" && r.w > 20 && r.h > 20 && r.fill !== "#FFFFFF");
    expect(Math.min(...others.map((r) => r.x))).toBeGreaterThan(biggest.x + biggest.w - 2);
  });

  test("banner parent labels: one grey band per branch, two label font sizes tall (Excel's render, treemap-banner)", async () => {
    const { spec, primitives } = await chartOf("treemap-banner");
    const g = spec.groups[0]!;
    expect(g.parentLabels).toBe("banner");
    expect(g.bannerFill).toBe("#A6A6A6");
    const bands = rects(primitives).filter((r) => r.fill === "#A6A6A6");
    expect(bands).toHaveLength(3);
    const size = g.series[0]!.labels!.text.size;
    for (const band of bands) near(band.h, size * 2, 0.01);
    expect(texts(primitives)).toEqual(expect.arrayContaining(["فاكهة", "خضار", "ألبان"]));
  });
});

describe("box and whisker, inclusive", () => {
  test("QUARTILE.INC when the series asks for it: Excel drew 6-10 and 13-18 (box-inclusive)", async () => {
    const { spec } = await chartOf("box-inclusive");
    const [alpha, beta] = spec.groups[0]!.series[0]!.boxes!;
    expect([alpha!.q1, alpha!.median, alpha!.q3]).toEqual([6, 8, 10]);
    expect([beta!.q1, beta!.median, beta!.q3]).toEqual([13, 15, 18]);
  });
});

describe("sunburst", () => {
  test("rings from the hole outward, largest branch first, labels along the radius", async () => {
    const { spec, primitives } = await chartOf("sunburst2");
    const g = spec.groups[0]!;
    expect(g.kind).toBe("sunburst");
    expect(g.tree!.map((n) => n.name)).toEqual(["C", "B", "A"]);
    const sectors = primitives.filter((p) => p.kind === "path" && p.fill !== undefined && /A/.test(p.path) && p.path.includes("Z") && p.fill !== "#FFFFFF");
    // 3 branches and 5 leaves.
    expect(sectors.length).toBe(8);
    const rotated = primitives.filter((p) => p.kind === "text" && p.rotate !== undefined);
    expect(rotated.length).toBeGreaterThan(0);
    expect(texts(primitives)).toEqual(expect.arrayContaining(["A", "B", "C", "a1", "a2", "b1", "b2", "c1"]));
  });

  test("the Arabic hierarchy keeps its names", async () => {
    const { spec, primitives } = await chartOf("sunburst");
    expect(spec.groups[0]!.tree!.map((n) => n.name)).toEqual(["فاكهة", "خضار", "ألبان"]);
    expect(texts(primitives)).toEqual(expect.arrayContaining(["فاكهة", "خضار"]));
  });
});

// ---------------------------------------------------------------- the cached form (PPTX, DOCX) and its bounds

const CX = 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"';

async function read(xml: string): Promise<ReturnType<typeof readChartModel>> {
  const { workbook } = await oracle();
  return readChartModel(parseXml(xml), themeForColorMap(workbook.theme, undefined));
}

function cached(layout: string, dims: string, extra = ""): string {
  return `<cx:chartSpace ${CX}><cx:chartData><cx:data id="0">${dims}</cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion>` +
    `<cx:series layoutId="${layout}"><cx:dataId val="0"/>${extra}</cx:series></cx:plotAreaRegion>` +
    `<cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/><cx:tickLabels/></cx:axis></cx:plotArea></cx:chart></cx:chartSpace>`;
}

describe("chartex with cached values (PPTX and DOCX embed the cache)", () => {
  test("a cached waterfall reads its points by index", async () => {
    const result = await read(cached(
      "waterfall",
      '<cx:strDim type="cat"><cx:lvl ptCount="3"><cx:pt idx="0">a</cx:pt><cx:pt idx="1">b</cx:pt><cx:pt idx="2">c</cx:pt></cx:lvl></cx:strDim>' +
      '<cx:numDim type="val"><cx:lvl ptCount="3" formatCode="General"><cx:pt idx="0">10</cx:pt><cx:pt idx="2">-4</cx:pt></cx:lvl></cx:numDim>',
    ));
    if (!("spec" in result)) throw new Error(result.unsupportedReason);
    expect(result.spec.groups[0]!.series[0]!.spans).toEqual([[0, 10], null, [10, 6]]);
  });

  test("hierarchy levels: the cache lists the innermost first, and empty outer cells take the name above", async () => {
    const result = await read(cached(
      "treemap",
      '<cx:strDim type="cat"><cx:lvl ptCount="3"><cx:pt idx="0">x</cx:pt><cx:pt idx="1">y</cx:pt><cx:pt idx="2">z</cx:pt></cx:lvl>' +
      '<cx:lvl ptCount="3"><cx:pt idx="0">P</cx:pt><cx:pt idx="2">Q</cx:pt></cx:lvl></cx:strDim>' +
      '<cx:numDim type="size"><cx:lvl ptCount="3"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt><cx:pt idx="2">4</cx:pt></cx:lvl></cx:numDim>',
    ));
    if (!("spec" in result)) throw new Error(result.unsupportedReason);
    const tree = result.spec.groups[0]!.tree!;
    expect(tree.map((n) => [n.name, n.value, (n.children ?? []).map((c) => c.name)])).toEqual([["Q", 4, ["z"]], ["P", 3, ["y", "x"]]]);
  });

  test("a histogram with a bin size and an overflow bin", async () => {
    const values = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 25].map((v, i) => `<cx:pt idx="${i}">${v}</cx:pt>`).join("");
    const result = await read(cached(
      "clusteredColumn",
      `<cx:numDim type="val"><cx:lvl ptCount="11">${values}</cx:lvl></cx:numDim>`,
      '<cx:layoutPr><cx:binning intervalClosed="r" overflow="10"><cx:binSize val="5"/></cx:binning></cx:layoutPr>',
    ));
    if (!("spec" in result)) throw new Error(result.unsupportedReason);
    const series = result.spec.groups[0]!.series[0]!;
    expect(series.categories).toEqual(["[1, 6]", "(6, 11]", "> 10"]);
    expect(series.values).toEqual([6, 4, 1]);
  });

  test("inclusive quartiles differ from exclusive ones", async () => {
    const dims = (): string =>
      '<cx:numDim type="val"><cx:lvl ptCount="4">' + [1, 2, 3, 4].map((v, i) => `<cx:pt idx="${i}">${v}</cx:pt>`).join("") + "</cx:lvl></cx:numDim>";
    const get = async (method: string) => {
      const r = await read(cached("boxWhisker", dims(), `<cx:layoutPr><cx:statistics quartileMethod="${method}"/></cx:layoutPr>`));
      if (!("spec" in r)) throw new Error(r.unsupportedReason);
      const box = r.spec.groups[0]!.series[0]!.boxes![0]!;
      return [box.q1, box.median, box.q3];
    };
    expect(await get("exclusive")).toEqual([1.25, 2.5, 3.75]);
    expect(await get("inclusive")).toEqual([1.75, 2.5, 3.25]);
  });

  test("a map chart stays a placeholder that names the type", async () => {
    const result = await read(cached("regionMap", ""));
    expect(result).toEqual({ unsupportedReason: "Map charts are not previewed" });
  });

  test("a part with neither a cache nor resolvable cells has no values to preview", async () => {
    const result = await read(cached("funnel", '<cx:strDim type="cat"><cx:f>_xlchart.v1.0</cx:f></cx:strDim><cx:numDim type="val"><cx:f>_xlchart.v1.1</cx:f></cx:numDim>'));
    expect(result).toEqual({ unsupportedReason: "this chart has no cached values to preview" });
  });

  test("a claimed point count past the bound is refused, not allocated", async () => {
    const result = await read(cached("funnel", `<cx:numDim type="val"><cx:lvl ptCount="${MAX_SERIES_POINTS + 1}"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim>`));
    expect("unsupportedReason" in result && result.unsupportedReason).toContain("more than the preview draws");
    const far = await read(cached("funnel", `<cx:numDim type="val"><cx:lvl ptCount="2"><cx:pt idx="${MAX_SERIES_POINTS + 5}">1</cx:pt></cx:lvl></cx:numDim>`));
    expect("unsupportedReason" in far && far.unsupportedReason).toContain("more than the preview draws");
  });

  /** `n` series of one layout, each over its own data set claiming `count` points with one value. */
  function many(layout: string, n: number, count: number, extra = ""): string {
    const sets = Array.from({ length: n }, (_, i) => `<cx:data id="${i}"><cx:numDim type="val"><cx:lvl ptCount="${count}"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">3</cx:pt></cx:lvl></cx:numDim></cx:data>`).join("");
    const sers = Array.from({ length: n }, (_, i) => `<cx:series layoutId="${layout}"><cx:dataId val="${i}"/>${extra}</cx:series>`).join("");
    return `<cx:chartSpace ${CX}><cx:chartData>${sets}</cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion>${sers}</cx:plotAreaRegion>` +
      `<cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/><cx:tickLabels/></cx:axis></cx:plotArea></cx:chart></cx:chartSpace>`;
  }
  const BINNED = '<cx:layoutPr><cx:binning><cx:binCount val="10"/></cx:binning></cx:layoutPr>';

  test("histograms count their bins, not their raw values, once binned: three over 90,000 values still draw", async () => {
    const result = await read(many("clusteredColumn", 3, 90_000, BINNED));
    if (!("spec" in result)) throw new Error(result.unsupportedReason);
    expect(result.spec.groups[0]!.series).toHaveLength(3);
  });

  test("raw values that stay held count in full: three box plots over 90,000 values are refused", async () => {
    const result = await read(many("boxWhisker", 3, 90_000));
    expect(result).toEqual({ unsupportedReason: "this chart has more than 250,000 points in all, more than the preview draws" });
  });

  test("binning does not hide the work: twelve histograms over 99,000 values are refused", async () => {
    const result = await read(many("clusteredColumn", 12, 99_000, BINNED));
    expect(result).toEqual({ unsupportedReason: "this chart reads more than 1,000,000 points in all, more than the preview draws" });
  });

  test("a bin size far smaller than the range cannot make millions of bins", async () => {
    const values = [0, 1e9].map((v, i) => `<cx:pt idx="${i}">${v}</cx:pt>`).join("");
    const result = await read(cached("clusteredColumn", `<cx:numDim type="val"><cx:lvl ptCount="2">${values}</cx:lvl></cx:numDim>`, '<cx:layoutPr><cx:binning><cx:binSize val="0.001"/></cx:binning></cx:layoutPr>'));
    if (!("spec" in result)) throw new Error(result.unsupportedReason);
    expect(result.spec.groups[0]!.series[0]!.values.length).toBeLessThanOrEqual(1000);
  });
});
