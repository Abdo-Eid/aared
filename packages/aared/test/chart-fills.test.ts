import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { readWorkbook } from "../src/core/xlsx/workbook.js";
import { parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { parseXml } from "../src/core/xml.js";
import { readChart } from "../src/core/chart/read.js";
import { layoutChart } from "../src/core/chart/draw.js";
import { pathBox } from "../src/core/chart/scene.js";
import type { ChartPathPrimitive } from "../src/core/chart/scene.js";
import { chartDescription, chartText, chartTitleText } from "../src/core/chart/text.js";
import type { ChartSpecModel } from "../src/core/chart/spec.js";
import type { FillOwner } from "../src/core/fill.js";
import type { ChartModel, SheetModel } from "../src/core/model.js";
import { themeForColorMap } from "../src/core/theme.js";
import { chartLayoutReady, chartsSettled, paintChart } from "../src/chart-svg.js";
import { readChartModel } from "../src/core/chart.js";
import { open as openDocx } from "../src/docx/index.js";
import { open as openPptx } from "../src/pptx/index.js";
import { open as openXlsx } from "../src/xlsx/index.js";
import { storedZip } from "./zip.js";

const FIXTURES = `${import.meta.dir}/../../../fixtures`;

const theme = themeForColorMap({
  colors: { dk1: "#000000", lt1: "#FFFFFF", dk2: "#44546A", lt2: "#E7E6E6", accent1: "#4472C4", accent2: "#ED7D31", accent3: "#A5A5A5", accent4: "#FFC000", accent5: "#5B9BD5", accent6: "#70AD47" },
  fonts: { major: { latin: "Calibri Light", complexScript: "", scripts: {} }, minor: { latin: "Calibri", complexScript: "", scripts: {} } },
}, undefined);

const NS = 'xmlns:c="c" xmlns:a="a" xmlns:r="r"';
const Q = ["Q1", "Q2", "Q3"];
const strCache = (values: string[]): string =>
  `<c:strRef><c:strCache><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}</c:strCache></c:strRef>`;
const numCache = (values: number[]): string =>
  `<c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numCache></c:numRef>`;

/** A one-series column chart whose series carries `spPr`, over the chart area `areaSpPr`. */
function columns(spPr: string, areaSpPr = "", extra = ""): string {
  return `<c:chartSpace ${NS}><c:chart>${extra}<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/>`
    + `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strCache(["Hours"])}</c:tx>${spPr}<c:cat>${strCache(Q)}</c:cat><c:val>${numCache([3, 5, 4])}</c:val></c:ser>`
    + `<c:axId val="1"/><c:axId val="2"/></c:barChart>`
    + `<c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>`
    + `</c:plotArea></c:chart>${areaSpPr}</c:chartSpace>`;
}

const owner: FillOwner = {
  relationships: [{ id: "rId7", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target: "ppt/media/fill.png", targetMode: "Internal" }],
  mediaContentType: () => "image/png",
  inherited: false,
};

function seriesFill(spPr: string, withOwner = true) {
  const spec = readChart(parseXml(columns(`<c:spPr>${spPr}</c:spPr>`)), theme, withOwner ? owner : undefined)!;
  return spec.groups[0]!.series[0]!.look.fill;
}

describe("chart fills keep their real definition", () => {
  test("a pattern fill is the preset with its two colors, not a 50% mix", () => {
    expect(seriesFill(`<a:pattFill prst="dkDnDiag"><a:fgClr><a:srgbClr val="156082"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>`))
      .toEqual({ kind: "pattern", preset: "dkDnDiag", foreground: "#156082", background: "#FFFFFF" });
  });

  test("a path gradient keeps its shape and fillToRect; a linear one keeps its angle and scaling", () => {
    const stops = `<a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="C9372C"/></a:gs></a:gsLst>`;
    expect(seriesFill(`<a:gradFill>${stops}<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`)).toMatchObject({
      kind: "path-gradient", path: "circle", focus: { left: 0.5, top: 0.5, right: 0.5, bottom: 0.5 },
      stops: [{ position: 0, color: "#FFFFFF" }, { position: 1, color: "#C9372C" }],
    });
    expect(seriesFill(`<a:gradFill>${stops}<a:lin ang="5400000" scaled="1"/></a:gradFill>`)).toMatchObject({ kind: "linear-gradient", angle: 90, scaled: true });
    // A one-stop gradient is that color.
    expect(seriesFill(`<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="123456"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>`)).toBe("#123456");
  });

  test("a picture fill names the chart part's media; without the part's relationships it is unsupported", () => {
    const blip = `<a:blipFill><a:blip r:embed="rId7"/><a:tile sx="50000" sy="50000" algn="tl" flip="none"/></a:blipFill>`;
    expect(seriesFill(blip)).toMatchObject({ kind: "picture", mediaId: "ppt/media/fill.png", contentType: "image/png", tile: { scaleX: 0.5, scaleY: 0.5 } });
    expect(seriesFill(blip, false)).toEqual({ kind: "unsupported", source: "image" });
  });

  test("a group fill and an unreadable gradient are marked unsupported, never flat grey", () => {
    expect(seriesFill(`<a:grpFill/>`)).toEqual({ kind: "unsupported", source: "group" });
    expect(seriesFill(`<a:gradFill><a:gsLst/></a:gradFill>`)).toEqual({ kind: "unsupported", source: "gradient-linear" });
  });

  test("chart-area fills read the same way", () => {
    const spec = readChart(parseXml(columns("", `<c:spPr><a:pattFill prst="pct50"><a:fgClr><a:srgbClr val="000000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill></c:spPr>`)), theme)!;
    expect(spec.look.fill).toMatchObject({ kind: "pattern", preset: "pct50" });
  });

  test("a line stroked with a pattern strokes in its foreground color (a line has no paint server here)", () => {
    const spec = readChart(parseXml(columns(`<c:spPr><a:ln w="12700"><a:pattFill prst="pct50"><a:fgClr><a:srgbClr val="AA0000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill></a:ln></c:spPr>`)), theme)!;
    expect(spec.groups[0]!.series[0]!.look.line).toMatchObject({ color: "#AA0000", width: 1 });
  });
});

describe("a shape with a paint-server fill carries its own box", () => {
  test("bars: the box is the rectangle; the merge never joins fills that need their own box", () => {
    const stops = `<a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="000000"/></a:gs></a:gsLst>`;
    const spec = readChart(parseXml(columns(`<c:spPr><a:gradFill>${stops}<a:lin ang="5400000"/></a:gradFill></c:spPr>`)), theme)!;
    const bars = layoutChart(spec, { width: 6, height: 4 }).primitives.filter((p): p is ChartPathPrimitive => p.kind === "path" && typeof p.fill === "object");
    // Three columns, one primitive each: fills that need their own box never merge.
    expect(bars).toHaveLength(3);
    for (const bar of bars) {
      const m = /^M(-?[\d.]+) (-?[\d.]+)h(-?[\d.]+)v(-?[\d.]+)h/.exec(bar.path)!;
      expect(bar.box).toEqual({ x: Number(m[1]), y: Number(m[2]), width: Number(m[3]), height: Number(m[4]) });
    }
  });

  test("a path's box follows its lines and the extremes of its circular arcs", () => {
    expect(pathBox("M10 20h30v40h-30Z")).toEqual({ x: 10, y: 20, width: 30, height: 40 });
    expect(pathBox("M1 2L11 2L11 7L1 7Z")).toEqual({ x: 1, y: 2, width: 10, height: 5 });
    // A circle drawn as two arcs: the arcs bulge past their end points.
    const circle = pathBox("M0 5A5 5 0 1 1 10 5A5 5 0 1 1 0 5Z");
    expect(circle.x).toBeCloseTo(0);
    expect(circle.y).toBeCloseTo(0);
    expect(circle.width).toBeCloseTo(10);
    expect(circle.height).toBeCloseTo(10);
    // A quarter pie from the centre (5,5): right, then clockwise to the bottom.
    expect(pathBox("M5 5L10 5A5 5 0 0 1 5 10Z")).toEqual({ x: 5, y: 5, width: 5, height: 5 });
  });

  test("a pie slice with a gradient gets the box of its sector", () => {
    const stops = `<a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="2E7D74"/></a:gs></a:gsLst>`;
    const pie = `<c:chartSpace ${NS}><c:chart><c:plotArea><c:layout/><c:pieChart><c:varyColors val="0"/>`
      + `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strCache(["Share"])}</c:tx><c:spPr><a:gradFill>${stops}<a:path path="circle"/></a:gradFill></c:spPr><c:cat>${strCache(Q)}</c:cat><c:val>${numCache([1, 1, 1, 1])}</c:val></c:ser>`
      + `</c:pieChart></c:plotArea></c:chart></c:chartSpace>`;
    const spec = readChart(parseXml(pie), theme)!;
    const slices = layoutChart(spec, { width: 5, height: 4 }).primitives.filter((p): p is ChartPathPrimitive => p.kind === "path" && p.box !== undefined);
    expect(slices.length).toBeGreaterThanOrEqual(1);
    for (const slice of slices) {
      expect(slice.box!.width).toBeGreaterThan(0);
      expect(slice.box!.height).toBeGreaterThan(0);
    }
  });
});

/** A recording stand-in for the SVG a chart paints into: the unit suite has no DOM, and the real output is covered by the visual suites. */
function fakeSvg() {
  const attributes: Record<string, string> = {};
  const children: Array<{ tag: string; attributes: Record<string, string>; textContent: string }> = [];
  const svg = {
    attributes,
    children,
    dataset: {} as Record<string, string>,
    classList: { add: () => undefined },
    setAttribute: (name: string, value: string) => { attributes[name] = value; },
    prepend: (node: (typeof children)[number]) => { children.unshift(node); },
    append: (node: (typeof children)[number]) => { children.push(node); },
    querySelector: () => null,
  };
  return svg;
}

async function paintAs(spec: ChartSpecModel, alt?: string) {
  const original = (globalThis as { document?: unknown }).document;
  (globalThis as { document?: unknown }).document = {
    createElementNS: (_ns: string, tag: string) => {
      const node = { tag, attributes: {} as Record<string, string>, textContent: "", setAttribute(name: string, value: string) { this.attributes[name] = value; } };
      return node;
    },
  };
  try {
    const svg = fakeSvg();
    paintChart(svg as unknown as SVGSVGElement, spec, { frame: { width: 4, height: 3 }, alt, unavailable: () => { throw new Error("layout unavailable"); } });
    await chartsSettled();
    return svg;
  } finally {
    (globalThis as { document?: unknown }).document = original;
  }
}

describe("a render that waits for the chart code paints charts with fills at once", () => {
  test("after chartLayoutReady() a chart whose fill needs the fill painter draws synchronously, marked where it cannot paint", async () => {
    const spec = readChart(parseXml(columns(`<c:spPr><a:grpFill/></c:spPr>`)), theme)!;
    await chartLayoutReady();
    const original = (globalThis as { document?: unknown }).document;
    (globalThis as { document?: unknown }).document = {
      createElementNS: (_ns: string, tag: string) => ({ tag, attributes: {} as Record<string, string>, textContent: "", setAttribute(name: string, value: string) { this.attributes[name] = value; } }),
    };
    try {
      const svg = fakeSvg();
      paintChart(svg as unknown as SVGSVGElement, spec, { frame: { width: 4, height: 3 }, unavailable: () => undefined });
      // No awaiting: DOCX never waits for `chartsSettled()`.
      expect(Number(svg.dataset.aaredChartPrimitives)).toBeGreaterThan(5);
      expect(svg.dataset.aaredFillFallback).toBe("group");
    } finally {
      (globalThis as { document?: unknown }).document = original;
    }
  });

  test("a malformed number in a fill costs the chart its drawing, marked, not the document around it", () => {
    const xml = columns(`<c:spPr><a:gradFill><a:gsLst><a:gs pos="abc"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></c:spPr>`);
    const read = readChartModel(parseXml(xml), theme);
    expect(read).toHaveProperty("unsupportedReason");
    expect((read as { unsupportedReason: string }).unsupportedReason).toContain("this chart has a damaged fill and is not previewed");
  });
});

describe("a painted chart is an image with a name and a description", () => {
  const titled = readChart(parseXml(columns("", "", `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Hours by quarter</a:t></a:r></a:p></c:rich></c:tx></c:title>`)), theme)!;
  const untitled = readChart(parseXml(columns("")), theme)!;

  test("the author's alt text wins, then the chart title, then 'Chart'", async () => {
    expect((await paintAs(titled, "Bar chart of hours")).attributes).toMatchObject({ role: "img", "aria-label": "Bar chart of hours" });
    expect((await paintAs(titled)).attributes["aria-label"]).toBe("Hours by quarter");
    expect((await paintAs(titled, "   ")).attributes["aria-label"]).toBe("Hours by quarter");
    expect((await paintAs(untitled)).attributes["aria-label"]).toBe("Chart");
  });

  test("a <desc> lists the plot, the series and the categories, and the svg points at it", async () => {
    const svg = await paintAs(titled);
    const desc = svg.children.find((child) => child.tag === "desc")!;
    expect(desc.textContent).toBe("Column chart. 1 series: Hours. 3 categories: Q1, Q2, Q3.");
    expect(svg.attributes["aria-describedby"]).toBe(desc.attributes.id!);
    // Drawn too: the geometry follows the name and the description.
    expect(Number(svg.dataset.aaredChartPrimitives)).toBeGreaterThan(5);
  });

  test("the description reads types, stacking, long lists and combinations", () => {
    const many = readChart(parseXml(columns("")), theme)!;
    many.groups[0]!.grouping = "stacked";
    many.groups[0]!.direction = "bar";
    many.groups[0]!.series[0]!.categories = Array.from({ length: 12 }, (_, i) => `C${i + 1}`);
    expect(chartDescription(many)).toBe("Stacked bar chart. 1 series: Hours. 12 categories: C1, C2, C3, C4, C5, C6, C7, C8 and 4 more.");
    many.groups.push({ ...many.groups[0]!, kind: "line", grouping: "standard", series: [] });
    expect(chartDescription(many)).toStartWith("Combination chart (stacked bar, line).");
  });
});

describe("chart words come from the spec, with no layout", () => {
  test("title, visible axis titles, series names and categories, once each, in reading order", () => {
    const xml = columns("", "", `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Hours by quarter</a:t></a:r></a:p></c:rich></c:tx></c:title>`)
      .replace("<c:crossAx val=\"2\"/></c:catAx>", `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title><c:crossAx val="2"/></c:catAx>`);
    const spec = readChart(parseXml(xml), theme)!;
    expect(chartTitleText(spec)).toBe("Hours by quarter");
    expect(chartText(spec)).toEqual(["Hours by quarter", "Quarter", "Hours", "Q1", "Q2", "Q3"]);
    spec.axes.find((axis) => axis.kind === "cat")!.deleted = true;
    expect(chartText(spec)).not.toContain("Quarter");
  });
});

describe("chart text is searchable in every format", () => {
  test("PPTX: a chart's category finds its slide", async () => {
    const deck = await openPptx(await Bun.file(`${FIXTURES}/post-production-review.pptx`).arrayBuffer());
    try {
      const matches = await deck.search("Ep 05");
      expect(matches.map((m) => [m.slide.index, m.text])).toEqual([[3, "Ep 05"]]);
      expect((await deck.search("planned")).some((m) => m.slide.index === 3)).toBe(true);
    } finally {
      deck.close();
    }
  });

  test("XLSX: a chart's title, series and categories are found at the chart's anchor cell, tagged with the chart", async () => {
    const workbook = await openXlsx(await Bun.file(`${FIXTURES}/post-production-tracker.xlsx`).arrayBuffer());
    try {
      const dashboard = workbook.sheets.get("Dashboard");
      const hits = (await dashboard.search("hours logged")).filter((m) => m.chart !== undefined);
      // The title ("Hours logged by episode") and the series name ("Hours logged").
      expect(hits.map((m) => m.text)).toEqual(["Hours logged by episode", "Hours logged"]);
      for (const hit of hits) expect(hit).toMatchObject({ sheet: { title: "Dashboard" }, chart: { name: "Chart 1" }, row: 18, col: 0, address: "A19" });
      // "Hours logged" is also written on the Notes sheet, which a Dashboard search does not return.
      expect((await dashboard.search("hours logged")).every((m) => m.sheet.title === "Dashboard")).toBe(true);
      expect((await workbook.sheets.get("Notes").search("hours logged")).map((m) => m.address)).toEqual(["C18"]);
      expect((await dashboard.search("Episode 5")).some((m) => m.chart !== undefined && m.text === "Episode 5")).toBe(true);
    } finally {
      workbook.close();
    }
  });

  test("DOCX: the chart in the flow and the one anchored in a paragraph are both searched", async () => {
    const document = await openDocx(await Bun.file(`${FIXTURES}/docx-coverage.docx`).arrayBuffer());
    try {
      const charts = (await document.search("الساعات المسجلة لكل حلقة")).filter((m) => m.kind === "chart");
      // Page 7 holds one chart in its own block and one anchored in a paragraph.
      expect(charts).toHaveLength(2);
      expect(charts.every((m) => m.text === "الساعات المسجلة لكل حلقة")).toBe(true);
      expect((await document.search("الحلقة الرابعة")).filter((m) => m.kind === "chart")).toHaveLength(2);
    } finally {
      document.close();
    }
  });
});

/** The fixture's parts as a record, with edits, as a new package: a real file with one change, never a file added to fixtures/. */
async function patched(fixture: string, edit: (files: Record<string, Uint8Array | string>) => void): Promise<OpcPackage> {
  const source = await OpcPackage.open(await Bun.file(`${FIXTURES}/${fixture}`).arrayBuffer());
  const files: Record<string, Uint8Array | string> = {};
  for (const name of source.partNames()) files[name] = await source.readPart(name);
  edit(files);
  return OpcPackage.open(storedZip(files));
}

const text = (part: Uint8Array | string | undefined): string => (typeof part === "string" ? part : new TextDecoder().decode(part));

describe("the author's alt text reaches the chart model", () => {
  test("PPTX: p:cNvPr/@descr, else its title", async () => {
    const edit = (attributes: string) => async () => {
      const pkg = await patched("post-production-review.pptx", (files) => {
        const part = Object.keys(files).find((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name) && text(files[name]).includes('name="Chart 0"'))!;
        files[part] = text(files[part]).replace(/(<p:cNvPr\b[^>]*name="Chart 0") descr=""/, `$1 ${attributes}`);
      });
      const slide = await parsePptxSlide(pkg, await indexPptx(pkg), 3);
      return slide.elements.find((element): element is ChartModel => element.kind === "chart")!;
    };
    expect((await edit(`descr="Logged versus planned hours"`)()).alt).toBe("Logged versus planned hours");
    expect((await edit(`title="Hours chart"`)()).alt).toBe("Hours chart");
    expect((await edit(`descr="" title=""`)()).alt).toBeUndefined();
  });

  test("XLSX: xdr:cNvPr/@descr", async () => {
    const pkg = await patched("post-production-tracker.xlsx", (files) => {
      const drawing = Object.keys(files).find((name) => name.startsWith("xl/drawings/") && name.endsWith(".xml"))!;
      files[drawing] = text(files[drawing]).replace(/(<xdr:cNvPr\b[^>]*name="Chart 1")/, `$1 descr="Hours logged for each episode"`);
    });
    const sheet: SheetModel = await parseXlsxSheet(pkg, await readWorkbook(pkg), 0);
    const chart = sheet.drawings.find((drawing) => drawing.kind === "chart");
    expect(chart?.alt).toBe("Hours logged for each episode");
  });

  test("a chart part's picture fill resolves its media through the chart part's own relationships", async () => {
    const pkg = await patched("post-production-review.pptx", (files) => {
      const chart = text(files["ppt/charts/chart1.xml"]);
      files["ppt/charts/chart1.xml"] = chart.replace(/<c:ser>([\s\S]*?)<c:spPr>[\s\S]*?<\/c:spPr>/, `<c:ser>$1<c:spPr><a:blipFill><a:blip xmlns:r="r" r:embed="rIdFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></c:spPr>`);
      files["ppt/charts/_rels/chart1.xml.rels"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFill" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/chart-fill.png"/></Relationships>`;
      files["ppt/media/chart-fill.png"] = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
    });
    const slide = await parsePptxSlide(pkg, await indexPptx(pkg), 3);
    const chart = slide.elements.find((element): element is ChartModel => element.kind === "chart")!;
    expect(chart.spec!.groups[0]!.series[0]!.look.fill).toMatchObject({ kind: "picture", mediaId: "ppt/media/chart-fill.png", contentType: "image/png" });
  });
});
