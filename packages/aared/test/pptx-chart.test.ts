import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { parseChartGraphic } from "../src/core/chart.js";
import { parseXml } from "../src/core/xml.js";
import { readDrawingTheme } from "../src/core/theme.js";
import type { ChartModel } from "../src/core/model.js";

const PRODUCTION = `${import.meta.dir}/../../../fixtures/post-production-review.pptx`;

async function productionChart(): Promise<ChartModel> {
  const pkg = await OpcPackage.open(await Bun.file(PRODUCTION).arrayBuffer());
  const deck = await indexPptx(pkg);
  const slide = await parsePptxSlide(pkg, deck, 3);
  const chart = slide.elements.find((element): element is ChartModel => element.kind === "chart");
  if (chart === undefined) throw new Error("fixture structure changed: expected a chart on slide 4");
  return chart;
}

describe("PPTX shared charts", () => {
  test("resolves the production bar chart to typed primitives without markup", async () => {
    const chart = await productionChart();
    expect(chart).toMatchObject({
      id: "4",
      name: "Chart 0",
      sourceChartId: "ppt/charts/chart1.xml",
      bounds: { x: 0.6, y: 1.45 },
    });
    const graphic = chart.graphic!;
    expect(graphic.viewBox).toEqual({ width: 1000, height: 600 });
    expect(graphic.primitives.length).toBeGreaterThan(10);
    const rects = graphic.primitives.filter((primitive) => primitive.kind === "rect");
    const texts = graphic.primitives.filter((primitive) => primitive.kind === "text");
    const lines = graphic.primitives.filter((primitive) => primitive.kind === "line");
    expect(rects.map((rect) => rect.kind === "rect" && rect.fill)).toContain("#2E7D74");
    expect(rects.map((rect) => rect.kind === "rect" && rect.fill)).toContain("#C9D6D3");
    expect(texts.map((text) => text.kind === "text" && text.text)).toEqual(
      expect.arrayContaining(["Ep 04", "Logged", "Planned", "39", "42"]),
    );
    expect(lines.length).toBeGreaterThan(0);
    // Typed primitives only: never raw SVG/HTML markup.
    expect(JSON.stringify(graphic)).not.toContain("<svg");
    expect(JSON.stringify(graphic)).not.toContain("<rect");
    for (const primitive of graphic.primitives) {
      expect(["rect", "line", "path", "text"]).toContain(primitive.kind);
    }
  });

  test("emits a horizontal bar chart with value labels", async () => {
    const pkg = await OpcPackage.open(await Bun.file(PRODUCTION).arrayBuffer());
    const theme = await readDrawingTheme(pkg);
    const chartSpace = parseXml(
      `<c:chartSpace><c:chart><c:plotArea><c:barChart><c:barDir val="bar"/><c:ser>` +
        `<c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Team</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
        `<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr>` +
        `<c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
        `<c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt></c:numCache></c:numRef></c:val>` +
        `</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`,
    );
    const graphic = parseChartGraphic(chartSpace, theme)!;
    const rects = graphic.primitives.filter((primitive) => primitive.kind === "rect");
    expect(rects).toHaveLength(2);
    expect(rects[1]).toMatchObject({ fill: "#FF0000" });
    // Horizontal bars extend further for larger values.
    if (rects[0]?.kind !== "rect" || rects[1]?.kind !== "rect") throw new Error("expected rects");
    expect(rects[1].width).toBeGreaterThan(rects[0].width);
  });

  test("emits a line chart as a stroked path with markers", async () => {
    const pkg = await OpcPackage.open(await Bun.file(PRODUCTION).arrayBuffer());
    const theme = await readDrawingTheme(pkg);
    const chartSpace = parseXml(
      `<c:chartSpace><c:chart><c:plotArea><c:lineChart><c:ser>` +
        `<c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Trend</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
        `<c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
        `<c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>` +
        `</c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>`,
    );
    const graphic = parseChartGraphic(chartSpace, theme)!;
    const paths = graphic.primitives.filter((primitive) => primitive.kind === "path");
    expect(paths).toHaveLength(1);
    expect(paths[0]?.kind === "path" && paths[0].path).toMatch(/^M .* L .* L /);
    expect(paths[0]).toMatchObject({ fill: "none" });
    expect(graphic.primitives.filter((primitive) => primitive.kind === "rect").length).toBeGreaterThanOrEqual(3);
  });

  test("emits a pie chart as slice paths with percentage labels", async () => {
    const pkg = await OpcPackage.open(await Bun.file(PRODUCTION).arrayBuffer());
    const theme = await readDrawingTheme(pkg);
    const chartSpace = parseXml(
      `<c:chartSpace><c:chart><c:plotArea><c:pieChart><c:ser>` +
        `<c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Web</c:v></c:pt><c:pt idx="1"><c:v>TV</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
        `<c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val>` +
        `</c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>`,
    );
    const graphic = parseChartGraphic(chartSpace, theme)!;
    const paths = graphic.primitives.filter((primitive) => primitive.kind === "path");
    expect(paths).toHaveLength(2);
    for (const path of paths) {
      expect(path.kind === "path" && path.path).toContain("A ");
    }
    const labels = graphic.primitives.filter((primitive) => primitive.kind === "text").map((primitive) => primitive.kind === "text" && primitive.text);
    expect(labels).toEqual(expect.arrayContaining(["Web 25%", "TV 75%"]));
  });

  test("returns undefined for unsupported chart types", async () => {
    const pkg = await OpcPackage.open(await Bun.file(PRODUCTION).arrayBuffer());
    const theme = await readDrawingTheme(pkg);
    const chartSpace = parseXml(`<c:chartSpace><c:chart><c:plotArea><c:scatterChart/></c:plotArea></c:chart></c:chartSpace>`);
    expect(parseChartGraphic(chartSpace, theme)).toBeUndefined();
  });
});
