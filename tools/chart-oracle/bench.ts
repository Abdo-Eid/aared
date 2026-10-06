/**
 * Speed of the chart engine's two halves on small and very large charts:
 * reading (the Worker half, `readChartModel`) and layout (the page half,
 * `layoutChart`), with the primitive count and the model's size.
 *
 *   bun tools/chart-oracle/bench.ts
 */
import { parseXml } from "../../packages/aared/src/core/xml.js";
import { readChartModel } from "../../packages/aared/src/core/chart.js";
import { layoutChart } from "../../packages/aared/src/core/chart/draw.js";
import { themeForColorMap } from "../../packages/aared/src/core/theme.js";

const theme = themeForColorMap({
  colors: { dk1: "#000000", lt1: "#FFFFFF", accent1: "#4472C4", accent2: "#ED7D31", accent3: "#A5A5A5", accent4: "#FFC000", accent5: "#5B9BD5", accent6: "#70AD47" },
  fonts: { major: { latin: "Calibri Light", complexScript: "", scripts: {} }, minor: { latin: "Calibri", complexScript: "", scripts: {} } },
}, undefined);

function chart(kind: "bar" | "line", nSeries: number, nPts: number): string {
  const points = (s: number) => Array.from({ length: nPts }, (_, i) => `<c:pt idx="${i}"><c:v>${(Math.sin(i / 7 + s) * 100 + 120).toFixed(2)}</c:v></c:pt>`).join("");
  const cats = Array.from({ length: nPts }, (_, i) => `<c:pt idx="${i}"><c:v>C${i}</c:v></c:pt>`).join("");
  const ser = Array.from({ length: nSeries }, (_, s) => `<c:ser><c:idx val="${s}"/><c:order val="${s}"/><c:tx><c:v>S${s}</c:v></c:tx><c:cat><c:strRef><c:strCache><c:ptCount val="${nPts}"/>${cats}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${nPts}"/>${points(s)}</c:numCache></c:numRef></c:val></c:ser>`).join("");
  const group = kind === "bar"
    ? `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser}<c:axId val="1"/><c:axId val="2"/></c:barChart>`
    : `<c:lineChart><c:grouping val="standard"/>${ser}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>`;
  return `<c:chartSpace><c:chart><c:plotArea><c:layout/>${group}<c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/></c:legend></c:chart></c:chartSpace>`;
}

const median = (xs: number[]): number => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)]!;
const time = (runs: number, task: () => void): number => {
  for (let i = 0; i < 3; i++) task();
  const samples: number[] = [];
  for (let i = 0; i < runs; i++) {
    const t0 = performance.now();
    task();
    samples.push(performance.now() - t0);
  }
  return median(samples);
};

const frame = { width: 8, height: 4.5 };
for (const [kind, series, points] of [["bar", 3, 12], ["bar", 4, 1000], ["line", 10, 5000], ["bar", 1, 50000]] as const) {
  const xml = parseXml(chart(kind, series, points));
  const runs = points > 1000 ? 7 : 101;
  const read = readChartModel(xml, theme);
  if (!("spec" in read)) throw new Error(read.unsupportedReason);
  const drawing = layoutChart(read.spec, frame);
  const readMs = time(runs, () => readChartModel(xml, theme));
  const layoutMs = time(runs, () => layoutChart(read.spec, frame));
  console.log(`${kind} ${series}x${points}: read ${readMs.toFixed(2)} ms, layout ${layoutMs.toFixed(2)} ms, ${drawing.primitives.length} primitives, model ${(JSON.stringify(read.spec).length / 1024).toFixed(0)} KB`);
}
