/**
 * Draw every chart in the oracle deck through aared's own route: read as the
 * Worker reads it (`parsePptxSlide`, the chart spec), laid out as the page lays
 * it out (`layoutChart`, at the chart's frame).
 *
 *   bun tools/chart-oracle/render.ts
 *
 * Writes into tools/chart-oracle/out/ (git-ignored):
 *   engine.json   one entry per slide: { bounds, viewBox, primitives }, or null
 *   svg/NN.svg    each chart as a standalone SVG, for eyeballing
 */
import { mkdirSync } from "node:fs";
import { OpcPackage } from "../../packages/aared/src/core/opc.js";
import { indexPptx, parsePptxSlide } from "../../packages/aared/src/core/pptx.js";
import { layoutChart } from "../../packages/aared/src/core/chart/draw.js";
import type { ChartPrimitive } from "../../packages/aared/src/core/chart/scene.js";
import type { ChartModel } from "../../packages/aared/src/core/model.js";

const HERE = import.meta.dir;
const OUT = `${HERE}/out`;

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");

function toSvg(primitives: ChartPrimitive[], width: number, height: number): string {
  let defs = "";
  let n = 0;
  const body = primitives.map((p) => {
    if (p.kind === "text") {
      const rot = p.rotate === undefined ? "" : ` transform="rotate(${p.rotate} ${p.x} ${p.y})"`;
      return `<text x="${p.x}" y="${p.y}" font-size="${p.fontSize}" font-family="${esc(p.fontFamily)}" fill="${p.fill}" text-anchor="${p.textAnchor}"${p.fontWeight === "bold" ? ' font-weight="bold"' : ""}${p.fontStyle === "italic" ? ' font-style="italic"' : ""}${rot}>${esc(p.text)}</text>`;
    }
    let fill = "none";
    if (typeof p.fill === "string") fill = p.fill;
    else if (p.fill !== undefined) {
      // An eyeballing aid, not the page painter (`src/fill-paint.ts`): a gradient is
      // drawn linear in its box, a pattern in its foreground, anything else unfilled.
      const f = p.fill;
      if (f.kind === "linear-gradient" || f.kind === "path-gradient") {
        const id = `g${++n}`;
        const r = ((f.kind === "linear-gradient" ? f.angle : 90) * Math.PI) / 180;
        defs += `<linearGradient id="${id}" x1="${0.5 - Math.cos(r) / 2}" y1="${0.5 - Math.sin(r) / 2}" x2="${0.5 + Math.cos(r) / 2}" y2="${0.5 + Math.sin(r) / 2}">${f.stops.map((s) => `<stop offset="${s.position}" stop-color="${s.color}"/>`).join("")}</linearGradient>`;
        fill = `url(#${id})`;
      } else if (f.kind === "pattern") fill = f.foreground;
    }
    const s = p.stroke;
    const stroke = s === undefined ? "" : ` stroke="${s.color}" stroke-width="${s.width}"${s.dash === undefined ? "" : ` stroke-dasharray="${s.dash.join(" ")}"`}${s.cap === undefined ? "" : ` stroke-linecap="${s.cap}"`}${s.join === undefined ? "" : ` stroke-linejoin="${s.join}"`}`;
    return `<path d="${p.path}" fill="${fill}"${stroke}/>`;
  }).join("\n");
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width * 2}" height="${height * 2}" viewBox="0 0 ${width} ${height}" style="direction:ltr"><defs>${defs}</defs><rect width="100%" height="100%" fill="white"/>${body}</svg>`;
}

mkdirSync(`${OUT}/svg`, { recursive: true });
const pkg = await OpcPackage.open(await Bun.file(`${HERE}/chart-oracle.pptx`).arrayBuffer());
const deck = await indexPptx(pkg);
const result: unknown[] = [];
const t0 = performance.now();
for (let i = 0; i < deck.slideParts.length; i++) {
  const slide = await parsePptxSlide(pkg, deck, i);
  const chart = slide.elements.find((e): e is ChartModel => e.kind === "chart");
  if (chart?.spec === undefined) {
    result.push(null);
    continue;
  }
  const drawing = layoutChart(chart.spec, chart.bounds);
  await Bun.write(`${OUT}/svg/${String(i + 1).padStart(2, "0")}.svg`, toSvg(drawing.primitives, drawing.width, drawing.height));
  result.push({ bounds: chart.bounds, viewBox: { width: drawing.width, height: drawing.height }, primitives: drawing.primitives });
}
console.log(`drew ${deck.slideParts.length} slides in ${(performance.now() - t0).toFixed(0)}ms`);
await Bun.write(`${OUT}/engine.json`, JSON.stringify(result));
