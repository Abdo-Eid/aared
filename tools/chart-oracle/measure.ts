/**
 * Lay the engine's text out in a real browser (Bun.WebView) and read back
 * every word's box, in points (each chart's SVG is drawn at 1px per point).
 * What the reader sees is the browser's own text, not the engine's width
 * estimate, so that is what gets compared with PowerPoint.
 *
 *   bun tools/chart-oracle/measure.ts      (after render.ts)
 *
 * Reads out/engine.json, writes out/measured.json: per slide, a list of
 * { text, x0, y0, x1, y1, size }.
 */

import { OUTPUT_ROOT } from "../shared/paths.js";

const OUT = `${OUTPUT_ROOT}/charts`;

interface Entry {
  viewBox: { width: number; height: number };
  primitives: Array<{ kind: string; x: number; y: number; text: string; fontSize: number; fontFamily: string; textAnchor: string; fontWeight?: string; fontStyle?: string; rotate?: number }>;
}

const engine = JSON.parse(await Bun.file(`${OUT}/engine.json`).text()) as Array<Entry | null>;
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");

function svg(entry: Entry | null, i: number): string {
  if (entry === null) return `<svg id="c${i}"></svg>`;
  const { width, height } = entry.viewBox;
  const body = entry.primitives.filter((p) => p.kind === "text").map((p) => {
    const rot = p.rotate === undefined ? "" : ` transform="rotate(${p.rotate} ${p.x} ${p.y})"`;
    return `<text x="${p.x}" y="${p.y}" font-size="${p.fontSize}" font-family="${esc(p.fontFamily)}" text-anchor="${p.textAnchor}"${p.fontWeight === "bold" ? ' font-weight="bold"' : ""}${p.fontStyle === "italic" ? ' font-style="italic"' : ""}${rot}>${esc(p.text)}</text>`;
  }).join("");
  return `<svg id="c${i}" xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" style="direction:ltr">${body}</svg>`;
}

const html = `<!doctype html><html><head><meta charset="utf-8"><style>body{margin:0} svg{display:block}</style></head><body>${engine.map(svg).join("")}</body></html>`;
const server = Bun.serve({ port: 0, fetch: () => new Response(html, { headers: { "content-type": "text/html" } }) });
try {
  await using view = new Bun.WebView({ width: 1200, height: 900 });
  await view.navigate(server.url.href);
  const result = await view.evaluate(`(async () => {
    await document.fonts.ready;
    const out = [];
    for (const svg of document.querySelectorAll("svg")) {
      const origin = svg.getBoundingClientRect();
      const words = [];
      for (const t of svg.querySelectorAll("text")) {
        const node = t.firstChild;
        if (!node) continue;
        const re = /\\S+/g;
        let m;
        while ((m = re.exec(node.textContent))) {
          const range = document.createRange();
          range.setStart(node, m.index);
          range.setEnd(node, m.index + m[0].length);
          const b = range.getBoundingClientRect();
          words.push({ text: m[0], x0: b.left - origin.left, y0: b.top - origin.top, x1: b.right - origin.left, y1: b.bottom - origin.top, size: Number(t.getAttribute("font-size")) });
        }
      }
      out.push(words);
    }
    return out;
  })()`);
  await Bun.write(`${OUT}/measured.json`, JSON.stringify(result));
  console.log(`measured ${engine.length} slides`);
} finally {
  server.stop(true);
}
