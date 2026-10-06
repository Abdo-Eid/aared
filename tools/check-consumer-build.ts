/**
 * Consumer-build probe: proves production builds of a minimal aared consumer,
 * with Vite and with webpack 5, both bundle the Worker correctly and open
 * every format.
 * Fixed-format Vite consumers also exercise preview/search/cleanup and assert
 * that their emitted module graphs omit sibling renderers and auto-routing.
 *
 *   bun tools/check-consumer-build.ts
 *
 * Background: `WorkerSession` starts its worker via
 * `new Worker(new URL("./worker.mjs", import.meta.url), { type: "module" })`.
 * A bundler only bundles that target when the pattern is statically
 * detectable; when Vite instead copied the file verbatim, the worker never
 * posted `ready`, and `open()` hung forever. Webpack 5 bundles the same
 * pattern with no loader (its "Web Workers" guide). For each bundler this
 * probe builds a fixture consumer with the bench's copy, asserts every
 * emitted script is Latin-1 (V8 keeps a script with any wider character two
 * bytes per character), serves the result, and opens one real fixture per
 * format headlessly; for Vite it also asserts the emitted worker is
 * self-contained (no relative chunk it could be separated from). The PPTX and
 * XLSX probes also render a unit holding a chart: chart layout is a separate
 * file the page fetches on the first chart (DECISION_LOG §71), so the consumer
 * bundler must emit it and the page must load it. Next.js 16, with Turbopack
 * and with webpack, was checked by hand (DECISION_LOG §56).
 * Exit non-zero on any failure.
 */
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, relative, resolve } from "node:path";

const ROOT = join(import.meta.dir, "..");
const VITE_BIN = join(ROOT, "bench", "node_modules", "vite", "bin", "vite.js");
const WEBPACK_BIN = join(ROOT, "bench", "node_modules", "webpack-cli", "bin", "cli.js");
const AARED_DIST = join(ROOT, "packages", "aared", "dist").replaceAll("\\", "/");
/** `chartUnit`: the unit to render, whose chart must draw (its SVG selector). */
const FIXTURES: Array<{ file: string; format: string; chartUnit?: { index: number; svg: string } }> = [
  { file: "post-production-review.pptx", format: "pptx", chartUnit: { index: 3, svg: "chart-geometry" } },
  { file: "post-production-tracker.xlsx", format: "xlsx", chartUnit: { index: 0, svg: "chart-graphic" } },
  { file: "docx-coverage.docx", format: "docx" },
];
const INDEX_HTML = `<!doctype html><html><body><script type="module" src="/main.js"></script></body></html>`;

/** The consumer: open each fixture, list its units, read the first one, and render the one with a chart. */
function consumerSource(singleFormat?: string): string {
  if (singleFormat !== undefined) {
    const fixture = FIXTURES.find(({ format }) => format === singleFormat)!;
    return `import { preview } from "${AARED_DIST}/${singleFormat}/index.mjs";
const events = [];
async function runProbe() {
  try {
    const target = document.body.appendChild(document.createElement("div"));
    target.style.cssText = "width:960px;height:600px";
    const viewer = await preview({ source: "/${fixture.file}", target });
    try {
      if (viewer.format !== "${singleFormat}") throw Error("wrong viewer format");
      ${fixture.chartUnit === undefined ? "" : `await viewer.goTo(${fixture.chartUnit.index});
      const svg = target.querySelector('[data-aared-element="${fixture.chartUnit.svg}"]');
      events.push("${singleFormat}:chart:" + (svg?.getAttribute("data-aared-chart-primitives") ?? "missing"));`}
      const matches = await viewer.search(${JSON.stringify(singleFormat === "docx" ? "nile" : singleFormat === "pptx" ? "production" : "0")});
      if (viewer.format === "pptx" && matches[0]) await viewer.goTo(matches[0].slide.index);
      if (viewer.format === "xlsx" && matches[0]) await viewer.xlsx.goToCell(matches[0].address);
      const ranges = CSS.highlights.get(target.dataset.aaredSearchHighlight);
      events.push("preview:search:" + (matches.length > 0 && ranges?.size > 0 ? "ok" : "missing"));
    } finally { viewer.destroy(); }
    if (target.childElementCount !== 0) throw Error("destroy left content");
    target.remove();
  } catch (error) { events.push("${singleFormat}:ERROR:" + String(error.message ?? error)); }
  window.__probe = { events };
}
void runProbe();`;
  }
  const openCalls = FIXTURES.map(({ file, format, chartUnit }) => {
    const opener = format === "pptx" ? "openPptx" : format === "xlsx" ? "openXlsx" : "openDocx";
    const renderChart = chartUnit === undefined ? "" : `
      const target = document.body.appendChild(document.createElement("div"));
      target.style.cssText = "width:960px;height:600px";
      const handle = await ${format === "pptx" ? "renderPptx({ slide" : "renderXlsx({ sheet"}: units[${chartUnit.index}], target });
      const svg = target.querySelector('[data-aared-element="${chartUnit.svg}"]');
      events.push("${format}:chart:" + (svg === null ? "missing" : svg.getAttribute("data-aared-chart-primitives") ?? "not-drawn"));
      handle.destroy();
      target.remove();`;
    return `
    try {
      const bytes = await fetch("/${file}").then((r) => r.arrayBuffer());
      events.push("${format}:fetched:" + bytes.byteLength);
      const doc = await ${opener}(bytes);
      const units = doc.${format === "docx" ? "sections" : format === "pptx" ? "slides" : "sheets"}.list();
      events.push("${format}:units:" + units.length);
      ${format === "docx" ? "" : `await units[0].read(); events.push("${format}:read:ok");`}${renderChart}
      doc.close();
    } catch (error) {
      events.push("${format}:ERROR:" + String(error && error.message ? error.message : error));
    }`;
  }).join("\n");
  return `import { open as openPptx, render as renderPptx } from "${AARED_DIST}/pptx/index.mjs";
import { open as openXlsx, render as renderXlsx } from "${AARED_DIST}/xlsx/index.mjs";
import { open as openDocx } from "${AARED_DIST}/docx/index.mjs";
import { preview } from "${AARED_DIST}/preview/index.mjs";
const events = [];
// Run after module evaluation: a bundler may merge statically imported format
// entries into this module, which preview imports dynamically. Top-level await
// would wait for that same module to finish evaluating.
async function runProbe() {
${openCalls}
try {
  const target = document.body.appendChild(document.createElement("div"));
  target.style.cssText = "width:960px;height:600px";
  const viewer = await preview({ source: "/docx-coverage.docx", target });
  try {
    const matches = await viewer.search("nile");
    const ranges = CSS.highlights.get(target.dataset.aaredSearchHighlight);
    events.push("preview:search:" + (matches.length > 0 && ranges?.size > 0 ? "ok" : "missing"));
  } finally { viewer.destroy(); target.remove(); }
} catch (error) {
  events.push("preview:ERROR:" + String(error && error.message ? error.message : error));
}
window.__probe = { events };
}
void runProbe();`;
}

/** Characters above U+00FF in a script, with some context, up to 20. */
function aboveLatin1(text: string): string[] {
  const hits: string[] = [];
  for (let i = 0; i < text.length && hits.length < 20; i += 1) {
    const code = text.charCodeAt(i);
    if (code > 0xff) hits.push(`U+${code.toString(16).toUpperCase().padStart(4, "0")} at ${i}: ${JSON.stringify(text.slice(Math.max(0, i - 40), i + 10))}`);
  }
  return hits;
}

/** Every `.js`/`.mjs` file under `dir`, as paths relative to it. */
function scripts(dir: string): string[] {
  const found: string[] = [];
  const walk = (current: string): void => {
    for (const entry of readdirSync(current, { withFileTypes: true })) {
      const path = join(current, entry.name);
      if (entry.isDirectory()) walk(path);
      else if (/\.[cm]?js$/.test(entry.name)) found.push(relative(dir, path));
    }
  };
  walk(dir);
  return found;
}

function run(label: string, command: string[], cwd: string): void {
  const build = Bun.spawnSync(command, { cwd, stderr: "pipe", stdout: "pipe" });
  if (build.exitCode !== 0) throw new Error(`${label} failed:\n${build.stderr.toString()}\n${build.stdout.toString()}`);
}

/** A production Vite build of the consumer. Returns the served directory. */
function buildVite(work: string, failures: string[]): string {
  writeFileSync(join(work, "index.html"), INDEX_HTML);
  writeFileSync(join(work, "vite.config.ts"), `export default {
    build: { emptyOutDir: true }, logLevel: "warn" as const,
    plugins: [{ name: "consumer-module-audit", generateBundle(_options, bundle) {
      this.emitFile({ type: "asset", fileName: "modules.json", source: JSON.stringify(
        Object.values(bundle).filter(item => item.type === "chunk").flatMap(chunk => Object.keys(chunk.modules))
      ) });
    } }],
  };`);
  run("vite build", ["node", VITE_BIN, "build"], work);
  const outDir = join(work, "dist");
  const assets = readdirSync(join(outDir, "assets"));
  console.log("emitted assets:", assets.join(", "));
  // The worker is one self-contained file: whatever Vite names it
  // (worker-*.js), it must reference no sibling chunk — a relative import
  // is a file the copy could be separated from, which is exactly the
  // verbatim-copy failure this probe guards against.
  const workerAsset = assets.find((name) => name.startsWith("worker-") && /\.[cm]?js$/.test(name));
  if (workerAsset === undefined) {
    failures.push("vite: no worker chunk emitted in consumer build");
  } else {
    const workerSource = readFileSync(join(outDir, "assets", workerAsset), "utf8");
    const relatives = [...workerSource.matchAll(/["'](\.[^"']+\.[cm]?js)["']/g)].map((m) => m[1]!.split("/").pop()!);
    if (relatives.length > 0) failures.push(`vite: worker references sibling chunk(s): ${relatives.join(", ")} (assets: ${assets.join(", ")})`);
    else console.log("worker is self-contained (no relative chunk references)");
  }
  return outDir;
}

/** A production webpack 5 build of the consumer, no worker loader. Returns the served directory. */
function buildWebpack(work: string): string {
  const outDir = join(work, "dist-webpack");
  writeFileSync(
    join(work, "webpack.config.mjs"),
    `export default { mode: "production", entry: "./main.js", output: { path: ${JSON.stringify(outDir)}, clean: true }, performance: { hints: false } };`,
  );
  run("webpack build", ["node", WEBPACK_BIN, "--config", "webpack.config.mjs"], work);
  writeFileSync(join(outDir, "index.html"), INDEX_HTML);
  console.log("emitted scripts:", scripts(outDir).join(", "));
  return outDir;
}

/** Every emitted script stays Latin-1; this checks what the consumer's own minifier printed, not just aared's dist. */
function checkLatin1(label: string, outDir: string, failures: string[]): void {
  const before = failures.length;
  for (const name of scripts(outDir)) {
    const hits = aboveLatin1(readFileSync(join(outDir, name), "utf8"));
    if (hits.length > 0) failures.push(`${label}: ${name} holds characters above U+00FF (V8 keeps it two-byte):\n  ${hits.join("\n  ")}`);
  }
  if (failures.length === before) console.log("every emitted script is Latin-1 (one byte per character in V8)");
}

/** Serve a build plus the fixtures and open every format in it, headless. */
async function probe(label: string, outDir: string, failures: string[], singleFormat?: string): Promise<void> {
  const missing404: string[] = [];
  const server = Bun.serve({
    port: 0,
    async fetch(request) {
      const path = new URL(request.url).pathname.slice(1) || "index.html";
      const file = Bun.file(join(outDir, path));
      if (!(await file.exists())) {
        missing404.push(path);
        return new Response("not found", { status: 404 });
      }
      return new Response(file);
    },
  });
  try {
    await using view = new Bun.WebView({ width: 800, height: 600 });
    await view.navigate(`http://localhost:${server.port}/index.html`);
    const deadline = Date.now() + 60000;
    let result: { events?: string[] } | undefined;
    while (Date.now() < deadline) {
      result = (await view.evaluate(`window.__probe ?? null`)) as typeof result;
      if (result !== null && result !== undefined) break;
      await Bun.sleep(500);
    }
    if (result === undefined || result === null) {
      failures.push(`${label}: probe page never reported (load or script failure, or open() hung)`);
      return;
    }
    for (const event of result.events ?? []) console.log(" ", event);
    const errors = (result.events ?? []).filter((event) => event.includes(":ERROR:"));
    if (errors.length > 0) failures.push(`${label}: ${errors.length} format(s) failed to open`);
    if (!(result.events ?? []).includes("preview:search:ok")) failures.push(`${label}: lazy viewer search highlighting failed`);
    // A drawn chart reports how many primitives it painted; anything else is the lazy layout file failing.
    for (const { format, chartUnit } of FIXTURES) {
      if (singleFormat !== undefined && format !== singleFormat) continue;
      if (chartUnit === undefined) continue;
      const drawn = (result.events ?? []).find((event) => event.startsWith(`${format}:chart:`));
      if (drawn === undefined || !/:chart:[1-9]\d*$/.test(drawn)) failures.push(`${label}: ${format} chart not drawn (${drawn ?? "no render"})`);
    }
    console.log("404s during run:", missing404.length === 0 ? "(none)" : missing404.join(", "));
    if (missing404.some((path) => path.endsWith(".mjs") || path.endsWith(".js"))) failures.push(`${label}: script chunk(s) failed to load at runtime`);
  } finally {
    server.stop();
  }
}

const work = mkdtempSync(join(tmpdir(), "aared-consumer-probe-"));
try {
  writeFileSync(join(work, "main.js"), consumerSource());
  // Fixtures are served from `public/` (Vite copies it into its build) and copied into webpack's output.
  mkdirSync(join(work, "public"));
  for (const { file } of FIXTURES) copyFileSync(join(ROOT, "fixtures", file), join(work, "public", file));
  const failures: string[] = [];

  console.log("== vite");
  const viteOut = buildVite(work, failures);
  checkLatin1("vite", join(viteOut, "assets"), failures);
  await probe("vite", viteOut, failures);

  console.log("== webpack");
  const webpackOut = buildWebpack(work);
  for (const { file } of FIXTURES) copyFileSync(join(work, "public", file), join(webpackOut, file));
  checkLatin1("webpack", webpackOut, failures);
  await probe("webpack", webpackOut, failures);

  for (const { format } of FIXTURES) {
    console.log(`== vite single-format ${format}`);
    writeFileSync(join(work, "main.js"), consumerSource(format));
    const singleOut = buildVite(work, failures);
    const modules = JSON.parse(readFileSync(join(singleOut, "modules.json"), "utf8")) as string[];
    for (const { format: other } of FIXTURES) {
      if (other === format) continue;
      if (modules.some((id) => {
        const normalized = id.replaceAll("\\", "/");
        return normalized.includes(`/dist/core/${other}-`) || normalized.includes(`/dist/${other}/index.mjs`);
      })) failures.push(`vite ${format}: included sibling renderer ${other}`);
    }
    if (modules.some(id => id.includes("auto-viewer-"))) failures.push(`vite ${format}: included auto-detect routing`);
    checkLatin1(`vite ${format}`, join(singleOut, "assets"), failures);
    await probe(`vite ${format}`, singleOut, failures, format);
    console.log(`${format}: production module graph contains no sibling renderers or auto-detect routing`);
  }

  // The probes open the three format entries. `element`, `preview` and
  // `react` are reached by other consumers and share chunks with them, so also
  // assert statically that every relative import any emitted module makes
  // resolves to a file that exists in `dist` (shared chunks belong under
  // `dist/core/`, where the worker URL also resolves).
  const unresolved: string[] = [];
  const scan = (dir: string): void => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const path = join(dir, entry.name);
      if (entry.isDirectory()) {
        scan(path);
      } else if (entry.name.endsWith(".mjs")) {
        const text = readFileSync(path, "utf8");
        // Minified dynamic imports quote with backticks: import(`./draw-x.mjs`).
        for (const match of text.matchAll(/(?:from|import)\s*\(?\s*["'`](\.{1,2}\/[^"'`]+)["'`]/g)) {
          if (!existsSync(resolve(dir, match[1]!))) unresolved.push(`${relative(AARED_DIST, path)} -> ${match[1]}`);
        }
      }
    }
  };
  scan(AARED_DIST);
  if (unresolved.length > 0) failures.push(`dist has unresolved relative import(s): ${unresolved.join(", ")}`);
  else console.log("every relative import in dist resolves (element, preview and react included)");

  if (failures.length > 0) throw new Error(`consumer build probe failed:\n- ${failures.join("\n- ")}`);
  console.log("CONSUMER-BUILD-OK");
} finally {
  rmSync(work, { recursive: true, force: true });
}
