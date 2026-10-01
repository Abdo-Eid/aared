/**
 * Consumer-build probe: proves production builds of a minimal aared consumer,
 * with Vite and with webpack 5, both bundle the Worker correctly and open
 * every format.
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
 * self-contained (no relative chunk it could be separated from). Next.js 16,
 * with Turbopack and with webpack, was checked by hand (DECISION_LOG §56).
 * Exit non-zero on any failure.
 */
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, relative, resolve } from "node:path";

const ROOT = join(import.meta.dir, "..");
const VITE_BIN = join(ROOT, "bench", "node_modules", "vite", "bin", "vite.js");
const WEBPACK_BIN = join(ROOT, "bench", "node_modules", "webpack-cli", "bin", "cli.js");
const AARED_DIST = join(ROOT, "packages", "aared", "dist").replaceAll("\\", "/");
const FIXTURES: Array<{ file: string; format: string }> = [
  { file: "post-production-review.pptx", format: "pptx" },
  { file: "post-production-tracker.xlsx", format: "xlsx" },
  { file: "docx-coverage.docx", format: "docx" },
];
const INDEX_HTML = `<!doctype html><html><body><script type="module" src="/main.js"></script></body></html>`;

/** The consumer: open each fixture, list its units, read the first one. */
function consumerSource(): string {
  const openCalls = FIXTURES.map(({ file, format }) => {
    const opener = format === "pptx" ? "openPptx" : format === "xlsx" ? "openXlsx" : "openDocx";
    return `
    try {
      const bytes = await fetch("/${file}").then((r) => r.arrayBuffer());
      events.push("${format}:fetched:" + bytes.byteLength);
      const doc = await ${opener}(bytes);
      const units = doc.${format === "docx" ? "sections" : format === "pptx" ? "slides" : "sheets"}.list();
      events.push("${format}:units:" + units.length);
      ${format === "docx" ? "" : `await units[0].read(); events.push("${format}:read:ok");`}
      doc.close();
    } catch (error) {
      events.push("${format}:ERROR:" + String(error && error.message ? error.message : error));
    }`;
  }).join("\n");
  return `import { open as openPptx } from "${AARED_DIST}/pptx/index.mjs";
import { open as openXlsx } from "${AARED_DIST}/xlsx/index.mjs";
import { open as openDocx } from "${AARED_DIST}/docx/index.mjs";
const events = [];
${openCalls}
window.__probe = { events };`;
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
  writeFileSync(join(work, "vite.config.ts"), `export default { build: { emptyOutDir: true }, logLevel: "warn" as const };`);
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
async function probe(label: string, outDir: string, failures: string[]): Promise<void> {
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
        for (const match of text.matchAll(/(?:from|import)\s*\(?\s*["'](\.{1,2}\/[^"']+)["']/g)) {
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
