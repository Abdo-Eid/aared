/**
 * Consumer-build probe: proves a production Vite build of a minimal aared
 * consumer both bundles the Worker correctly and opens every format.
 *
 *   bun tools/check-consumer-build.ts
 *
 * Background: `WorkerSession` starts its worker via
 * `new Worker(new URL("./worker.mjs", import.meta.url))`. Vite only bundles
 * that target when the pattern is statically detectable; when it instead
 * copies the file verbatim, the worker never posts `ready`, and `open()`
 * hangs forever. This probe builds a fixture consumer with the repo's Vite,
 * asserts the emitted worker is self-contained (no relative chunk it could
 * be separated from) and that every emitted script is Latin-1 (V8 keeps a
 * script with any wider character two bytes per character), serves the
 * result, and opens one real fixture per format headlessly. Exit non-zero on
 * any failure.
 */
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, relative, resolve } from "node:path";

const ROOT = join(import.meta.dir, "..");
const VITE_BIN = join(ROOT, "bench", "node_modules", "vite", "bin", "vite.js");
const AARED_DIST = join(ROOT, "packages", "aared", "dist").replaceAll("\\", "/");
const FIXTURES: Array<{ file: string; format: string }> = [
  { file: "post-production-review.pptx", format: "pptx" },
  { file: "post-production-tracker.xlsx", format: "xlsx" },
  { file: "docx-coverage.docx", format: "docx" },
];

const work = mkdtempSync(join(tmpdir(), "aared-consumer-probe-"));
try {
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

  writeFileSync(join(work, "index.html"), `<!doctype html><html><body><script type="module" src="/main.ts"></script></body></html>`);
  writeFileSync(
    join(work, "main.ts"),
    `import { open as openPptx } from "${AARED_DIST}/pptx/index.mjs";
import { open as openXlsx } from "${AARED_DIST}/xlsx/index.mjs";
import { open as openDocx } from "${AARED_DIST}/docx/index.mjs";
const events: string[] = [];
${openCalls}
(window as unknown as { __probe: unknown }).__probe = { events };`,
  );
  writeFileSync(
    join(work, "vite.config.ts"),
    `export default { build: { emptyOutDir: true }, logLevel: "warn" as const };`,
  );
  mkdirSync(join(work, "public"));
  for (const { file } of FIXTURES) copyFileSync(join(ROOT, "fixtures", file), join(work, "public", file));

  const build = Bun.spawnSync(["node", VITE_BIN, "build"], { cwd: work, stderr: "pipe", stdout: "pipe" });
  if (build.exitCode !== 0) {
    throw new Error(`vite build failed:\n${build.stderr.toString()}\n${build.stdout.toString()}`);
  }

  const outDir = join(work, "dist");
  const assets = readdirSync(join(outDir, "assets"));
  console.log("emitted assets:", assets.join(", "));
  const failures: string[] = [];
  // The worker is one self-contained file: whatever Vite names it
  // (worker-*.js), it must reference no sibling chunk — a relative import
  // is a file the copy could be separated from, which is exactly the
  // verbatim-copy failure this probe guards against.
  const workerAsset = assets.find((name) => name.startsWith("worker-") && /\.[cm]?js$/.test(name));
  if (workerAsset === undefined) {
    failures.push("no worker chunk emitted in consumer build");
  } else {
    const workerSource = await Bun.file(join(outDir, "assets", workerAsset)).text();
    const relatives = [...workerSource.matchAll(/["'](\.[^"']+\.[cm]?js)["']/g)].map((m) => m[1]!.split("/").pop()!);
    if (relatives.length > 0) {
      failures.push(`worker references sibling chunk(s): ${relatives.join(", ")} (assets: ${assets.join(", ")})`);
    } else {
      console.log("worker is self-contained (no relative chunk references)");
    }
  }

  // Every emitted script stays Latin-1: one character above U+00FF makes V8
  // hold the whole script two bytes per character (src/core/chars.ts). This
  // checks what the consumer's own minifier printed, not just aared's dist.
  const aboveLatin1 = (text: string): string[] => {
    const hits: string[] = [];
    for (let i = 0; i < text.length && hits.length < 20; i += 1) {
      const code = text.charCodeAt(i);
      if (code > 0xff) hits.push(`U+${code.toString(16).toUpperCase().padStart(4, "0")} at ${i}: ${JSON.stringify(text.slice(Math.max(0, i - 40), i + 10))}`);
    }
    return hits;
  };
  for (const name of assets.filter((n) => /\.[cm]?js$/.test(n))) {
    const hits = aboveLatin1(readFileSync(join(outDir, "assets", name), "utf8"));
    if (hits.length > 0) failures.push(`${name} holds characters above U+00FF (V8 keeps it two-byte):\n  ${hits.join("\n  ")}`);
  }
  if (workerAsset !== undefined && !failures.some((f) => f.startsWith(workerAsset))) console.log("worker source is Latin-1 (one byte per character in V8)");

  // The probe above opens the three format entries. `element`, `preview` and
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
    let probe: { events?: string[] } | undefined;
    while (Date.now() < deadline) {
      probe = (await view.evaluate(`window.__probe ?? null`)) as typeof probe;
      if (probe !== null && probe !== undefined) break;
      await Bun.sleep(500);
    }
    if (probe === undefined || probe === null) throw new Error("probe page never reported (load or script failure)");
    for (const event of probe.events ?? []) console.log(" ", event);
    const errors = (probe.events ?? []).filter((event) => event.includes(":ERROR:"));
    if (errors.length > 0) failures.push(`${errors.length} format(s) failed to open`);
    console.log("404s during run:", missing404.length === 0 ? "(none)" : missing404.join(", "));
    if (missing404.some((path) => path.endsWith(".mjs") || path.endsWith(".js"))) {
      failures.push("script chunk(s) failed to load at runtime");
    }
    if (failures.length > 0) throw new Error(`consumer build probe failed:\n- ${failures.join("\n- ")}`);
    console.log("CONSUMER-BUILD-OK");
  } finally {
    server.stop();
  }
} finally {
  rmSync(work, { recursive: true, force: true });
}
