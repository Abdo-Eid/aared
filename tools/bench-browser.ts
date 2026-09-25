/**
 * The browser pass of the benchmark harness: drives the bench's measurement
 * page (`bench/measure.html`, built into `bench/dist`) in a real browser via
 * `Bun.WebView`, one renderer × fixture at a time, and turns the samples into
 * `Cell`s. The page only times; sampling, statistics and memory readings live
 * here, because memory needs the browser's DevTools protocol, not page code
 * (`tools/cdp.ts` reads the page and each Worker in its own session).
 *
 *   bun tools/bench-measure.ts --run            # every cell
 *   bun tools/bench-measure.ts --run pptx-coverage.pptx:aared   # one cell
 *   bun tools/bench-measure.ts --size           # download size per renderer
 *   bun tools/bench-measure.ts --calibrate      # check the memory reading
 */

import { existsSync } from "node:fs";
import { join } from "node:path";
import {
  FIXTURES,
  LARGE_FIXTURES,
  METHOD,
  RENDERERS_BY_FORMAT,
  buildCells,
  formatOfFixture,
  kbRange,
  pendingCell,
  stat,
  type BenchFormat,
  type Cell,
  type SearchCell,
  type SearchTimings,
  type SizeFileEntry,
  type SizeRow,
} from "./bench-measure";
import { Browser, partBytes, readingBytes, workerBytes, type HeapReading, type Page } from "./cdp";

const ROOT = join(import.meta.dir, "..");
const DIST = join(ROOT, "bench", "dist");
const FIXTURE_DIRS = [join(ROOT, "fixtures"), join(ROOT, "tools", "bench-search", "out")];
const TIMEOUT_MS = 90_000;

interface Request {
  path: string;
  /** `Sec-Fetch-Dest`: `script`, `worker`, `empty` (a fetch, e.g. WASM)… */
  dest: string;
}

interface Server {
  url: string;
  /** Every path the server was asked for since the last `clear()`, in order. */
  requests: Request[];
  clear(): void;
  stop(): void;
}

/**
 * Static server for the built bench plus every fixture under /f/. `noStore`
 * (the size pass) forbids caching, so every file a renderer needs is asked
 * for again in each fresh page and shows up in the request log.
 */
function serve(options: { noStore?: boolean } = {}): Server {
  const requests: Request[] = [];
  const headers: HeadersInit = options.noStore === true ? { "cache-control": "no-store" } : {};
  const server = Bun.serve({
    port: 0,
    async fetch(request) {
      const path = decodeURIComponent(new URL(request.url).pathname);
      requests.push({ path, dest: request.headers.get("sec-fetch-dest") ?? "" });
      if (path.startsWith("/f/")) {
        for (const dir of FIXTURE_DIRS) {
          const file = Bun.file(join(dir, path.slice(3)));
          if (await file.exists()) return new Response(file, { headers });
        }
        return new Response("no such fixture", { status: 404 });
      }
      const file = Bun.file(join(DIST, path === "/" ? "measure.html" : path));
      return (await file.exists()) ? new Response(file, { headers }) : new Response("not found", { status: 404 });
    },
  });
  return {
    url: `http://localhost:${server.port}/measure.html`,
    requests,
    clear: () => {
      requests.length = 0;
    },
    stop: () => server.stop(true),
  };
}

const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));
const q = (value: string): string => JSON.stringify(value);

/** Run a page expression, failing (not hanging) if it takes too long. */
async function call<T>(page: Page, expression: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Error(`timed out after ${TIMEOUT_MS / 1000}s: ${expression.slice(0, 80)}`)), TIMEOUT_MS);
  });
  try {
    return await Promise.race([page.evaluate<T>(expression), timeout]);
  } finally {
    clearTimeout(timer);
  }
}

interface Delta {
  total: number;
  page: number;
  worker: number;
}

const delta = (before: HeapReading, after: HeapReading): Delta => ({
  total: readingBytes(after) - readingBytes(before),
  page: partBytes(after.page) - partBytes(before.page),
  worker: workerBytes(after) - workerBytes(before),
});

const isLarge = (fixture: string): boolean => (LARGE_FIXTURES as readonly string[]).includes(fixture);

/** Measure one renderer on one fixture. Never throws: a failure becomes an error cell. */
export async function measureCell(browser: Browser, server: Server, fixture: string, renderer: string): Promise<Cell> {
  const cell = pendingCell(fixture, renderer);
  const format = formatOfFixture(fixture);
  const large = isLarge(fixture);
  const open = `__m.open(${q(format)}, ${q(renderer)}, ${q(fixture)})`;
  try {
    // Cold: the first open in a fresh page, several fresh pages. The memory
    // it adds includes loading the renderer's code (every adapter imports
    // its library on open), the same for every renderer.
    const cold: number[] = [];
    const first: Delta[] = [];
    for (let run = 0; run < METHOD.coldRuns; run += 1) {
      const page = await browser.open(server.url);
      try {
        await call(page, `__m.preload(${q(fixture)})`);
        await call(page, `__m.prime(${q(fixture)})`);
        await page.settle();
        const before = await page.heap();
        const { ms, units } = await call<{ ms: number; units: number }>(page, open);
        cold.push(ms);
        cell.units = units;
        await page.settle();
        first.push(delta(before, await page.heap()));
        cell.elements ??= await call<number>(page, "__m.elements()");
      } finally {
        page.close();
      }
    }
    cell.open_cold_ms = stat(cold);
    cell.heap_first_kb = kbRange(first.map((d) => d.total));
    cell.heap_first_page_kb = kbRange(first.map((d) => d.page));
    cell.heap_first_worker_kb = kbRange(first.map((d) => d.worker));

    // Warm, navigation and memory share one page.
    const page = await browser.open(server.url);
    try {
      await call(page, `__m.preload(${q(fixture)})`);
      await call(page, `__m.prime(${q(fixture)})`);
      const warm: number[] = [];
      const moves: number[] = [];
      const passes: number[] = [];
      for (let run = 0; run < METHOD.warmupDiscarded + METHOD.runs; run += 1) {
        const { ms } = await call<{ ms: number }>(page, open);
        const measured = run >= METHOD.warmupDiscarded;
        if (measured) warm.push(ms);
        if (measured && !large && cell.units !== null && cell.units > 0 && run - METHOD.warmupDiscarded < METHOD.navigationRuns) {
          const pass = await call<number[]>(page, "__m.navigateAll()");
          moves.push(...pass);
          passes.push(pass.reduce((sum, ms) => sum + ms, 0));
        }
        await call(page, "__m.destroy()");
      }
      cell.open_warm_ms = stat(warm);
      cell.nav_ms = stat(moves);
      cell.full_ms = stat(passes);

      // Memory: live bytes before open, after open (settled), after destroy.
      // The page settles first, so no run starts with the last warm open's
      // garbage or a Worker that is still shutting down.
      const opened: Delta[] = [];
      const residual: number[] = [];
      await page.settle();
      for (let run = 0; run < METHOD.memoryRuns; run += 1) {
        const before = await page.heap();
        await call(page, open);
        await page.settle();
        opened.push(delta(before, await page.heap()));
        await call(page, "__m.destroy()");
        await page.settle();
        residual.push(readingBytes(await page.heap()) - readingBytes(before));
      }
      cell.heap_open_kb = kbRange(opened.map((d) => d.total));
      cell.heap_open_page_kb = kbRange(opened.map((d) => d.page));
      cell.heap_open_worker_kb = kbRange(opened.map((d) => d.worker));
      cell.heap_residual_kb = kbRange(residual);
      // Workers a renderer keeps alive with no document open (aared parks one
      // for the next open): part of every repeat-open baseline, so reported.
      cell.heap_idle_worker_kb = kbRange([workerBytes(await page.heap())]);

      // After a full read: every slide or sheet, or the document scrolled to
      // its end. Skipped on the large decks and workbooks, where a baseline's
      // full pass takes minutes.
      if (!large || format === "docx") {
        const visited: number[] = [];
        for (let run = 0; run < (large ? 1 : METHOD.visitRuns); run += 1) {
          const before = await page.heap();
          await call(page, open);
          await call(page, "__m.visitAll()");
          await page.settle();
          visited.push(readingBytes(await page.heap()) - readingBytes(before));
          await call(page, "__m.destroy()");
          await page.settle();
        }
        cell.heap_visited_kb = kbRange(visited);
      }
    } finally {
      page.close();
    }
    cell.status = "measured";
  } catch (error) {
    cell.status = "error";
    cell.error = error instanceof Error ? error.message : String(error);
  }
  return cell;
}

/** Time aared's `search()` on one large fixture, over fresh sessions in fresh pages. */
export async function measureSearch(browser: Browser, server: Server, fixture: string): Promise<SearchCell> {
  const format = formatOfFixture(fixture);
  const cell: SearchCell = {
    fixture,
    format,
    size_kb: 0,
    open_ms: null,
    cold_ms: null,
    warm_ms: null,
    hit_ms: null,
    hits: null,
    status: "measured",
    error: null,
  };
  const runs: SearchTimings[] = [];
  try {
    for (let run = 0; run < METHOD.coldRuns; run += 1) {
      const page = await browser.open(server.url);
      try {
        cell.size_kb = (await call<number>(page, `__m.preload(${q(fixture)})`)) / 1024;
        runs.push(await call<SearchTimings>(page, `__m.search(${q(format)}, ${q(fixture)})`));
      } finally {
        page.close();
      }
    }
    cell.open_ms = stat(runs.map((r) => r.openMs));
    cell.cold_ms = stat(runs.map((r) => r.coldMs));
    cell.warm_ms = stat(runs.map((r) => r.warmMs));
    cell.hit_ms = stat(runs.map((r) => r.hitMs));
    cell.hits = runs[0]?.hits ?? null;
  } catch (error) {
    cell.status = "error";
    cell.error = error instanceof Error ? error.message : String(error);
  }
  return cell;
}

export interface RunOptions {
  /** `fixture:renderer` pairs to run; every cell when empty. */
  only: readonly string[];
  log: (line: string) => void;
}

function requireDist(): void {
  if (!existsSync(join(DIST, "measure.html"))) {
    throw new Error("bench/dist/measure.html is missing — run `bun run --cwd bench build` first");
  }
}

/** Search timings for every large fixture that has been generated. */
export async function runSearchPass(browser: Browser, log: (line: string) => void): Promise<{ cells: SearchCell[]; skipped: string[] }> {
  requireDist();
  const skipped: string[] = [];
  const server = serve();
  const cells: SearchCell[] = [];
  try {
    for (const fixture of LARGE_FIXTURES) {
      if (!FIXTURE_DIRS.some((dir) => existsSync(join(dir, fixture)))) {
        skipped.push(fixture);
        continue;
      }
      const cell = await measureSearch(browser, server, fixture);
      cells.push(cell);
      log(
        cell.status === "error"
          ? `${fixture}: ERROR ${cell.error}`
          : `${fixture}: cold ${cell.cold_ms!.median.toFixed(0)} ms, warm ${cell.warm_ms!.median.toFixed(1)} ms, ${cell.hits} hits`,
      );
    }
  } finally {
    server.stop();
  }
  return { cells, skipped };
}

/**
 * A hung page can leave the whole browser unanswering DevTools, and every cell
 * after it then fails for the same reason. (A second browser can't be started
 * in the same process, so the recovery is a new run of the rest.)
 */
const isWedged = (error: string | null): boolean => error !== null && /timed out|DevTools/.test(error);

export async function runBrowserPass(
  browser: Browser,
  options: RunOptions,
): Promise<{ cells: Cell[]; skipped: string[]; remaining: string[] }> {
  requireDist();
  const skipped: string[] = [];
  const available = (fixture: string): boolean => FIXTURE_DIRS.some((dir) => existsSync(join(dir, fixture)));
  const wanted = buildCells().filter((cell) => {
    if (!available(cell.fixture)) {
      if (!skipped.includes(cell.fixture)) skipped.push(cell.fixture);
      return false;
    }
    return options.only.length === 0 || options.only.includes(`${cell.fixture}:${cell.renderer}`);
  });
  const server = serve();
  const cells: Cell[] = [];
  let remaining: string[] = [];
  try {
    for (const [index, pending] of wanted.entries()) {
      const started = performance.now();
      const cell = await measureCell(browser, server, pending.fixture, pending.renderer);
      if (cell.status === "error" && isWedged(cell.error)) {
        // Stop here: measuring on would only record the same failure for every cell after it.
        options.log(`[${index + 1}/${wanted.length}] ${cell.fixture} × ${cell.renderer}: ${cell.error} — the browser stopped answering; stopping the pass`);
        remaining = wanted.slice(index).map((c) => `${c.fixture}:${c.renderer}`);
        break;
      }
      cells.push(cell);
      const seconds = ((performance.now() - started) / 1000).toFixed(0);
      const open = cell.open_warm_ms === null ? "—" : `${cell.open_warm_ms.median.toFixed(0)} ms`;
      const heap = cell.heap_open_kb === null ? "—" : `${cell.heap_open_kb.median.toFixed(0)} KB`;
      const first = cell.heap_first_kb === null ? "—" : `${cell.heap_first_kb.median.toFixed(0)} KB`;
      options.log(
        `[${index + 1}/${wanted.length}] ${cell.fixture} × ${cell.renderer}: ${cell.status === "error" ? `ERROR ${cell.error}` : `warm open ${open}, memory +${heap} (first open +${first})`} (${seconds}s)`,
      );
    }
  } finally {
    server.stop();
  }
  return { cells, skipped, remaining };
}

async function sized(path: string, dest: string): Promise<SizeFileEntry> {
  const { gzipSync, brotliCompressSync, constants } = await import("node:zlib");
  const buf = new Uint8Array(await Bun.file(join(DIST, path)).arrayBuffer());
  return {
    path,
    kind: path.endsWith(".wasm") ? "wasm" : dest === "worker" ? "worker" : "main",
    bytes: buf.byteLength,
    gzip: gzipSync(buf).byteLength,
    brotli: brotliCompressSync(buf, { params: { [constants.BROTLI_PARAM_QUALITY]: 11 } }).byteLength,
  };
}

/**
 * Download size: every file the browser fetched to open a fixture, in a
 * fresh page with caching off, from the same production build the timing
 * pass serves (one minifier for every renderer, each with its dependencies).
 * The page's own code is loaded before the log starts, so only what the
 * renderer pulls in counts; each file counts once.
 */
export async function runSizePass(
  browser: Browser,
  versionOf: (renderer: string) => string | null,
  log: (line: string) => void,
): Promise<SizeRow[]> {
  requireDist();
  const server = serve({ noStore: true });
  const rows: SizeRow[] = [];
  try {
    for (const fixture of FIXTURES) {
      const format: BenchFormat = formatOfFixture(fixture);
      for (const renderer of RENDERERS_BY_FORMAT[format]) {
        const row: SizeRow = { renderer, format, fixture, version: versionOf(renderer), files: [], status: "measured", error: null };
        const page = await browser.open(server.url);
        try {
          await call(page, `__m.preload(${q(fixture)})`);
          await sleep(100);
          server.clear();
          await call(page, `__m.open(${q(format)}, ${q(renderer)}, ${q(fixture)})`);
          await page.settle();
          const seen = new Map<string, string>();
          for (const request of server.requests) {
            if (request.path.startsWith("/assets/") && !seen.has(request.path)) seen.set(request.path, request.dest);
          }
          for (const [path, dest] of seen) row.files.push(await sized(path, dest));
          await call(page, "__m.destroy()");
        } catch (error) {
          row.status = "error";
          row.error = error instanceof Error ? error.message : String(error);
        } finally {
          page.close();
        }
        rows.push(row);
        const gzip = row.files.reduce((sum, file) => sum + file.gzip, 0);
        log(`${fixture} × ${renderer}: ${row.status === "error" ? `ERROR ${row.error}` : `${row.files.length} files, ${(gzip / 1024).toFixed(1)} KiB gzip`}`);
      }
    }
  } finally {
    server.stop();
  }
  return rows;
}

/**
 * Check the memory reading before trusting it: a control that holds exactly
 * 10 MiB on the page, then one that holds it in a Worker, must each read as
 * about 10 MiB — counted once, in the right place.
 */
export async function runCalibration(browser: Browser, log: (line: string) => void): Promise<boolean> {
  requireDist();
  const server = serve();
  let ok = true;
  try {
    for (const where of ["page", "worker"] as const) {
      const page = await browser.open(server.url);
      try {
        await page.settle();
        const before = await page.heap();
        const bytes = await call<number>(page, `__m.calibrate(${q(where)})`);
        await page.settle();
        const d = delta(before, await page.heap());
        const measured = where === "page" ? d.page : d.worker;
        const error = (measured - bytes) / bytes;
        const pass = Math.abs(error) <= 0.05 && (where === "page" ? Math.abs(d.worker) < bytes * 0.05 : true);
        ok &&= pass;
        log(
          `${where}: held ${(bytes / 1024).toFixed(0)} KiB, read ${(measured / 1024).toFixed(0)} KiB in the ${where} (page ${(d.page / 1024).toFixed(0)}, Workers ${(d.worker / 1024).toFixed(0)}) — ${pass ? "ok" : "WRONG"}`,
        );
        await call(page, "__m.release()");
      } finally {
        page.close();
      }
    }
  } finally {
    server.stop();
  }
  return ok;
}
