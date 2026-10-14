import { REPO_ROOT, OUTPUT_ROOT } from "../shared/paths.js";

/**
 * Deterministic benchmark harness.
 *
 * A competitor's numbers change only with its version, so `--run` and
 * `--size` measure aared alone and merge it into the latest result of that
 * kind, keeping every other renderer's cells (each cell records when, and on
 * which browser, it was measured). `--all` measures every renderer into a new
 * file: the first time, or after a competitor's version changes.
 *
 * - `--run [--all] [fixture:renderer …]`: open / navigation / memory in a
 *   real browser (`tools/benchmarks/browser.ts`) into `tools/out/benchmarks/<stamp>.json`.
 *   Named pairs measure just those cells, merged the same way.
 * - `--size [--all]`: download size — every file the browser fetched to open
 *   each fixture, from the bench's production build — into
 *   `tools/out/benchmarks/size-<stamp>.json`.
 * - `--search`: aared `search()` on the large generated fixtures.
 * - `--calibrate`: check the memory reading against controls of known size.
 * - `--render-md` / `--render-site`: regenerate the `<!-- generated:* -->`
 *   blocks in `docs/BENCHMARKS.md` and the home page's data file from the
 *   latest JSON of each kind.
 *
 * No hand-filled numbers: the docs tables are generated or they are TBD.
 */

/** The pinned real fixtures (`fixtures/`), measured for open, navigation, memory and size. */
export const FIXTURES = [
  "pptx-coverage.pptx",
  "post-production-review.pptx",
  "post-production-tracker.xlsx",
  "xlsx-coverage.xlsx",
  "docx-coverage.docx",
] as const;

/**
 * Large generated fixtures (`tools/benchmarks/search/make-fixtures.py`, untracked),
 * measured for open and memory only: how cost grows with file size. Skipped
 * with a note when they haven't been generated.
 */
export const LARGE_FIXTURES = ["large-300-slides.pptx", "large-13-sheets.xlsx", "large-40x.docx"] as const;

export type BenchFormat = "pptx" | "xlsx" | "docx";
export const BENCH_FORMATS: readonly BenchFormat[] = ["pptx", "xlsx", "docx"];

/** Which renderers open which format — the bench adapters that exist. */
export const RENDERERS_BY_FORMAT: Readonly<Record<BenchFormat, readonly string[]>> = {
  pptx: ["aared", "pptxviewjs", "silurus", "aiden0z", "mdurbar", "pptx-browser", "extend", "omni"],
  xlsx: ["aared", "sheetjs", "xlscalc", "spreadsheet-preview", "extend", "omni"],
  docx: ["aared", "docx-preview", "docx-renderer", "apollo", "extend", "omni"],
};

/**
 * The npm package behind each renderer, as the bench installs it (its version
 * is recorded with every size). A renderer that ships one package per format
 * names each.
 */
export const RENDERER_PACKAGES: Readonly<Record<string, string | Readonly<Record<BenchFormat, string>>>> = {
  aared: "@aared/viewer",
  pptxviewjs: "pptxviewjs",
  silurus: "@silurus/ooxml",
  aiden0z: "@aiden0z/pptx-renderer",
  mdurbar: "pptx-viewer",
  omni: "omni-doc-viewer",
  "docx-preview": "docx-preview",
  "docx-renderer": "docx-renderer",
  sheetjs: "xlsx",
  "spreadsheet-preview": "@datagridxl/spreadsheet-preview",
  xlscalc: "@xlscalc/xlsx-preview",
  apollo: "@apollo-design/docx-preview",
  "pptx-browser": "pptx-browser",
  extend: { pptx: "@extend-ai/react-pptx", xlsx: "@extend-ai/react-xlsx", docx: "@extend-ai/react-docx" },
};

export function formatOfFixture(fixture: string): BenchFormat {
  const extension = fixture.split(".").pop();
  if (extension === "pptx" || extension === "xlsx" || extension === "docx") return extension;
  throw new Error(`unknown fixture format: ${fixture}`);
}

/** Summary of repeated timings, in the unit of the samples. */
export interface Stat {
  n: number;
  median: number;
  p95: number;
  min: number;
  max: number;
}

export function stat(samples: readonly number[]): Stat | null {
  if (samples.length === 0) return null;
  const sorted = [...samples].sort((a, b) => a - b);
  const at = (q: number): number => sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(q * sorted.length) - 1))]!;
  return { n: sorted.length, median: at(0.5), p95: at(0.95), min: sorted[0]!, max: sorted[sorted.length - 1]! };
}

export const METHOD = {
  runs: 11,
  warmupDiscarded: 2,
  coldRuns: 5,
  memoryRuns: 5,
  navigationRuns: 5,
  visitRuns: 3,
  viewport: "1280x720 stage for every renderer; DOCX in layout mode (pagination.mode = layout)",
  fonts: "no font resolver and nothing fetched: fixtures and code are served from localhost, text uses the system fonts",
  report: "median + p95; cold and warm separately",
  open: "adapter entry to first paint, then two animation frames (the frame commit). Includes the renderer's dynamic import and, for aared, the Worker parse. Cold = first open in a fresh page (modules not yet evaluated, browser cache warm); warm = repeat opens in one page, first 2 discarded",
  navigation: "moving to every slide or sheet once, each move timed to the committed frame; per-move and whole-pass times over 5 passes. DOCX has no comparable navigation in the baselines",
  memory:
    "Chrome DevTools Runtime.getHeapUsage (JS heap + embedder heap — DOM, layout, styles — + buffers, WASM memory included), read in the page and in each of its Workers through their own DevTools sessions, after forced GCs once the page has settled. Repeat open: after open minus before open, in a page that has already opened the file (the renderer's code is loaded; a Worker started per open counts in full, one kept between opens sits in the baseline and is reported as the idle Worker), median over 5 runs, split into page and Workers. First open: the same in a fresh page, so loading the renderer's code counts too. After a full read: every slide or sheet visited, or the document scrolled to its end (not on the large decks and workbooks). Left after destroy. The bench's own side reads (a sheet's RTL flag, a document's widest section) happen before anything is measured. Pixels held by canvas or the GPU are not counted",
  size: "download size: every file the browser fetched to open each real fixture, in a fresh page with caching off, from the bench's Vite production build (one minifier for every renderer, dependencies bundled, versions as installed by bench/package.json); each file counted once, gzip level 6 and brotli level 11, split into main-thread code, Workers and WASM. Per format, the union over that format's fixtures",
} as const;

/** A memory reading in KB: median and range over the runs, and every sample. */
export interface KbRange {
  n: number;
  median: number;
  min: number;
  max: number;
  samples: number[];
}

export function kbRange(bytes: readonly number[]): KbRange | null {
  const kb = bytes.map((b) => b / 1024);
  const s = stat(kb);
  return s === null ? null : { n: s.n, median: s.median, min: s.min, max: s.max, samples: kb.map((v) => Math.round(v)) };
}

export interface Cell {
  fixture: string;
  format: BenchFormat;
  renderer: string;
  /** Slides or sheets in the fixture; 0 for a document. */
  units: number | null;
  /** First open in a fresh page, ms. */
  open_cold_ms: Stat | null;
  /** Repeat opens in one page, ms. */
  open_warm_ms: Stat | null;
  /** One move to another slide or sheet, ms (pooled over every move of every pass). */
  nav_ms: Stat | null;
  /** All slides or sheets once, ms (per pass). */
  full_ms: Stat | null;
  /** Retained after a repeat open (page + Workers), and its page and Worker parts, in KB. */
  heap_open_kb: KbRange | null;
  heap_open_page_kb: KbRange | null;
  heap_open_worker_kb: KbRange | null;
  /** Retained after the first open in a fresh page, the renderer's code included, in KB. */
  heap_first_kb: KbRange | null;
  heap_first_page_kb: KbRange | null;
  heap_first_worker_kb: KbRange | null;
  /** Retained after a full read (every unit, or scrolled to the end), in KB. */
  heap_visited_kb: KbRange | null;
  /** Left after destroying, in KB. */
  heap_residual_kb: KbRange | null;
  /** Workers alive with no document open (a Worker kept for the next open), in KB. */
  heap_idle_worker_kb: KbRange | null;
  /** Elements the open painted into the stage. */
  elements: number | null;
  status: "pending" | "measured" | "error";
  error: string | null;
  /** When and where this cell was measured; absent in older files, whose header says. */
  measured_on?: MeasuredOn;
}

/** When, on which browser and OS, a cell or size row was measured. */
export interface MeasuredOn {
  date: string;
  browser: string;
  os: string;
}

/** A cell's or size row's identity across results: `fixture:renderer`. */
export const cellKey = (item: { fixture: string; renderer: string }): string => `${item.fixture}:${item.renderer}`;

/** What a plain `--run` measures: every aared cell, as `fixture:renderer` pairs. */
export function aaredPairs(): string[] {
  return buildCells().filter((cell) => cell.renderer === "aared").map(cellKey);
}

/**
 * Put freshly measured cells (or size rows) into a base result in place of the
 * ones they re-measure, in the canonical order (fixtures, then each format's
 * renderers, aared first): re-measuring aared keeps every competitor's numbers.
 */
export function mergeMeasured<T extends { fixture: string; renderer: string }>(base: readonly T[], fresh: readonly T[]): T[] {
  const byKey = new Map(base.map((item) => [cellKey(item), item]));
  for (const item of fresh) byKey.set(cellKey(item), item);
  const order = new Map(buildCells().map((cell, index) => [cellKey(cell), index]));
  const rank = (item: T): number => order.get(cellKey(item)) ?? Number.MAX_SAFE_INTEGER;
  return [...byKey.values()].sort((a, b) => rank(a) - rank(b));
}

/**
 * Which renderers were measured when, newest first, for a table's header — a
 * merged result holds aared's latest cells beside competitors measured once.
 * Empty when everything was measured together. Items without `measured_on`
 * date from the file's own header (`fallback`).
 */
export function describeMeasured(
  items: ReadonlyArray<{ renderer: string; measured_on?: MeasuredOn }>,
  fallback: { date: string; browser: string },
): string {
  const groups = new Map<string, Map<string, number>>();
  const totals = new Map<string, number>();
  for (const item of items) {
    const label = `${(item.measured_on?.date ?? fallback.date).slice(0, 10)} on ${item.measured_on?.browser ?? fallback.browser}`;
    let renderers = groups.get(label);
    if (renderers === undefined) groups.set(label, (renderers = new Map()));
    renderers.set(item.renderer, (renderers.get(item.renderer) ?? 0) + 1);
    totals.set(item.renderer, (totals.get(item.renderer) ?? 0) + 1);
  }
  if (groups.size <= 1) return "";
  return [...groups]
    .sort(([a], [b]) => b.localeCompare(a))
    .map(([label, renderers]) => {
      // Never hide a renderer whose cells span dates inside "the others":
      // a targeted rerun must not make its untouched fixtures look current.
      const partial = [...renderers].filter(([renderer, count]) => count < totals.get(renderer)!);
      const complete = [...renderers.keys()].filter((renderer) => !partial.some(([name]) => name === renderer));
      const names = [
        ...partial.map(([renderer, count]) => `${renderer} (${count} of ${totals.get(renderer)} cells)`),
        ...(complete.length > 3 ? ["the other renderers"] : complete),
      ];
      return `${names.join(", ")} measured ${label}`;
    })
    .join("; ");
}

/** One file a renderer fetched to open a fixture. */
export interface SizeFileEntry {
  path: string;
  kind: "main" | "worker" | "wasm";
  bytes: number;
  gzip: number;
  brotli: number;
}

/** What one renderer fetched to open one fixture. */
export interface SizeRow {
  renderer: string;
  format: BenchFormat;
  fixture: string;
  version: string | null;
  files: SizeFileEntry[];
  status: "measured" | "error";
  error: string | null;
  /** When and where this row was measured; absent in older files, whose header says. */
  measured_on?: MeasuredOn;
}

/** Totals over a set of files, each counted once; `main` / `worker` / `wasm` are gzip bytes. */
export interface SizeTotals {
  files: number;
  bytes: number;
  gzip: number;
  brotli: number;
  main: number;
  worker: number;
  wasm: number;
}

export interface SizeFormatRow {
  renderer: string;
  format: BenchFormat;
  fixtures: string[];
  version: string | null;
  totals: SizeTotals;
}

export interface SizeSummary {
  /** One row per renderer × format: everything it fetched across that format's fixtures. */
  formats: SizeFormatRow[];
  /** aared opening one file of each format in one page. */
  aaredAll: SizeTotals | null;
  /** The lightest other renderer per format, stitched into one page (a shared file counts once). */
  stitched: { totals: SizeTotals; parts: Array<{ renderer: string; format: BenchFormat }> } | null;
}

export function sizeTotals(files: readonly SizeFileEntry[]): SizeTotals {
  const unique = new Map<string, SizeFileEntry>();
  for (const file of files) if (!unique.has(file.path)) unique.set(file.path, file);
  const totals: SizeTotals = { files: 0, bytes: 0, gzip: 0, brotli: 0, main: 0, worker: 0, wasm: 0 };
  for (const file of unique.values()) {
    totals.files += 1;
    totals.bytes += file.bytes;
    totals.gzip += file.gzip;
    totals.brotli += file.brotli;
    totals[file.kind] += file.gzip;
  }
  return totals;
}

export function summarizeSize(rows: readonly SizeRow[]): SizeSummary {
  const ok = rows.filter((row) => row.status === "measured" && row.files.length > 0);
  const formats: SizeFormatRow[] = [];
  const filesOf = new Map<string, SizeFileEntry[]>();
  for (const format of BENCH_FORMATS) {
    // Every fixture of this format the pass tried: a renderer counts only when
    // it opened all of them, or its row would weigh a smaller set of files.
    const tried = new Set(rows.filter((row) => row.format === format).map((row) => row.fixture));
    for (const renderer of RENDERERS_BY_FORMAT[format]) {
      const mine = ok.filter((row) => row.format === format && row.renderer === renderer);
      if (mine.length === 0 || mine.length < tried.size) continue;
      const files = mine.flatMap((row) => row.files);
      filesOf.set(`${renderer}:${format}`, files);
      formats.push({
        renderer,
        format,
        fixtures: mine.map((row) => row.fixture),
        version: mine[0]!.version,
        totals: sizeTotals(files),
      });
    }
  }
  const aaredFiles = BENCH_FORMATS.map((format) => filesOf.get(`aared:${format}`));
  const aaredAll = aaredFiles.every((files) => files !== undefined) ? sizeTotals(aaredFiles.flatMap((files) => files!)) : null;
  const parts: Array<{ renderer: string; format: BenchFormat }> = [];
  for (const format of BENCH_FORMATS) {
    const others = formats.filter((row) => row.format === format && row.renderer !== "aared");
    if (others.length === 0) return { formats, aaredAll, stitched: null };
    const lightest = others.reduce((least, row) => (row.totals.gzip < least.totals.gzip ? row : least));
    parts.push({ renderer: lightest.renderer, format });
  }
  const stitched = { totals: sizeTotals(parts.flatMap((part) => filesOf.get(`${part.renderer}:${part.format}`)!)), parts };
  return { formats, aaredAll, stitched };
}

/** What the measurement page reports for one search session (mirrors `bench/src/measure.ts`). */
export interface SearchTimings {
  openMs: number;
  coldMs: number;
  warmMs: number;
  hitMs: number;
  hits: number;
}

/** aared's `search()` on a large fixture, medians over fresh sessions. */
export interface SearchCell {
  fixture: string;
  format: BenchFormat;
  size_kb: number;
  open_ms: Stat | null;
  /** First search: reads and indexes every unit. */
  cold_ms: Stat | null;
  /** Repeat search: scans what the first one kept. */
  warm_ms: Stat | null;
  /** A common word (many matches), after the miss searches. */
  hit_ms: Stat | null;
  hits: number | null;
  status: "measured" | "error";
  error: string | null;
}

export function pendingCell(fixture: string, renderer: string): Cell {
  return {
    fixture,
    format: formatOfFixture(fixture),
    renderer,
    units: null,
    open_cold_ms: null,
    open_warm_ms: null,
    nav_ms: null,
    full_ms: null,
    heap_open_kb: null,
    heap_open_page_kb: null,
    heap_open_worker_kb: null,
    heap_first_kb: null,
    heap_first_page_kb: null,
    heap_first_worker_kb: null,
    heap_visited_kb: null,
    heap_residual_kb: null,
    heap_idle_worker_kb: null,
    elements: null,
    status: "pending",
    error: null,
  };
}

/** One pending cell per fixture and each renderer that opens its format. */
export function buildCells(fixtures: readonly string[] = [...FIXTURES, ...LARGE_FIXTURES]): Cell[] {
  return fixtures.flatMap((fixture) =>
    RENDERERS_BY_FORMAT[formatOfFixture(fixture)].map((renderer) => pendingCell(fixture, renderer)),
  );
}

/** The package a renderer opens `format` with, or undefined when it names none. */
export function rendererPackage(renderer: string, format: BenchFormat): string | undefined {
  const entry = RENDERER_PACKAGES[renderer];
  return typeof entry === "string" ? entry : entry?.[format];
}

/**
 * The installed version of a renderer's package, resolved from `bench/` the
 * way the bench build resolves it. Fails loudly when resolution leaves the
 * workspace (Bun's global auto-install cache would hand back a different,
 * npm-latest version than the one the bench runs).
 */
export async function installedVersion(root: string, renderer: string, format: BenchFormat): Promise<string | null> {
  const { dirname, join, relative, isAbsolute } = await import("node:path");
  const { realpathSync, existsSync, readFileSync } = await import("node:fs");
  const name = rendererPackage(renderer, format);
  if (name === undefined) return null;
  const benchDir = join(root, "bench");
  const inWorkspace = (path: string): boolean => {
    const rel = relative(realpathSync(root), realpathSync(path));
    return !rel.startsWith("..") && !isAbsolute(rel);
  };
  // aared deliberately has no root export: resolve the same format subpath
  // its bench adapter imports, then find that package's own manifest.
  let dir = dirname(Bun.resolveSync(renderer === "aared" ? `${name}/${format}` : name, benchDir));
  for (;;) {
    const manifest = join(dir, "package.json");
    if (existsSync(manifest)) {
      const pkg = JSON.parse(readFileSync(manifest, "utf8")) as { name?: string; version?: string };
      if (pkg.name === name) {
        if (!inWorkspace(manifest)) throw new Error(`${name} resolved outside the workspace (${manifest}); run \`bun install\` in bench/`);
        return pkg.version ?? null;
      }
    }
    const parent = dirname(dir);
    if (parent === dir) return null;
    dir = parent;
  }
}

const kib = (bytes: number): string => (bytes / 1024).toFixed(1);

/** Human-readable size rows for docs/BENCHMARKS.md. */
export function sizeTableRows(summary: SizeSummary): string {
  const rows = summary.formats.map((row) => {
    const t = row.totals;
    return `| ${row.format} | ${row.renderer} | ${row.version ?? "—"} | ${kib(t.gzip)} | ${kib(t.brotli)} | ${kib(t.main)} | ${t.worker > 0 ? kib(t.worker) : "—"} | ${t.wasm > 0 ? kib(t.wasm) : "—"} | ${t.files} |`;
  });
  if (summary.aaredAll !== null) {
    const t = summary.aaredAll;
    rows.push(`| all three | aared | — | ${kib(t.gzip)} | ${kib(t.brotli)} | ${kib(t.main)} | ${kib(t.worker)} | — | ${t.files} |`);
  }
  if (summary.stitched !== null) {
    const t = summary.stitched.totals;
    const names = summary.stitched.parts.map((part) => `${part.renderer} (${part.format})`).join(" + ");
    rows.push(`| all three | ${names} | — | ${kib(t.gzip)} | ${kib(t.brotli)} | ${kib(t.main)} | ${t.worker > 0 ? kib(t.worker) : "—"} | ${t.wasm > 0 ? kib(t.wasm) : "—"} | ${t.files} |`);
  }
  return rows.join("\n");
}

/** Shared by the tools dispatcher and the direct CLI, before any browser is launched. */
export function validateBenchmarkArgs(argv: readonly string[]): void {
  const flags = new Set(argv.filter((arg) => arg.startsWith("--")));
  const known = new Set(["--run", "--size", "--search", "--calibrate", "--render-md", "--render-site", "--all", "--help"]);
  for (const flag of flags) if (!known.has(flag)) throw new Error(`Unknown benchmark option: ${flag}`);
  const browserModes = ["--run", "--size", "--search", "--calibrate"].filter((mode) => flags.has(mode));
  const report = flags.has("--render-md") || flags.has("--render-site");
  if (browserModes.length > 1 || (browserModes.length > 0 && report)) throw new Error("Choose one benchmark measurement mode, or report generation.");
  if (flags.has("--all") && !flags.has("--run") && !flags.has("--size")) throw new Error("--all requires --run or --size.");
  const pairs = argv.filter((arg) => !arg.startsWith("--"));
  if (pairs.length > 0 && (!flags.has("--run") || flags.has("--all"))) throw new Error("fixture:renderer selections require --run without --all.");
  const valid = new Set(buildCells().map(cellKey));
  for (const pair of pairs) if (!valid.has(pair)) throw new Error(`Unknown benchmark selection '${pair}' (want a known fixture:renderer pair).`);
}

/** Artifacts remain useful on failure, but the shell must not report a successful measurement. */
export function benchmarkExitCode(cells: readonly { status: string }[], remaining = 0): number {
  return cells.length === 0 || cells.some((cell) => cell.status === "error") || remaining > 0 ? 1 : 0;
}

export function sameBenchmarkMachine(a: Record<string, string | number>, b: Record<string, string | number>): boolean {
  return ["cpu", "cores", "ram_gb", "os", "platform", "arch"].every((key) => a[key] === b[key]);
}

async function main(): Promise<void> {
  validateBenchmarkArgs(process.argv.slice(2));
  const args = new Set(process.argv.slice(2));
  if (args.has("--help")) args.clear();
  const { mkdir, readFile, writeFile } = await import("node:fs/promises");
  const { join } = await import("node:path");
  const root = REPO_ROOT;
  const outDir = join(OUTPUT_ROOT, "benchmarks");
  const os = await import("node:os");
  const machine = {
    bun: process.versions.bun ?? "unknown",
    node: process.version,
    platform: process.platform,
    arch: process.arch,
    os: `${os.type()} ${os.release()}`,
    cpu: os.cpus()[0]?.model.trim() ?? "unknown",
    cores: os.cpus().length,
    ram_gb: Math.round(os.totalmem() / 2 ** 30),
  };

  if (args.has("--run") || args.has("--search") || args.has("--size") || args.has("--calibrate")) {
    for (const pair of process.argv.slice(2).filter((arg) => !arg.startsWith("--"))) {
      const fixture = pair.split(":")[0]!;
      if (!(await Bun.file(join(root, "fixtures", fixture)).exists()) && !(await Bun.file(join(OUTPUT_ROOT, "search-fixtures", fixture)).exists())) {
        throw new Error(`Requested benchmark fixture is missing: ${fixture}. Generate large fixtures with tools/benchmarks/search/make-fixtures.py.`);
      }
    }
    const { runBrowserPass, runCalibration, runSearchPass, runSizePass } = await import("./browser");
    const { Browser } = await import("../shared/cdp");
    await mkdir(outDir, { recursive: true });
    const browser = await Browser.launch();
    try {
      const product = await browser.product();
      const log = (line: string): void => console.log(line);
      // Everything measured now carries its own date and browser, so a merged
      // result can say which renderers were measured when.
      const measuredOn: MeasuredOn = { date: new Date().toISOString(), browser: product, os: machine.os };
      const all = args.has("--all");
      if (args.has("--calibrate")) {
        const ok = await runCalibration(browser, log);
        console.log(ok ? "memory reading calibrated" : "memory reading is WRONG — do not publish");
        if (!ok) process.exitCode = 1;
        return;
      }
      if (args.has("--size")) {
        const versions = new Map<string, string | null>();
        for (const format of BENCH_FORMATS) {
          for (const renderer of RENDERERS_BY_FORMAT[format]) versions.set(`${renderer}:${format}`, await installedVersion(root, renderer, format));
        }
        const measured = (await runSizePass(
          browser,
          (renderer, format) => versions.get(`${renderer}:${format}`) ?? null,
          log,
          (renderer) => all || renderer === "aared",
        )).map((row) => ({ ...row, measured_on: measuredOn }));
        const header = { version: 1, kind: "size", timestamp: measuredOn.date, machine: { ...machine, browser: product }, method: METHOD.size };
        // Without `--all`, aared's rows go into the latest size file in place of its old ones.
        const latest = all ? null : await latestResult<SizeFile>(outDir, "size-", 1);
        const base = latest !== null && sameBenchmarkMachine(latest.file.machine, machine) ? latest : null;
        const rows = base === null ? measured : mergeMeasured(base.file.rows, measured);
        if (base !== null) Object.assign(header, { timestamp: base.file.timestamp, machine: base.file.machine });
        const path = base === null ? join(outDir, `size-${stamp()}.json`) : join(outDir, base.name);
        await writeFile(path, JSON.stringify({ ...header, rows, summary: summarizeSize(rows) }, null, 2));
        const failed = measured.filter((row) => row.status === "error");
        console.log(`wrote ${path} (${measured.length} rows measured${base === null ? "" : `, merged into ${base.name}`}, ${failed.length} failed)`);
        process.exitCode = benchmarkExitCode(measured);
        return;
      }
      if (args.has("--search")) {
        const { cells, skipped } = await runSearchPass(browser, log);
        const path = join(outDir, `search-${stamp()}.json`);
        await writeFile(path, JSON.stringify({ version: 1, kind: "search", timestamp: new Date().toISOString(), machine: { ...machine, browser: product }, method: METHOD, cells }, null, 2));
        console.log(`wrote ${path}${skipped.length > 0 ? ` (not generated, skipped: ${skipped.join(", ")} — run tools/benchmarks/search/make-fixtures.py)` : ""}`);
        process.exitCode = benchmarkExitCode(cells);
        return;
      }
      // `--run`: aared's cells; named `fixture:renderer` pairs: just those;
      // both merged into the latest run. `--run --all`: every cell, a new run.
      const pairs = process.argv.slice(2).filter((arg) => !arg.startsWith("--"));
      const only = all ? [] : pairs.length > 0 ? pairs : aaredPairs();
      const { cells, skipped, remaining } = await runBrowserPass(browser, { only, log });
      const measured = cells.map((cell) => ({ ...cell, measured_on: measuredOn }));
      const header = { version: 2, kind: "run", timestamp: measuredOn.date, machine: { ...machine, browser: product }, method: METHOD };
      const latest = all ? null : await latestResult<RunFile>(outDir, null, 2);
      const base = latest !== null && sameBenchmarkMachine(latest.file.machine, machine) ? latest : null;
      const written = base === null ? measured : mergeMeasured(base.file.cells, measured);
      if (base !== null) Object.assign(header, { timestamp: base.file.timestamp, machine: base.file.machine });
      const path = base === null ? join(outDir, `${stamp()}.json`) : join(outDir, base.name);
      await writeFile(path, JSON.stringify({ ...header, cells: written }, null, 2));
      const failed = measured.filter((cell) => cell.status === "error");
      process.exitCode = benchmarkExitCode(measured, remaining.length);
      console.log(`wrote ${path} (${measured.length} cells measured${base === null ? "" : `, merged into ${base.name}`}, ${failed.length} failed${skipped.length > 0 ? `; not generated, skipped: ${skipped.join(", ")}` : ""})`);
      if (remaining.length > 0) {
        console.log(`stopped early: the browser stopped answering. Measure the rest into this file with a new run:\n  bun tools/benchmarks/measure.ts --run ${remaining.join(" ")}`);
        process.exitCode = 1;
      }
      return;
    } finally {
      browser.close();
    }
  }

  if (args.has("--render-md") || args.has("--render-site")) {
    const run = await latestResult<RunFile>(outDir, null, 2);
    const size = await latestResult<SizeFile>(outDir, "size-", 1);
    const search = await latestResult<SearchFile>(outDir, "search-", 1);

    if (args.has("--render-site")) {
      // The site's "By the numbers" section reads this file at build time;
      // nothing on the page is typed in by hand.
      const dataDir = join(root, "site", "src", "data");
      await mkdir(dataDir, { recursive: true });
      const payload = {
        generated: new Date().toISOString(),
        sources: { run: run?.name ?? null, size: size?.name ?? null, search: search?.name ?? null },
        machine: run?.file.machine ?? null,
        measured: run === null ? "" : describeMeasured(run.file.cells, { date: run.file.timestamp, browser: String(run.file.machine.browser) }),
        method: { runs: METHOD.runs, warmupDiscarded: METHOD.warmupDiscarded, coldRuns: METHOD.coldRuns, memoryRuns: METHOD.memoryRuns, navigationRuns: METHOD.navigationRuns },
        size: size?.file.summary ?? null,
        // The size file is merged separately from the run file (aared's rows re-measured over competitors'), so it names its own split.
        sizeMeasured: size === null ? "" : describeMeasured(size.file.rows, { date: size.file.timestamp, browser: String(size.file.machine.browser) }),
        run: (run?.file.cells ?? []).filter((cell) => cell.status === "measured"),
        search: (search?.file.cells ?? []).filter((cell) => cell.status === "measured"),
      };
      const out = join(dataDir, "benchmarks.json");
      await writeFile(out, `${JSON.stringify(payload, null, 1)}\n`);
      console.log(`wrote ${out} (${payload.run.length} run cells, ${payload.size?.formats.length ?? 0} size rows, ${payload.search.length} search cells)`);
      if (!args.has("--render-md")) return;
    }

    const cells: Cell[] = run?.file.cells ?? buildCells();
    const measured = cells.filter((c) => c.status === "measured");
    const med = (s: Stat | null, digits = 0): string => (s === null ? "—" : s.median.toFixed(digits));
    const medP95 = (s: Stat | null, digits = 0): string => (s === null ? "—" : `${s.median.toFixed(digits)} / ${s.p95.toFixed(digits)}`);
    const mb = (r: KbRange | null): string => (r === null ? "—" : (r.median / 1024).toFixed(1));
    const openRows =
      measured.length === 0
        ? `| *no measured runs yet* | — | — | — | — | — | — | — | — | — | — | — |`
        : cells
            .map((c) =>
              c.status === "error"
                ? `| ${c.fixture} | ${c.renderer} | *${c.error}* | | | | | | | | | |`
                : `| ${c.fixture} | ${c.renderer} | ${med(c.open_cold_ms)} | ${medP95(c.open_warm_ms)} | ${med(c.nav_ms, 1)} | ${med(c.full_ms)} | ${mb(c.heap_open_kb)} (${mb(c.heap_open_page_kb)} + ${mb(c.heap_open_worker_kb)}) | ${mb(c.heap_first_kb)} | ${mb(c.heap_visited_kb)} | ${mb(c.heap_residual_kb)} | ${c.heap_idle_worker_kb !== null && c.heap_idle_worker_kb.median > 0 ? mb(c.heap_idle_worker_kb) : "—"} | ${c.elements ?? "—"} |`,
            )
            .join("\n");
    const searchCells = search?.file.cells ?? [];
    const searchRows =
      searchCells.length === 0
        ? `| *no measured searches yet* | — | — | — | — | — |`
        : searchCells
            .map((c) =>
              c.status === "error"
                ? `| ${c.fixture} | *${c.error}* | | | | |`
                : `| ${c.fixture} | ${(c.size_kb / 1024).toFixed(1)} MB | ${med(c.open_ms)} | ${med(c.cold_ms)} | ${med(c.warm_ms, 1)} | ${med(c.hit_ms, 1)} (${c.hits} matches) |`,
            )
            .join("\n");
    const sizeRows = size === null ? `| *no measured sizes yet* | — | — | — | — | — | — | — | — |` : sizeTableRows(size.file.summary);
    const mdPath = join(root, "docs", "BENCHMARKS.md");
    let md = await readFile(mdPath, "utf8");
    // A merged result says which renderers were measured when (and on which browser).
    const when = (items: ReadonlyArray<{ renderer: string; measured_on?: MeasuredOn }>, file: { timestamp: string; machine: Record<string, string | number> }): string => {
      const line = describeMeasured(items, { date: file.timestamp, browser: String(file.machine.browser) });
      return line === "" ? "" : ` ${line[0]!.toUpperCase()}${line.slice(1)}.`;
    };
    const machineLine = run === null ? "" : `\n> Run \`${run.name}\` on ${describeMachine(run.file.machine)}.${when(run.file.cells, run.file)}\n`;
    const sizeLine = size === null ? "" : `\n> Size \`${size.name}\` from the bench's production build.${when(size.file.rows, size.file)}\n`;
    md = md.replace(
      /<!-- generated:open -->[\s\S]*?<!-- \/generated:open -->/,
      `<!-- generated:open -->${machineLine}\n| Fixture | Renderer | cold open ms (med) | warm open ms (med / p95) | one move ms (med) | all units ms (med) | memory, repeat open MB (page + Workers) | memory, first open MB | memory after a full read MB | left after destroy MB | idle Worker kept MB | elements painted |\n|---|---|---|---|---|---|---|---|---|---|---|---|\n${openRows}\n<!-- /generated:open -->`,
    );
    md = md.replace(
      /<!-- generated:search -->[\s\S]*?<!-- \/generated:search -->/,
      `<!-- generated:search -->\n| Fixture | Size | open ms | first search ms | repeat search ms | common word ms |\n|---|---|---|---|---|---|\n${searchRows}\n<!-- /generated:search -->`,
    );
    md = md.replace(
      /<!-- generated:size -->[\s\S]*?<!-- \/generated:size -->/,
      `<!-- generated:size -->${sizeLine}\n| Format | Renderer | version | gzip KiB | brotli KiB | main thread KiB (gzip) | Workers KiB (gzip) | WASM KiB (gzip) | files |\n|---|---|---|---|---|---|---|---|---|\n${sizeRows}\n<!-- /generated:size -->`,
    );
    await writeFile(mdPath, md);
    console.log(`rendered ${measured.length} run rows, size from ${size?.name ?? "none yet"}, search from ${search?.name ?? "none yet"}`);
    return;
  }

  console.log(
    [
      "usage: bun tools/benchmarks/measure.ts <mode>   (every browser mode needs `bun run --cwd bench build`)",
      "  --run                  open / navigation / memory for aared, merged into the latest run (competitors kept)",
      "  --run fixture:renderer …   just those cells, merged the same way (redo a failed cell)",
      "  --run --all            every renderer, into a new run file (first run, or a competitor's version changed)",
      "  --size [--all]         download size: the files aared (every renderer with --all) fetches to open each fixture",
      "  --search               aared search() on the large generated fixtures",
      "  --calibrate            check the memory reading against 10 MiB controls on the page and in a Worker",
      "  --render-md            regenerate the tables in docs/BENCHMARKS.md from the latest results",
      "  --render-site          write site/src/data/benchmarks.json for the home page's numbers section",
    ].join("\n"),
  );
}

interface RunFile {
  version?: number;
  timestamp: string;
  machine: Record<string, string | number>;
  cells: Cell[];
}

interface SizeFile {
  version?: number;
  timestamp: string;
  machine: Record<string, string | number>;
  rows: SizeRow[];
  summary: SizeSummary;
}

/**
 * The newest result of one kind (`prefix` `size-` or `search-`; `null` for
 * runs) taken with the current method: run files from version 2 (Workers read
 * in their own sessions), size files from version 1 (fetched files, one
 * build). Older files are ignored, never mixed in.
 */
async function latestResult<T extends { version?: number }>(
  outDir: string,
  prefix: string | null,
  minVersion: number,
): Promise<{ name: string; file: T } | null> {
  const { readdir, readFile } = await import("node:fs/promises");
  const { join } = await import("node:path");
  const names = (await readdir(outDir).catch(() => [] as string[]))
    .filter((name) => name.endsWith(".json"))
    .filter((name) => (prefix === null ? !name.startsWith("size-") && !name.startsWith("search-") : name.startsWith(prefix)))
    .sort()
    .reverse();
  for (const name of names) {
    const file = JSON.parse(await readFile(join(outDir, name), "utf8")) as T;
    if ((file.version ?? 0) >= minVersion) return { name, file };
  }
  return null;
}

interface SearchFile {
  version?: number;
  timestamp: string;
  machine: Record<string, string | number>;
  cells: SearchCell[];
}

function describeMachine(machine: Record<string, string | number>): string {
  return `${machine.cpu} (${machine.cores} threads), ${machine.ram_gb} GB, ${machine.os}, ${machine.browser}, Bun ${machine.bun}`;
}

function stamp(): string {
  const d = new Date();
  const pad = (n: number): string => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

if (import.meta.main) {
  await main();
}
