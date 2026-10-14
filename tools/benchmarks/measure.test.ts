import { describe, expect, test } from "bun:test";
import {
  BENCH_FORMATS,
  FIXTURES,
  LARGE_FIXTURES,
  METHOD,
  RENDERERS_BY_FORMAT,
  aaredPairs,
  buildCells,
  cellKey,
  describeMeasured,
  formatOfFixture,
  kbRange,
  installedVersion,
  mergeMeasured,
  rendererPackage,
  sizeTotals,
  stat,
  summarizeSize,
  type BenchFormat,
  type Cell,
  type SizeFileEntry,
  type SizeRow,
} from "./measure";

describe("bench-measure contract", () => {
  test("one pending cell per fixture and each renderer that opens its format", () => {
    const cells = buildCells();
    const expected = [...FIXTURES, ...LARGE_FIXTURES].reduce(
      (sum, fixture) => sum + RENDERERS_BY_FORMAT[formatOfFixture(fixture)].length,
      0,
    );
    expect(cells).toHaveLength(expected);
    expect(cells.every((c) => c.status === "pending")).toBe(true);
    expect(cells.every((c) => c.open_warm_ms === null && c.heap_open_kb === null && c.heap_first_kb === null)).toBe(true);
    // A renderer never appears on a format it doesn't open.
    expect(cells.some((c) => c.renderer === "docx-preview" && c.format !== "docx")).toBe(false);
  });

  test("every renderer names the package its version is read from, per format it opens", () => {
    for (const [format, renderers] of Object.entries(RENDERERS_BY_FORMAT) as Array<[BenchFormat, readonly string[]]>) {
      for (const renderer of renderers) expect(rendererPackage(renderer, format), `${renderer} (${format})`).toBeString();
    }
  });

  test("aared is measured on every format", () => {
    for (const renderers of Object.values(RENDERERS_BY_FORMAT)) {
      expect(renderers[0]).toBe("aared");
    }
  });

  test("aared's installed version resolves through its format exports, never a nonexistent root entry", async () => {
    const manifest = await Bun.file(`${import.meta.dir}/../../packages/aared/package.json`).json();
    for (const format of BENCH_FORMATS) {
      expect(await installedVersion(`${import.meta.dir}/../..`, "aared", format)).toBe(manifest.version);
    }
  });

  test("method pins runs, viewport, fonts, and says where memory is read", () => {
    expect(METHOD.runs).toBe(11);
    expect(METHOD.warmupDiscarded).toBe(2);
    expect(METHOD.viewport).toContain("1280x720");
    expect(METHOD.fonts).toContain("nothing fetched");
    expect(METHOD.memory).toContain("each of its Workers");
    expect(METHOD.size).toContain("fetched");
  });

  test("stat reports median, p95 and range", () => {
    expect(stat([])).toBeNull();
    const s = stat([5, 1, 3, 2, 4, 9, 8, 7, 6, 10])!;
    expect(s.n).toBe(10);
    expect(s.median).toBe(5);
    expect(s.p95).toBe(10);
    expect([s.min, s.max]).toEqual([1, 10]);
  });

  test("kbRange keeps every sample, in KB", () => {
    expect(kbRange([])).toBeNull();
    const r = kbRange([2048, 1024, 3072])!;
    expect([r.n, r.median, r.min, r.max]).toEqual([3, 2, 1, 3]);
    expect(r.samples).toEqual([2, 1, 3]);
  });

  test("the harness lists the renderers the bench registry opens, per format", async () => {
    const { BASELINES } = await import("../../bench/src/adapters");
    for (const format of BENCH_FORMATS) {
      expect(RENDERERS_BY_FORMAT[format].slice(1).sort(), format).toEqual(Object.keys(BASELINES[format]).sort());
    }
  });
});

const file = (path: string, gzip: number, kind: SizeFileEntry["kind"] = "main"): SizeFileEntry => ({
  path,
  kind,
  bytes: gzip * 3,
  gzip,
  brotli: gzip - 1,
});

const row = (renderer: string, fixture: string, files: SizeFileEntry[]): SizeRow => ({
  renderer,
  format: formatOfFixture(fixture),
  fixture,
  version: "1.0.0",
  files,
  status: "measured",
  error: null,
});

describe("measuring aared only, keeping the competitors", () => {
  const cell = (fixture: string, renderer: string, date?: string, browser = "HeadlessChrome/154.0.0.0"): Cell => ({
    ...buildCells().find((c) => c.fixture === fixture && c.renderer === renderer)!,
    status: "measured",
    ...(date === undefined ? {} : { measured_on: { date, browser, os: "Windows_NT 10.0.26300" } }),
  });

  test("a plain run measures every aared cell and nothing else", () => {
    const pairs = aaredPairs();
    expect(pairs).toHaveLength(FIXTURES.length + LARGE_FIXTURES.length);
    expect(pairs.every((pair) => pair.endsWith(":aared"))).toBe(true);
  });

  test("re-measured cells replace their old ones in place, in the canonical order, and the rest stay", () => {
    const base = [cell("pptx-coverage.pptx", "mdurbar"), cell("pptx-coverage.pptx", "aared"), cell("docx-coverage.docx", "omni")];
    const fresh = [cell("pptx-coverage.pptx", "aared", "2026-10-01T18:00:00Z"), cell("xlsx-coverage.xlsx", "aared", "2026-10-01T18:00:00Z")];
    const merged = mergeMeasured(base, fresh);
    expect(merged.map(cellKey)).toEqual([
      "pptx-coverage.pptx:aared",
      "pptx-coverage.pptx:mdurbar",
      "xlsx-coverage.xlsx:aared",
      "docx-coverage.docx:omni",
    ]);
    expect(merged[0]!.measured_on?.date).toBe("2026-10-01T18:00:00Z");
  });

  test("a merged result says which renderers were measured when, newest first", () => {
    const fallback = { date: "2026-09-24T23:58:05Z", browser: "HeadlessChrome/153.0.0.0" };
    const others = ["pptxviewjs", "silurus", "aiden0z", "mdurbar"].map((renderer) => cell("pptx-coverage.pptx", renderer));
    expect(describeMeasured([cell("pptx-coverage.pptx", "aared", "2026-10-01T18:00:00Z"), ...others], fallback)).toBe(
      "aared measured 2026-10-01 on HeadlessChrome/154.0.0.0; the other renderers measured 2026-09-24 on HeadlessChrome/153.0.0.0",
    );
    // Measured together: nothing to add to the header.
    expect(describeMeasured(others, fallback)).toBe("");
  });

  test("partial reruns disclose the older cells of the same renderer instead of hiding them among competitors", () => {
    const fallback = { date: "2026-09-24T23:58:05Z", browser: "HeadlessChrome/153.0.0.0" };
    const others = ["pptxviewjs", "silurus", "aiden0z", "mdurbar"].map((renderer) => cell("pptx-coverage.pptx", renderer));
    expect(describeMeasured([
      cell("pptx-coverage.pptx", "aared", "2026-10-01T18:00:00Z"),
      cell("docx-coverage.docx", "aared"), ...others,
    ], fallback)).toBe(
      "aared (1 of 2 cells) measured 2026-10-01 on HeadlessChrome/154.0.0.0; aared (1 of 2 cells), the other renderers measured 2026-09-24 on HeadlessChrome/153.0.0.0",
    );
  });
});

describe("download size summary", () => {
  test("a file counts once, and the split follows its kind", () => {
    const totals = sizeTotals([file("/assets/a.js", 10), file("/assets/a.js", 10), file("/assets/w.js", 20, "worker"), file("/assets/p.wasm", 30, "wasm")]);
    expect(totals).toEqual({ files: 3, bytes: 180, gzip: 60, brotli: 57, main: 10, worker: 20, wasm: 30 });
  });

  test("per format is the union over that format's fixtures", () => {
    const summary = summarizeSize([
      row("aared", "pptx-coverage.pptx", [file("/assets/pptx.js", 30), file("/assets/worker.js", 40, "worker"), file("/assets/geometry.js", 2)]),
      row("aared", "post-production-review.pptx", [file("/assets/pptx.js", 30), file("/assets/worker.js", 40, "worker")]),
    ]);
    const pptx = summary.formats.find((r) => r.renderer === "aared" && r.format === "pptx")!;
    expect(pptx.totals.gzip).toBe(72);
    expect(pptx.fixtures).toEqual(["pptx-coverage.pptx", "post-production-review.pptx"]);
  });

  test("all three and the stitched alternative share files once, and skip failed rows", () => {
    const summary = summarizeSize([
      row("aared", "pptx-coverage.pptx", [file("/assets/pptx.js", 30), file("/assets/worker.js", 40, "worker")]),
      row("aared", "xlsx-coverage.xlsx", [file("/assets/xlsx.js", 10), file("/assets/worker.js", 40, "worker")]),
      row("aared", "docx-coverage.docx", [file("/assets/docx.js", 12), file("/assets/worker.js", 40, "worker")]),
      row("mdurbar", "pptx-coverage.pptx", [file("/assets/mdurbar.js", 29)]),
      row("pptxviewjs", "pptx-coverage.pptx", [file("/assets/pptxviewjs.js", 200), file("/assets/jszip.js", 30)]),
      row("sheetjs", "xlsx-coverage.xlsx", [file("/assets/xlsx-sheetjs.js", 160)]),
      row("docx-preview", "docx-coverage.docx", [file("/assets/docx-preview.js", 21), file("/assets/jszip.js", 30)]),
      { ...row("docx-renderer", "docx-coverage.docx", []), status: "error", error: "crashed" },
    ]);
    expect(summary.aaredAll?.gzip).toBe(92);
    expect(summary.stitched?.parts).toEqual([
      { renderer: "mdurbar", format: "pptx" },
      { renderer: "sheetjs", format: "xlsx" },
      { renderer: "docx-preview", format: "docx" },
    ]);
    expect(summary.stitched?.totals.gzip).toBe(29 + 160 + 21 + 30);
    expect(summary.formats.some((r) => r.renderer === "docx-renderer")).toBe(false);
  });

  test("a renderer that failed on one of its format's files gets no row", () => {
    const summary = summarizeSize([
      row("aared", "pptx-coverage.pptx", [file("/assets/pptx.js", 30)]),
      row("aared", "post-production-review.pptx", [file("/assets/pptx.js", 30)]),
      row("mdurbar", "pptx-coverage.pptx", [file("/assets/mdurbar.js", 29)]),
      { ...row("mdurbar", "post-production-review.pptx", []), status: "error", error: "crashed" },
    ]);
    expect(summary.formats.map((r) => r.renderer)).toEqual(["aared"]);
  });
});
