import { describe, expect, test } from "bun:test";
import {
  FIXTURES,
  LARGE_FIXTURES,
  METHOD,
  RENDERERS_BY_FORMAT,
  buildCells,
  formatBytes,
  formatOfFixture,
  kbRange,
  rendererPackage,
  sizeTotals,
  stat,
  summarizeSize,
  type BenchFormat,
  type SizeFileEntry,
  type SizeRow,
} from "./bench-measure";

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

  test("formatBytes uses KiB above 1024", () => {
    expect(formatBytes(512)).toBe("512 B");
    expect(formatBytes(2048)).toBe("2.0 KiB");
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
