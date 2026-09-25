/**
 * The home page's "By the numbers" data: pure functions over the generated
 * results in `src/data/benchmarks.json` (`bun tools/bench-measure.ts
 * --render-site`). Nothing here is typed in by hand — every figure and every
 * caption is computed from a measurement, so a slower result reads as slower.
 */

export type Format = "pptx" | "xlsx" | "docx";

export interface Stat {
  median: number;
  p95: number;
  min: number;
  max: number;
  n: number;
}

/** A memory reading in KiB (1024 bytes), as the harness records it. */
export interface KbRange {
  median: number;
  min: number;
  max: number;
  n: number;
}

export interface RunCell {
  fixture: string;
  format: Format;
  renderer: string;
  units: number | null;
  open_cold_ms: Stat | null;
  open_warm_ms: Stat | null;
  nav_ms: Stat | null;
  full_ms: Stat | null;
  heap_open_kb: KbRange | null;
  heap_open_page_kb: KbRange | null;
  heap_open_worker_kb: KbRange | null;
  heap_first_kb: KbRange | null;
  heap_first_page_kb?: KbRange | null;
  heap_first_worker_kb?: KbRange | null;
  heap_visited_kb: KbRange | null;
  heap_residual_kb: KbRange | null;
  /** Workers kept alive with no document open (aared parks one for the next open). */
  heap_idle_worker_kb: KbRange | null;
  elements: number | null;
}

/** Download totals over a set of files, each counted once, in bytes; `main` / `worker` / `wasm` are gzip bytes. */
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
  format: Format;
  fixtures: string[];
  version: string | null;
  totals: SizeTotals;
}

export interface SizeSummary {
  formats: SizeFormatRow[];
  aaredAll: SizeTotals | null;
  stitched: { totals: SizeTotals; parts: Array<{ renderer: string; format: Format }> } | null;
}

export interface SearchCell {
  fixture: string;
  format: Format;
  size_kb: number;
  open_ms: Stat | null;
  cold_ms: Stat | null;
  warm_ms: Stat | null;
  hit_ms: Stat | null;
  hits: number | null;
}

export interface Benchmarks {
  generated: string;
  machine: Record<string, string | number> | null;
  method: { runs: number; warmupDiscarded: number; coldRuns: number; memoryRuns: number; navigationRuns: number };
  size: SizeSummary | null;
  run: RunCell[];
  search: SearchCell[];
}

export const FORMAT_LABEL: Record<Format, string> = { pptx: "PowerPoint", xlsx: "Excel", docx: "Word" };
export const FORMATS: readonly Format[] = ["pptx", "xlsx", "docx"];

/** One bar: a renderer and its value in the chart's unit. */
export interface Bar {
  label: string;
  value: number;
  /** aared's own bar. */
  ours: boolean;
  /** Small text beside the value (a ratio, a note). */
  note?: string;
  /** Share of the bar (0–1) drawn in the second colour: the part held in Workers. */
  split?: number;
  /** The breakdown, shown on hover and read by screen readers, instead of more text on the row. */
  tooltip?: string;
}

export interface BarGroup {
  title: string;
  bars: Bar[];
}

const isLarge = (fixture: string): boolean => fixture.startsWith("large-");

/** Bars sorted best (smallest) first. */
function sorted(bars: Bar[]): Bar[] {
  return [...bars].sort((a, b) => a.value - b.value);
}

const median = (values: readonly number[]): number => {
  const s = [...values].sort((a, b) => a - b);
  const mid = s.length >> 1;
  return s.length % 2 === 1 ? s[mid]! : (s[mid - 1]! + s[mid]!) / 2;
};

export type Measure = "open_warm" | "open_cold" | "heap" | "heap_first";

/** KiB, as recorded, to decimal MB, as shown. */
const kibToMb = (kib: number): number => (kib * 1024) / 1e6;

function pick(cell: RunCell, measure: Measure): number | null {
  switch (measure) {
    case "open_warm":
      return cell.open_warm_ms?.median ?? null;
    case "open_cold":
      return cell.open_cold_ms?.median ?? null;
    case "heap":
      return cell.heap_open_kb === null ? null : kibToMb(cell.heap_open_kb.median);
    case "heap_first":
      return cell.heap_first_kb === null ? null : kibToMb(cell.heap_first_kb.median);
  }
}

/** A per-cell KiB reading's median across cells, in MB, when any cell has one above zero. */
function cellsMb(cells: readonly RunCell[], read: (cell: RunCell) => KbRange | null): number | null {
  const values = cells.map((cell) => read(cell)?.median ?? 0).filter((kib) => kib > 0);
  return values.length === 0 ? null : kibToMb(median(values));
}

/** A repeat-open bar's Worker share (drawn as the bar's second colour) and its breakdown for the tooltip. */
function workerSplit(cells: readonly RunCell[], total: number, first: boolean): { split?: number; tooltip?: string } {
  const inWorkers = cellsMb(cells, (cell) => (first ? cell.heap_first_worker_kb ?? null : cell.heap_open_worker_kb));
  // A first open starts its own Worker, so no idle one is kept beside it.
  const idle = first ? null : cellsMb(cells, (cell) => cell.heap_idle_worker_kb);
  if (inWorkers === null && idle === null) return {};
  const lines = [`Page ${formatMb(Math.max(0, total - (inWorkers ?? 0)))}`, `Workers ${formatMb(inWorkers ?? 0)}`];
  if (idle !== null) lines.push(`Plus ${formatMb(idle)} in a Worker kept idle between opens`);
  return {
    ...(inWorkers === null || total <= 0 ? {} : { split: Math.min(1, inWorkers / total) }),
    tooltip: lines.join(" · "),
  };
}

/**
 * One measure across renderers, per format. `large` picks the big generated
 * files, otherwise the real ones (a renderer measured on several files gets the
 * median of its medians). With `growth`, each bar notes how much the large file
 * cost relative to the real ones.
 */
export function groups(data: Benchmarks, measure: Measure, large: boolean, growth = false): BarGroup[] {
  const out: BarGroup[] = [];
  for (const format of FORMATS) {
    const cells = data.run.filter((cell) => cell.format === format);
    const renderers = [...new Set(cells.map((cell) => cell.renderer))];
    const bars: Bar[] = [];
    for (const renderer of renderers) {
      const mine = cells.filter((cell) => cell.renderer === renderer);
      const values = (isLargeSet: boolean): number[] =>
        mine
          .filter((cell) => isLarge(cell.fixture) === isLargeSet)
          .map((cell) => pick(cell, measure))
          .filter((v): v is number => v !== null);
      const now = values(large);
      if (now.length === 0) continue;
      const value = median(now);
      let note: string | undefined;
      if (growth && large) {
        const before = values(false);
        if (before.length > 0) {
          const ratio = value / median(before);
          note = ratio >= 10 ? `×${ratio.toFixed(0)}` : `×${ratio.toFixed(1)}`;
        }
      }
      const heap = measure === "heap" || measure === "heap_first";
      const workers = heap ? workerSplit(mine.filter((cell) => isLarge(cell.fixture) === large), value, measure === "heap_first") : {};
      bars.push({ label: renderer, value, ours: renderer === "aared", ...(note === undefined ? {} : { note }), ...workers });
    }
    if (bars.length > 1) out.push({ title: FORMAT_LABEL[format], bars: sorted(bars) });
  }
  return out;
}

/** Bytes to decimal KB, as shown. */
const kb = (bytes: number): number => bytes / 1000;

/**
 * One line for the whole size chart: which bars include a Worker or WASM,
 * and how much (the largest a renderer fetched in any format), said once
 * instead of on every bar.
 */
export function sizeSplitNote(data: Benchmarks): string {
  if (data.size === null) return "";
  const worker = new Map<string, number>();
  const wasm = new Map<string, number>();
  for (const row of data.size.formats) {
    if (row.totals.worker > 0) worker.set(row.renderer, Math.max(worker.get(row.renderer) ?? 0, row.totals.worker));
    if (row.totals.wasm > 0) wasm.set(row.renderer, Math.max(wasm.get(row.renderer) ?? 0, row.totals.wasm));
  }
  const parts = [
    ...[...worker].map(([renderer, bytes]) => `${renderer}'s ${formatKb(kb(bytes))} parse Worker`),
    ...[...wasm].map(([renderer, bytes]) => `${renderer}'s ${formatKb(kb(bytes))} of WASM`),
  ];
  return parts.length === 0 ? "" : `Bars include ${parts.join(" and ")}.`;
}

/** Gzipped download per format: everything each renderer fetched to open that format's files. */
export function sizeGroups(data: Benchmarks): BarGroup[] {
  if (data.size === null) return [];
  const out: BarGroup[] = [];
  for (const format of FORMATS) {
    const bars: Bar[] = data.size.formats
      .filter((row) => row.format === format)
      .map((row) => ({ label: row.renderer, value: kb(row.totals.gzip), ours: row.renderer === "aared" }));
    if (bars.length > 1) out.push({ title: FORMAT_LABEL[format], bars: sorted(bars) });
  }
  return out;
}

/** aared opening one file of each format in one page, gzipped KB. */
export function aaredAllGzipKb(data: Benchmarks): number | null {
  return data.size?.aaredAll == null ? null : kb(data.size.aaredAll.gzip);
}

/** The lightest other renderer per format stitched into one page (a shared file counted once), gzipped KB. */
export function stitchedGzipKb(data: Benchmarks): { total: number; parts: Array<{ label: string; format: Format }> } | null {
  const stitched = data.size?.stitched;
  if (stitched == null) return null;
  return { total: kb(stitched.totals.gzip), parts: stitched.parts.map((part) => ({ label: part.renderer, format: part.format })) };
}

/**
 * Why a renderer that downloads less than aared can: what it leaves out,
 * from docs/COMPETITORS.md. Shown only when that renderer is the lighter one.
 */
export const LIGHTER_BECAUSE: Readonly<Record<string, string>> = {
  mdurbar:
    "mdurbar is lighter because it draws about half of Office's preset shapes and turns the rest into rectangles, has no right-to-left text handling, and parses on the page; on the coverage deck it also paints one slide's table on every slide. aared ships Office's full preset-shape catalog.",
};

/** The reason note for every format where a lighter renderer beats aared on download size. */
export function lighterNotes(chart: readonly BarGroup[]): string[] {
  const notes: string[] = [];
  for (const group of chart) {
    const ours = group.bars.find((bar) => bar.ours);
    const lightest = fastestOther(group.bars);
    if (ours === undefined || lightest === null || lightest.value >= ours.value) continue;
    const reason = LIGHTER_BECAUSE[lightest.label];
    if (reason !== undefined) notes.push(`${group.title}: ${reason}`);
  }
  return notes;
}

/** aared's own numbers on the large generated files, for the search chart. */
export function searchGroups(data: Benchmarks): BarGroup[] {
  return data.search
    .filter((cell) => cell.cold_ms !== null && cell.warm_ms !== null)
    .map((cell) => ({
      title: `${FORMAT_LABEL[cell.format]} · ${(cell.size_kb / 1024).toFixed(1)} MB`,
      bars: [
        { label: "first search", value: cell.cold_ms!.median, ours: true, note: "reads every unit" },
        { label: "repeat search", value: cell.warm_ms!.median, ours: true, note: "scans what it kept" },
      ],
    }));
}

/** "48 ms", "1.2 s" — the value with a unit that keeps it readable. */
export function formatMs(ms: number): string {
  if (ms >= 1000) return `${(ms / 1000).toFixed(ms >= 10000 ? 0 : 1)} s`;
  if (ms >= 100) return `${ms.toFixed(0)} ms`;
  if (ms >= 10) return `${ms.toFixed(0)} ms`;
  return `${ms.toFixed(1)} ms`;
}

/** Decimal units: 1 KB = 1000 bytes, 1 MB = 1000 KB. */
export const formatKb = (value: number): string => (value >= 1000 ? `${(value / 1000).toFixed(1)} MB` : `${value.toFixed(value >= 100 ? 0 : 1)} KB`);
export const formatMb = (mb: number): string => (mb < 0.1 ? "<0.1 MB" : `${mb.toFixed(mb >= 10 ? 0 : 1)} MB`);

/** aared's open time on one fixture group, for the headline stats. */
export function aaredOpenMs(data: Benchmarks, format: Format, large: boolean): number | null {
  const cells = data.run.filter((c) => c.renderer === "aared" && c.format === format && isLarge(c.fixture) === large && c.open_warm_ms !== null);
  return cells.length === 0 ? null : median(cells.map((c) => c.open_warm_ms!.median));
}

export function largestUnits(data: Benchmarks, format: Format): number | null {
  const cells = data.run.filter((c) => c.renderer === "aared" && c.format === format && isLarge(c.fixture) && c.units !== null);
  return cells.length === 0 ? null : Math.max(...cells.map((c) => c.units!));
}

/** The fastest bar that isn't aared, for a neutral "fastest other" caption. */
export function fastestOther(bars: readonly Bar[]): Bar | null {
  const others = bars.filter((bar) => !bar.ours);
  return others.length === 0 ? null : others.reduce((least, bar) => (bar.value < least.value ? bar : least));
}
