/**
 * Times `search()` on the large generated fixtures (see make-fixtures.py).
 *
 *   python tools/benchmarks/search/make-fixtures.py
 *   bun tools/benchmarks/search/bench.ts [pptx|xlsx|docx ...]
 *
 * Per format: open, first search (cold: parses/indexes every unit), repeat
 * searches (warm), a common-word query, and how long an abort takes to land.
 * "miss" is a query that matches nothing, so every unit is scanned. (XLSX
 * `search()` is per sheet, so its figures are a loop over every sheet.)
 */
import { open as openDocx } from "../../../packages/aared/src/docx/index.ts";
import { open as openPptx } from "../../../packages/aared/src/pptx/index.ts";
import { open as openXlsx } from "../../../packages/aared/src/xlsx/index.ts";
import { OUTPUT_ROOT } from "../../shared/paths.js";

const OUT = `${OUTPUT_ROOT}/search-fixtures`;
const MISS = "zzqq-no-such-text";

interface Searchable {
  search(query: string, options?: { signal?: AbortSignal }): Promise<readonly unknown[]>;
  close(): void;
}

/** A workbook is searched one sheet at a time; this times the whole file, so it searches every sheet in tab order. */
async function openWorkbook(bytes: ArrayBuffer): Promise<Searchable> {
  const workbook = await openXlsx(bytes);
  return {
    async search(query, options) {
      const matches: unknown[] = [];
      for (const sheet of workbook.sheets.list()) matches.push(...(await sheet.search(query, options)));
      return matches;
    },
    close: () => workbook.close(),
  };
}

const formats: Record<string, { file: string; open: (bytes: ArrayBuffer) => Promise<Searchable>; common: string }> = {
  pptx: { file: "large-300-slides.pptx", open: openPptx as never, common: "the" },
  xlsx: { file: "large-13-sheets.xlsx", open: openWorkbook, common: "the" },
  docx: { file: "large-40x.docx", open: openDocx as never, common: "the" },
};

const mb = (bytes: number): string => (bytes / 1e6).toFixed(0);
const ms = (value: number): string => value.toFixed(1).padStart(8);

async function time<T>(run: () => T | Promise<T>): Promise<[number, T]> {
  const start = performance.now();
  const result = await run();
  return [performance.now() - start, result];
}

async function median(run: () => unknown, times = 5): Promise<number> {
  const samples: number[] = [];
  for (let i = 0; i < times; i++) samples.push((await time(async () => await run()))[0]);
  return samples.sort((a, b) => a - b)[Math.floor(times / 2)]!;
}

const requested = process.argv.slice(2);
for (const [name, { file, open, common }] of Object.entries(formats)) {
  if (requested.length > 0 && !requested.includes(name)) continue;
  const path = `${OUT}/${file}`;
  if (!(await Bun.file(path).exists())) {
    console.log(`${name}: missing ${file} — run make-fixtures.py first`);
    continue;
  }
  const bytes = await Bun.file(path).arrayBuffer();
  Bun.gc(true);
  const rssBefore = process.memoryUsage().rss;
  const [openMs, doc] = await time(() => open(bytes));
  const [coldMs, cold] = await time(async () => await doc.search(MISS));
  const warmMs = await median(async () => await doc.search(MISS));
  const [hitMs, hits] = await time(async () => await doc.search(common));
  const warmHitMs = await median(async () => await doc.search(common));
  // Abort: fire during a second full scan on a fresh document.
  const abortDoc = await open(bytes);
  const controller = new AbortController();
  const started = performance.now();
  const pending = abortDoc.search(MISS, { signal: controller.signal }).then(
    () => "finished",
    () => "aborted",
  );
  setTimeout(() => controller.abort(), 5);
  const outcome = await pending;
  const abortMs = performance.now() - started;
  const rssAfter = process.memoryUsage().rss;
  console.log(
    [
      `${name.padEnd(5)} ${file}`,
      `  open          ${ms(openMs)} ms`,
      `  miss  cold    ${ms(coldMs)} ms   (${cold.length} matches)`,
      `  miss  warm    ${ms(warmMs)} ms   (median of 5)`,
      `  "${common}" cold  ${ms(hitMs)} ms   (${hits.length} matches)`,
      `  "${common}" warm  ${ms(warmHitMs)} ms   (median of 5)`,
      `  abort at 5ms  ${ms(abortMs)} ms   (${outcome})`,
      `  rss           ${mb(rssBefore)} -> ${mb(rssAfter)} MB`,
    ].join("\n"),
  );
  doc.close();
  abortDoc.close();
}
process.exit(0);
