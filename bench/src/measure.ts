import { BASELINES } from "./adapters";
import { openAaredDocx, openAaredPptx, openAaredXlsx } from "./adapters/aared";
import { maxSectionWidthPx } from "./adapters/docx";
import type { OpenAdapter, OpenedDoc } from "./adapters/types";
import { sheetRtlFlags } from "./adapters/xlsx";
import type { Format } from "./lib/format";

/**
 * The measurement page `tools/bench-measure.ts` drives in a real browser.
 * One stage, one renderer at a time; every renderer goes through the same
 * adapter contract the bench UI uses (bytes in, first paint out, then move
 * between units, then destroy). The page only times things — sampling,
 * medians, and memory readings belong to the harness, which sees the whole
 * browser, not one page.
 */

const AARED: Record<Format, OpenAdapter | undefined> = {
  pdf: undefined,
  pptx: openAaredPptx,
  xlsx: openAaredXlsx,
  docx: openAaredDocx,
};

function adapterFor(format: Format, renderer: string): OpenAdapter {
  const adapter = renderer === "aared" ? AARED[format] : BASELINES[format][renderer];
  if (adapter === undefined) throw new Error(`no ${renderer} adapter for ${format}`);
  return adapter;
}

const host = document.getElementById("host") as HTMLElement;
const fixtures = new Map<string, Uint8Array>();
let current: OpenedDoc | null = null;
/** What a calibration control is holding (see `calibrate`). */
let held: { buffer: ArrayBuffer | null; worker: Worker | null } = { buffer: null, worker: null };
const CALIBRATION_BYTES = 10 * 1024 * 1024;

async function bytesOf(fixture: string): Promise<Uint8Array> {
  let bytes = fixtures.get(fixture);
  if (bytes === undefined) {
    // The harness serves fixtures under /f/, both the pinned real ones and
    // the large generated ones.
    const response = await fetch(`/f/${fixture}`);
    if (!response.ok) throw new Error(`fixture ${fixture}: HTTP ${response.status}`);
    bytes = new Uint8Array(await response.arrayBuffer());
    fixtures.set(fixture, bytes);
  }
  return bytes;
}

/** Wait for the browser to commit what was just painted. */
const painted = (): Promise<void> =>
  new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));

/** How many units there are to move between: slides or sheets; a document has none. */
function unitCount(doc: OpenedDoc): number {
  if (doc.format === "pptx") return doc.pageCount;
  if (doc.format === "xlsx") return doc.sheets.length;
  return 0;
}

const api = {
  /** Fetch a fixture into memory so its transfer is never part of a timing. */
  async preload(fixture: string): Promise<number> {
    return (await bytesOf(fixture)).byteLength;
  },

  /**
   * Do the bench's own side reads for a fixture (a workbook's sheet RTL
   * flags, a document's widest section) before anything is measured, so no
   * renderer's time or memory carries harness work. Not called by the size
   * pass, which only counts what the renderer downloads.
   */
  async prime(fixture: string): Promise<void> {
    const bytes = await bytesOf(fixture);
    if (fixture.endsWith(".xlsx")) await sheetRtlFlags(bytes);
    if (fixture.endsWith(".docx")) await maxSectionWidthPx(bytes);
  },

  /** Open one document; the time is the adapter's own, entry to first paint, then the frame commit. */
  async open(format: Format, renderer: string, fixture: string): Promise<{ ms: number; units: number }> {
    const bytes = await bytesOf(fixture);
    const adapter = adapterFor(format, renderer);
    const t0 = performance.now();
    const doc = await adapter(bytes, host, new AbortController().signal);
    await painted();
    const ms = performance.now() - t0;
    current = doc;
    return { ms, units: unitCount(doc) };
  },

  /** Move to every slide or sheet once, timing each move to the committed frame. */
  async navigateAll(): Promise<number[]> {
    const doc = current;
    if (doc === null) throw new Error("navigateAll() before open()");
    const times: number[] = [];
    const count = unitCount(doc);
    for (let index = 0; index < count; index += 1) {
      const t0 = performance.now();
      if (doc.format === "pptx") await doc.goToSlide(index);
      else if (doc.format === "xlsx") await doc.renderSheet(index);
      await painted();
      times.push(performance.now() - t0);
    }
    return times;
  },

  /**
   * Look at the whole file once, the way a reader would: every slide or
   * sheet, or a document scrolled through a stage at a time to its end. For
   * the memory reading taken after a full read, not for timing.
   */
  async visitAll(): Promise<number> {
    const doc = current;
    if (doc === null) throw new Error("visitAll() before open()");
    if (doc.format === "pptx" || doc.format === "xlsx") return (await api.navigateAll()).length;
    const scroller = [host, ...host.querySelectorAll<HTMLElement>("*")].find(
      (element) => element.scrollHeight > element.clientHeight + 1,
    );
    if (scroller === undefined) return 0;
    let steps = 0;
    for (let top = 0; top < scroller.scrollHeight; top += scroller.clientHeight) {
      scroller.scrollTop = top;
      await painted();
      steps += 1;
    }
    return steps;
  },

  /** Elements the open painted into the stage: what each renderer built, for the record. */
  elements(): number {
    return host.querySelectorAll("*").length;
  },

  destroy(): void {
    current?.destroy();
    current = null;
    host.replaceChildren();
  },

  /**
   * Calibration controls for the harness's memory reading: hold exactly
   * 10 MiB on the page, or in a Worker, and check the reading sees it once.
   */
  async calibrate(where: "page" | "worker"): Promise<number> {
    if (where === "page") {
      held.buffer = new ArrayBuffer(CALIBRATION_BYTES);
      new Uint8Array(held.buffer).fill(1);
      return CALIBRATION_BYTES;
    }
    const source = `const kept = new ArrayBuffer(${CALIBRATION_BYTES}); new Uint8Array(kept).fill(1); self.kept = kept; postMessage("ready");`;
    const url = URL.createObjectURL(new Blob([source], { type: "text/javascript" }));
    const worker = new Worker(url);
    await new Promise<void>((resolve) => {
      worker.onmessage = () => resolve();
    });
    URL.revokeObjectURL(url);
    held.worker = worker;
    return CALIBRATION_BYTES;
  },

  release(): void {
    held.worker?.terminate();
    held = { buffer: null, worker: null };
  },

  /**
   * aared's `search()` on one document, in a fresh session: the first search
   * reads and indexes every unit (cold), repeats scan what it kept (warm), and
   * a common word shows the cost of many matches. The baselines have no search.
   */
  async search(format: Exclude<Format, "pdf">, fixture: string): Promise<SearchTimings> {
    const bytes = await bytesOf(fixture);
    const MISS = "zzqq-no-such-text";
    const t0 = performance.now();
    const doc =
      format === "pptx"
        ? await (await import("aared/pptx")).open(bytes)
        : format === "xlsx"
          ? await (await import("aared/xlsx")).open(bytes)
          : await (await import("aared/docx")).open(bytes);
    const openMs = performance.now() - t0;
    try {
      // A workbook is searched one sheet at a time; the whole file is what is timed here, so every sheet is.
      const searchAll = async (query: string): Promise<number> => {
        if (!("sheets" in doc)) return (await doc.search(query)).length;
        let hits = 0;
        for (const sheet of doc.sheets.list()) hits += (await sheet.search(query)).length;
        return hits;
      };
      const time = async (query: string): Promise<{ ms: number; hits: number }> => {
        const start = performance.now();
        const hits = await searchAll(query);
        return { ms: performance.now() - start, hits };
      };
      const cold = await time(MISS);
      const warmRuns: number[] = [];
      for (let run = 0; run < 5; run += 1) warmRuns.push((await time(MISS)).ms);
      warmRuns.sort((a, b) => a - b);
      const hit = await time("the");
      return { openMs, coldMs: cold.ms, warmMs: warmRuns[2]!, hitMs: hit.ms, hits: hit.hits };
    } finally {
      doc.close();
    }
  },
};

export interface SearchTimings {
  openMs: number;
  coldMs: number;
  warmMs: number;
  hitMs: number;
  hits: number;
}

export type MeasureApi = typeof api;
(window as unknown as { __m: MeasureApi }).__m = api;
