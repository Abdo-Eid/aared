import { afterEach, expect, spyOn, test } from "bun:test";
import { open as openPptx } from "../src/pptx/index.js";
import { open as openXlsx } from "../src/xlsx/index.js";
import { WorkerSession } from "../src/core/session.js";

const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;

const spies: Array<{ mockRestore(): void }> = [];
afterEach(() => {
  for (const spy of spies.splice(0)) spy.mockRestore();
});

test("PPTX search reads each slide once, then answers from its text index", async () => {
  const deck = await openPptx(await Bun.file(PPTX_FIXTURE).arrayBuffer());
  try {
    const unit = spyOn(WorkerSession.prototype, "unit");
    spies.push(unit);
    const first = await deck.search("الإنتاج");
    expect(first.length).toBeGreaterThan(0);
    expect(unit).toHaveBeenCalledTimes(deck.slides.list().length);
    // Search never paints, so its reads skip the font-resolver round trip.
    for (const [, options] of unit.mock.calls) expect(options).toMatchObject({ fonts: false });

    const second = await deck.search("الإنتاج");
    expect(second).toEqual(first);
    expect(unit).toHaveBeenCalledTimes(deck.slides.list().length);

    // A different query reuses the same index too.
    await deck.search("no-such-text-anywhere");
    expect(unit).toHaveBeenCalledTimes(deck.slides.list().length);
  } finally {
    deck.close();
  }
});

test("XLSX search reads each sheet once, then answers from its cell index", async () => {
  const workbook = await openXlsx(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const unit = spyOn(WorkerSession.prototype, "unit");
    spies.push(unit);
    const first = await workbook.search("budget");
    expect(first.length).toBeGreaterThan(0);
    expect(unit).toHaveBeenCalledTimes(workbook.sheets.list().length);
    for (const [, options] of unit.mock.calls) expect(options).toMatchObject({ fonts: false });

    expect(await workbook.search("budget")).toEqual(first);
    expect(unit).toHaveBeenCalledTimes(workbook.sheets.list().length);
  } finally {
    workbook.close();
  }
});

test("an aborted search leaves no partial index behind", async () => {
  const deck = await openPptx(await Bun.file(PPTX_FIXTURE).arrayBuffer());
  try {
    const controller = new AbortController();
    controller.abort();
    const error = await deck.search("الإنتاج", { signal: controller.signal }).catch((e: unknown) => e);
    expect((error as Error).name).toBe("AbortError");
    // The next, un-aborted search still reads and finds everything.
    expect((await deck.search("الإنتاج")).length).toBeGreaterThan(0);
  } finally {
    deck.close();
  }
});
