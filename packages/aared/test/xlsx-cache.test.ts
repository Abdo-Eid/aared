import { expect, test } from "bun:test";
import type { SheetModel, SheetRangeModel } from "../src/core/model.js";
import type { WorkerCall, WorkerResponse, WorkerResult } from "../src/core/rpc.js";
import { rejection, within } from "./harness.js";

const FIXTURE = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;

async function withCacheWorker(run: (client: {
  call: (request: WorkerCall, sessionId?: number) => Promise<WorkerResult>;
  post: (message: unknown) => void;
  parses: (sheet: number) => number;
  paused: Promise<void>;
  cancelledReplies: WorkerResponse[];
}) => Promise<void>): Promise<void> {
  const worker = new Worker(new URL("./xlsx-cache.worker.ts", import.meta.url));
  const counts = new Map<string, number>();
  const pending = new Map<number, { resolve: (value: WorkerResult) => void; reject: (error: Error) => void }>();
  let ready!: () => void;
  const booted = new Promise<void>((resolve) => { ready = resolve; });
  let paused!: () => void;
  const pause = new Promise<void>((resolve) => { paused = resolve; });
  const cancelledReplies: WorkerResponse[] = [];
  let next = 0;
  worker.onmessage = (event: MessageEvent<WorkerResponse | { type: "test-sheet-parse"; part: string } | { type: "test-paused" }>) => {
    const message = event.data;
    if (message.type === "ready") ready();
    else if (message.type === "test-sheet-parse") counts.set(message.part, (counts.get(message.part) ?? 0) + 1);
    else if (message.type === "test-paused") paused();
    else {
      if (message.requestId === -1) cancelledReplies.push(message);
      const waiting = pending.get(message.requestId);
      if (waiting === undefined) return;
      pending.delete(message.requestId);
      if (message.ok) waiting.resolve(message.result);
      else waiting.reject(Object.assign(new Error(message.error.message), { code: message.error.code }));
    }
  };
  worker.onerror = (event) => {
    for (const entry of pending.values()) entry.reject(new Error(event.message));
    pending.clear();
  };
  const call = (request: WorkerCall, sessionId = 1): Promise<WorkerResult> => {
    const requestId = ++next;
    const response = new Promise<WorkerResult>((resolve, reject) => { pending.set(requestId, { resolve, reject }); });
    worker.postMessage({ ...request, sessionId, requestId });
    return within(`Worker ${request.type} response`, response, 5000);
  };
  try {
    await within("cache Worker boot", booted, 5000);
    await call({ type: "open", bytes: await Bun.file(FIXTURE).arrayBuffer() });
    await run({ call, post: (message) => worker.postMessage(message), parses: (sheet) => counts.get(`xl/worksheets/sheet${sheet + 1}.xml`) ?? 0, paused: pause, cancelledReplies });
  } finally {
    worker.terminate();
  }
}

test("XLSX Worker range reads slice one retained parse and reuse it for search", async () => {
  await withCacheWorker(async ({ call, parses }) => {
    const range = { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } };
    const first = await call({ type: "unit", index: 1, range }) as SheetRangeModel;
    expect(first.range).toEqual(range);
    expect(first.rows.length).toBeGreaterThan(0);
    expect(parses(1)).toBe(1);
    expect(await call({ type: "unit", index: 1, range })).toEqual(first);
    const other = { start: { row: 3, col: 1 }, end: { row: 5, col: 3 } };
    const window = await call({ type: "unit", index: 1, range: other }) as SheetRangeModel;
    expect(window.range).toEqual(other);
    expect(window.rows.length).toBeGreaterThan(0);
    for (const row of window.rows) {
      expect(row.index >= 3 && row.index <= 5).toBe(true);
      expect(row.cells.every((cell) => cell.col >= 1 && cell.col <= 3)).toBe(true);
    }
    const words = await call({ type: "text", index: 1 });
    expect(words).toHaveProperty("text");
    expect(parses(1)).toBe(1);
  });
});

test("XLSX Worker searches neither fill the sheet cache nor evict the displayed sheets", async () => {
  await withCacheWorker(async ({ call, parses }) => {
    await call({ type: "unit", index: 0 });
    await call({ type: "unit", index: 1 });
    await call({ type: "text", index: 2 });
    await call({ type: "text", index: 2 });
    expect(parses(2)).toBe(2);
    await call({ type: "unit", index: 0 });
    await call({ type: "unit", index: 1 });
    expect([parses(0), parses(1)]).toEqual([1, 1]);
  });
});

test("XLSX Worker search hits leave recency alone while rendered reads refresh it", async () => {
  await withCacheWorker(async ({ call, parses }) => {
    await call({ type: "unit", index: 0 });
    await call({ type: "unit", index: 1 });
    await call({ type: "text", index: 0 });
    await call({ type: "unit", index: 2 });
    await call({ type: "unit", index: 1 });
    expect(parses(1)).toBe(1);
    await call({ type: "unit", index: 0 });
    expect(parses(0)).toBe(2);
    // A rendered hit refreshes recency: 0 survives the next tab switch.
    await call({ type: "unit", index: 2 });
    await call({ type: "unit", index: 0 });
    expect(parses(0)).toBe(2);
  });
});

test("closing an XLSX session drops its retained sheets on the same real Worker", async () => {
  await withCacheWorker(async ({ call, post, parses }) => {
    const first = await call({ type: "unit", index: 0 });
    post({ type: "close", sessionId: 1 });
    await call({ type: "open", bytes: await Bun.file(FIXTURE).arrayBuffer() }, 2);
    expect(await rejection(call({ type: "unit", index: 0 }))).toMatchObject({ code: "closed" });
    expect(await call({ type: "unit", index: 0 }, 2)).toEqual(first);
    expect(parses(0)).toBe(2);
  });
});

test("a cancelled XLSX parse is not retained as a rejected or partial sheet", async () => {
  await withCacheWorker(async ({ call, post, parses, paused, cancelledReplies }) => {
    post({ type: "test-pause" });
    post({ type: "unit", sessionId: 1, requestId: -1, index: 0 });
    await within("paused sheet parse", paused, 5000);
    post({ type: "cancel", sessionId: 1, requestId: -1 });
    post({ type: "test-resume" });
    const sheet = await call({ type: "unit", index: 0 }) as SheetModel;
    expect(sheet.name).toBe("الميزانية");
    expect(sheet.rows.length).toBeGreaterThan(0);
    expect(parses(0)).toBe(2);
    expect(cancelledReplies).toEqual([]);
    expect(await call({ type: "unit", index: 0 })).toEqual(sheet);
    expect(parses(0)).toBe(2);
  });
});
