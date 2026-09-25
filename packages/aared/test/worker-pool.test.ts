import { afterEach, beforeEach, expect, test } from "bun:test";
import { configureWorkerPool, drainWorkerPool } from "../src/core/session.js";
import { detect } from "../src/preview/index.js";
import { open } from "../src/pptx/index.js";
import { rejection } from "./harness.js";

/**
 * The Worker pool, against a fake Worker: a closed document's clean Worker
 * serves the next open (and a detect serves the open after it), an idle one
 * stops on its timer, and anything cancelled, closed mid-read or built by a
 * caller's factory is never reused.
 */

type Message = { type: string; sessionId?: number; requestId?: number };

const PPTX_INFO = { format: "pptx", count: 0, titles: [], theme: { colors: {}, fonts: { major: "", minor: "" } } };

class PoolWorker {
  static instances: PoolWorker[] = [];
  messages: Message[] = [];
  terminated = false;
  private handler: ((event: MessageEvent<unknown>) => void) | null = null;

  constructor(readonly url: unknown) {
    PoolWorker.instances.push(this);
  }

  set onmessage(handler: ((event: MessageEvent<unknown>) => void) | null) {
    this.handler = handler;
    queueMicrotask(() => this.handler?.({ data: { type: "ready" } } as MessageEvent<unknown>));
  }

  set onerror(_handler: unknown) {}

  postMessage(message: Message): void {
    this.messages.push(message);
    // Media reads never answer: they stay in flight until the test acts.
    if (message.type === "open") this.respond(message, PPTX_INFO);
    if (message.type === "detect") this.respond(message, "pptx");
  }

  terminate(): void {
    this.terminated = true;
  }

  private respond(message: Message, result: unknown): void {
    queueMicrotask(() =>
      this.handler?.({ data: { type: "response", sessionId: message.sessionId, requestId: message.requestId, ok: true, result } } as MessageEvent<unknown>),
    );
  }
}

let originalWorker: PropertyDescriptor | undefined;
let timers: Array<() => void> = [];
let restoreClock: () => void = () => {};

beforeEach(() => {
  PoolWorker.instances = [];
  timers = [];
  originalWorker = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  Object.defineProperty(globalThis, "Worker", { configurable: true, value: PoolWorker });
  restoreClock = configureWorkerPool({
    set: (run) => {
      timers.push(run);
      return run;
    },
    clear: (handle) => {
      timers = timers.filter((run) => run !== handle);
    },
  });
});

afterEach(() => {
  drainWorkerPool();
  restoreClock();
  if (originalWorker === undefined) delete (globalThis as { Worker?: unknown }).Worker;
  else Object.defineProperty(globalThis, "Worker", originalWorker);
});

test("a closed document's Worker serves the next open", async () => {
  const first = await open(new ArrayBuffer(8));
  first.close();
  const second = await open(new ArrayBuffer(8));
  expect(PoolWorker.instances).toHaveLength(1);
  const worker = PoolWorker.instances[0]!;
  expect(worker.terminated).toBe(false);
  const opens = worker.messages.filter((message) => message.type === "open");
  expect(opens).toHaveLength(2);
  expect(opens[0]!.sessionId).not.toBe(opens[1]!.sessionId);
  expect(worker.messages.some((message) => message.type === "close" && message.sessionId === opens[0]!.sessionId)).toBe(true);
  // Reuse cancelled the idle timer.
  expect(timers).toHaveLength(0);
  second.close();
});

test("an idle Worker stops on its timer, and the next open starts a new one", async () => {
  (await open(new ArrayBuffer(8))).close();
  expect(timers).toHaveLength(1);
  timers[0]!();
  expect(PoolWorker.instances[0]!.terminated).toBe(true);
  (await open(new ArrayBuffer(8))).close();
  expect(PoolWorker.instances).toHaveLength(2);
});

test("two documents open at once get two Workers; the newer idle one is kept", async () => {
  const [a, b] = await Promise.all([open(new ArrayBuffer(8)), open(new ArrayBuffer(8))]);
  expect(PoolWorker.instances).toHaveLength(2);
  a.close();
  b.close();
  expect(PoolWorker.instances[0]!.terminated).toBe(true);
  expect(PoolWorker.instances[1]!.terminated).toBe(false);
  (await open(new ArrayBuffer(8))).close();
  expect(PoolWorker.instances).toHaveLength(2);
});

test("detection and the open after it share one Worker", async () => {
  expect(await detect(new ArrayBuffer(8))).toBe("pptx");
  (await open(new ArrayBuffer(8))).close();
  expect(PoolWorker.instances).toHaveLength(1);
  expect(PoolWorker.instances[0]!.messages.map((message) => message.type)).toEqual(["detect", "close", "open", "close"]);
});

test("a Worker URL is its own pool; a factory's Workers are never reused", async () => {
  (await open(new ArrayBuffer(8), { worker: "blob:aared-test" })).close();
  (await open(new ArrayBuffer(8), { worker: new URL("blob:aared-test") })).close();
  (await open(new ArrayBuffer(8))).close();
  expect(PoolWorker.instances.map((worker) => worker.url)).toHaveLength(2);

  let built = 0;
  const factory = (): Worker => {
    built += 1;
    return new PoolWorker("factory") as unknown as Worker;
  };
  (await open(new ArrayBuffer(8), { worker: factory })).close();
  (await open(new ArrayBuffer(8), { worker: factory })).close();
  expect(built).toBe(2);
  expect(PoolWorker.instances.filter((worker) => worker.url === "factory").every((worker) => worker.terminated)).toBe(true);
});

test("a Worker with a cancelled or abandoned read is stopped, never reused", async () => {
  const cancelled = await open(new ArrayBuffer(8));
  const controller = new AbortController();
  const read = cancelled.media.get("ppt/media/image1.png", { signal: controller.signal });
  controller.abort();
  await rejection(read);
  cancelled.close();
  expect(PoolWorker.instances[0]!.terminated).toBe(true);

  const abandoned = await open(new ArrayBuffer(8));
  const inFlight = abandoned.media.get("ppt/media/image1.png");
  abandoned.close();
  expect(await rejection(inFlight)).toMatchObject({ code: "closed" });
  expect(PoolWorker.instances[1]!.terminated).toBe(true);

  (await open(new ArrayBuffer(8))).close();
  expect(PoolWorker.instances).toHaveLength(3);
});

test("a parked Worker never serves a different Worker constructor", async () => {
  (await open(new ArrayBuffer(8))).close();
  class OtherWorker extends PoolWorker {}
  Object.defineProperty(globalThis, "Worker", { configurable: true, value: OtherWorker });
  (await open(new ArrayBuffer(8))).close();
  expect(PoolWorker.instances).toHaveLength(2);
  expect(PoolWorker.instances[0]!.terminated).toBe(true);
  expect(PoolWorker.instances[1]).toBeInstanceOf(OtherWorker);
});
