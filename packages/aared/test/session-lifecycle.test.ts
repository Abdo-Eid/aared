import { expect, test } from "bun:test";
import { AaredError, isAaredError, isAbort } from "../src/core/index.js";
import { drainWorkerPool } from "../src/core/session.js";
import { open } from "../src/pptx/index.js";
import { rejection } from "./harness.js";

const ARABIC_FIXTURE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

async function rejectionMessage(promise: Promise<unknown>): Promise<string> {
  try {
    await promise;
  } catch (error) {
    return error instanceof Error ? error.message : String(error);
  }
  throw new Error("expected the request to reject");
}

test("keeps a caller-owned ArrayBuffer usable after transfer", async () => {
  const bytes = await Bun.file(ARABIC_FIXTURE).arrayBuffer();
  const length = bytes.byteLength;
  const deck = await open(bytes);
  try {
    expect(bytes.byteLength).toBe(length);
    expect(new Uint8Array(bytes, 0, 2)).toEqual(new Uint8Array([0x50, 0x4b]));
  } finally {
    deck.close();
  }
});

test("a typed array opens over its own range, and the caller's bytes stay usable", async () => {
  const file = new Uint8Array(await Bun.file(ARABIC_FIXTURE).arrayBuffer());
  // The file sits in the middle of a bigger buffer, as in a Node Buffer pool:
  // only the view's own range may be read.
  const pool = new Uint8Array(file.byteLength + 64).fill(0xff);
  pool.set(file, 32);
  const view = pool.subarray(32, 32 + file.byteLength);
  const deck = await open(view);
  try {
    expect(deck.info.count).toBeGreaterThan(0);
    expect(view[0]).toBe(0x50);
    expect(pool[0]).toBe(0xff);
  } finally {
    deck.close();
  }
});

test("a source open() cannot read is a typed invalid-argument, not the platform's raw TypeError", async () => {
  const bytes = new Uint8Array(await Bun.file(ARABIC_FIXTURE).arrayBuffer());
  for (const wrong of [null, 42, { name: "deck.pptx" }]) {
    const error = await rejection(open(wrong as unknown as ArrayBuffer));
    expect(isAaredError(error) && error.code).toBe("invalid-argument");
    expect(error.message).toContain("aared/pptx: a source is a File, Blob, ArrayBuffer, typed array or URL string; received");
  }
  // A Blob is read by what is used from it, so one from another frame still opens.
  const foreign = { size: bytes.byteLength, arrayBuffer: async () => bytes.slice().buffer };
  const deck = await open(foreign as unknown as Blob);
  try {
    expect(deck.info.count).toBeGreaterThan(0);
  } finally {
    deck.close();
  }
});

test("aborts URL source loading through fetch", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "fetch");
  let suppliedSignal: AbortSignal | undefined;
  let fetchStarted!: () => void;
  const started = new Promise<void>((resolve) => {
    fetchStarted = resolve;
  });
  Object.defineProperty(globalThis, "fetch", {
    configurable: true,
    value: (_source: string, options?: RequestInit) => {
      suppliedSignal = options?.signal ?? undefined;
      fetchStarted();
      return new Promise<Response>((_resolve, reject) => {
        options?.signal?.addEventListener("abort", () => reject(new Error("fetch aborted")), { once: true });
      });
    },
  });
  const controller = new AbortController();
  try {
    const opening = open("https://example.test/deck.pptx", { signal: controller.signal });
    await started;
    controller.abort();
    const error = await rejection(opening);
    expect(suppliedSignal).toBe(controller.signal);
    expect(error).toMatchObject({ name: "AbortError", message: "aared/core: request was cancelled" });
  } finally {
    if (original === undefined) delete (globalThis as { fetch?: typeof fetch }).fetch;
    else Object.defineProperty(globalThis, "fetch", original);
  }
});

test("cancels a worker slide read", async () => {
  const deck = await open(await Bun.file(ARABIC_FIXTURE).arrayBuffer());
  try {
    const controller = new AbortController();
    const read = deck.slides.get(1).read({ signal: controller.signal });
    controller.abort();
    expect(await rejectionMessage(read)).toContain("request was cancelled");
  } finally {
    deck.close();
  }
});

test("cancels open and media requests after they reach the worker", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  const messages: Array<{ type: string; requestId?: number }> = [];
  let openReceived!: () => void;
  const opened = new Promise<void>((resolve) => {
    openReceived = resolve;
  });
  let mediaReceived!: () => void;
  const mediaRequested = new Promise<void>((resolve) => {
    mediaReceived = resolve;
  });
  class ControlledWorker {
    private message: ((event: MessageEvent<unknown>) => void) | null = null;
    private static openCount = 0;

    set onmessage(handler: ((event: MessageEvent<unknown>) => void) | null) {
      this.message = handler;
      queueMicrotask(() => this.message?.({ data: { type: "ready" } } as MessageEvent<unknown>));
    }

    set onerror(_handler: ((event: ErrorEvent) => void) | null) {}

    postMessage(message: { type: string; sessionId?: number; requestId?: number }): void {
      messages.push(message);
      if (message.type === "open") {
        ControlledWorker.openCount++;
        if (ControlledWorker.openCount === 1) openReceived();
        else {
          queueMicrotask(() => this.message?.({
            data: {
              type: "response",
              sessionId: message.sessionId,
              requestId: message.requestId,
              ok: true,
              result: {
                format: "pptx",
                count: 0,
                titles: [],
                theme: { colors: {}, fonts: { major: "", minor: "" } },
              },
            },
          } as MessageEvent<unknown>));
        }
        return;
      }
      if (message.type === "media") mediaReceived();
    }

    terminate(): void {}
  }
  Object.defineProperty(globalThis, "Worker", { configurable: true, value: ControlledWorker });
  try {
    const openController = new AbortController();
    const opening = open(new ArrayBuffer(8), { signal: openController.signal });
    await opened;
    openController.abort();
    expect(await rejection(opening)).toMatchObject({ name: "AbortError" });
    expect(messages.some((message) => message.type === "cancel" && message.requestId === 1)).toBe(true);

    messages.length = 0;
    const documentOpening = open(new ArrayBuffer(8));
    const deck = await documentOpening;
    const mediaController = new AbortController();
    const media = deck.media.get("ppt/media/image1.png", { signal: mediaController.signal });
    await mediaRequested;
    mediaController.abort();
    expect(await rejection(media)).toMatchObject({ name: "AbortError" });
    expect(messages.some((message) => message.type === "cancel" && message.requestId === 2)).toBe(true);
    deck.close();
  } finally {
    if (original === undefined) delete (globalThis as { Worker?: typeof Worker }).Worker;
    else Object.defineProperty(globalThis, "Worker", original);
  }
});

test("resolves each requested font family once on the main side", async () => {
  const requested: string[] = [];
  const deck = await open(await Bun.file(ARABIC_FIXTURE).arrayBuffer(), {
    fontResolver: (family: string) => {
      requested.push(family);
      return undefined;
    },
  });
  try {
    await deck.slides.get(1).read();
    await deck.slides.get(1).read();
    expect(requested).toContain("Arial");
    expect(new Set(requested).size).toBe(requested.length);
  } finally {
    deck.close();
  }
});

test("a failing font resolver never fails a read, and is asked again next time", async () => {
  // A font is cosmetic: the family falls back (and the render's warnings name it).
  const calls: string[] = [];
  let attempt = 0;
  const deck = await open(await Bun.file(ARABIC_FIXTURE).arrayBuffer(), {
    fontResolver: (family: string) => {
      calls.push(family);
      if (family !== "Arial") return undefined;
      attempt++;
      if (attempt === 1) throw new Error("resolver threw");
      return Promise.reject(new Error("resolver rejected"));
    },
  });
  try {
    const first = await deck.slides.get(1).read();
    expect(first.elements.length).toBeGreaterThan(0);
    // A rejection is not cached either: the second read asks for Arial again, and still succeeds.
    await deck.slides.get(1).read();
    expect(calls.filter((family) => family === "Arial")).toHaveLength(2);
    // Families that resolved (here to `undefined`) stay cached: asked once each.
    const others = calls.filter((family) => family !== "Arial");
    expect(new Set(others).size).toBe(others.length);
  } finally {
    deck.close();
  }
});

/** A stub Worker that boots, opens a one-slide deck, and answers `unit` with `unitReply`. */
function stubWorker(unitReply: (message: { sessionId: number; requestId: number }) => unknown): {
  Worker: unknown;
  crash: (message: string) => void;
} {
  let current: { message: ((event: MessageEvent<unknown>) => void) | null; error: ((event: ErrorEvent) => void) | null } | null = null;
  class StubWorker {
    private message: ((event: MessageEvent<unknown>) => void) | null = null;
    private error: ((event: ErrorEvent) => void) | null = null;
    constructor() {
      current = this as unknown as typeof current;
    }
    set onmessage(handler: ((event: MessageEvent<unknown>) => void) | null) {
      this.message = handler;
      queueMicrotask(() => this.message?.({ data: { type: "ready" } } as MessageEvent<unknown>));
    }
    set onerror(handler: ((event: ErrorEvent) => void) | null) {
      this.error = handler;
    }
    postMessage(message: { type: string; sessionId: number; requestId: number }): void {
      const reply = message.type === "open"
        ? { ok: true, result: { format: "pptx", count: 1, titles: ["Slide 1"], sizes: [{ width: 10, height: 7.5 }], theme: { colors: {}, fonts: { major: "", minor: "" } } } }
        : message.type === "unit" ? unitReply(message) : undefined;
      if (reply === undefined) return;
      queueMicrotask(() => this.message?.({ data: { type: "response", sessionId: message.sessionId, requestId: message.requestId, ...(reply as object) } } as MessageEvent<unknown>));
    }
    terminate(): void {}
  }
  return {
    Worker: StubWorker,
    crash: (message) => current?.error?.({ message } as ErrorEvent),
  };
}

async function withWorker(stub: unknown, run: () => Promise<void>): Promise<void> {
  const original = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  Object.defineProperty(globalThis, "Worker", { configurable: true, value: stub });
  try {
    await run();
  } finally {
    if (original === undefined) delete (globalThis as { Worker?: typeof Worker }).Worker;
    else Object.defineProperty(globalThis, "Worker", original);
    drainWorkerPool();
  }
}

test("after the Worker crashes, later calls report the crash, never `closed`", async () => {
  const stub = stubWorker(() => undefined);
  await withWorker(stub.Worker, async () => {
    const deck = await open(new ArrayBuffer(8));
    const pending = deck.slides.get(0).read();
    stub.crash("Uncaught RangeError: out of memory.");
    const inFlight = await rejection(pending);
    expect(inFlight).toMatchObject({ code: "worker-error" });
    expect(inFlight.message).toBe("aared/pptx: worker failed to start or stopped unexpectedly (Uncaught RangeError: out of memory)");
    expect((inFlight as AaredError).cause).toMatchObject({ message: "Uncaught RangeError: out of memory." });
    // The next call on the same document: still the crash, not a use-after-close bug.
    const later = await rejection(deck.slides.get(0).read());
    expect(later).toMatchObject({ code: "worker-error" });
    deck.close();
  });
});

test("an internal Worker failure arrives with its original error as `cause`, in the caller's scope", async () => {
  const stub = stubWorker(() => ({
    ok: false,
    error: {
      code: "internal",
      message: "aared/core: the viewer failed while reading this file (TypeError: x is undefined)",
      cause: { name: "TypeError", message: "x is undefined", stack: "TypeError: x is undefined\n    at parse (pptx.ts:1:1)" },
    },
  }));
  await withWorker(stub.Worker, async () => {
    const deck = await open(new ArrayBuffer(8));
    const error = await rejection(deck.slides.get(0).read());
    expect(error).toBeInstanceOf(AaredError);
    expect(error).toMatchObject({ code: "internal", message: "aared/pptx: the viewer failed while reading this file (TypeError: x is undefined)" });
    const cause = (error as AaredError).cause as Error;
    expect(cause).toBeInstanceOf(Error);
    expect(cause.name).toBe("TypeError");
    expect(cause.stack).toContain("at parse (pptx.ts:1:1)");
    deck.close();
  });
});

test("a download that drops partway is a network-error, not the platform's raw TypeError", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "fetch");
  Object.defineProperty(globalThis, "fetch", {
    configurable: true,
    value: async () => new Response(new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array([0x50, 0x4b]));
        controller.error(new TypeError("network connection was lost"));
      },
    })),
  });
  try {
    const error = await rejection(open("https://example.test/deck.pptx"));
    expect(error).toMatchObject({ code: "network-error" });
    expect(error.message).toBe('aared/pptx: the download of "https://example.test/deck.pptx" stopped partway: network connection was lost');
  } finally {
    if (original === undefined) delete (globalThis as { fetch?: typeof fetch }).fetch;
    else Object.defineProperty(globalThis, "fetch", original);
  }
});

test("closing rejects an in-flight slide read", async () => {
  const deck = await open(await Bun.file(ARABIC_FIXTURE).arrayBuffer());
  const read = deck.slides.get(0).read();
  deck.close();
  expect(await rejectionMessage(read)).toContain("document is closed");
});

test("rejects promptly when posting a worker request throws", async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  class FailingWorker {
    private message: ((event: MessageEvent<{ type: "ready" }>) => void) | null = null;

    set onmessage(handler: ((event: MessageEvent<{ type: "ready" }>) => void) | null) {
      this.message = handler;
      queueMicrotask(() => this.message?.({ data: { type: "ready" } } as MessageEvent<{ type: "ready" }>));
    }

    set onerror(_handler: ((event: ErrorEvent) => void) | null) {}

    postMessage(): void {
      throw new Error("test post failure");
    }

    terminate(): void {}
  }

  Object.defineProperty(globalThis, "Worker", { configurable: true, value: FailingWorker });
  try {
    expect(await rejectionMessage(open(new ArrayBuffer(8)))).toContain("failed to post a Worker request: test post failure");
  } finally {
    if (original === undefined) delete (globalThis as { Worker?: typeof Worker }).Worker;
    else Object.defineProperty(globalThis, "Worker", original);
  }
});

test("reports worker boot failure as worker-error, never as aborted", async () => {
  // A dead worker used to report "aborted", which every abort check
  // (notably usePreview's) swallowed — the UI loaded forever instead of
  // erroring. Boot failure must stay out of the cancellation taxonomy.
  const original = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  class BootFailingWorker {
    private failure: ((event: ErrorEvent) => void) | null = null;

    set onmessage(_handler: unknown) {}

    set onerror(handler: ((event: ErrorEvent) => void) | null) {
      this.failure = handler;
      queueMicrotask(() => this.failure?.({} as ErrorEvent));
    }

    postMessage(): void {}

    terminate(): void {}
  }

  Object.defineProperty(globalThis, "Worker", { configurable: true, value: BootFailingWorker });
  try {
    const error = await rejection(open(await Bun.file(ARABIC_FIXTURE).arrayBuffer()));
    expect(error).toBeInstanceOf(AaredError);
    expect((error as AaredError).code).toBe("worker-error");
    expect(isAbort(error)).toBe(false);
    expect(isAaredError(error) && error.code === "aborted").toBe(false);
  } finally {
    if (original === undefined) delete (globalThis as { Worker?: typeof Worker }).Worker;
    else Object.defineProperty(globalThis, "Worker", original);
  }
});

test("reports a Worker the browser refuses to construct as a typed worker-error", async () => {
  // Scripts on another origin (a CDN) make `new Worker()` throw a
  // SecurityError synchronously, before any `onerror`; it used to escape raw.
  const original = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  const refusal = new Error("Failed to construct 'Worker': Script at 'https://cdn.example/worker.js' cannot be accessed from origin 'https://app.example'.");
  refusal.name = "SecurityError";
  class RefusedWorker {
    constructor() {
      throw refusal;
    }
  }
  Object.defineProperty(globalThis, "Worker", { configurable: true, value: RefusedWorker });
  try {
    const error = await rejection(open(await Bun.file(ARABIC_FIXTURE).arrayBuffer()));
    expect(error).toBeInstanceOf(AaredError);
    expect(error).toMatchObject({ code: "worker-error", cause: refusal });
    const message = (error as AaredError).message;
    expect(message.startsWith("aared/pptx: the browser refused to start the Worker (Failed to construct 'Worker'")).toBe(true);
    expect(message).toContain('"worker" option');
    expect(isAbort(error)).toBe(false);
  } finally {
    if (original === undefined) delete (globalThis as { Worker?: typeof Worker }).Worker;
    else Object.defineProperty(globalThis, "Worker", original);
    drainWorkerPool();
  }
});
