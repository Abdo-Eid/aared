import { expect, test } from "bun:test";
import { AaredError, isAaredError, isAbort } from "../src/core/index.js";
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
