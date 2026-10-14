import { expect, test } from "bun:test";
import type { WorkerRequest, WorkerResponse } from "../src/core/rpc.js";

const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

interface WorkerScopeForTest {
  onmessage: ((event: MessageEvent<WorkerRequest>) => void) | null;
  postMessage(message: WorkerResponse, transfer?: Transferable[]): void;
}

function responseFor(
  responses: WorkerResponse[],
  onResponse: (callback: () => void) => void,
  requestId: number,
  sessionId: number,
): Promise<Extract<WorkerResponse, { type: "response" }>> {
  const existing = responses.find(
    (response): response is Extract<WorkerResponse, { type: "response" }> =>
      response.type === "response" && response.requestId === requestId && response.sessionId === sessionId,
  );
  if (existing !== undefined) return Promise.resolve(existing);
  return new Promise((resolve) => {
    onResponse(() => {
      const response = responses.find(
        (candidate): candidate is Extract<WorkerResponse, { type: "response" }> =>
          candidate.type === "response" && candidate.requestId === requestId && candidate.sessionId === sessionId,
      );
      if (response !== undefined) resolve(response);
    });
  });
}

test("worker cancellation stops after a controlled inflate checkpoint without answering", async () => {
  const scope = globalThis as unknown as WorkerScopeForTest;
  const originalOnmessage = scope.onmessage;
  const originalPostMessage = scope.postMessage;
  const originalDecompressionStream = globalThis.DecompressionStream;
  const responses: WorkerResponse[] = [];
  const responseListeners = new Set<() => void>();
  let releaseInflate!: () => void;
  const inflateReleased = new Promise<void>((resolve) => {
    releaseInflate = resolve;
  });
  let reachInflate!: () => void;
  const inflateReached = new Promise<void>((resolve) => {
    reachInflate = resolve;
  });

  scope.postMessage = (message) => {
    responses.push(message);
    for (const listener of responseListeners) listener();
  };
  const onResponse = (listener: () => void): void => {
    responseListeners.add(listener);
  };

  try {
    await import("../src/core/worker.js");
    const dispatch = (request: WorkerRequest): void => {
      scope.onmessage?.({ data: request } as MessageEvent<WorkerRequest>);
    };
    const sessionId = 41;
    dispatch({ type: "open", sessionId, requestId: 1, bytes: await Bun.file(FIXTURE_PATH).arrayBuffer() });
    expect((await responseFor(responses, onResponse, 1, sessionId)).ok).toBe(true);

    class CheckpointDecompressionStream {
      constructor(_format: string) {
        return new TransformStream<Uint8Array, Uint8Array>({
          async transform(chunk, controller) {
            reachInflate();
            await inflateReleased;
            controller.enqueue(chunk);
          },
        });
      }
    }
    Object.defineProperty(globalThis, "DecompressionStream", { configurable: true, value: CheckpointDecompressionStream });

    dispatch({ type: "unit", sessionId, requestId: 2, index: 0 });
    await inflateReached;
    dispatch({ type: "cancel", sessionId, requestId: 2 });
    releaseInflate();

    Object.defineProperty(globalThis, "DecompressionStream", { configurable: true, value: originalDecompressionStream });
    dispatch({ type: "unit", sessionId, requestId: 3, index: 0 });
    expect((await responseFor(responses, onResponse, 3, sessionId)).ok).toBe(true);
    // Let the released read finish: a cancelled request never answers, so
    // without the cancel, request 2's response would be here by now.
    await new Promise((resolve) => setTimeout(resolve, 50));
    expect(responses.some((response) => response.type === "response" && response.requestId === 2)).toBe(false);
    dispatch({ type: "close", sessionId });

    const cancelledOpenSession = 42;
    dispatch({ type: "open", sessionId: cancelledOpenSession, requestId: 4, bytes: await Bun.file(FIXTURE_PATH).arrayBuffer() });
    expect((await responseFor(responses, onResponse, 4, cancelledOpenSession)).ok).toBe(true);
    // A caller can abort after the worker posts success but before it receives
    // that response. The late cancel must still remove this session.
    dispatch({ type: "cancel", sessionId: cancelledOpenSession, requestId: 4 });
    dispatch({ type: "unit", sessionId: cancelledOpenSession, requestId: 5, index: 0 });
    const closedSession = await responseFor(responses, onResponse, 5, cancelledOpenSession);
    expect(closedSession.ok).toBe(false);
    if (closedSession.ok) throw new Error("expected closed session failure");
    expect(closedSession.error.message).toContain("document session is closed");
    // A usage error, never a cancellation that `isAbort()` would swallow.
    expect(closedSession.error.code).toBe("closed");

    // Every failure leaves the Worker typed: an aared bug is `internal` (never
    // "the file may be damaged"), with the original error for the bug report;
    // a corrupt deflate stream is the file's fault.
    const typedSession = 43;
    dispatch({ type: "open", sessionId: typedSession, requestId: 6, bytes: await Bun.file(FIXTURE_PATH).arrayBuffer() });
    expect((await responseFor(responses, onResponse, 6, typedSession)).ok).toBe(true);
    // An unexpected throw on aared's side (here: the decompressor cannot even be built).
    class ExplodingDecompressionStream {
      constructor(_format: string) {
        throw new RangeError("decompressor exploded");
      }
    }
    Object.defineProperty(globalThis, "DecompressionStream", { configurable: true, value: ExplodingDecompressionStream });
    dispatch({ type: "unit", sessionId: typedSession, requestId: 7, index: 1 });
    const bug = await responseFor(responses, onResponse, 7, typedSession);
    Object.defineProperty(globalThis, "DecompressionStream", { configurable: true, value: originalDecompressionStream });
    if (bug.ok) throw new Error("expected the broken decompressor to fail the read");
    expect(bug.error.code).toBe("internal");
    expect(bug.error.message).toBe("@aared/viewer/core: the viewer failed while reading this file (RangeError: decompressor exploded)");
    expect(bug.error.message).not.toContain("re-saving");
    expect(bug.error.cause).toMatchObject({ name: "RangeError", message: "decompressor exploded" });
    expect(bug.error.cause?.stack).toContain("decompressor exploded");

    class CorruptDecompressionStream {
      constructor(_format: string) {
        return new TransformStream<Uint8Array, Uint8Array>({
          transform() {
            throw new TypeError("invalid block type");
          },
        });
      }
    }
    Object.defineProperty(globalThis, "DecompressionStream", { configurable: true, value: CorruptDecompressionStream });
    dispatch({ type: "unit", sessionId: typedSession, requestId: 8, index: 2 });
    const corrupt = await responseFor(responses, onResponse, 8, typedSession);
    Object.defineProperty(globalThis, "DecompressionStream", { configurable: true, value: originalDecompressionStream });
    if (corrupt.ok) throw new Error("expected the corrupt stream to fail the read");
    expect(corrupt.error.code).toBe("invalid-document");
    expect(corrupt.error.message).toContain("entry data is not valid deflate (invalid block type)");
    dispatch({ type: "close", sessionId: typedSession });
  } finally {
    Object.defineProperty(globalThis, "DecompressionStream", { configurable: true, value: originalDecompressionStream });
    scope.onmessage = originalOnmessage;
    scope.postMessage = originalPostMessage;
  }
});
