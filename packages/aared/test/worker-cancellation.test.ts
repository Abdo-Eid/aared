import { expect, test } from "bun:test";
import type { WorkerRequest, WorkerResponse } from "../src/core/rpc.js";

const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/arabic-pptx-rtl-coverage.pptx`;

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

test("worker cancellation stops after a controlled inflate checkpoint before XML decoding", async () => {
  const scope = globalThis as unknown as WorkerScopeForTest;
  const originalOnmessage = scope.onmessage;
  const originalPostMessage = scope.postMessage;
  const originalDecompressionStream = globalThis.DecompressionStream;
  const originalTextDecoder = globalThis.TextDecoder;
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
    let decodeCount = 0;
    class CountingTextDecoder extends originalTextDecoder {
      override decode(...args: Parameters<TextDecoder["decode"]>): string {
        decodeCount++;
        return super.decode(...args);
      }
    }
    Object.defineProperty(globalThis, "DecompressionStream", { configurable: true, value: CheckpointDecompressionStream });
    Object.defineProperty(globalThis, "TextDecoder", { configurable: true, value: CountingTextDecoder });

    dispatch({ type: "unit", sessionId, requestId: 2, index: 0 });
    await inflateReached;
    dispatch({ type: "cancel", sessionId, requestId: 2 });
    releaseInflate();

    Object.defineProperty(globalThis, "DecompressionStream", { configurable: true, value: originalDecompressionStream });
    Object.defineProperty(globalThis, "TextDecoder", { configurable: true, value: originalTextDecoder });
    dispatch({ type: "unit", sessionId, requestId: 3, index: 0 });
    expect((await responseFor(responses, onResponse, 3, sessionId)).ok).toBe(true);
    expect(decodeCount).toBe(0);
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
  } finally {
    Object.defineProperty(globalThis, "DecompressionStream", { configurable: true, value: originalDecompressionStream });
    Object.defineProperty(globalThis, "TextDecoder", { configurable: true, value: originalTextDecoder });
    scope.onmessage = originalOnmessage;
    scope.postMessage = originalPostMessage;
  }
});
