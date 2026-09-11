/**
 * Internal worker session machinery shared by the format entrypoints.
 *
 * This module is deliberately NOT re-exported from `aared/core`: it is
 * plumbing, not public API (the docs-parity test only scans entrypoint
 * files, and these names must stay out of the API reference page). Each format's
 * `open()` builds a `WorkerSession` and wraps it in format-native
 * collections and handles.
 *
 * Headless-safe: no DOM access. The `Worker` global is only touched when
 * a session is actually opened, never at import time (SSR-safe imports).
 */

import type { DocumentInfo, MediaOptions } from "./index.js";
import type { CellRange } from "./model.js";
import type { SerializedError, WorkerCall, WorkerRequest, WorkerResponse, WorkerResult } from "./rpc.js";
import { AaredError, type AaredErrorCode } from "./errors.js";
import { abortError, throwIfAborted } from "./zip.js";

export type { DocumentInfo };
export type { CellRange };

export interface OpenSessionOptions {
  fontResolver?: FontResolver;
  signal?: AbortSignal;
}

export type FontResolver = (
  requestedFamily: string,
) => Promise<FontFace> | undefined;

/** Narrow numeric window for an XLSX `unit` request; other formats reject it. */
export interface UnitRange {
  start: { row: number; col: number };
  end: { row: number; col: number };
}

export interface UnitRequestOptions {
  signal?: AbortSignal;
  /** XLSX-only: numeric window previously validated from an A1 reference. */
  range?: UnitRange;
}

interface PendingRequest {
  sessionId: number;
  resolve: (result: WorkerResult) => void;
  reject: (reason: Error) => void;
  signal?: AbortSignal;
  abort?: () => void;
}

function closedError(scope: string): AaredError {
  return new AaredError("aborted", `${scope}: document is closed`);
}

function deserializeError(serialized: SerializedError): Error {
  // Worker-side cancellation arrives as a named AbortError without a code;
  // restore the typed form so `error.code === "aborted"` holds everywhere.
  if (serialized.code !== undefined) {
    return new AaredError(serialized.code, serialized.message);
  }
  if (serialized.name === "AbortError") {
    return new AaredError("aborted", serialized.message);
  }
  const error = new Error(serialized.message);
  error.name = serialized.name;
  return error;
}

function postError(scope: string, error: unknown): Error {
  const detail = error instanceof Error ? error.message : String(error);
  return new AaredError("aborted", `${scope}: failed to post a Worker request: ${detail}`);
}

class WorkerBridge {
  private readonly worker: Worker;
  private readonly pending = new Map<number, PendingRequest>();
  private nextRequestId = 1;
  private stopped = false;
  private ready: Promise<void>;
  private resolveReady!: () => void;
  private rejectReady!: (reason: Error) => void;

  constructor(private readonly scope: string) {
    if (typeof Worker === "undefined") {
      throw new AaredError("unsupported-feature", `${scope}: open() requires a browser Worker`);
    }
    this.worker = new Worker(new URL("./worker.mjs", import.meta.url), { type: "module" });
    this.ready = new Promise<void>((resolve, reject) => {
      this.resolveReady = resolve;
      this.rejectReady = reject;
    });
    this.worker.onmessage = (event: MessageEvent<WorkerResponse>) => this.receive(event.data);
    this.worker.onerror = () => this.fail(new AaredError("aborted", `${scope}: worker failed to start or stopped unexpectedly`));
  }

  waitUntilReady(): Promise<void> {
    return this.ready;
  }

  request<T extends WorkerResult>(
    sessionId: number,
    message: WorkerCall,
    signal?: AbortSignal,
    transfer?: Transferable[],
  ): Promise<T> {
    if (this.stopped) return Promise.reject(closedError(this.scope));
    if (signal?.aborted === true) return Promise.reject(abortError());
    const requestId = this.nextRequestId++;
    const request = { ...message, sessionId, requestId } as WorkerRequest;
    return new Promise<T>((resolve, reject) => {
      const pending: PendingRequest = {
        sessionId,
        resolve: (result) => resolve(result as T),
        reject,
        ...(signal === undefined ? {} : { signal }),
      };
      if (signal !== undefined) {
        pending.abort = () => {
          const active = this.takePending(requestId);
          if (active === undefined) return;
          try {
            this.worker.postMessage({ type: "cancel", sessionId, requestId } satisfies WorkerRequest);
          } catch {
            // The caller has already been cancelled; a failed notification cannot change that.
          }
          active.reject(abortError());
        };
        signal.addEventListener("abort", pending.abort, { once: true });
      }
      this.pending.set(requestId, pending);
      try {
        this.worker.postMessage(request, transfer ?? []);
      } catch (error) {
        this.takePending(requestId)?.reject(postError(this.scope, error));
      }
    });
  }

  close(sessionId: number): void {
    if (this.stopped) return;
    try {
      this.worker.postMessage({ type: "close", sessionId } satisfies WorkerRequest);
    } catch {
      // Termination below still releases this document when the close message cannot be sent.
    }
    this.stop(closedError(this.scope));
  }

  private receive(message: WorkerResponse): void {
    if (this.stopped) return;
    if (message.type === "ready") {
      this.resolveReady();
      return;
    }
    const candidate = this.pending.get(message.requestId);
    if (candidate === undefined || candidate.sessionId !== message.sessionId) return;
    const pending = this.takePending(message.requestId)!;
    if (message.ok) pending.resolve(message.result);
    else pending.reject(deserializeError(message.error));
  }

  private fail(error: Error): void {
    if (this.stopped) return;
    this.rejectReady(error);
    this.stop(error);
  }

  private stop(error: Error): void {
    if (this.stopped) return;
    this.stopped = true;
    for (const pending of this.pending.values()) {
      if (pending.signal !== undefined && pending.abort !== undefined) {
        pending.signal.removeEventListener("abort", pending.abort);
      }
      pending.reject(error);
    }
    this.pending.clear();
    this.worker.terminate();
  }

  private takePending(requestId: number): PendingRequest | undefined {
    const pending = this.pending.get(requestId);
    if (pending === undefined) return undefined;
    this.pending.delete(requestId);
    if (pending.signal !== undefined && pending.abort !== undefined) {
      pending.signal.removeEventListener("abort", pending.abort);
    }
    return pending;
  }
}

function requestedFontFamilies(model: unknown): string[] {
  const families = new Set<string>();
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    if (value === null || typeof value !== "object") return;
    for (const [key, child] of Object.entries(value)) {
      if (key === "fontFamily" && typeof child === "string" && child !== "") families.add(child);
      else visit(child);
    }
  };
  visit(model);
  return [...families];
}

let nextSessionId = 1;

/**
 * Internal byte normalization shared by `WorkerSession.open`,
 * `WorkerSession.detectFormat`, and `aared/preview` (which normalizes once
 * so URL sources are fetched exactly once across detect + open). Not public
 * API: importable from `./session.js` inside the package only, never
 * re-exported from an entrypoint, never named in the API reference page.
 */
export async function sourceBytes(
  scope: string,
  source: File | Blob | ArrayBuffer | string,
  signal?: AbortSignal,
): Promise<ArrayBuffer> {
  throwIfAborted(signal);
  if (typeof source === "string") {
    let response: Response;
    try {
      response = await fetch(source, signal === undefined ? {} : { signal });
    } catch (error) {
      throwIfAborted(signal);
      if (error instanceof Error && error.name === "AbortError") throw error;
      const detail = error instanceof Error ? error.message : String(error);
      throw new AaredError("network-error", `${scope}: failed to fetch "${source}": ${detail}`);
    }
    throwIfAborted(signal);
    if (!response.ok) {
      throw new AaredError("network-error", `${scope}: failed to fetch "${source}": HTTP ${response.status}`);
    }
    try {
      const bytes = await response.arrayBuffer();
      throwIfAborted(signal);
      return bytes;
    } catch (error) {
      throwIfAborted(signal);
      throw error;
    }
  }
  if (source instanceof ArrayBuffer) {
    const bytes = source.slice(0);
    throwIfAborted(signal);
    return bytes;
  }
  const bytes = await source.arrayBuffer();
  throwIfAborted(signal);
  return bytes;
}

function whileUnaborted<T>(task: Promise<T>, signal?: AbortSignal): Promise<T> {
  if (signal?.aborted === true) return Promise.reject(abortError());
  if (signal === undefined) return task;
  return new Promise<T>((resolve, reject) => {
    const abort = () => finish(() => reject(abortError()));
    const finish = (settle: () => void) => {
      signal.removeEventListener("abort", abort);
      settle();
    };
    signal.addEventListener("abort", abort, { once: true });
    task.then(
      (result) => finish(() => {
        if (signal.aborted) reject(abortError());
        else resolve(result);
      }),
      (error: unknown) => finish(() => {
        if (signal.aborted) reject(abortError());
        else reject(error);
      }),
    );
  });
}

/**
 * One opened document's worker session: the shared lifecycle (open /
 * unit / media / close, font resolution, abort wiring) behind every
 * format-native document. Format modules wrap this in collections and
 * handles; they never touch `Worker` or the RPC protocol directly.
 */
export class WorkerSession {
  private closed = false;
  private readonly closeWaiters = new Set<(error: Error) => void>();
  private readonly resolvedFonts = new Map<string, Promise<FontFace | undefined>>();

  private constructor(
    readonly info: DocumentInfo,
    private readonly bridge: WorkerBridge,
    private readonly sessionId: number,
    private readonly fontResolver: FontResolver | undefined,
    private readonly scope: string,
  ) {}

  static async open(
    scope: string,
    source: File | Blob | ArrayBuffer | string,
    options?: OpenSessionOptions,
  ): Promise<WorkerSession> {
    const signal = options?.signal;
    const bytes = await sourceBytes(scope, source, signal);
    throwIfAborted(signal);
    const bridge = new WorkerBridge(scope);
    const sessionId = nextSessionId++;
    try {
      await whileUnaborted(bridge.waitUntilReady(), signal);
      throwIfAborted(signal);
      const info = await bridge.request<DocumentInfo>(sessionId, { type: "open", bytes }, signal, [bytes]);
      return new WorkerSession(info, bridge, sessionId, options?.fontResolver, scope);
    } catch (error) {
      bridge.close(sessionId);
      throw error;
    }
  }

  get isClosed(): boolean {
    return this.closed;
  }

  assertOpen(): void {
    if (this.closed) throw closedError(this.scope);
  }

  async unit<T extends WorkerResult>(index: number, options?: UnitRequestOptions): Promise<T> {
    // Check before creating the bridge request: otherwise the rejected
    // request promise floats unhandled while `whileOpen` rejects separately.
    if (this.closed) throw closedError(this.scope);
    if (options?.signal?.aborted === true) throw abortError();
    const call: WorkerCall = options?.range === undefined
      ? { type: "unit", index }
      : { type: "unit", index, range: options.range };
    const model = await this.whileOpen(
      this.bridge.request<T>(this.sessionId, call, options?.signal),
      options?.signal,
    );
    await this.resolveFonts(model, options?.signal);
    if (this.closed) throw closedError(this.scope);
    if (options?.signal?.aborted) throw abortError();
    return model;
  }

  media(mediaId: string, options?: MediaOptions): Promise<ArrayBuffer> {
    if (this.closed) return Promise.reject(closedError(this.scope));
    if (options?.signal?.aborted === true) return Promise.reject(abortError());
    return this.whileOpen(
      this.bridge.request<ArrayBuffer>(this.sessionId, { type: "media", mediaId }, options?.signal),
      options?.signal,
    );
  }

  /** Format sniffing for `<aared-viewer>`: reads content types only, keeps no session. */
  static async detectFormat(
    scope: string,
    source: File | Blob | ArrayBuffer | string,
    signal?: AbortSignal,
  ): Promise<DocumentInfo["format"]> {
    const bytes = await sourceBytes(scope, source, signal);
    throwIfAborted(signal);
    const bridge = new WorkerBridge(scope);
    try {
      await whileUnaborted(bridge.waitUntilReady(), signal);
      throwIfAborted(signal);
      return await bridge.request<DocumentInfo["format"]>(0, { type: "detect", bytes }, signal, [bytes]);
    } finally {
      // Session 0 is never installed; closing only terminates the bridge.
      bridge.close(0);
    }
  }

  resolveFontsFor(model: unknown, signal?: AbortSignal): Promise<void> {
    return this.resolveFonts(model, signal);
  }

  close(): void {
    if (this.closed) return;
    this.closed = true;
    const error = closedError(this.scope);
    for (const reject of this.closeWaiters) reject(error);
    this.closeWaiters.clear();
    this.bridge.close(this.sessionId);
  }

  private whileOpen<T>(task: Promise<T>, signal?: AbortSignal): Promise<T> {
    if (this.closed) return Promise.reject(closedError(this.scope));
    if (signal?.aborted === true) return Promise.reject(abortError());
    return new Promise<T>((resolve, reject) => {
      const close = (error: Error) => {
        finish(() => reject(error));
      };
      const abort = () => finish(() => reject(abortError()));
      const finish = (settle: () => void) => {
        this.closeWaiters.delete(close);
        signal?.removeEventListener("abort", abort);
        settle();
      };
      this.closeWaiters.add(close);
      signal?.addEventListener("abort", abort, { once: true });
      task.then(
        (result) => {
          finish(() => {
            if (this.closed) reject(closedError(this.scope));
            else resolve(result);
          });
        },
        (error: unknown) => {
          finish(() => reject(error));
        },
      );
    });
  }

  private resolveFonts(model: unknown, signal?: AbortSignal): Promise<void> {
    if (this.fontResolver === undefined) return Promise.resolve();
    if (this.closed) return Promise.reject(closedError(this.scope));
    if (signal?.aborted === true) return Promise.reject(abortError());
    return this.whileOpen(
      Promise.all(
        requestedFontFamilies(model).map((family) => {
          let resolved = this.resolvedFonts.get(family);
          if (resolved === undefined) {
            resolved = Promise.resolve(this.fontResolver!(family));
            this.resolvedFonts.set(family, resolved);
          }
          return resolved;
        }),
      ).then(() => undefined),
      signal,
    );
  }
}

export type { AaredErrorCode };
export { AaredError };
