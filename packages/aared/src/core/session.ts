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

import type { DocumentFormat, DocumentInfo, FontResolver, MediaCollection, MediaOptions, WorkerInit } from "./index.js";
import type { CellRange, SlideLayoutModel } from "./model.js";
import type { SerializedError, WorkerCall, WorkerRequest, WorkerResponse, WorkerResult } from "./rpc.js";
import { AaredError, abortError, isAbort, throwIfAborted, type AaredErrorCode } from "./errors.js";
import { collectDeferredGeometry, collectStrings } from "./collect.js";

export type { DocumentInfo, FontResolver };
export type { CellRange };

export interface OpenSessionOptions {
  fontResolver?: FontResolver;
  signal?: AbortSignal;
  worker?: WorkerInit;
  /**
   * The preset geometry table (`core/presets.ts`), supplied only by
   * `aared/pptx`'s `open()`. It is threaded in rather than imported here
   * because this module is bundled into every format's chunks: a static
   * import would make xlsx/docx apps carry every preset outline. Without it
   * deferred geometry is not resolved.
   */
  geometryPrograms?: Readonly<Record<string, string>>;
}

/** Narrow numeric window for an XLSX `unit` request; other formats reject it. */
export interface UnitRange {
  start: { row: number; col: number };
  end: { row: number; col: number };
}

export interface UnitRequestOptions {
  signal?: AbortSignal;
  /**
   * Set `false` for reads that only inspect the model (search) and never
   * paint it: skips the font-resolver round trip, which can trigger font
   * loads, and skips deferred-geometry evaluation. Defaults to resolving
   * fonts.
   */
  fonts?: boolean;
  /**
   * Set `false` for reads that never paint (search): skips deferred-geometry
   * evaluation, which would otherwise load the evaluator and walk the preset
   * table for outlines a search never uses. Defaults to resolving geometry.
   */
  geometry?: boolean;
  /**
   * Set `false` for reads that never paint (search): the Worker serves a
   * cached unit but doesn't keep this one in its bounded cache, so a search
   * over every unit doesn't push out the ones being viewed.
   */
  retain?: boolean;
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

/** A read after `close()`: a usage error, never a cancellation (so `isAbort()` doesn't swallow it). */
function closedError(scope: string): AaredError {
  return new AaredError("closed", `${scope}: document is closed`);
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
  return new AaredError("worker-error", `${scope}: failed to post a Worker request: ${detail}`);
}

class WorkerBridge {
  private readonly worker: Worker;
  private readonly pending = new Map<number, PendingRequest>();
  private nextRequestId = 1;
  private stopped = false;
  private ready: Promise<void>;
  private resolveReady!: () => void;
  private rejectReady!: (reason: Error) => void;
  /**
   * Nothing was cancelled, failed to post, or closed while in flight. A
   * cancelled or abandoned request can keep the Worker busy until its next
   * abort checkpoint (a synchronous XML parse can't be interrupted), so only a
   * clean bridge is idle enough to serve the next document.
   */
  private clean = true;
  /** The `Worker` constructor this bridge was built with: a parked bridge never serves a swapped one. */
  private readonly workerClass: unknown;

  constructor(private scope: string, workerInit?: WorkerInit, private readonly poolKey?: string) {
    if (typeof Worker === "undefined") {
      throw new AaredError("unsupported-feature", `${scope}: open() requires a browser Worker`);
    }
    // Keep the default branch in the exact `new Worker(new URL(...))` shape:
    // Vite only bundles the worker target when `new URL()` appears literally
    // inside `new Worker()` (see "Web Workers / Import with Constructors" in
    // the Vite docs). The `workerInit ?? new URL(...)` form this replaced
    // silently defeated detection, so production builds copied worker.mjs
    // verbatim and every open() hung. Never fold
    // the branches back together.
    if (typeof workerInit === "function") {
      this.worker = workerInit();
    } else if (typeof workerInit === "string" || workerInit instanceof URL) {
      this.worker = new Worker(workerInit, { type: "module" });
    } else {
      this.worker = new Worker(new URL("./worker.mjs", import.meta.url), { type: "module" });
    }
    this.workerClass = Worker;
    this.ready = new Promise<void>((resolve, reject) => {
      this.resolveReady = resolve;
      this.rejectReady = reject;
    });
    // A parked bridge may fail to boot with nobody waiting on it; callers that
    // do wait still see the rejection.
    this.ready.catch(() => undefined);
    this.worker.onmessage = (event: MessageEvent<WorkerResponse>) => this.receive(event.data);
    // A dead worker is never a cancellation: it used to report "aborted",
    // which made every abort check (notably usePreview's) swallow boot
    // failures and left the UI loading forever. "worker-error" surfaces.
    this.worker.onerror = () => this.fail(new AaredError("worker-error", `${this.scope}: worker failed to start or stopped unexpectedly`));
  }

  /** Take the bridge for a new document: its messages name the new entrypoint. */
  lease(scope: string): void {
    this.scope = scope;
  }

  /** Alive, idle and clean, and built with the current `Worker`: fit to serve another document. */
  reusable(): boolean {
    return !this.stopped && this.clean && this.pending.size === 0 && this.workerClass === (globalThis as { Worker?: unknown }).Worker;
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
          this.clean = false;
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
        this.clean = false;
        this.takePending(requestId)?.reject(postError(this.scope, error));
      }
    });
  }

  /**
   * End one document's session: the Worker drops everything the session owns
   * and only that session's pending reads reject. A clean bridge then goes
   * back to the pool for the next open; any other one is terminated.
   */
  close(sessionId: number): void {
    if (this.stopped) return;
    try {
      this.worker.postMessage({ type: "close", sessionId } satisfies WorkerRequest);
    } catch {
      // Termination below still releases this document when the close message cannot be sent.
      this.clean = false;
    }
    const error = closedError(this.scope);
    const owned = [...this.pending].filter(([, pending]) => pending.sessionId === sessionId).map(([requestId]) => requestId);
    for (const requestId of owned) this.takePending(requestId)?.reject(error);
    if (owned.length > 0) this.clean = false;
    if (this.poolKey !== undefined && this.reusable()) parkBridge(this, this.poolKey);
    else this.terminate(error);
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
    this.terminate(error);
  }

  /** Stop the Worker for good, rejecting whatever is pending, and leave the pool. */
  terminate(error: Error = closedError(this.scope)): void {
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
    if (this.poolKey !== undefined) evictBridge(this, this.poolKey);
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

/*
 * The Worker pool. Booting a Worker means a new isolate and compiling the
 * parser again (~0.5 MB and some milliseconds each time), so a Worker
 * outlives its document: after `close()` one idle, clean Worker per init is
 * parked for the next `open()` (and the facades' detect-then-open share one),
 * and stops after WORKER_IDLE_MS or when a newer one parks. Two documents open
 * at once still get two Workers. A `() => Worker` factory's Workers are the
 * caller's: used once, terminated on close.
 */

const WORKER_IDLE_MS = 15_000;

interface PoolClock {
  idleMs: number;
  set(run: () => void, ms: number): unknown;
  clear(handle: unknown): void;
}

const systemClock: PoolClock = {
  idleMs: WORKER_IDLE_MS,
  set: (run, ms) => {
    const handle = setTimeout(run, ms);
    // Never keep a process (a test run, a server render) alive for an idle Worker.
    (handle as { unref?: () => void }).unref?.();
    return handle;
  },
  clear: (handle) => clearTimeout(handle as ReturnType<typeof setTimeout>),
};

let poolClock = systemClock;

interface Parked {
  bridge: WorkerBridge;
  timer: unknown;
}

const idleBridges = new Map<string, Parked>();

/** Pool key for a worker init; a factory's Workers are never pooled. */
function workerPoolKey(init: WorkerInit | undefined): string | undefined {
  if (typeof init === "function") return undefined;
  if (init === undefined) return "default";
  return `url:${typeof init === "string" ? init : init.href}`;
}

/** A parked Worker for this init when one is idle, else a new one. */
function acquireBridge(scope: string, init: WorkerInit | undefined): WorkerBridge {
  const key = workerPoolKey(init);
  if (key !== undefined) {
    const parked = idleBridges.get(key);
    if (parked !== undefined) {
      idleBridges.delete(key);
      poolClock.clear(parked.timer);
      if (parked.bridge.reusable()) {
        parked.bridge.lease(scope);
        return parked.bridge;
      }
      parked.bridge.terminate();
    }
  }
  return new WorkerBridge(scope, init, key);
}

function parkBridge(bridge: WorkerBridge, key: string): void {
  const previous = idleBridges.get(key);
  if (previous !== undefined) {
    idleBridges.delete(key);
    poolClock.clear(previous.timer);
    previous.bridge.terminate();
  }
  const entry: Parked = { bridge, timer: undefined };
  entry.timer = poolClock.set(() => {
    if (idleBridges.get(key) === entry) bridge.terminate();
  }, poolClock.idleMs);
  idleBridges.set(key, entry);
}

function evictBridge(bridge: WorkerBridge, key: string): void {
  const parked = idleBridges.get(key);
  if (parked?.bridge !== bridge) return;
  idleBridges.delete(key);
  poolClock.clear(parked.timer);
}

/** Test hook: swap the pool's idle clock; returns the restore. Not exported from any entrypoint. */
export function configureWorkerPool(next: Partial<PoolClock>): () => void {
  const previous = poolClock;
  poolClock = { ...poolClock, ...next };
  return () => {
    poolClock = previous;
  };
}

/** Test hook: terminate every parked Worker now. Not exported from any entrypoint. */
export function drainWorkerPool(): void {
  for (const { bridge } of [...idleBridges.values()]) bridge.terminate();
}

function requestedFontFamilies(model: unknown): string[] {
  const families = new Set<string>();
  collectStrings(model, "fontFamily", families);
  return [...families];
}

let nextSessionId = 1;

/**
 * Internal byte normalization shared by `WorkerSession.open`,
 * `WorkerSession.detectFormat`, `aared/preview` and `<aared-viewer>` (the
 * facades normalize once so URL sources are fetched exactly once across
 * detect + open, and fail with the same `network-error` everywhere).
 * `onProgress` reports bytes as they arrive — streamed for a URL, once before
 * and once after for in-memory sources. Not public API: importable from
 * `./session.js` inside the package only, never re-exported from an
 * entrypoint, never named in the API reference page.
 */
export async function sourceBytes(
  scope: string,
  source: File | Blob | ArrayBuffer | string,
  signal?: AbortSignal,
  onProgress?: (loaded: number, total: number | undefined) => void,
): Promise<ArrayBuffer> {
  throwIfAborted(signal);
  if (typeof source === "string") {
    let response: Response;
    try {
      response = await fetch(source, signal === undefined ? {} : { signal });
    } catch (error) {
      throwIfAborted(signal);
      if (isAbort(error)) throw error;
      const detail = error instanceof Error ? error.message : String(error);
      throw new AaredError("network-error", `${scope}: failed to fetch "${source}": ${detail}`);
    }
    throwIfAborted(signal);
    if (!response.ok) {
      throw new AaredError("network-error", `${scope}: failed to fetch "${source}": HTTP ${response.status}`);
    }
    try {
      const bytes = await readResponse(response, onProgress);
      throwIfAborted(signal);
      return bytes;
    } catch (error) {
      throwIfAborted(signal);
      throw error;
    }
  }
  const total = source instanceof ArrayBuffer ? source.byteLength : source.size;
  onProgress?.(0, total);
  const bytes = source instanceof ArrayBuffer ? source.slice(0) : await source.arrayBuffer();
  throwIfAborted(signal);
  onProgress?.(bytes.byteLength, total);
  return bytes;
}

/** A response body as bytes, reporting each chunk when someone is listening. */
async function readResponse(
  response: Response,
  onProgress: ((loaded: number, total: number | undefined) => void) | undefined,
): Promise<ArrayBuffer> {
  const header = response.headers.get("content-length");
  const parsed = header === null ? Number.NaN : Number(header);
  const total = Number.isFinite(parsed) && parsed >= 0 ? parsed : undefined;
  if (onProgress === undefined || response.body === null) {
    onProgress?.(0, total);
    const bytes = await response.arrayBuffer();
    onProgress?.(bytes.byteLength, total ?? bytes.byteLength);
    return bytes;
  }
  onProgress(0, total);
  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let loaded = 0;
  while (true) {
    const chunk = await reader.read();
    if (chunk.done) break;
    chunks.push(chunk.value);
    loaded += chunk.value.byteLength;
    onProgress(loaded, total);
  }
  const bytes = new Uint8Array(loaded);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes.buffer;
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

/** How a refusal names each format when a document is opened with the wrong entrypoint. */
const FORMAT_NOUN: Record<DocumentFormat, string> = {
  pptx: "a PowerPoint presentation",
  xlsx: "a spreadsheet workbook",
  docx: "a Word document",
};

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
  /** PPTX layouts already fetched (or in flight), by layout id. */
  private readonly layouts = new Map<string, Promise<SlideLayoutModel>>();
  /** Bytes of layout/master media, by media id. */
  private readonly sharedMediaBytes = new Map<string, Promise<ArrayBuffer>>();

  private constructor(
    readonly info: DocumentInfo,
    private readonly bridge: WorkerBridge,
    private readonly sessionId: number,
    private readonly fontResolver: FontResolver | undefined,
    private readonly scope: string,
    private readonly geometryPrograms: Readonly<Record<string, string>> | undefined,
  ) {}

  static async open(
    scope: string,
    source: File | Blob | ArrayBuffer | string,
    options?: OpenSessionOptions,
  ): Promise<WorkerSession> {
    const signal = options?.signal;
    const bytes = await sourceBytes(scope, source, signal);
    throwIfAborted(signal);
    const bridge = acquireBridge(scope, options?.worker);
    const sessionId = nextSessionId++;
    try {
      await whileUnaborted(bridge.waitUntilReady(), signal);
      throwIfAborted(signal);
      const info = await bridge.request<DocumentInfo>(sessionId, { type: "open", bytes }, signal, [bytes]);
      return new WorkerSession(info, bridge, sessionId, options?.fontResolver, scope, options?.geometryPrograms);
    } catch (error) {
      bridge.close(sessionId);
      throw error;
    }
  }

  /**
   * Open a document that must be `format`: the open every format entrypoint
   * wraps. Any other format is closed and refused with a pointer to the
   * entrypoint that does open it.
   */
  static async openAs(
    format: DocumentFormat,
    source: File | Blob | ArrayBuffer | string,
    options?: OpenSessionOptions,
  ): Promise<WorkerSession> {
    const scope = `aared/${format}`;
    const session = await WorkerSession.open(scope, source, options);
    if (session.info.format !== format) {
      const detected = session.info.format;
      session.close();
      throw new AaredError(
        "unsupported-format",
        `${scope}: expected ${FORMAT_NOUN[format]}; detected "${detected}". Open it with aared/${detected} instead`,
      );
    }
    return session;
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
    const call: WorkerCall = {
      type: "unit",
      index,
      ...(options?.range === undefined ? {} : { range: options.range }),
      ...(options?.retain === false ? { retain: false as const } : {}),
    };
    const model = await this.whileOpen(
      this.bridge.request<T>(this.sessionId, call, options?.signal),
      options?.signal,
    );
    await Promise.all([
      options?.fonts === false ? undefined : this.resolveFonts(model, options?.signal),
      options?.geometry === false ? undefined : this.resolveGeometry(model, options?.signal),
    ]);
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

  /**
   * PPTX only: the shared graphics of one layout. Every slide on a layout
   * gets the SAME object, fetched once per session.
   */
  layout(layoutId: string, options?: { signal?: AbortSignal }): Promise<SlideLayoutModel> {
    return this.shared(
      this.layouts,
      layoutId,
      () => this.bridge
        .request<SlideLayoutModel>(this.sessionId, { type: "layout", layoutId })
        .then(async (layout) => {
          await Promise.all([this.resolveFonts(layout), this.resolveGeometry(layout)]);
          return layout;
        }),
      options?.signal,
    );
  }

  /**
   * Bytes of media that many slides paint (layout/master pictures), fetched
   * once per session. The same buffer goes to every caller: read-only, and
   * never transfer it. Slide-owned media stays on `media()`.
   */
  sharedMedia(mediaId: string, options?: MediaOptions): Promise<ArrayBuffer> {
    return this.shared(
      this.sharedMediaBytes,
      mediaId,
      () => this.bridge.request<ArrayBuffer>(this.sessionId, { type: "media", mediaId }),
      options?.signal,
    );
  }

  /**
   * One fetch per key, shared by every caller. The fetch is not bound to any
   * caller's signal (each caller only stops *waiting* when its own signal
   * aborts, so one cancelled reader can't fail the rest), and a failed fetch
   * is dropped so the next caller retries.
   */
  private shared<T>(cache: Map<string, Promise<T>>, key: string, fetch: () => Promise<T>, signal?: AbortSignal): Promise<T> {
    if (this.closed) return Promise.reject(closedError(this.scope));
    let pending = cache.get(key);
    if (pending === undefined) {
      const started = fetch();
      pending = started;
      cache.set(key, started);
      started.catch(() => {
        if (cache.get(key) === started) cache.delete(key);
      });
    }
    return this.whileOpen(pending, signal);
  }

  /** Format sniffing for `<aared-viewer>`: reads content types only, keeps no session. */
  static async detectFormat(
    scope: string,
    source: File | Blob | ArrayBuffer | string,
    options?: { signal?: AbortSignal; worker?: WorkerInit },
  ): Promise<DocumentInfo["format"]> {
    const signal = options?.signal;
    const bytes = await sourceBytes(scope, source, signal);
    throwIfAborted(signal);
    const bridge = acquireBridge(scope, options?.worker);
    try {
      await whileUnaborted(bridge.waitUntilReady(), signal);
      throwIfAborted(signal);
      return await bridge.request<DocumentInfo["format"]>(0, { type: "detect", bytes }, signal, [bytes]);
    } finally {
      // Session 0 is never installed; closing hands the Worker back to the
      // pool (so the open that follows detection reuses it), or stops it when
      // a request was cancelled.
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

  /**
   * Evaluate the geometry the worker deferred (custGeom, and presets beyond
   * its hand-coded fast path) against the preset table `aared/pptx` supplied
   * at open. The evaluator is imported only when a model actually needs it —
   * never from the flat worker bundle. An evaluator that fails to load leaves
   * the worker's visible fallback in place rather than failing the slide.
   */
  private resolveGeometry(model: unknown, signal?: AbortSignal): Promise<void> {
    const programs = this.geometryPrograms;
    if (programs === undefined) return Promise.resolve();
    const pending = collectDeferredGeometry(model);
    if (pending.length === 0) return Promise.resolve();
    return this.whileOpen(
      import("./drawingml-resolve.js")
        .then((engine) => engine.resolveDeferredGeometry(pending, programs))
        .catch(() => undefined)
        .finally(() => {
          for (const geometry of pending) delete geometry.deferred;
        }),
      signal,
    );
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

/** `document.media` for every format: a session's media route behind the public `MediaCollection` shape. */
export class SessionMedia implements MediaCollection {
  constructor(private readonly session: WorkerSession) {}

  get(mediaId: string, options?: MediaOptions): Promise<ArrayBuffer> {
    return this.session.media(mediaId, options);
  }
}

export type { AaredErrorCode };
export { AaredError };
