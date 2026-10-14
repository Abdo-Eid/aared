/**
 * The benchmark's own DevTools connection to the browser `Bun.WebView`
 * drives, so memory can be read from the page *and* from each Worker.
 *
 * `view.cdp()` is scoped to the view's page session: a `sessionId` passed in
 * its params is ignored, so a "Worker" reading taken through it is the page
 * read a second time (the page counted twice, the Worker never). Here the
 * browser is started with a DevTools port, the harness holds one browser-level
 * WebSocket, attaches to each page with flat sessions and auto-attaches its
 * Workers, and sends every command with a top-level `sessionId`. Each session
 * proves what it is (`typeof WorkerGlobalScope`) before its first reading.
 */

type Json = Record<string, unknown>;

/** One target's live memory, bytes: JS heap, embedder heap (DOM, layout, styles), buffers (ArrayBuffers, WASM). */
export interface HeapPart {
  js: number;
  embedder: number;
  buffers: number;
}

export interface HeapReading {
  page: HeapPart;
  workers: HeapPart[];
}

export const partBytes = (part: HeapPart): number => part.js + part.embedder + part.buffers;
export const readingBytes = (reading: HeapReading): number =>
  partBytes(reading.page) + reading.workers.reduce((sum, part) => sum + partBytes(part), 0);
export const workerBytes = (reading: HeapReading): number => reading.workers.reduce((sum, part) => sum + partBytes(part), 0);

const COMMAND_TIMEOUT_MS = 90_000;
const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

async function freePort(): Promise<number> {
  const probe = Bun.serve({ port: 0, fetch: () => new Response() });
  const port = probe.port;
  probe.stop(true);
  if (port === undefined) throw new Error("no free port for the DevTools connection");
  return port;
}

export class Browser {
  #next = 0;
  readonly #pending = new Map<number, { resolve: (value: Json) => void; reject: (error: Error) => void }>();
  /** Worker session → the page session it belongs to. */
  readonly #workerOf = new Map<string, string>();
  #pages = 0;

  private constructor(
    private readonly socket: WebSocket,
    /** The first view: it owns the browser process, so it stays open until `close()`. */
    private readonly first: Bun.WebView,
  ) {
    socket.onmessage = (event) => this.#receive(JSON.parse(String(event.data)) as Json);
  }

  /** Start the browser with a DevTools port. Must create the process's first `Bun.WebView`. */
  static async launch(): Promise<Browser> {
    const port = await freePort();
    const first = new Bun.WebView({
      width: 400,
      height: 300,
      // `url: false` always spawns a fresh browser: never attach to one the user has open.
      backend: { type: "chrome", url: false, argv: [`--remote-debugging-port=${port}`] },
    });
    await first.navigate("about:blank");
    const version = (await (await fetch(`http://127.0.0.1:${port}/json/version`)).json()) as { webSocketDebuggerUrl: string };
    const socket = new WebSocket(version.webSocketDebuggerUrl);
    await new Promise<void>((resolve, reject) => {
      socket.onopen = () => resolve();
      socket.onerror = () => reject(new Error("could not open the DevTools connection"));
    });
    return new Browser(socket, first);
  }

  /** The browser's user agent product, e.g. `HeadlessChrome/153.0.0.0`. */
  async product(): Promise<string> {
    const agent = String(await this.first.evaluate("navigator.userAgent"));
    return agent.match(/(?:Headless)?Chrome\/[\d.]+/)?.[0] ?? "unknown";
  }

  send<T = Json>(method: string, params: Json = {}, sessionId?: string): Promise<T> {
    const id = ++this.#next;
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.#pending.delete(id);
        reject(new Error(`DevTools ${method} timed out`));
      }, COMMAND_TIMEOUT_MS);
      this.#pending.set(id, {
        resolve: (value) => {
          clearTimeout(timer);
          resolve(value as T);
        },
        reject: (error) => {
          clearTimeout(timer);
          reject(error);
        },
      });
      this.socket.send(JSON.stringify(sessionId === undefined ? { id, method, params } : { id, method, params, sessionId }));
    });
  }

  #receive(message: Json): void {
    if (message.method === "Target.attachedToTarget") {
      const params = message.params as { sessionId: string; targetInfo: { type: string } };
      const parent = message.sessionId as string | undefined;
      if (params.targetInfo.type === "worker" && parent !== undefined) {
        const page = this.#workerOf.get(parent) ?? parent;
        this.#workerOf.set(params.sessionId, page);
        // A Worker may start Workers of its own; they belong to the same page.
        void this.send("Target.setAutoAttach", { autoAttach: true, waitForDebuggerOnStart: false, flatten: true }, params.sessionId).catch(() => {});
      }
      return;
    }
    if (message.method === "Target.detachedFromTarget") {
      this.#workerOf.delete((message.params as { sessionId: string }).sessionId);
      return;
    }
    const id = message.id as number | undefined;
    if (id === undefined) return;
    const pending = this.#pending.get(id);
    if (pending === undefined) return;
    this.#pending.delete(id);
    const error = message.error as { message: string } | undefined;
    if (error !== undefined) pending.reject(new Error(error.message));
    else pending.resolve((message.result ?? {}) as Json);
  }

  /** Worker sessions currently attached under a page session. */
  workersOf(pageSession: string): string[] {
    return [...this.#workerOf].filter(([, page]) => page === pageSession).map(([session]) => session);
  }

  /** Open a page at `url` in a new view and attach to it (and to every Worker it starts). */
  async open(url: string, width = 1280, height = 720): Promise<Page> {
    const token = `bench-page-${++this.#pages}`;
    const view = new Bun.WebView({ width, height, backend: { type: "chrome", url: false } });
    const tagged = `${url}${url.includes("?") ? "&" : "?"}${token}`;
    await view.navigate(tagged);
    const { targetInfos } = await this.send<{ targetInfos: Array<{ type: string; url: string; targetId: string }> }>("Target.getTargets");
    const target = targetInfos.find((info) => info.type === "page" && info.url.endsWith(token));
    if (target === undefined) {
      view.close();
      throw new Error(`could not find the page ${tagged} among the browser's targets`);
    }
    const { sessionId } = await this.send<{ sessionId: string }>("Target.attachToTarget", { targetId: target.targetId, flatten: true });
    await this.send("Target.setAutoAttach", { autoAttach: true, waitForDebuggerOnStart: false, flatten: true }, sessionId);
    return new Page(this, view, sessionId);
  }

  close(): void {
    this.socket.close();
    this.first.close();
    Bun.WebView.closeAll();
  }
}

/** One page under measurement, with its Workers. */
export class Page {
  readonly #verified = new Set<string>();

  constructor(
    private readonly browser: Browser,
    private readonly view: Bun.WebView,
    private readonly session: string,
  ) {}

  /** Run a page expression. */
  evaluate<T>(expression: string): Promise<T> {
    return this.view.evaluate(expression) as Promise<T>;
  }

  private workers(): string[] {
    return this.browser.workersOf(this.session);
  }

  /** Each session answers from the right realm before it is ever read: the page is no Worker, a Worker is. */
  async #verify(session: string, worker: boolean): Promise<void> {
    if (this.#verified.has(session)) return;
    const { result } = await this.browser.send<{ result: { value: string } }>(
      "Runtime.evaluate",
      { expression: "typeof WorkerGlobalScope", returnByValue: true },
      session,
    );
    const expected = worker ? "function" : "undefined";
    if (result.value !== expected) {
      throw new Error(`DevTools session routing is wrong: a ${worker ? "Worker" : "page"} session answered typeof WorkerGlobalScope = ${result.value}`);
    }
    this.#verified.add(session);
  }

  async #part(session: string, worker: boolean): Promise<HeapPart> {
    await this.#verify(session, worker);
    const usage = await this.browser.send<{ usedSize: number; embedderHeapUsedSize: number; backingStorageSize: number }>(
      "Runtime.getHeapUsage",
      {},
      session,
    );
    return { js: usage.usedSize, embedder: usage.embedderHeapUsedSize, buffers: usage.backingStorageSize };
  }

  /** Force a GC in the page and in every Worker, so readings compare live memory, not garbage. */
  async collect(): Promise<void> {
    await this.browser.send("HeapProfiler.collectGarbage", {}, this.session);
    for (const worker of this.workers()) {
      await this.browser.send("HeapProfiler.collectGarbage", {}, worker).catch(() => {});
    }
  }

  /** Page and Worker memory, each read in its own session. */
  async heap(): Promise<HeapReading> {
    const page = await this.#part(this.session, false);
    const workers: HeapPart[] = [];
    for (const worker of this.workers()) {
      try {
        workers.push(await this.#part(worker, true));
      } catch (error) {
        // A Worker that ended between listing and reading holds nothing; a
        // routing failure is a real error.
        if (error instanceof Error && error.message.startsWith("DevTools session routing")) throw error;
      }
    }
    return { page, workers };
  }

  /**
   * Let the page come to rest: Workers a destroy just ended have detached
   * (the set holds still for a while), then two GCs.
   */
  async settle(ms = 300): Promise<void> {
    await sleep(ms);
    let last = this.workers().join();
    let stableSince = performance.now();
    const deadline = performance.now() + 3000;
    while (performance.now() < deadline && performance.now() - stableSince < 250) {
      await sleep(50);
      const now = this.workers().join();
      if (now !== last) {
        last = now;
        stableSince = performance.now();
      }
    }
    await this.collect();
    await this.collect();
  }

  close(): void {
    this.view.close();
  }
}
