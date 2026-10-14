/**
 * Shared test harness and helpers for aared tests. This module is imported
 * by test files but is not itself a test (no `.test.` in the name).
 */

import { test } from "bun:test";
import { resolve } from "node:path";

/**
 * The visual-test gate. Browser proof is valid but must be explicit: visual
 * files use this instead of repeating the environment check, so the
 * documented file-selected command keeps running every browser test.
 */
export const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;

/**
 * Await a promise and return any rejection as an Error. Used to test that
 * functions reject as expected.
 */
export async function rejection(promise: Promise<unknown>): Promise<Error> {
  try {
    await promise;
  } catch (error) {
    return error instanceof Error ? error : new Error(String(error));
  }
  throw new Error("expected the request to reject");
}

/**
 * Build an entry point as browser-compatible ESM. Used by visual tests to
 * bundle aared components and the worker for testing in a headless browser.
 * Identical entry points build once per run: the sources do not change while
 * the tests run, and every visual file rebuilds the same few entries (each
 * format plus the worker), so the cache only removes repeated work. Each
 * test still gets a fresh server and a fresh WebView; only the build text is
 * shared, which carries no per-test state.
 */
const browserModuleCache = new Map<string, Promise<string>>();

async function buildBrowserModule(entrypoint: string): Promise<string> {
  const build = await Bun.build({ entrypoints: [entrypoint], target: "browser", format: "esm" });
  if (!build.success) throw new Error(build.logs.map((log) => log.message).join("\n"));
  return build.outputs[0]!.text();
}

export async function browserModule(entrypoint: string): Promise<string> {
  // Callers spell the same entry differently (`join(dir, "../src/x.ts")` vs a
  // template literal), so key on the resolved path or one entry bundles twice.
  const key = resolve(entrypoint);
  const cached = browserModuleCache.get(key);
  if (cached !== undefined) return cached;
  const pending = buildBrowserModule(entrypoint);
  browserModuleCache.set(key, pending);
  // A failed build must not poison later tests: drop it so a retry rebuilds.
  pending.catch(() => {
    if (browserModuleCache.get(key) === pending) browserModuleCache.delete(key);
  });
  return pending;
}

/** How many distinct entry points are cached (harness proof, not shipped code). */
export function browserModuleCacheSize(): number {
  return browserModuleCache.size;
}

/** Drop all cached builds (harness proof isolation only; the suite never calls this). */
export function clearBrowserModuleCache(): void {
  browserModuleCache.clear();
}

/**
 * Serve static visual-test assets: `modules` (path to ESM source, served as
 * JavaScript), `files` (path to bytes, served as-is), and `page` (the HTML
 * document for every other path). `onRequest` runs first so tests with
 * custom needs (request logging, media types, extra headers) keep them
 * without duplicating the static handling; `pageHeaders` sets the fallback
 * page's headers (for example a content-security-policy).
 */
export function serveVisual(options: {
  modules?: Record<string, string>;
  files?: Record<string, ArrayBuffer | Uint8Array>;
  page: string;
  pageHeaders?: Record<string, string>;
  onRequest?: (url: URL, request: Request) => Response | undefined | Promise<Response | undefined>;
}): ReturnType<typeof Bun.serve> {
  const modules = options.modules ?? {};
  const files = options.files ?? {};
  const pageHeaders = { "content-type": "text/html; charset=utf-8", ...options.pageHeaders };
  return Bun.serve({
    port: 0,
    async fetch(request) {
      const url = new URL(request.url);
      const custom = await options.onRequest?.(url, request);
      if (custom !== undefined) return custom;
      const module = modules[url.pathname];
      if (module !== undefined) {
        return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      const file = files[url.pathname];
      if (file !== undefined) {
        const body = file instanceof Uint8Array ? (file.slice().buffer as ArrayBuffer) : file;
        return new Response(body);
      }
      return new Response(options.page, { headers: pageHeaders });
    },
  });
}

/**
 * Run `fn` against a fresh page: a server (its own by default) plus a fresh
 * WebView navigated to it. Isolation is per call — no shared page, no shared
 * state; the only shared work is the memoized module text passed in. The
 * server and the WebView are both torn down before this returns or throws,
 * including for a server the caller supplied.
 */
export async function withVisualPage<T>(
  options: {
    modules?: Record<string, string>;
    files?: Record<string, ArrayBuffer | Uint8Array>;
    /** Required unless `server` is given (the runner's own server serves it). */
    page?: string;
    pageHeaders?: Record<string, string>;
    onRequest?: (url: URL, request: Request) => Response | undefined | Promise<Response | undefined>;
    width?: number;
    height?: number;
    /**
     * A server the test already owns, instead of letting the runner build one
     * (a bespoke route set, or one whose behavior the test mutates). The runner
     * still drives it and still stops it on the way out, because the page it
     * navigates to lives exactly as long as this call. Pass `keepServer` only
     * when one server must outlive a single call.
     */
    server?: ReturnType<typeof Bun.serve>;
    /** Leave the server running after this call (a later call reuses it). */
    keepServer?: boolean;
  },
  fn: (view: import("bun").WebView) => Promise<T>,
): Promise<T> {
  if (options.server === undefined && options.page === undefined) {
    throw new Error("aared visual harness: withVisualPage needs a page or a server");
  }
  const server = options.server ?? serveVisual(options as { page: string });
  try {
    await using view = new Bun.WebView({ width: options.width ?? 1280, height: options.height ?? 800 });
    await view.navigate(server.url.href);
    return await fn(view);
  } finally {
    // The WebView is disposed above before this runs, so nothing is still
    // talking to the server when it stops.
    if (options.keepServer !== true) server.stop(true);
  }
}

/**
 * Time out a task if it does not complete within `milliseconds`. Throws a
 * descriptive error if the timeout is exceeded.
 */
export async function within<T>(label: string, task: Promise<T>, milliseconds: number): Promise<T> {
  let timeout: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      task,
      new Promise<T>((_, reject) => {
        timeout = setTimeout(() => reject(new Error(`aared visual test: ${label} did not complete within ${milliseconds}ms`)), milliseconds);
      }),
    ]);
  } finally {
    if (timeout !== undefined) clearTimeout(timeout);
  }
}

/**
 * Parse PNG dimensions from a WebView or rasterizer capture. Reads the
 * width and height fields from the PNG header, which appear at offsets 16
 * and 20 as big-endian 32-bit integers.
 */
export function pngSize(bytes: Uint8Array): { width: number; height: number } {
  const signature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  if (!signature.every((value, index) => bytes[index] === value)) {
    throw new Error("aared visual review: WebView capture was not a PNG");
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return { width: view.getUint32(16), height: view.getUint32(20) };
}
