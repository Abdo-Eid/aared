/**
 * Shared test harness and helpers for aared tests. This module is imported
 * by test files but is not itself a test (no `.test.` in the name).
 */

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
 */
export async function browserModule(entrypoint: string): Promise<string> {
  const build = await Bun.build({ entrypoints: [entrypoint], target: "browser", format: "esm" });
  if (!build.success) throw new Error(build.logs.map((log) => log.message).join("\n"));
  return build.outputs[0]!.text();
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
