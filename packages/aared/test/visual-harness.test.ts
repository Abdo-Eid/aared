import { expect, test } from "bun:test";
import {
  browserModule,
  browserModuleCacheSize,
  clearBrowserModuleCache,
  serveVisual,
} from "./harness.js";

test("visual harness: identical browser builds happen once per run", async () => {
  clearBrowserModuleCache();
  expect(browserModuleCacheSize()).toBe(0);
  const entry = `${import.meta.dir}/../src/pptx/index.ts`;
  const [first, second] = await Promise.all([browserModule(entry), browserModule(entry)]);
  expect(first).toBe(second);
  expect(first.length).toBeGreaterThan(0);
  expect(browserModuleCacheSize()).toBe(1);
  // A second entry builds separately; the first stays cached.
  const other = await browserModule(`${import.meta.dir}/../src/xlsx/index.ts`);
  expect(other.length).toBeGreaterThan(0);
  expect(other).not.toBe(first);
  expect(browserModuleCacheSize()).toBe(2);
  expect(await browserModule(entry)).toBe(first);
  expect(browserModuleCacheSize()).toBe(2);
});

test("visual harness: a failed build does not poison the cache", async () => {
  clearBrowserModuleCache();
  const missing = `${import.meta.dir}/does-not-exist.ts`;
  await expect(browserModule(missing)).rejects.toThrow();
  expect(browserModuleCacheSize()).toBe(0);
  await expect(browserModule(missing)).rejects.toThrow();
  expect(browserModuleCacheSize()).toBe(0);
});

test("visual harness: production painter flags have isolated cache keys", async () => {
  clearBrowserModuleCache();
  const entry = `${import.meta.dir}/../src/pptx/renderer.ts`;
  const development = await browserModule(entry);
  const production = await browserModule(entry, { production: true });
  expect(production.length).toBeLessThan(development.length);
  // Bun may retain a false branch in JS; the flag, not string removal,
  // controls runtime chrome. The shipped `dist` bundles are checked
  // separately, in build-output.test.ts.
  expect(development).toContain("__AARED_MARK_FALLBACKS__");
  expect(production).not.toContain("__AARED_MARK_FALLBACKS__");
  expect(browserModuleCacheSize()).toBe(2);
  expect(await browserModule(entry, { production: true })).toBe(production);
});

test("visual harness: static servers are isolated but share content", async () => {
  const page = "<!doctype html><p>page</p>";
  const first = serveVisual({ modules: { "/a.mjs": "export const a = 1;" }, files: { "/f.bin": new Uint8Array([1, 2, 3]).buffer as ArrayBuffer }, page });
  const second = serveVisual({ modules: { "/a.mjs": "export const a = 1;" }, files: { "/f.bin": new Uint8Array([1, 2, 3]).buffer as ArrayBuffer }, page });
  try {
    expect(first.url.port).not.toBe(second.url.port);
    for (const server of [first, second]) {
      const module = await fetch(new URL("/a.mjs", server.url.href));
      expect(module.headers.get("content-type")).toContain("text/javascript");
      expect(await module.text()).toBe("export const a = 1;");
      const file = await fetch(new URL("/f.bin", server.url.href));
      expect([...new Uint8Array(await file.arrayBuffer())]).toEqual([1, 2, 3]);
      const fallback = await fetch(new URL("/anything-else", server.url.href));
      expect(await fallback.text()).toBe(page);
    }
    // Stopping one server leaves the other serving.
    first.stop(true);
    expect((await fetch(new URL("/a.mjs", second.url.href))).status).toBe(200);
  } finally {
    first.stop(true);
    second.stop(true);
  }
});

test("visual harness: custom routes and page headers win over static handling", async () => {
  const server = serveVisual({
    modules: { "/a.mjs": "static" },
    page: "<!doctype html><p>page</p>",
    pageHeaders: { "content-security-policy": "default-src 'self'" },
    onRequest: (url) => {
      if (url.pathname === "/a.mjs") return new Response("custom", { headers: { "content-type": "text/javascript" } });
      return undefined;
    },
  });
  try {
    expect(await (await fetch(new URL("/a.mjs", server.url.href))).text()).toBe("custom");
    const fallback = await fetch(new URL("/", server.url.href));
    expect(fallback.headers.get("content-security-policy")).toBe("default-src 'self'");
  } finally {
    server.stop(true);
  }
});
