import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);

/**
 * Standalone visual bench. No aliases into other repos, no shims: every
 * renderer here is an npm dependency of this package. Two bench-only
 * module rewrites, both confined to third-party Node-first packaging:
 * - `konva` (via docx-renderer) ships a node-first `main` (`index-node.js`,
 *   which requires the native `canvas` binary) with the browser build behind
 *   the `browser` field — Rollup takes `main`, so pin the import to the
 *   browser build explicitly.
 * - `canvas` / `@napi-rs/canvas*` are Node-native binaries some dependency
 *   graphs reference statically; nothing executes them in the browser, so
 *   they resolve to an empty stub instead of breaking the bundle. A column
 *   that truly needed them would fail visibly at runtime, not silently.
 * Product code is unaffected by either.
 */
export default defineConfig({
  plugins: [react(), tailwindcss()],
  publicDir: fileURLToPath(new URL("../fixtures", import.meta.url)),
  optimizeDeps: {
    // @silurus/ooxml resolves its parser WASM as
    // `new URL("<format>_parser_bg.wasm", import.meta.url)`. Pre-bundling
    // rewrites import.meta.url to `.vite/deps/`, so the fetch 404s and the
    // dev server's SPA fallback returns index.html — the WASM magic check
    // then fails on `<!do`. Excluding keeps the real file URL (known Vite
    // issue #8427/#19310; production build is unaffected — it emits the
    // .wasm as a real asset). ESM-only package, so exclusion is safe.
    exclude: [
      "@silurus/ooxml",
      "@silurus/ooxml/pptx",
      "@silurus/ooxml/xlsx",
      "@silurus/ooxml/docx",
    ],
  },
  resolve: {
    alias: [
      // Absolute path required: per vite docs, relative alias values "will
      // be used as-is and will not be resolved" — a bare
      // "konva/lib/index.js" falls back to the node-first `main`.
      // require.resolve on the subpath is manager-agnostic (.bun/pnpm/npm).
      { find: /^konva$/, replacement: require.resolve("konva/lib/index.js") },
      {
        find: /^canvas$/,
        replacement: fileURLToPath(new URL("./src/lib/node-shim.ts", import.meta.url)),
      },
      {
        find: /^@napi-rs\/canvas/,
        replacement: fileURLToPath(new URL("./src/lib/node-shim.ts", import.meta.url)),
      },
    ],
  },
  build: {
    rollupOptions: {
      // The bench app, plus the page `tools/bench-measure.ts` drives.
      input: {
        main: fileURLToPath(new URL("./index.html", import.meta.url)),
        measure: fileURLToPath(new URL("./measure.html", import.meta.url)),
      },
    },
  },
  server: {
    port: 5173,
    open: false,
  },
});
