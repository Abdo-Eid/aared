import { defineConfig, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";

const require = createRequire(import.meta.url);

/**
 * The chart corpora (`tools/chart-oracle/`) as bench samples, at
 * `/chart-oracle/<file>`: every chart type aared draws, Office-made, without
 * adding test workbooks to `fixtures/` (the public directory).
 */
const CHART_ORACLE = fileURLToPath(new URL("../tools/chart-oracle/", import.meta.url));
const CHART_ORACLE_FILES = ["chart-oracle.pptx", "types-oracle.xlsx", "chartex-oracle.xlsx"];
function chartOracleSamples(): Plugin {
  return {
    name: "chart-oracle-samples",
    configureServer(server) {
      server.middlewares.use((request, response, next) => {
        const name = request.url?.match(/^\/chart-oracle\/([^/?#]+)/)?.[1];
        if (name === undefined || !CHART_ORACLE_FILES.includes(name)) return next();
        response.setHeader("Content-Type", "application/octet-stream");
        response.end(readFileSync(join(CHART_ORACLE, name)));
      });
    },
    generateBundle() {
      for (const name of CHART_ORACLE_FILES) this.emitFile({ type: "asset", fileName: `chart-oracle/${name}`, source: readFileSync(join(CHART_ORACLE, name)) });
    },
  };
}

/**
 * Standalone visual bench. No aliases into other repos, no shims: every
 * renderer here is an npm dependency of this package. Three bench-only
 * module rewrites, all confined to third-party packaging faults:
 * - `konva` (via docx-renderer) ships a node-first `main` (`index-node.js`,
 *   which requires the native `canvas` binary) with the browser build behind
 *   the `browser` field — Rollup takes `main`, so pin the import to the
 *   browser build explicitly.
 * - `canvas` / `@napi-rs/canvas*` are Node-native binaries some dependency
 *   graphs reference statically; nothing executes them in the browser, so
 *   they resolve to an empty stub instead of breaking the bundle. A column
 *   that truly needed them would fail visibly at runtime, not silently.
 * - `@apollo-design/docx-preview` maps its `development` export condition to
 *   `./src/docx-preview.ts`, which the published package doesn't ship; Vite
 *   dev resolves that condition first and fails, so pin the shipped ESM build.
 * Product code is unaffected by any of these.
 */
export default defineConfig({
  plugins: [react(), tailwindcss(), chartOracleSamples()],
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
      // Same `new URL("<name>.wasm" | "<worker>.js", import.meta.url)`
      // shape for Extend's parsers and their Workers.
      "@extend-ai/react-pptx",
      "@extend-ai/react-docx",
      "@extend-ai/react-xlsx",
    ],
    // Excluding a package also skips converting its CommonJS dependencies,
    // so pre-bundle Extend's CJS-only ones (react-dom/server included) explicitly.
    include: [
      "react-dom/server",
      "@extend-ai/react-pptx > regl",
      "@extend-ai/react-xlsx > regl",
      "@extend-ai/react-docx > utif",
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
        find: /^@apollo-design\/docx-preview$/,
        replacement: join(dirname(require.resolve("@apollo-design/docx-preview/package.json")), "dist/docx-preview.mjs"),
      },
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
  // Workers keep Vite's default `iife` format — the setup a default consumer
  // builds aared's single-file Worker with (DECISION_LOG §20). Extend's XLSX
  // Worker lazy-imports its WASM glue, which `iife` can only emit inlined.
  worker: { rollupOptions: { output: { inlineDynamicImports: true } } },
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
