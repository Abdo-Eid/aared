import { defineConfig } from "tsdown";

export default defineConfig({
  entry: {
    "core/index": "src/core/index.ts",
    "core/worker": "src/core/worker.ts",
    "pptx/index": "src/pptx/index.ts",
    "xlsx/index": "src/xlsx/index.ts",
    "docx/index": "src/docx/index.ts",
    "preview/index": "src/preview/index.ts",
    "react/index": "src/react/index.ts",
    "element/index": "src/element/index.ts",
  },
  format: "esm",
  dts: true,
  clean: true,
  // `react` is an optional peer of the `aared/react` entry only: never
  // bundle it — the consumer's React renders the hooks.
  // pptx/xlsx/docx stay separate lazy chunks because element/index.ts
  // (and preview/index.ts) reach them only through dynamic `import()`:
  // rolldown code-splits those into per-entry chunks automatically. That
  // separation is structural, not an externals setting — `external`
  // (now `deps.neverBundle`) only governs npm package dependencies, of
  // which this package has none, so no setting is needed here. (An empty
  // `external: []` used to sit here as documentation; tsdown 0.22 warns
  // that the option is deprecated, so it was removed.)
  // Shared chunks must live under dist/core/: WorkerSession starts the
  // worker with `new URL("./worker.mjs", import.meta.url)`, which resolves
  // relative to the emitted session chunk. Left at the dist root (the
  // default), the URL 404s because the worker entry emits to
  // dist/core/worker.mjs — the bench's "worker failed to start" failure.
  outputOptions: {
    chunkFileNames: "core/[name]-[hash].mjs",
  },
});
