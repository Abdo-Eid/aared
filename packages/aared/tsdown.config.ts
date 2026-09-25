import { defineConfig } from "tsdown";

// Shared chunks must live under dist/core/: WorkerSession starts the
// worker with `new URL("./worker.mjs", import.meta.url)`, which resolves
// relative to the emitted session chunk. Left at the dist root (the
// default), the URL 404s because the worker entry emits to
// dist/core/worker.mjs — the bench's "worker failed to start" failure.
// Shipped JS keeps no doc comments: `aared/worker?raw`, dev servers and
// no-bundler pages run dist verbatim, and a comment holding one character
// above U+00FF makes V8 store the whole script two bytes per character (see
// src/core/chars.ts). Annotations stay so `/* @__PURE__ */` reaches
// consumer bundlers; the `.d.mts` types are emitted separately and keep
// their docs.
const shippedComments = { legal: true, annotation: true, jsdoc: false } as const;

const coreChunks = {
  chunkFileNames: "core/[name]-[hash].mjs",
  comments: shippedComments,
};

export default defineConfig([
  {
    entry: {
      "core/index": "src/core/index.ts",
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
    // Shipped minified: a consumer bundler re-minifies anyway, but the
    // `aared/worker?raw` Blob recipe, no-bundler and CDN pages run dist as-is.
    minify: true,
    // Development diagnostics (`src/diagnostics.ts`: PPTX dashed fallback
    // outlines, the XLSX sheet notice bar) paint when running from source
    // (unit/visual tests, harnesses) but the shipped bundles strip them —
    // builds like the bench render fallback geometry as an invisible box and
    // sheets without a notice bar. Tests import from `src`, where the
    // identifier below is undeclared and the `typeof` guard keeps them on.
    define: { __AARED_MARK_FALLBACKS__: "false" },
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
    outputOptions: coreChunks,
  },
  {
    // The worker is one self-contained file: the same `src/core/worker.ts`
    // source with every format parser flattened into it (`codeSplitting:
    // false`), so it needs no sibling chunks — deploy `dist/core/worker.mjs`
    // inline it into a Blob URL, or let a consumer bundler copy it verbatim.
    // The per-format dynamic `import()`s in the source resolve in place.
    // Must stay `clean: false` — the main config above owns `dist`.
    entry: { "core/worker": "src/core/worker.ts" },
    format: "esm",
    dts: false,
    clean: false,
    minify: true,
    // Worker stack traces are unreadable minified; the map points at src.
    sourcemap: true,
    outputOptions: { codeSplitting: false, comments: shippedComments },
  },
]);
