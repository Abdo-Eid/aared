import { defineConfig } from "tsdown";

// Shared chunks must live under dist/core/: WorkerSession starts the
// worker with `new URL("./worker.mjs", import.meta.url)`, which resolves
// relative to the emitted session chunk. Left at the dist root (the
// default), the URL 404s because the worker entry emits to
// dist/core/worker.mjs — the bench's "worker failed to start" failure.
// Shipped JS keeps no doc comments: `@aared/viewer/worker?raw`, dev servers and
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
    // Shared review code crosses the test/tools boundary. Declaration emit
    // must see shipped source only, never write tooling .d.ts beside its source.
    dts: { tsconfig: "tsconfig.build.json" },
    clean: true,
    // Shipped minified: a consumer bundler re-minifies anyway, but the
    // `@aared/viewer/worker?raw` Blob recipe, no-bundler and CDN pages run dist as-is.
    minify: true,
    // Development diagnostics (`src/diagnostics.ts`: PPTX dashed fallback
    // outlines, the XLSX sheet notice bar) paint when running from source
    // (unit/visual tests, harnesses) but the shipped bundles strip them —
    // builds like the bench render fallback geometry as an invisible box and
    // sheets without a notice bar. Tests import from `src`, where the
    // identifier below is undeclared and the `typeof` guard keeps them on.
    define: { __AARED_MARK_FALLBACKS__: "false" },
    // pptx/xlsx/docx stay separate chunks: their fixed viewer routes depend
    // only on their own renderer. Element and preview reach those routes
    // through auto-viewer.ts's dynamic imports:
    // rolldown code-splits those into per-entry chunks automatically. That
    // separation is structural, not an externals setting. The package has no
    // npm dependencies to bundle, and `react` is an optional peer of the
    // `@aared/viewer/react` entry that is imported, not bundled.
    outputOptions: coreChunks,
  },
  {
    // The worker is one self-contained file: the same `src/core/worker.ts`
    // source with every format parser flattened into it (`codeSplitting:
    // false`), so it needs no sibling chunks and can be copied or inlined into
    // a Blob URL as is.
    // Worker imports are static and symbol-level: with splitting disabled, dynamic
    // imports cannot defer bytes and only obstruct whole-worker tree shaking.
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
