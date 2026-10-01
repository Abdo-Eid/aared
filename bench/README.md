# bench

Visual bench for aared: open real Office/PDF files and **see** them, with the
open speed on screen. Third-party renderers per format (registry in
`src/adapters/`); aared's own renderers drop into the second column.

```bash
bun install          # from the repo root (workspace)
bun run dev          # serves bench/ at http://localhost:5173
```

- Left rail: one-click sample buttons from canonical `../fixtures/` (Arabic deck
  first), file upload + drag-drop, per-format "what to look for" checks, and
  a run history with milliseconds.
- Header renderer picker: `aared` plus every baseline registered for the open
  format — PPTX: pptxviewjs, silurus, aiden0z, mdurbar, pptx-browser,
  extend, omni; XLSX: sheetjs, xlscalc, spreadsheet-preview, extend, omni;
  DOCX: docx-preview, docx-renderer, apollo, extend, omni. `extend` is
  Extend's `@extend-ai/react-{pptx,docx,xlsx}`, one React component per
  format, each mounted in its own React root.
  Single renders the pick; compare pits it against aared. PDF stays control.
- Timing is `performance.now()` from bytes-in-hand to first meaningful
  paint, shown in the footer chip and the run history.
- Deterministic numbers live outside the UI: `bun tools/bench-measure.ts
  --size` writes entry weights to `tools/bench-results/` and regenerates the
  tables in `docs/BENCHMARKS.md` via `--render-md`. The WebView N-run pass is
  still TBD there.

Known baseline gaps (not bench bugs): pptxviewjs paints canvas with no text
layer and drops paragraph-level RTL (guarded so it can't clobber
`window.FileReader` anymore); SheetJS CE drops fills/fonts/borders;
spreadsheet-preview is the community build (badge stays, no key in repo);
mdurbar mis-resolves slide-master relationships on the Arabic deck and paints
slide 2's table on all 6 slides — parsed slide count stays correct, so trust
the pixels, not the pager; pptx-browser 4.1.5 throws on every paragraph
with an `<a:pPr>` (`defRPr` used before its declaration in `render.js`), so
text shapes are skipped and early slides paint blank; extend's XLSX grid is
canvas by default and doesn't mirror right-to-left sheets; apollo's
`paginate` reflow breaks pages where its own measurement says, not where
Word does. Per-format notes live in the checks panel.

Build notes: every competitor is a lazy `import()` so one broken renderer
can't poison the others. Bench-only module handling in `vite.config.ts`
(product code unaffected): `konva` (via docx-renderer) is pinned to its
browser build — its node-first `main` pulls a native binary Rollup can't
parse — `canvas` / `@napi-rs/canvas*` resolve to an empty stub, and
`@silurus/ooxml` is excluded from dep pre-bundling so its
`new URL("<format>_parser_bg.wasm", import.meta.url)` keeps the real file
URL (otherwise dev serves `index.html` for the `.wasm` and instantiation
fails on `<!do`; known Vite issue #8427/#19310). The three `@extend-ai`
packages are excluded for the same reason (their WASM and Workers load by
`import.meta.url`), which also skips converting their CommonJS
dependencies, so `regl`, `utif` and `react-dom/server` are pre-bundled
explicitly. Workers keep Vite's default `iife` format (the setup a default
consumer builds aared's Worker with); Extend's XLSX Worker lazy-imports its
WASM glue, so Worker builds inline dynamic imports. `@apollo-design/docx-preview` is pinned to its built
`dist/docx-preview.mjs`: its `development` export condition points at an
unshipped `src/` file.
