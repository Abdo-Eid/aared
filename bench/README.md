# bench

Visual bench for aared: open real Office/PDF files and **see** them, with the
open speed on screen. Third-party renderers per format (registry in
`src/adapters/`); aared's own renderers drop into the second column.

```bash
bun install          # from the repo root (workspace)
bun run dev          # serves bench/ at http://localhost:5173
```

- Left rail: one-click sample buttons from canonical `../fixtures/` (Arabic deck
  first, then the three chart corpora from `../tools/chart-oracle/`, served at
  `/chart-oracle/` by a plugin in `vite.config.ts`) and file upload + drag-drop.
- Header renderer picker: `aared` plus every baseline registered for the open
  format — PPTX: pptxviewjs, silurus, aiden0z, mdurbar, pptx-browser,
  extend, omni; XLSX: sheetjs, xlscalc, spreadsheet-preview, extend, omni;
  DOCX: docx-preview, docx-renderer, apollo, extend, omni. `extend` is
  Extend's `@extend-ai/react-{pptx,docx,xlsx}`, one React component per
  format, each mounted in its own React root.
  Single renders the pick; compare pits it against aared. PDF stays control.
- Search: the aared card has a search button beside Download for PPTX, XLSX
  and DOCX. The panel starts closed; opening it focuses the native search
  input. The compact bar shows an active/total counter and previous/next
  arrows; Enter advances and Shift+Enter goes back. Soft highlights mark all
  matches and a stronger outlined highlight marks the current one. Arrows wrap
  and scroll to the active match across slides, cells, or DOCX occurrences. XLSX searches
  the current sheet and refreshes results when the sheet changes. Closing the
  panel clears its query, counter and highlights; reopening starts empty.
  The input uses the browser's native clear control. This is bench UI, built
  on `viewer.search()`; no search panel ships in the package or baseline cards.
- Timing is `performance.now()` from bytes-in-hand to first meaningful
  paint, shown in the footer chip. `?fonts=aared` gives the aared column
  `aared-fonts`' substitutes (off by default, so timings carry no font
  downloads).
- The interactive aared card uses `usePreview()` and takes slide/sheet position
  directly from its hook. DOCX paging calls `viewer.docx.goToPage()`; DOM page
  observation still supplies the scroll-based counter. The frame owns display
  controls only; each card owns its renderer's lifetime.
- Measurement adapters use format-local `preview()` with the supplied abort
  signal and one `destroy()`. XLSX tabs stay off and DOCX stays in layout mode;
  timing still includes dynamic import, opening and first paint.
- Deterministic numbers live outside the UI: `measure.html` is the page the
  harness drives, and `bun tools/bench-measure.ts` runs it in a real browser
  (`--run` for the open-time and memory passes, `--size` for every file a
  renderer fetches, `--search`), writes to `tools/bench-results/` (local,
  gitignored) and regenerates the tables in `docs/BENCHMARKS.md` via
  `--render-md`. `docs/BENCHMARKS.md` has the method.

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
Word does.

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
