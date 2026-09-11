# HANDOFF — aared

Onboarding for whoever picks this up next. Read this first, then the docs it
points to.

## What this is

A client-side previewer for `.pptx`, `.xlsx` and `.docx` files —
**view-only**, Arabic/RTL-correct by design, published as a headless core
plus a `<aared-viewer>` web component. See
[`AGENTS.md`](AGENTS.md) for the working charter and
[`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) for how it's designed to
work. PDF is deliberately out of scope — pair with `pdf.js`.

**The one-line architecture:** parse in a Worker (OPC → DrawingML/styles/
theme → a resolved model), render on the main thread straight to real DOM/
SVG — text is never drawn into a canvas or shaped by anything other than the
browser itself. Full reasoning:
[`docs/DECISION_LOG.md §3`](docs/DECISION_LOG.md#3-text-is-rendered-by-the-browser-never-rasterized--the-load-bearing-decision).

## Current state: API redesign complete, fidelity expansion remains

- **Design complete.** Distilled into `docs/{ARCHITECTURE,API,DECISION_LOG,COMPETITORS}.md`.
  Raw session log preserved at
  [`docs/archive/ms-previewrs/DESIGN.md`](docs/archive/ms-previewrs/DESIGN.md)
  (archived — see banner at the top of that file).
- **M0 established the scaffold** (`AGENTS.md`, this file, `README.md`, the
  Bun workspace, `LICENSE`/`NOTICE`, typed package boundaries, seeded fixtures,
  and the benchmark harness). See
  [`docs/ROADMAP.md`](docs/ROADMAP.md).
- **PPTX foundation works headlessly.** `core` reads an OPC ZIP lazily, indexes
  slide parts without inflating all slides, then parses one requested slide's
  DrawingML text/tables. It preserves paragraph RTL independently from table
  RTL and explicit LTR cells; resolves theme colors/fonts and table-style
  fills; resolves placeholder geometry through slide layout/master inheritance;
  and returns a structured-cloneable, inch-based slide model. Theme defaults
  resolve runs that omit a typeface, preventing host-page fonts from changing
  their metrics. DrawingML `a:tint` resolves in linear-light color space,
  matching PowerPoint's table bands. It parses common shape geometry, solid
  fills/outlines, transforms, pictures, charts, notes, and text-body layout.
  `aared/pptx` paints that model to DOM/SVG inside a fixed authored scene with
  one responsive scale, including browser-rendered Arabic text and explicit
  unsupported-geometry fallbacks.
- **All three formats open through format-specific entrypoints.**
  `aared/pptx`, `aared/xlsx`, and `aared/docx` each export `open()` that
  returns a format-native document (`PptxDocument`, `XlsxDocument`,
  `DocxDocument`). `aared/core` is shared types/errors only — it opens
  nothing. Each document exposes collection grammar (`slides.get()`/
  `sheets.get()`/`sections.list()`/`headings.list()`), typed errors
  (`AaredError` with 10 codes), `MediaCollection`, `Appearance`/dark mode,
  and a lifecycle from `open()` → collection `list()`/`get()` → `read()`
  (where applicable) → `render()` → `destroy()` → `close()`.
- **One-call embed facade.** `aared/preview` exports `preview(source,
  target, options?)`, `detect()`, and `PreviewHandle`: format detection runs
  in the Worker (the same probe `<aared-viewer>` uses), only the detected
  format entrypoint is lazily imported, and the handle owns navigation,
  zoom, render-event forwarding, and teardown. One teardown verb on this
  path — `destroy()` covers render output + Worker session + painted DOM,
  deliberately no separate `close()` (the facade owns the whole session).
  `type` asserts a format and skips detection. Supersedes two old
  deferrals; see [`docs/DECISION_LOG.md §12`](docs/DECISION_LOG.md).
- **Shared async object render envelope.** Every format renders as
  `await render({ <unit>, target, ... })` with a format-native noun
  (`slide` / `sheet` / `document`) — DOCX's old positional sync form is
  gone (pre-release break, not an alias). Passing the owning
  document/handle defaults `mediaResolver` to the session everywhere, so
  only detached-model callers wire media by hand. See
  [`docs/DECISION_LOG.md §13`](docs/DECISION_LOG.md).
- **React hooks over the facade.** `aared/react` exports `usePreview()`
  plus `usePptx()` / `useXlsx()` / `useDocx()` — the same hook with `type`
  asserted (detection skipped, only that format chunk loads) and `info`
  narrowed to the format type. Same open → render → navigate → destroy
  boundary as the facade; search, range reads, and cell navigation stay on
  the headless entrypoints. See
  [`docs/DECISION_LOG.md §14`](docs/DECISION_LOG.md).
- **The default viewer and benchmark route work across formats.**
  `<aared-viewer>` detects format via `WorkerSession.detectFormat()`, then
  lazily imports the matching entrypoint, validates an optional explicit
  `type`, supports unit navigation (`goTo`), zoom, and emits
  `slidechange`/`sheetchange`/`pagechange` lifecycle events. The benchmark
  opens its aared column through `aared/preview` (Worker-side auto-detect,
  not try-order across entrypoints) and renders it through `usePreview()`
  in a dedicated card — that card is the reference consumer integration;
  baseline and aared cards share all chrome below the open step
  (`ViewerFrame` + `useViewerFrame`). XLSX provides a
  virtualized RTL grid with cell navigation, ARIA semantics, keyboard
  control, and adaptive dark mode; DOCX provides flow HTML with bidi
  paragraphs, headings, tables, lists, images, footnotes, and one-time
  headers/footers.
- **Current PPTX limits are deliberate and visible.** Common preset geometry,
  solid paint, and linear gradients render to DOM/SVG. Unknown/custom geometry
  and unsupported fills render bounds-preserving marked fallbacks rather than
  disappearing or pretending to be supported. Pattern/image fills, path
  gradients, embedded fonts, and animations remain follow-on work, not behavior
  hidden behind a silent fallback.
- **Verified title-layout correction:** a bottom-anchored 60pt Arabic title
  clipped because browser CSS applied DrawingML's 90% line spacing as a
  single-line box smaller than the glyph ink. The parser now resolves the
  inherited paragraph default size and each text-frame inset independently;
  the renderer preserves the source `0.9` in the model but clamps its browser
  line-height to `1.1`. The real-fixture DOM range and paired PowerPoint/WebView
  capture confirm the descender is contained. The probe in
  `tools/probe/title-insets/` established master, layout, and slide inset
  precedence. Do not work around this class of defect by disabling text
  clipping globally.
- **The current viewer route is verified:** `bun run typecheck`, `bun test`,
  `bun run test:visual`, and `bun run build` pass. Tests cover the public Worker
  path, lazy slides/media, semantic real-fixture parsing, and WebView DOM/SVG
  output. Native PowerPoint exports pair with same-size aared captures for
  review before accepting text fidelity; the LibreOffice conversion is a
  companion artifact only. API docs parity verifies exports and documented
  callables.
- **All three implementation routes are present.** Fidelity review and
  deliberate unsupported-feature closure remain active; see the roadmap for
  scope and acceptance gates.

## Run it

```bash
bun install            # workspace install
bun run dev               # bench/ visual harness at http://localhost:4321
bun run typecheck       # tsc --noEmit
bun test                # bun:test — core and renderer coverage for all formats
bun run build            # tsdown — real dist/*/index.mjs + .d.mts output
```

`bun run test:visual` runs focused Bun.WebView DOM-output and capture tests.
Use `bun run test:visual:artifacts` to retain native/aared PNG pairs and
their review pages (same-size PowerPoint/aared slides; per-sheet
Excel/aared grids). LibreOffice/`pdftoppm` output is optional and is
not a typography oracle.

## The gate ladder (fill in as milestones land — see `AGENTS.md § Testing rule`)

| Gate | Proves | Status |
|---|---|---|
| `typecheck` | Workspace + Worker/core/renderer types compile | ✅ current route |
| `build` | `tsdown` produces real `dist/*/index.mjs` + `.d.mts` per subpath export | ✅ M0 |
| unit tests | Lazy OPC, Worker RPC, real-fixture parsing, and renderer behavior | ✅ current route |
| `docs-parity` | Every public export named in the API reference page, every doc-named identifier is real | ✅ M1 |
| PPTX unit tests | Lazy OPC, Worker route, DrawingML text/table/common-shape parsing | ✅ current route |
| PPTX renderer DOM test | Bun.WebView re-reads Arabic bullets, RTL table order/colors, fixed-scene geometry, font isolation, transforms, and SVG output | ✅ current route |
| PPTX visual review | Same-size native PowerPoint and WebView DOM/SVG captures inspected side by side, including an Arabic deck with a table | ⏳ required before M1 acceptance |
| XLSX unit tests + visual regression | `xlsx` styles/grid/RTL/headings, plus per-sheet native Excel-print/aared A4 capture pairs | ✅ current route |
| DOCX unit tests + visual regression | `docx` flow rendering, bidi | ✅ current route |

## Repo map

```
AGENTS.md               working charter — rules, capability-delivery philosophy, ground-truth probe
HANDOFF.md               this file
README.md                public pitch
LICENSE / NOTICE         Apache-2.0
docs/                    ARCHITECTURE · API · DECISION_LOG · COMPETITORS (canonical)
                          ROADMAP.md (open/active)
                          archive/ms-previewrs/ (superseded design session — banner-marked)
                          research/ (reserved, not yet created)
fixtures/                real generated .pptx/.xlsx/.docx/.pdf test files
bench/                   visual open-speed harness (real app; baseline renderers + aared compare column, whose card demonstrates the `usePreview` integration)
tools/powerpoint-render/ native PowerPoint slide-export pipeline
tools/excel-render/      native Excel sheet-export pipeline
packages/
  aared/                 core + pptx + xlsx + docx + preview + react + element, one package, subpath exports
  aared-fonts/           OFL fonts + default font resolver
```

**Which docs are authoritative vs. historical:**

- **Standalone continuation:** this file — current status, commands, gate
  ladder, gotchas, open points.
- **Public-facing:** `README.md`.
- **Current technical truth:** the API reference page in the docs site
  (`site/src/content/docs/api.mdx`; `bun run docs:dev`, open `/api`) owns
  public syntax;
  `docs/ARCHITECTURE.md` owns designed/implemented mechanisms;
  `docs/DECISION_LOG.md` records why; `docs/COMPETITORS.md` owns the
  external landscape.
- **Open work:** `docs/ROADMAP.md`.
- **Historical:** `docs/archive/ms-previewrs/` — the evaluation and design
  work that produced this project. Banner-marked, kept for provenance, not
  current truth.

## Gotchas & load-bearing facts

- **Text is never rasterized. Full stop.** See `AGENTS.md § Rules #1`. This
  single rule is why the project exists — re-read
  [`docs/DECISION_LOG.md §1` and `§3`](docs/DECISION_LOG.md) before
  reaching for canvas `fillText` to solve a rendering problem, no matter how
  tempting it looks for one edge case.
- **Moving a directory containing a symlink across drives can silently
  dereference it.** During M0, `docs/archive/ms-previewrs/renderer-bench/`
  contained a symlink (`.zynda`) pointing at an unrelated product repo
  (`D:\work\zynda-main-dev`, including its `.env.local`). A cross-drive `mv`
  of the parent folder dereferenced that symlink and physically copied the
  entire target repo — secrets included — into what was about to become a
  public archive. Caught and deleted before anything was committed or
  pushed. **Lesson:** before any `mv`/`cp` of a directory tree into this
  repo, check for symlinks first (`find <dir> -type l`), and after any
  cross-drive move, verify with `diff -rq <src> <dst>` before deleting the
  source.
- **`fixtures/` are real generated Office files, not synthetic.** Seeded
  from the archived evaluation work — see
  [`docs/archive/ms-previewrs/tests/`](docs/archive/ms-previewrs/tests/)
  for the originals (kept there too; `fixtures/` holds independent copies,
  not the only copies).
- **`core` must stay headless-testable.** No `document`/`window` reference
  inside `packages/aared/src/core/` — that keeps unit tests fast; DOM/SVG
  output is exercised separately with Bun.WebView.
- **Script-generated XLSX must keep the reserved fills slot.** ECMA-376
  reserves `fills[0]` (`none`) and `fills[1]` (`gray125`). A generated
  predecessor of the coverage fixture put its custom dark title fill at
  index 1; Excel "repaired" the file by replacing it with `gray125`, which
  turned the white title text invisible in both Excel and aared (aared drops
  non-solid pattern fills).   The merged `arabic-xlsx-rtl-coverage.xlsx`
  keeps `gray125` at index 1 (verified by inspection: no external links,
  all defined names internal). **Lesson:** fixtures are authored in real
  Office, never hand-generated — aared only ever previews.
- **Sheet RTL orders columns; cell reading order orders text.** The XLSX grid
  sets `dir="rtl"` on the table to mirror column order, but cells without an
  explicit `readingOrder` must not inherit it — Excel treats an unspecified
  order as context. Inheriting sheet RTL visually swapped `180,000 EGP` into
  `EGP 180,000` and `13 Sep 2026` into `Sep 2026 13`. Context-ordered cells
  now render with `dir="auto"`; explicit `readingOrder` 1/2 still maps to
  `ltr`/`rtl`. Pinned by direction assertions in the XLSX visual DOM test.
- **Excel review compares print, not screen.** `tools/excel-render/`
  prints every visible sheet to A4 (hardcoded portrait fit-to-page with
  headings, owned by tooling — not read from the file) via
  `ExportAsFixedFormat`, then rasterizes with PyMuPDF at 150 DPI; the
  capture test fits the aared grid with `headers: true` onto the same A4
  box. An earlier `CopyPicture` revision proved the screen grid instead —
  truer to scrolling, but page-less, window-dependent (hidden Excel
  captures blank), and a different question than the A4 review. Remaining
  known gaps on the print side: the printed page header ("Public") has no
  aared counterpart, print-title rows repeat on native part pages but not
  in aared bands (Shot Log row 4), and the shared chart module draws
  data-driven axes rather than Excel's configured gridlines/ticks.
  Overfull sheets exercise the shrink path via the tracker Dashboard
  (pinned `scale <= 1`); the gate fixture fits at scale 1, pinned by
  assertion. The capture harness hides grid scrollbars (print pages have
  none) and re-renders fresh per part so drawing shift/clip never
  accumulates across parts.
- **The `headers` renderer option is sticky Excel-like chrome.** Row
  numbers skip hidden rows (…6, 8…), letters skip hidden columns (…C, E…),
  headings stick past frozen panes on the inline-start edge (physical right
  in RTL), and virtualization spacers span the extra column. Off by
  default; the scrolling grid is byte-identical without it.
- **A virtualized grid must restore keyboard focus after re-render.** The
  XLSX window renderer rebuilds rows with `replaceChildren()`, so a
  selection move that scrolls destroys the focused cell and focus falls to
  `<body>` — the keydown listener sits *below* the destroyed element, so
  every later arrow key is silently dropped (navigation "works once").
  Selection re-application therefore re-focuses the selected cell (or the
  stable host as fallback) whenever focus fell through, and keydown
  handling re-focuses the new cell after `#ensureVisible`. Pinned by
  `test/xlsx-keyboard.test.ts` (gated on `AARED_VISUAL=1`).
- **Shared build chunks must stay under `dist/core/`.** `WorkerSession`
  starts the worker with `new URL("./worker.mjs", import.meta.url)`, so the
  emitted session chunk must sit *next to* the worker entry. tsdown hoists
  shared chunks to the output root by default, which put the session chunk
  at `dist/session-*.mjs` while the worker stayed at `dist/core/worker.mjs`
  — the URL 404'd and every `open()` failed with "worker failed to start".
  `tsdown.config.ts` pins `outputOptions.chunkFileNames` to `core/…` to keep
  them together. Verify after build changes: the relative `worker.mjs` URL
  must resolve from the session chunk's directory.
- **`usePreview()` mounts on identity, not equality.** `source` and the
  function resolvers are identity-keyed: a fresh `ArrayBuffer` (or a new
  options object, if it held functions) every render remounts the preview
  in a loop — memoize buffers, hold the same `File`/`Blob`/URL across
  renders. Plain-data options (`type`, `index`, `appearance`, `docx`) are
  JSON-keyed inside the hook, so inline literals are safe. Aborted mounts
  never surface as errors; StrictMode mount → cleanup → remount resolves
  exactly once with no leaked worker.

## Open points

Pulled from [`docs/ROADMAP.md`](docs/ROADMAP.md) — that file is
the live version; this is a snapshot for orientation.

- The public route now detects PPTX/XLSX/DOCX in the Worker, returns canonical
  slide/sheet/section models, and dispatches the default viewer to each format.
  It has real-fixture worker coverage and browser-DOM checks for all three
  renderers. PowerPoint remains the acceptance reference for PPTX text;
  LibreOffice is a companion conversion only. Native Excel review for XLSX now
  exists (`tools/excel-render/` pairs per-sheet A4 print pages with aared A4
  fits in the XLSX capture test; needs Python with PyMuPDF next to Excel);
  native Office review for DOCX and remaining fidelity gaps are the next work.
- For current PPTX fidelity, common preset shapes use their authored bounds,
  solid paint, transforms, linear gradients, pictures, and charts.
  Custom/unknown geometry and path gradients remain visibly marked fallbacks;
  pattern/image fills, embedded fonts, and animations are not implemented.
- `tools/powerpoint-render/` provides native PowerPoint review artifacts;
  `tools/excel-render/` provides native Excel A4 print artifacts (per-sheet
  pages paired with aared A4 fits; human-review only, never pixel diffs);
  `bench/` runs the public aared Worker-to-renderer route beside baselines.
- The next phase is a paired visual-inspection pass: inspect the real fixtures
  in `bench/` and native Office/browser artifacts together, then trace every
  meaningful mismatch through OOXML, the Worker model, and rendered DOM/SVG
  before changing code. Record root cause, blast radius, deliberate
  non-changes, and residual limits rather than treating screenshots as local
  CSS defects.
- After inspection fixes land, run a final polish and optimization pass. Keep
  changes measured and scoped to confirmed issues, reconcile behavior across
  PPTX/XLSX/DOCX, then finish with the full build/typecheck/unit/visual gate and
  current package documentation. Two integration follow-ups were explored but
  are proposed, not scheduled: a `workerUrl` override (plus inline-worker
  docs) for restrictive bundlers/CSP setups, and `isAbort`/`messageFor`
  error helpers for consumers.
- No CI yet (`.github/workflows/ci.yml` is a to-do — typecheck + unit tests
  at minimum, once there's something to typecheck against beyond stubs).

## Conventions

- **TypeScript strict, `tsdown` build, `bun:test` for unit tests, `Bun.WebView`
  (Bun 1.4+) for visual-regression screenshots — no Playwright, no separate
  linter/formatter — Changesets for releases.** Revised from the original
  `DESIGN.md` D21 tooling picks after review; not up for casual
  reconsideration mid-milestone.
- **One real fixture beats a pile of mocks.** Prefer testing against actual
  generated Office files over synthetic structures — see `AGENTS.md §
  Testing rule`.
- **Every format capability gets an Arabic fixture before it's "done."** Not
  a nice-to-have; the reason this project exists.

## Maintaining this handoff

Treat this file as a standalone continuation package. Whenever a milestone
lands, update its current state, gate ladder, repo map, gotchas, and open
points in the same change. Someone should be able to read only this file
(and the docs it points to) and pick the project up with no other context.
