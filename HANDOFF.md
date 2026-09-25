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

## Current state: API complete, fidelity + pre-publish gates remain

- **All three formats ship** through format-specific entrypoints
  (`aared/pptx`, `aared/xlsx`, `aared/docx`): `open()` → collections
  (`slides.get()` / `sheets.get()` / `sections.list()` / `headings.list()`)
  → `read()` (where applicable) → `render()` → `destroy()` → `close()`,
  with typed `AaredError` (10 codes), `MediaCollection`, and theme-only
  `Appearance`. `aared/core` is shared types/errors only.
- **Embed layers:** `aared/preview` (`preview()` / `detect()` /
  `PreviewHandle`, one teardown verb), `aared/react` (`usePreview()` +
  `usePptx` / `useXlsx` / `useDocx`), and `<aared-viewer>` (detect, `goTo`,
  zoom, `indexchange` + `slidechange` / `sheetchange`). The bench's
  `AaredViewerCard` is the reference consumer integration.
- **Rendering contract:** Worker parses, main thread paints DOM/SVG (never
  canvas). Content is source-faithful in every theme; dark `theme` drives
  chrome only (DECISION_LOG §25). PPTX master/layout surface built (§17);
  DOCX layout mode paginates structurally, never Word-exact (§10, §22).
  Unsupported geometry/fills render marked fallbacks, never silent
  approximations; deliberate limits live in ROADMAP's fidelity-gaps table.
- **Gates green:** `typecheck`, unit tests, `build` (incl. site + bench).
  What's left is the gate ladder below and
  [`docs/ROADMAP.md`](docs/ROADMAP.md).

## Run it

Commands live in one place: `README.md` § Develop. Artifacts + review pages:
`bun run test:visual:artifacts`. LibreOffice/`pdftoppm` output is optional
and is not a typography oracle.

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
| `bench-measure` (run rarely: a full pass is ~30 min on an idle machine, so it is not in CI or the normal gate; the published numbers are the last full run's) | Download size (`--size`, files fetched from the bench's production build), open / navigation / memory (`--run`, a real browser driving `bench/measure.html`, page and each Worker read in their own DevTools sessions) and search (`--search`) results generate `docs/BENCHMARKS.md` and the home page's numbers (`--render-site`); no hand-filled numbers. `--calibrate` first; re-run on an idle machine; a wedged browser stops the pass and prints the `--run --merge …` command for the rest | ✅ |

## Repo map

```
AGENTS.md               working charter — rules, capability delivery, ground-truth probe
HANDOFF.md              this file
README.md               public pitch
docs/                   ARCHITECTURE · DECISION_LOG · COMPETITORS (canonical) + ROADMAP.md (open work)
                        archive/ (superseded snapshots — banner-marked; the first design
                        session's ms-previewrs/ notes are in git history before 5a32bf0)
fixtures/               real .pptx/.xlsx/.docx files, authored in real Office, never hand-generated
                        (a generated file once hit the reserved gray125 fills slot and Excel "repaired" it)
bench/                  visual open-speed harness (reference usePreview integration: AaredViewerCard)
tools/*-render/         native PowerPoint/Excel/Word review pipelines (+ aared capture)
packages/aared/         core + pptx + xlsx + docx + preview + react + element (one package, subpath exports)
packages/aared-fonts/   OFL fonts + default font resolver
```

Authoritative sources: public syntax → the API reference page
(`site/src/content/docs/api.mdx`; `bun run docs:dev`, `/api`); mechanisms →
`ARCHITECTURE.md`; why → `DECISION_LOG.md`; landscape → `COMPETITORS.md`.
Open work → `ROADMAP.md`. Archive → provenance only.

## Gotchas & load-bearing facts

(Charter rules live in `AGENTS.md` — never-rasterize, headless `core`,
read-only models, canonical import homes. Below is what bit us in practice.)

- **Packaging traps (all bitten before):** shared chunks must emit under
  `dist/core/` (`tsdown.config.ts` pins it) or the worker URL 404s; the
  worker constructor must stay literally
  `new Worker(new URL("./worker.mjs", import.meta.url), { type: "module" })`
  or Vite verbatim-copies it and `open()` hangs (§19); a dead worker
  reports `worker-error`, never `aborted` (§19); `test/worker.test.ts` and
  `tools/check-consumer-build.ts` need a prior `bun run build`.
- **Bench serves `dist`, not source.** After any build, restart the bench
  dev server + hard-refresh; one server on the known port, never parallel
  auto-port servers.
- **The Worker is pooled.** A closed document's clean Worker serves the next
  `open()` (and a facade's detect-then-open share one); it stops after 15 s
  idle (`session.ts`, `acquireBridge`, DECISION_LOG §38). Tests that swap
  `globalThis.Worker` call `drainWorkerPool()` in `finally` (a parked Worker
  never serves a different constructor, but a parked fake lingers
  otherwise). Anything the Worker keeps at module level must stay bounded
  and own its strings.
- **Renderer defaults are classes** (`src/class-sheet.ts`, §36): inline
  style is only what a node authors, so renderer tests read defaults with
  `getComputedStyle`, never `el.style`. XLSX: an empty cell has no text span,
  no `dir`, no alignment; a row's fill columns are one cell
  (`data-aared-fill="tracks"`) and a fill row one cell
  (`data-aared-fill="row"`) with gradient gridlines. DOCX: no
  `paragraph-text` wrapper; layout pages are `content-visibility: auto` and
  lay out their tabs when first shown (`settlePageTabStops`, §37).
- **Replacing a unit must not move the reader.** A slide goes into its
  container at its full-width box before it is fitted, and `measureFitRoom`
  restores scroll positions after setting content aside (§40). Any new
  renderer that swaps content in a container that sizes to its content needs
  the same care.
- **Worker sessions keep nothing the page holds.** A string sliced from a
  part keeps the whole part text alive (V8/JSC), so anything a session or a
  module-level table keeps past a parse goes through `ownString` / `ownTree`
  (`core/xml.ts`). DOCX streams `w:body` and caches no sections (§35).
- **Shipped JS is Latin-1.** No character above U+00FF in a shipped string,
  template or regex literal (a raw one or a `\u` escape — Vite's esbuild
  prints escapes back raw): one would make V8 hold the whole Worker script
  two bytes per character. Use `src/core/chars.ts`; messages use plain
  ASCII punctuation. Guarded by `test/shipped-latin1.test.ts` (needs a build)
  and `tools/check-consumer-build.ts`.
- **`Bun.WebView.cdp()` only talks to its page.** A `sessionId` passed in its
  params is ignored, so a "Worker" reading through it is the page again (the
  bug DECISION_LOG §33 fixed). Read anything outside the page through
  `tools/cdp.ts` (browser-level socket, flat sessions, top-level
  `sessionId`), and run `--calibrate` after touching it.
- **Geometry is deferred, not shipped in the worker.** The flat worker
  carries only the 11 hand-coded presets; everything else resolves on the
  main thread against `presetPrograms` (`src/core/presets.ts`, generated by
  `scripts/compile-presets.ts`). The table must stay in the `aared/pptx`
  entry only: `pptx/index.ts` imports it and passes it to its session as
  `geometryPrograms` — never import it from `core/session.ts`, which every
  format bundles (DECISION_LOG §26 addendum). Search reads pass
  `fonts: false, geometry: false`, so they never evaluate outlines.
- **Every PPTX fill goes through `src/pptx/fill-paint.ts`.** Shapes, table
  cells and the slide background share it; never paint a fill kind in only
  one of them. A fill it cannot paint returns `undefined` and the caller must
  mark it (dashed in dev) — silent `fill="none"` is the bug DECISION_LOG §28
  fixed. Gradient stops go through `officeGradientStops` (linear light,
  smoothstep) everywhere, WordArt included.
- **Fallback markers and the XLSX notice bar are dev-only.** One flag,
  `MARK_FALLBACKS` in `src/diagnostics.ts`: `appendShapeGeometry` paints
  the dashed marker and `paintSheet` the sheet notice bar only when running
  from source (unit/visual tests, harnesses); the shipped bundles define
  `__AARED_MARK_FALLBACKS__` as `false` via tsdown, so builds like the bench
  render fallback geometry as an invisible box and sheets with no notice
  bar. Shapes with no authored geometry at all (`missing` — text-only
  placeholders) are never marked anywhere. Check from built output (dist is
  minified, so the flag's name is gone once it folds to `false`):
  `worker.mjs` contains no preset names, the table appears only in
  `dist/pptx/index.mjs`, and no dist `.mjs` (the `.map`s quote the source) contains the dev-only notice text
  ("Formulas show the values saved by Excel").
- **Freeze the fixture before comparing screenshots.** `Ctrl+S`, lockfile
  gone, re-verify `pgMar`/key paragraphs — unsaved Word state produces
  ghost discrepancies.
- **OOXML ≠ CSS vocabulary.** `w:vertAlign` `superscript` maps to
  `super`/`sub`; browsers silently drop the rest. Pin with
  computed-style assertions, not eyeballing.
- **Line-box clipping is a metrics defect.** DrawingML `0.9` stays in the
  model, the renderer clamps browser line-height to `1.1`; inset precedence
  is nearest-present master → layout → slide
  (`tools/probe/title-insets/`). Never disable clipping globally.
- **XLSX specifics:** context-ordered cells render `dir="auto"` (sheet
  `dir="rtl"` mirrors columns only — inheriting it reordered `180,000 EGP`
  and `13 Sep 2026`); `headers` chrome is sticky/exclusion-aware and off by
  default; virtualized re-renders must restore focus to the selected cell
  (`xlsx-keyboard.test.ts`); Excel review compares A4 **print**
  (`ExportAsFixedFormat` + PyMuPDF, `headers: true`, `fill: false`), never
  screen — print-title rows and the "Public" page header have no aared
  counterpart. The grid paints its first window in the same task once it is
  attached (a frame later showed a short sheet without its fill, then grew
  it), and a `max-height` clipper's room is its limit, not its current
  height (§40). Fill lines come from `syncFillLines()` custom properties: it
  must run before the first window, on scale and on `resize` (dpr).
- **DOCX specifics:** `alignFor` swaps `left`/`right` in RTL paragraphs; a
  header hangs from its `w:header` distance and is never clipped; line boxes
  follow run sizes (`lineFontSizePt`) at ~1.15em (`WORD_SINGLE_LINE`);
  first-block spacing drops after breaks (`dropsTopSpacing`); shared
  line-box constants live in `core/docx/pagination.ts`, the floor + Arabic
  test in `core/typography.ts`; one `MediaUrls` owner per render, one Blob
  URL per part.
- **Bench CSS resets bite.** Tailwind's `img { display: block }` broke
  inline images — state `display` explicitly and check new inline elements
  in the bench, not just the harness. (`usePreview()` mount identity rules
  live in the API reference, React hooks section.)

## Publishing and how development goes after it (DECISION_LOG §42)

**Where things stand.** `@aared/viewer@0.0.0` on npm is a name-holding
placeholder (README only; org `aared`, owner `mego0eid`). The repo's packages
are still named `aared` / `aared-fonts`, `private: true`, `0.0.0`; git tag
`v0.0.0` marks the reservation. Nothing of the real code is public.

**Before the first real release (`0.1.0`), in one piece of work:**

1. Resolve the "Before first publish" list in ROADMAP (model-contract tier,
   the rest of the list).
2. Rename: `aared` → `@aared/viewer`, `aared-fonts` → `@aared/fonts` (reserve
   the name first with a placeholder: one 2-file package, dry-run, then the
   maintainer publishes). Imports become `@aared/viewer/pptx`, `/xlsx`,
   `/docx`, `/preview`, `/react`, `/element`, `/worker`. Touch:
   both `package.json` names and `exports`, every import in `bench/`,
   `site/`, `tools/` and the tests, `tsdown.config.ts`,
   `tools/check-consumer-build.ts`, the API reference page (docs-parity), the
   README install block, the site install card and playground snippets. The
   `<aared-viewer>` tag name does not change.
3. Drop `private`, add `publishConfig.access: "public"` and a `files` list;
   run `bun run build`, unit + visual tests, `bun tools/check-consumer-build.ts`
   and `npm pack --dry-run` (check the tarball has `dist/`, no fixtures or
   sources you did not intend).
4. The maintainer, in their own terminal (2FA on; no token is ever given to an
   agent): `npm publish`. Then `git tag -a v0.1.0` on the release commit.

**After it (the rules change):**

- **No more free breaking changes.** AGENTS.md "Pre-release breaking changes"
  ends at `0.1.0`. In `0.x`, a breaking change is a **minor** bump with a
  migration note (in the release note and the API page); fixes and additions
  are patch/minor. `1.0.0` only when the API is frozen and the ROADMAP gates
  (PPTX visual review accepted, model contract decided) are closed.
- **Releases are one commit + one annotated tag `vX.Y.Z`**, made after the
  publish succeeds; the version in `package.json` matches the tag. Tags are
  pushed once a remote exists.
- **Published versions are immutable.** A bad release is fixed with a new
  patch (`npm deprecate` the bad one); never plan on unpublishing.
- **Fonts** (`@aared/fonts`) versions independently of the viewer.
- Keep docs in step with each release: API page, README, ROADMAP status,
  DECISION_LOG entry for anything that overturned a documented decision.

## Open points

[`docs/ROADMAP.md`](docs/ROADMAP.md) is the live list; shortest version:

1. **Last two gradient kinds** — off-center `path="circle"` and
   `path="shape"` gradients still render as marked fallbacks. Add one of
   each by hand on a new slide at the end of `pptx-coverage.pptx` (COM
   cannot author them; give the slide a unique Arabic title and a manifest
   entry), then probe PowerPoint's render before painting them
   (DECISION_LOG §28).
2. **Final polish pass** — confirmed defects only, then the full
   build/typecheck/unit/visual gate + docs.
3. **Pre-publish decisions** — resolved-model contract tier; bundle size
   settled (single worker).
4. **Backlog** — decorations API, search highlighting, headless thumbnails,
   embedded fonts (M4), the first real npm release (steps above).
5. **DOCX residuals** — measured pagination instead of the estimator
   (DECISION_LOG §21; estimator stays for headless), page-3 drift
   (per-font single-line metrics, block-image spacing), promote the
   anchor-comparison script into `tools/word-render`.

## Fonts — residuals only (consumer note lives in `README.md` § Fonts)

- **Embedded fonts are not implemented** (PPTX `*.fntdata`, DOCX
  `*.odttf`; ROADMAP M4). XLSX expected not applicable — unconfirmed.
- **Unverified:** substitutes register under their own names while
  renderers paint the requested name, so they may never be selected where
  the requested font is missing. Check before relying on it.

## Maintaining this handoff

Distill landed work away; keep live status, gates, gotchas, and open
points. Someone should be able to read only this file (and the docs it
points to) and pick the project up with no other context.
