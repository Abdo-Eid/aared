# AGENTS.md

> ⚠ **Development status: pre-release, unpublished (`private: true`, v0.0.0).**
> PPTX implementation is in progress; the Worker-backed parser, focused
> DOM/SVG renderer, and end-to-end viewer route exist. See
> [`docs/ROADMAP.md`](docs/ROADMAP.md) for status and open work, and
> [`docs/DECISION_LOG.md`](docs/DECISION_LOG.md) for what landed.

## What this repo is

`aared` (عارض) is a family of packages that preview real `.pptx`, `.xlsx`
and `.docx` files client-side, in the browser, correctly for Arabic/RTL
content — **view-only**, never editing. One engine:

- **Parse** in a Worker: OPC (zip + rels + content types) → DrawingML/
  styles/theme → a resolved, structured-cloneable model.
- **Render** on the main thread: paint that model to real DOM/SVG — never
  canvas, never rasterized text. This is the load-bearing decision; see
  [`docs/DECISION_LOG.md §3`](docs/DECISION_LOG.md#3-text-is-rendered-by-the-browser-never-rasterized--the-load-bearing-decision).

```text
   .pptx/.xlsx/.docx  ──parse (worker)──►  resolved UnitModel  ──render──►  DOM/SVG
                                                                              (never canvas)
```

Full detail: [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md).

## Environment & commands

Bun required. TypeScript, `tsdown` build (no plain-JS shortcut — this is a
published package, types ship for real).

```bash
bun install            # workspace install
bun run typecheck       # tsc --noEmit across all packages
bun test                # bun:test — unit tests
bun run test:visual     # bun:test, visual-regression suite (uses Bun.WebView)
bun run build            # tsdown, the two packages (bench and the site build on their own: bench's vite build, docs:build)
```

**Windows agent automation:** on this Codex setup, Word COM export and
`Bun.WebView` capture succeed with approved execution outside the agent sandbox
but fail inside it (`E_ACCESSDENIED` / `Chrome process closed the pipe`). Use
the command runner's scoped escalation for native Word/browser automation;
rasterization and token comparison can run inside the sandbox. This workflow
was verified on October 6, 2026. A sandbox launch failure does not establish a
renderer defect or require Bash; see `tools/word-render/README.md`.

## Package map

| Package | Role | Deps |
|---|---|---|
| `aared` (subpath `core`, plus `worker`) | OPC + XML + DrawingML + theme + worker RPC; `worker` is the Worker script | none (`DecompressionStream` does the unzip) |
| `aared` (subpaths `pptx`/`xlsx`/`docx`) | Per-format renderers, each consuming `core`'s resolved model | none beyond `core` |
| `aared` (subpaths `preview`/`react`/`element`) | One-call embed, React hooks, `<aared-viewer>` web component | lazy-`import()` the format subpaths |
| `aared-fonts` | A Google-Fonts-backed font resolver with calibrated substitutes (no font files ship). Frozen and unpublished (DECISION_LOG §53): the default is the browser's own fonts | none |

Renderers reach `core` only through the resolved `UnitModel` it emits —
never by parsing XML themselves. See
[`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md#package-layout).

```text
AGENTS.md               this charter: rules, commands, load-bearing facts
README.md               public pitch, install, quick start, Develop
docs/                   ARCHITECTURE · DECISION_LOG · COMPETITORS (canonical) +
                        ROADMAP.md (status, gates, open work, release checklist);
                        archive/ (superseded snapshots, banner-marked)
fixtures/               real .pptx/.xlsx/.docx authored in real Office, never
                        hand-generated; README.md has the authoring rules
bench/                  visual open-speed harness; AaredViewerCard is the
                        reference usePreview() integration
tools/*-render/         native PowerPoint/Excel/Word review pipelines + aared capture
packages/aared/         core + pptx + xlsx + docx + preview + react + element
packages/aared-fonts/   font resolver with calibrated substitutes (frozen, unpublished)
site/                   Astro docs site (the API reference lives here)
```

## API surface

Full reference: the API reference page in the docs site
(`site/src/content/docs/api.mdx`; `bun run docs:dev`, open `/api`).
Kept honest by the
`docs-parity` test (see Testing rule, below): every public export must be
named there.

## Rules (do not break these)

1. **Text is rendered by the browser — DOM or SVG — never rasterized by
   `core` or a format renderer.** This is the one rule everything else in
   this project follows from. See
   [`docs/DECISION_LOG.md §3`](docs/DECISION_LOG.md#3-text-is-rendered-by-the-browser-never-rasterized--the-load-bearing-decision).
   If you find yourself reaching for `<canvas>` `fillText` or a WASM text
   shaper (a WASM decoder for EMF/WMF images would be the one acceptable
   exception, and none exists yet), stop — that's the exact failure mode that
   motivated this project.
2. **XML/ZIP parsing never touches the main thread.** `core` runs inside a
   Worker; the main thread only receives resolved models and paints. See
   [`docs/ARCHITECTURE.md § Performance model`](docs/ARCHITECTURE.md#performance-model).
3. **`core` has zero DOM dependency.** It must stay testable headless
   (plain Bun, no browser) so unit tests don't require spinning up a
   `Bun.WebView`. DOM-only code belongs in the format renderers or the
   element.
4. **Every format module ships at least one Arabic fixture with a table**
   before a milestone touching it is considered done — not optional polish.
   This is the specific gap the project exists to close; see
   [`docs/DECISION_LOG.md §1`](docs/DECISION_LOG.md#1-why-this-project-exists-at-all).
5. **View-only means view-only.** No formula recalculation, no
   round-tripping, no write path — ever. If a feature request implies
   editing, it's out of scope by design, not by omission; see
   [`docs/DECISION_LOG.md §4`](docs/DECISION_LOG.md#4-view-only-permanently--not-a-phased-in-decision).
6. When stuck on a real-world OOXML edge case, read
   [Ground-truth probe](#ground-truth-probe) below before guessing from the
   ECMA-376 spec text alone — the spec and what Office actually writes
   diverge more often than you'd expect.

## Load-bearing code facts

- **Resolved models are read-only.** XLSX cells sharing a style id share one
  style object (same for shared hyperlinks; structured clone preserves it) —
  clone before mutating. Parsed XML trees are frozen the same way
  (`sheetData`/`sharedStrings` stream rather than materialize).
- **Shipped JS stays Latin-1.** A character above U+00FF in any shipped
  literal makes V8 hold the whole script (the Worker's included) two bytes
  per character; the few aared needs come from `src/core/chars.ts`. A `\u`
  escape does not survive the build, and rolldown folds
  `String.fromCharCode(…)` of a constant into the raw character (it did,
  DECISION_LOG §34). Messages use plain ASCII punctuation. Guarded by
  `test/shipped-latin1.test.ts` (reads `dist`, so build first) and
  `tools/check-consumer-build.ts`.
- **Imports have a canonical home:** abort helpers from `core/errors.ts`
  (not `zip.ts`); `FontResolver` from `core/index.ts`. The Worker reaches
  format parsers only via dynamic `import()` — never static.
- **Packaging traps, all hit before:** shared chunks must emit under
  `dist/core/` (`tsdown.config.ts` pins it) or the Worker URL 404s; the
  Worker constructor stays literally
  `new Worker(new URL("./worker.mjs", import.meta.url), { type: "module" })`
  or Vite copies it verbatim and `open()` hangs (`tools/check-consumer-build.ts`
  builds a consumer with Vite and webpack; Next.js was checked by hand, §56);
  a dead Worker reports
  `worker-error`, never `aborted` (DECISION_LOG §19). `test/worker.test.ts`,
  `tools/check-consumer-build.ts`, the bench and the site all read `dist`
  (the site and bench through the `exports` map, with no source fallback),
  so build first; the deploy workflow does the same. After a build, restart
  the bench dev server and hard-refresh — one server on the known port.
- **The Worker is pooled.** A closed document's clean Worker serves the next
  `open()`; it stops after 15 s idle (`acquireBridge`, DECISION_LOG §38).
  Tests that swap `globalThis.Worker` call `drainWorkerPool()` in `finally`.
  Anything a Worker session or module-level table keeps past a parse is
  bounded and copied with `ownString`/`ownTree` (`core/xml.ts`): a string
  sliced from a part keeps the whole part text alive (§35).
- **Geometry is deferred, not shipped in the Worker.** The Worker carries the
  12 hand-coded presets; the rest resolve on the main thread against
  `presetPrograms` (`src/core/presets.ts`, generated by
  `scripts/compile-presets.ts`). That table stays in the `aared/pptx` entry
  (passed to its session as `geometryPrograms`), never imported from
  `core/session.ts`, which every format bundles (§26).
- **Search reads words, never models.** PPTX and XLSX search ask the Worker
  for a unit's text (`session.text()`, the `text` request): a slide's words
  come from `readPptxSlideText`, a walk of the slide XML that must equal the
  full model's paragraph text (`test/pptx-text.test.ts` checks every fixture
  slide); a sheet's from `sheetText`. A change to what a slide paints as text
  (a new element kind, a run rule) changes that walk too (§80).
- **Every PPTX fill, and every chart fill in any format, goes through
  `src/fill-paint.ts`** — shapes, table cells, the background and charts
  alike. A fill it cannot paint returns `undefined`
  and the caller marks it; a silent `fill="none"` is the bug §28 fixed.
  Gradient stops go through `officeGradientStops` everywhere, WordArt too.
- **Development diagnostics are one flag.** `MARK_FALLBACKS`
  (`src/diagnostics.ts`) paints the dashed fallback outline and the XLSX
  notice bar only from source; shipped bundles define it `false`. Shapes with
  no authored geometry (`missing`, text-only placeholders) are never marked.
  Check built output, not names (dist is minified): `worker.mjs` holds no
  preset names, the preset table appears only in the PPTX renderer chunk, and
  no dist `.mjs` contains the notice text ("Formulas show the values saved
  by Excel").
- **Renderer defaults are classes** (`src/class-sheet.ts`, §36): inline style
  is only what a node authors, so tests read defaults with
  `getComputedStyle`, never `el.style`.
- **Replacing a unit must not move the reader.** A slide goes in at its
  full-width box before it is fitted, and `measureFitRoom` restores scroll
  positions after setting content aside (§40). A new renderer that swaps
  content inside a content-sized container needs the same care.
- **Delivered fonts load when painted** (§50). Anything that measures painted
  text or reads which font painted calls `awaitDocumentFonts(model)` first;
  before it, a substitute has not been fetched at all.
- **`Bun.WebView.cdp()` only talks to its page.** A `sessionId` in its params
  is ignored, so a "Worker" read through it is the page again (§33). Read
  anything else through `tools/cdp.ts` and run `--calibrate` after touching
  it.
- **Host CSS resets bite.** The bench's Tailwind `img { display: block }`
  broke inline images: state `display` explicitly, and check new inline
  elements in the bench, not only the test harness.

## Capability delivery

**Good enough beats perfect.** The target is what a normal reader notices,
not the last pixel or decimal. Once a capability works at a normal-use
level, stop: record the cosmetic residue in `docs/ARCHITECTURE.md` § Known
limits instead of
chasing it, don't refit or restructure for sub-percent or rare cases, and
when the effort starts outgrowing what a reader would see, say so and
propose a stopping point. This never excuses sloppy shared code or a
half-usable capability (point 2 below); it limits polish, not correctness.

"Good enough" covers only differences a reader would not notice: a pixel, a
rounding, a rare edge case. It never covers a difference a reader sees at a
glance. A layout Office shows that aared shows differently (text on one side
of a picture instead of both, a picture on its own band instead of beside
the text) is a bug to fix, not residue to record. "CSS can't do it" is not a
reason to stop: build it in the renderer with browser text (rule 1), or ask
before accepting the difference.

"One thing at a time" means one coherent, *usable* capability — not the
smallest possible diff. A capability may span `core` parsing, a format
renderer, the element, and fixtures/tests when those pieces are required for
a normal usable result.

1. Before implementing, state the complete capability: what's missing today,
   what a consumer should be able to do afterward, which layers it touches,
   and how it'll be proven.
2. Deliver to a practical normal-use level. Don't call a thin primitive
   "done" while deferring the obvious pieces that make it actually usable
   (e.g. a PPTX shape renderer that ignores rotation isn't a finished shape
   renderer).
3. "Smallest correct change" means avoid unrelated work and unneeded
   abstraction — it does not mean narrowing the requested capability.
4. **Route completeness:** assess every capability against every format it
   touches. A format may be marked "unchanged" for a capability only when
   that's actually true (shared `core` machinery covers it) — document and
   test that reuse rather than silently skipping the format.
5. If work lands as a foundation rather than the complete capability, say so
   explicitly in `docs/ROADMAP.md` — don't record it as finished.

## Root-cause discipline

Do not treat a failing screenshot, test, or bug report as a local patching
exercise. Before changing code, trace the source OOXML, resolved Worker model,
renderer DOM, embedding styles, and native Office output needed to explain the
behavior. Identify the root cause, not merely the clipped glyph or wrong pixel.

1. Inspect the relevant architecture and surrounding code; trace data and
   control flow end to end. Search for shared helpers, defaults, inheritance,
   duplicated logic, and host-environment coupling that can produce the same
   class of failure elsewhere. The shared homes (units, colour, theme fonts,
   appearance tokens, media URLs, text and table painting, …) are listed in
   [`docs/ARCHITECTURE.md § Shared modules`](docs/ARCHITECTURE.md#shared-modules):
   put a rule two formats need there once, and pass what differs as a
   parameter, instead of writing a second copy.
2. Assess blast radius before choosing a change. Check every caller and format
   route the shared code affects, preserve public contracts unless a documented
   change is necessary, and add coverage for related edge cases rather than
   only the reported fixture.
3. Prefer a focused structural correction over a call-site workaround when it
   materially improves correctness, consistency, or maintainability. Do not
   refactor unrelated code or introduce abstractions without a demonstrated
   need.
4. Validate the original scenario, the relevant embedding context, and the
   affected shared paths. Renderer tests must re-read actual DOM/SVG output;
   visual fidelity questions require the paired native PowerPoint review. When
   the package and Office disagree, use the Ground-truth probe rather than
   inventing a CSS approximation.
5. Before declaring completion, record the root cause, blast radius, deliberate
   non-changes, and remaining limits in the doc that owns each: DECISION_LOG
   for why, ARCHITECTURE for how it works now and its accepted limits,
   ROADMAP for status and the gaps still open.

## Commits

- **Batched, never atomic.** One commit per completed piece of work (code +
  tests + docs together) — never one commit per file, and keep the count
  low.
- **Messages never name milestones or steps.** Describe the change itself,
  never `M1` / `step N` labels.
- **Code comments, test names, and user-visible errors never name milestones
  or steps either.** Describe the capability or its current limitation;
  `docs/ROADMAP.md` owns sequencing labels.
- **Only commit when explicitly asked.** Before staging, always inspect
  `git status` and stage only the intended files — never sweep in someone
  else's in-progress work.
- **"Commit" means one commit for all the work.** When the user says commit,
  put everything achieved up to that point in a single commit — never split
  into multiple commits or stage by pieces without being asked.

## Authoring API design

The public API serves both application developers and AI agents building on
top of it. Prefer dedicated, stateless functions; plain resolved data
(`UnitModel`, not a live class graph you can't `JSON.stringify` or pass
across a worker boundary); one canonical shape per concept, not multiple
overloads; explicit units (inches, matching Office's own convention) and
defaults; structured, actionable errors over silent `undefined`. See
the API reference page (`site/src/content/docs/api.mdx`) for the shape this
produces.

## Pre-release breaking changes

(The npm placeholder `@aared/viewer@0.0.0` only holds the name and does not
end this stance; the first real release, `0.1.0`, does. The release
checklist is in [`docs/ROADMAP.md`](docs/ROADMAP.md) "Before first publish";
DECISION_LOG §42.)

While `aared` is pre-release (`private: true`, v0.0.0, unpublished), **best
design wins over backward compatibility.** Prefer the single canonical shape
— one async object-form `render({ <unit>, target, ... })` per format, one
teardown story, one error model — over preserving a signature that shipped
days ago. Never add a compatibility alias or overload to avoid breaking an
in-development caller; update the callers (package, element, bench, tests)
and the docs (the API reference page, `README.md`) in the same piece of work, and
record the overturned decision in
[`docs/DECISION_LOG.md`](docs/DECISION_LOG.md) with its rationale. This
stance expires at first publish: once a version is public, breaking changes
need a migration note, not just a log entry.

**After `0.1.0` the rules change:**

- In `0.x` a breaking change is a **minor** bump with a migration note (in
  the release note and the API page); fixes and additions are patch or
  minor. `1.0.0` only once the API is frozen and the ROADMAP gates (PPTX
  visual review accepted, model contract decided) are closed.
- A release is one commit plus one annotated tag `vX.Y.Z`, made after the
  publish succeeds; `package.json`'s version matches the tag.
- Published versions are immutable: fix a bad one with a new patch and
  `npm deprecate` it; never plan on unpublishing.
- The maintainer publishes from their own terminal with 2FA; no npm token is
  ever given to an agent.
- Each release keeps the docs in step: API page, README, ROADMAP status, and
  a DECISION_LOG entry for anything that overturned a documented decision.

## Testing rule

Every renderer capability ships a test that **re-reads rendered output from
the DOM/model it produced** — never trusts an in-memory intermediate. Use
real fixtures (`fixtures/`), not synthetic mocks, for anything format-shape
related; synthetic fixtures are fine only for pure-function edge cases
(e.g. a number-format string parser).

`docs-parity` (in `bun test`): every public export, type, and enum must be named in the API reference page
(`site/src/content/docs/api.mdx`), and a fixed list of the callables the page
documents must appear in it. The test does not check the reverse (a name left
on the page after its export is removed), so removing an export means editing
the page by hand. A runnable example on that page, or on
the Start page (`site/src/content/docs/start.mdx`), is fenced ` ```ts check `
(or `tsx check`) and must paste whole (its own imports; `file`, `container`,
`el` are given): `test/api-snippets.test.ts` compiles every such block, and
every playground snippet, against the source under the repo's strict
settings. Signature listings stay unmarked. Text both pages need lives once
in `site/src/partials/` and is imported (the "Choosing a layer" table).

Docs-only changes (markdown/site text, no code) skip verification runs —
no typecheck, tests, or build — except a change to a checked example on the
API or Start page, which runs `bun test packages/aared/test/api-snippets.test.ts`.

**Visual review:** a `bun:test` suite drives `Bun.WebView` (built into Bun
1.4+ — no Playwright/Puppeteer dependency) to load aared's DOM/SVG render and
call `.screenshot()`. `tools/powerpoint-render/` exports the same fixture with
native PowerPoint, at the same dimensions; `tools/excel-render/` prints every
visible sheet to A4 with native Excel and rasterizes with PyMuPDF, paired
against an aared A4 fit (scaled to equal display width in review);
`tools/word-render/` exports a DOCX through native Word to per-page PNGs plus
the exact position and size of every text span (read layout numbers from
that, not from a screenshot), and its `capture.ts` renders the same pages
through aared for side-by-side review. Inspect
the paired captures and `review.html` from `bun run test:visual:artifacts`
for typography, bidi placement, and wrapping. This machine has Office's
fonts, so those captures never show the `aared-fonts` substitutes; for any
font change add `--absent-fonts` (and `capture.ts --absent-fonts`), which
renders each unit as a device without Office's fonts would, as a third
column (network). LibreOffice conversion, when
useful, is never a PPTX or XLSX oracle because it substitutes fonts. Plain
unit tests do not require Office; the capture tests do (PowerPoint/Excel on
Windows, plus Python with PyMuPDF for the Excel route; Word and PyMuPDF for
`tools/word-render/`, which is run by hand, not by a test).

## Ground-truth probe

For a rendering bug where it's unclear what a piece of real-world OOXML
actually means — the spec text and what Office actually writes/expects
diverge more than the docs let on — **reverse-engineer it by interacting
with real Office**, adapted from shariha's repair-probe workflow (there,
used to diagnose PowerPoint "Repair" prompts on export; here, used to
understand what a feature *is* on read, since aared never writes):

1. Save the confusing fixture into a gitignored, **nested** git repo at
   `tools/probe/<topic>/` (not the main repo — keep the noise contained).
2. Unzip every part and commit that snapshot ("before").
3. Open the file in the real application (PowerPoint/Excel/Word), make the
   specific interaction that's in question (toggle RTL on a paragraph,
   apply a conditional format, whatever's ambiguous), save.
4. Re-unzip over the same directory and `git diff`. Ignore incidental
   Office save noise (`docProps/custom.xml`, thumbnails, whitespace);
   focus on the XML that changed for the interaction you made.
5. That diff is ground truth. Codify the finding as a comment at the parse
   site referencing the probe topic, and add a fixture + test that pins the
   behavior — don't leave the discovery only in an uncommitted probe
   directory.
6. Clean up after finishing: once the finding is codified (comment +
   fixture + test), delete `tools/probe/<topic>/`. The comment and the test
   are the durable record, not the directory — write the comment so it
   stands alone (what was done in Office, what changed in the XML), never
   as a bare pointer to a directory that no longer exists. One finished
   probe must never linger as dead weight next to the next one.

The probe root is scratch evidence, not a deliverable — don't move, reset,
or delete another contributor's in-progress probe without checking with
them first. A probe counts as finished only when step 5 is done; until
then it stays. A finished probe has no value of its own — the comment,
fixture, and test reproduce its result without the snapshot trees — so an
ignored probe directory is safe to lose (fresh clone, folder delete).
Only an unfinished probe matters. New probes follow the standard shape in
`tools/probe/README.md`: one role script (`probe.py`, `author.py`, or
`extend.py`), a README, and shared helpers imported from
`tools/probe/_common/` — never a second copy of the COM or temp-copy
boilerplate. (The font bench lives at `tools/font-calibration/`, not
under `probe/`.)

## Docs map

Start here, then read what a task touches. There is no separate handoff
document: each fact lives in the one doc that owns it.

- `README.md` — public pitch, install, quick start, Develop (commands)
- `site/src/content/docs/api.mdx` — designed public API and syntax (docs site)
- `docs/ARCHITECTURE.md` — how the engine is designed to work, including the
  Office behaviours found the hard way ("Format rules") and the accepted
  differences ("Known limits")
- `docs/DECISION_LOG.md` — why, chronologically
- `docs/COMPETITORS.md` — the external landscape
- `docs/ROADMAP.md` — open work only: status, the open gate, what's next, open
  gaps, the release checklist
- `fixtures/README.md` — what each fixture covers, and how to author one
- `docs/archive/` — superseded snapshots, banner-marked; provenance only
- `docs/research/` — not created yet; reserved for future unchecked
  audits/proposals, processed point-by-point into the canonical docs and
  never moved there wholesale (same lifecycle shariha uses for its own
  `docs/research/`)
