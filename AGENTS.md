# AGENTS.md

> ⚠ **Development status: pre-release, unpublished (`private: true`, v0.0.0).**
> PPTX implementation is in progress; the Worker-backed parser, focused
> DOM/SVG renderer, and end-to-end viewer route exist. See
> [`docs/ROADMAP.md`](docs/ROADMAP.md) for what's actually built.

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
bun run build            # tsdown, all packages
```

## Package map

| Package | Role | Deps |
|---|---|---|
| `aared` (subpath `core`) | OPC + XML + DrawingML + theme + worker RPC | jszip-equivalent only if `DecompressionStream` isn't enough; otherwise none |
| `aared` (subpaths `pptx`/`xlsx`/`docx`) | Per-format renderers, each consuming `core`'s resolved model | none beyond `core` |
| `aared` (subpath `element`) | `<aared-viewer>` web component | lazy-`import()`s the format subpaths |
| `aared-fonts` | OFL font files + CDN-backed default font resolver | none |

Renderers reach `core` only through the resolved `UnitModel` it emits —
never by parsing XML themselves. See
[`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md#package-layout).

## API surface

Full reference: the API reference page in the docs site
(`site/src/content/docs/api.mdx`; `bun run docs:dev`, open `/api`).
Kept honest by a
`docs-parity` test once real exports exist (see Testing rule, below) — every
public export must be named there.

## Rules (do not break these)

1. **Text is rendered by the browser — DOM or SVG — never rasterized by
   `core` or a format renderer.** This is the one rule everything else in
   this project follows from. See
   [`docs/DECISION_LOG.md §3`](docs/DECISION_LOG.md#3-text-is-rendered-by-the-browser-never-rasterized--the-load-bearing-decision).
   If you find yourself reaching for `<canvas>` `fillText` or a WASM text
   shaper for anything other than EMF/WMF image decoding, stop — that's the
   exact failure mode that motivated this project.
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
  per character; the few aared needs come from `src/core/chars.ts`.
- **Imports have a canonical home:** abort helpers from `core/errors.ts`
  (not `zip.ts`); `FontResolver` from `core/index.ts`. The Worker reaches
  format parsers only via dynamic `import()` — never static.

## Capability delivery

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
   non-changes, and remaining limits in the appropriate docs and handoff.

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
end this stance; the first real release, `0.1.0`, does. See
[`HANDOFF.md`](HANDOFF.md) "Publishing" and DECISION_LOG §42.)

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

## Testing rule

Every renderer capability ships a test that **re-reads rendered output from
the DOM/model it produced** — never trusts an in-memory intermediate. Use
real fixtures (`fixtures/`), not synthetic mocks, for anything format-shape
related; synthetic fixtures are fine only for pure-function edge cases
(e.g. a number-format string parser).

`docs-parity` (once `core`/`pptx`/`xlsx`/`docx`/`element` have real exports):
every public export, type, and enum must be named in the API reference page
(`site/src/content/docs/api.mdx`); every
callable identifier named there must exist as a real export. Wire
this in as part of M1, not deferred indefinitely — see
[`docs/ROADMAP.md`](docs/ROADMAP.md).

Docs-only changes (markdown/site text, no code) skip verification runs —
no typecheck, tests, or build.

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
for typography, bidi placement, and wrapping. LibreOffice conversion, when
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

The probe root is scratch evidence, not a deliverable — don't move, reset,
or delete another contributor's in-progress probe without checking with
them first.

## Docs map

- `README.md` — public pitch + quick start (marked pre-release until M1 ships)
- `HANDOFF.md` — contributor onboarding, current status, gotchas
- `site/src/content/docs/api.mdx` — designed public API and syntax (docs site; `docs/API.md` is a pointer stub)
- `docs/ARCHITECTURE.md` — how the engine is designed to work
- `docs/DECISION_LOG.md` — why, chronologically
- `docs/COMPETITORS.md` — the external landscape
- `docs/ROADMAP.md` — queued milestones and current status
- `docs/archive/` — superseded snapshots, banner-marked; provenance only
- `docs/research/` — not created yet; reserved for future unchecked
  audits/proposals, processed point-by-point into the canonical docs and
  never moved there wholesale (same lifecycle shariha uses for its own
  `docs/research/`)
