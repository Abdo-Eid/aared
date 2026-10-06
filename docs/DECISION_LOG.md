# Decision Log

> Purpose: a chronological record of what was decided and why, with the
> reasoning behind each call. Read this when you ask "why is it shaped this
> way" months from now. For the current architecture, see
> [ARCHITECTURE.md](ARCHITECTURE.md). For the public API, see the API
> reference page ([`site/src/content/docs/api.mdx`](../site/src/content/docs/api.mdx)).
> For what's queued next, see [ROADMAP.md](ROADMAP.md).

The log runs from the first design session (2026-09-03) through Oct 2026. The
raw decision table of that session, with every alternative considered, was
`docs/archive/ms-previewrs/DESIGN.md`; that archive has left the tree and is
in git history (`git show 5a32bf0^:docs/archive/ms-previewrs/DESIGN.md`).

Entries were compacted in Oct 2026: each keeps its decision, reason and the
rules code cites; measurement tables, test lists and step-by-step stories
were cut, and superseded entries became short stubs. The full original text
is in git history before that change. Section numbers are not contiguous
(§41 was never used). The probe evidence an entry cites (`tools/probe/...`)
is local to the maintainer's machine, since that folder is gitignored: each
entry carries the measured result itself, and the probe name is only a label.

## 1. Why this project exists at all

The trigger was concrete, not speculative: evaluating client-side renderers
for an existing chat product's document-delivery panel surfaced two
disqualifying findings, preserved as primary evidence in the archived
`ms-previewrs/PPTXVIEWJS-NOTES.md` and `ms-previewrs/SHEETJS-LIMITS.md`
(in git history before `5a32bf0`, like `DESIGN.md` above):

- `pptxviewjs` paints PPTX to `<canvas>`, which drops paragraph-level RTL
  (`algn="r"`, `rtl="1"`) with no way to patch it back in — canvas output is
  pixels, and pixels have no CSS escape hatch.
- SheetJS Community Edition returns an *empty* `cell.s` even when asked for
  styles (`cellStyles: true`), silently — it doesn't fail, it just returns
  nothing usable, which is the trap.

Neither is a bug in those libraries so much as a structural consequence of
how they're built. Searching for an alternative confirmed a real gap: no
lightweight, view-only, browser-native library spans PPTX/XLSX/DOCX *through
DOM/SVG output* with correct Arabic/RTL — single-format DOM renderers exist,
and Canvas/WASM engines span all three at suite weight (see
[COMPETITORS.md](COMPETITORS.md) and the measured snapshot in
[BENCHMARKS.md](BENCHMARKS.md)). The choices are "own an entire office suite in the
browser" (Univer, LibreOffice-WASM, ~300 MB) or "accept broken RTL." aared is
the attempt at the missing middle.

## 2. PDF was scoped out, deliberately

`pdf.js` already solves client-side PDF rendering well. The gap that
motivated this project doesn't exist for PDF — its layout is already fixed
and its glyphs are already positioned by the time the file exists. Spending
effort here would be solving an already-solved problem instead of the actual
gap. Consumers who need PDF too pair aared with `pdf.js` directly.

## 3. Text is rendered by the browser, never rasterized — the load-bearing decision

Every other technical decision in this project is downstream of this one.

The alternative — draw text into a `<canvas>` or a WASM-side rasterizer —
was rejected because it means reimplementing the browser's own text stack
(HarfBuzz-equivalent shaping, the Unicode bidi algorithm, kashida justification,
ligatures, font fallback) to get back what `<div>`/`<svg>` text already does
for free. This is not a hypothetical cost: it's the exact, specific way
`pptxviewjs` failed (§1). Once text is pixels, a direction bug has no CSS
fix — only a patch to the rendering library itself, which for a third-party
dependency means waiting on someone else or forking it.

Consequence: WASM is scoped to a **leaf tool**, not the rendering core. The
only planned use is EMF/WMF → SVG conversion (a real, common need in Office
files, with no reasonable JS decoder available). ZIP decompression and XML
parsing both have adequate native browser APIs (`DecompressionStream`,
`DOMParser`) and don't need WASM at all.

## 4. View-only, permanently — not a phased-in decision

Considered and rejected: building toward eventual editing. Editing requires
a formula engine, round-tripping (preserving everything not touched, the way
shariha's PPTX editor does), and lossless-write guarantees — an entirely
different, much larger project. Scoping to view-only from the start halves
the problem and keeps the fidelity contract honest: aared promises to show
you what's in the file, not to let you change it.

## 5. Build order: PPTX → XLSX → DOCX, not alphabetical or "easiest first"

DOCX is structurally different from the other two: PPTX and XLSX both have
*bounded* layout problems (absolutely-positioned shapes; a grid). DOCX is a
flow-layout problem — line breaking, pagination, floats — which is, honestly,
building a word processor's layout engine. That's unbounded scope if taken
on first, before the shared `core` layer (OPC, DrawingML, charts, theme) has
even been proven out on two bounded formats. PPTX goes first over XLSX
because Arabic/RTL correctness — the whole reason this project exists — is
most visibly, immediately testable there (a slide either has right-aligned
Arabic text with correct bullet placement, or it doesn't).

## 6. Packaging: one package with lazy subpath exports, not `@scope/*` packages

The instinct toward `@aared/core`, `@aared/pptx`, `@aared/xlsx` as separate
published packages (shariha's own `@shariha/*` pattern) was considered and
rejected for a project this size. The reasoning: bundle size is decided by
*where code-splits happen*, not by *npm package boundaries* — a single
package with `aared/pptx`, `aared/xlsx`, etc. as subpath exports, each
`import()`ed lazily by the web component the first time that file type is
opened, gets the same "pay only for what you use" outcome without paying the
real cost of multi-package projects at small-team scale: version skew
between `core` and a renderer package that must always move together. Kept
consistent with shariha's proven pattern where it *does* apply — `core` has
zero rendering dependencies, and format renderers reach it through one
surface, not by depth-first coupling. `aared-fonts` is the one deliberate
exception: font files are large, optional, and genuinely independent of
release cadence, so it's a separate package on purpose.

## 7. DOCX pagination is explicitly not promised

Faithful pagination — matching Word's exact page breaks — needs the same
layout engine that made DOCX the hardest format (§5). Promising it would
either delay DOCX indefinitely or produce a renderer that's "mostly right,"
which is worse than being upfront: aared renders DOCX as flow-layout HTML
(matching the `docx-preview` tier of fidelity) and states plainly that page
count won't match Word. This is recorded as part of the fidelity contract in
[ARCHITECTURE.md](ARCHITECTURE.md#fidelity-contract) so it's a documented
boundary, not a discovered one.

## 8. PowerPoint is the text reference; LibreOffice is a companion artifact

Considered and rejected: asserting pixel-identical output against some
ground truth. No two OOXML renderers agree pixel-for-pixel even when both
are "correct" — shariha's own decision log (§15 in its `DECISION_LOG.md`)
independently reaches the same conclusion for a related problem: SVG/pdf-lib,
resvg, and Chromium-HTML painters are different painters, and demanding
identity across them is the wrong bar. But LibreOffice is not PowerPoint:
its font substitution can change Arabic glyph metrics, line wrapping, and
baseline placement. Using its raster to judge PPTX text would accept the wrong
output or reject the right one.

aared captures its own real DOM/SVG output in `Bun.WebView` and uses native
PowerPoint automation to export same-size reference PNGs. The paired artifacts
and review page are inspected for typography, bidi placement, and wrapping;
DOM tests re-read generated output to pin semantic layout properties. A
`soffice --headless --convert-to pdf` → `pdftoppm` conversion remains useful as
a companion artifact for broad visual inspection, never as the text-fidelity
oracle. See [ARCHITECTURE.md](ARCHITECTURE.md#quality-strategy).

## 9. PowerPoint probes settle renderer disagreements

When package XML, a browser capture, and native PowerPoint disagree, do not
pick the most convenient CSS approximation. The ignored nested repository in
`tools/probe/` captures a before/after package snapshot around one deliberate
Office interaction, while PowerPoint's object model can report the effective
style values it paints. This has already established two renderer rules:

- DrawingML `a:tint` blends in linear-light color space and retains the source
  proportion named by its value; gamma-encoded RGB interpolation produced the
  wrong table bands.
- A run with no explicit typeface still has an Office theme default. Resolving
  it in the Worker prevents the embedding application's font from changing
  text metrics or clipping Arabic titles.

The bottom-anchored title issue was a paragraph line-box defect, not a reason
to rasterize text or remove clipping from every text box. The
`tools/probe/title-insets/` experiment set distinct nondefault bottom margins
on master, layout, and slide placeholders: PowerPoint selected each nearest
present attribute. The parser now follows that rule for insets. The browser
still cannot use a 90%-of-font single-line box when it is smaller than Arabic
glyph ink, so the renderer preserves the source value in its model and applies
a minimum `1.1` CSS line-height. The real fixture pins the resolved model and
rendered containment.

## 10. DOCX layout mode: structural pagination without Word parity

Stale. Opted DOCX into physical pages (`pagination: { mode: "layout" }`)
beside flow rendering (§7), because page markers alone cannot paginate a real
document: Word printed four pages where the markers implied two. The parser
resolves `DocumentSection[]` and a headless paginator fills section boxes with
estimated block heights, so break positions match structurally, not exactly.
Its layout rules were later re-derived from native Word's own numbers.
Superseded by §22.

## 11. DOCX layout fidelity corrections from bench comparison

Stale. A bench comparison against Word fixed three defects, each traced to its
OOXML source: paragraph borders (`w:pBdr`) were dropped, list indents painted
twice, and headers/footers stacked inside the page consumed room the paginator
never reserved. It also made inline images inherit their paragraph's
alignment. The header clipping, border/indent geometry and list geometry
were later re-derived from native Word and changed. Superseded by §22.

## 12. `aared/preview`: superseding the deferred facade decisions

The bench's aared adapter re-implemented what the library already knew: try
each format's `open()` in turn, then hand-manage generation guards, teardown
and DOM cleanup — while `<aared-viewer>` did the same internally. Two earlier
deferrals ("no `preview()` convenience API", "no all-format auto-detect
entrypoint") rested on one rationale: one-format consumers must not pay for
all formats.

`aared/preview` supersedes both without breaking that rationale: the facade
lives behind its own subpath (root `.` stays `aared/core`), and format
entrypoints load via runtime `import()`, so only the detected format loads.
Detection reuses `WorkerSession.detectFormat` (Worker-side probe). The facade
is extraction, not a fork of the element's detect, lazy-open and
generation-guard shape.

Deliberate boundaries, held from review against the settled rows:

- `preview()` covers open → render → navigate → destroy only. Search, range
  reads, and media stay on the headless entrypoints.
- One teardown verb (`destroy()` does render + session + painted DOM)
  instead of the headless `destroy()`/`close()` split — the facade owns the
  whole session, so the split would be ceremony. The split stays where
  callers own the session.
- Fonts stay opt-in (`fontResolver` passthrough, never a CDN default). The
  rule, stated once: aared makes no network request of its own unless the
  application passes a resolver or a URL (see §53).
- `core` stays DOM-free: the facade is its own subpath importing the
  internal session plumbing, never part of `aared/core`.

## 13. DOCX `render()` joins the shared async object envelope

DOCX rendered positionally and synchronously — `render(sections,
container, options?)` — while PPTX/XLSX rendered as `await render({ slide |
sheet, target, ... })`. Every consumer (bench, element, `preview()`) had to
branch on call shape, not just on format; the shared-grammar design (same
envelope, format-native nouns) exists to eliminate that.

The old form was removed, not aliased (AGENTS.md "Pre-release breaking
changes"): `await render({ document, target, ... })` everywhere. Two gains
beyond symmetry:

- Taking the `DocxDocument` (not a bare section list) lets media default to
  the owning session, the same rule as PPTX/XLSX handle inputs. The old
  signature *structurally* couldn't do this — it never received the
  document — which is why bench hand-wired `mediaResolver` for DOCX only.
  Detached `DocumentSection[]` input stays for model-only callers, with an
  explicit resolver required for images.
- Async uniformity costs nothing (`await` on a sync paint) and lets the
  `preview()` facade delegate without format branches; it also leaves room
  for layout pagination to move work off-thread later.

The format-native noun stays (`document`, not `unit`): a slide isn't a
sheet isn't a document, and unifying the noun would erase meaning. What's
shared is the envelope.

## 14. `aared/react`: the mount lifecycle as hooks, not docs prose

`preview()` removed the per-format open/render boilerplate, but every React
embed still rewrote the same effect: `AbortController` per mount, capture
the handle for cleanup, swallow aborted rejections, mirror `index`/`zoom`
into state, destroy on unmount. That logic is fixed — it never varies by
app — so it belongs in the package, not in every consumer.

`aared/react` exports one hook, `usePreview()`, plus `usePptx()`,
`useXlsx()`, and `useDocx()` — the same hook with `type` asserted, so
detection is skipped and only that format chunk loads. The wrappers add no
lifecycle; they exist because "preview this one PPTX" should read as
`usePptx`, not `usePreview({ type: "pptx" })`. Same boundary as the facade
(open → render → navigate → destroy); headless power stays headless.

Two packaging choices, both about not taxing non-React consumers:

- `react` (`>=18`) is an *optional* peer behind the `aared/react`
  subpath. Vanilla and element consumers never install it; the entrypoint
  is import-time safe, and `react` itself is never bundled.
- Plain-data options are keyed by value inside the effect so inline
  literals (`appearance={{ theme }}`) don't remount; `source` and the
  function resolvers are identity-keyed and documented as must-be-stable.
  StrictMode mount → cleanup → remount resolves exactly once.

## 15. `capabilities` cut before first publish; `search()` and the headless boundary kept

The API redesign gave each document a `capabilities` object
(`textSelection`, `media`, a per-format block, `editing: false`) so
consumers could avoid format-name conditionals. Auditing what the headless
surface is for turned up nothing that justified it:

- **Constant.** The value was hard-coded and identical for every file, so a
  consumer could not branch on it usefully.
- **No consumer.** Nothing in the package read it; only tests and the API
  page did.
- **Already wrong once.** `presentation.notes: true` claimed something
  nothing rendered. A flag kept honest by hand, with no reader to notice
  drift, is a liability.
- **Cheap to add, costly to remove.** Pre-release is the cheapest moment to
  cut.

It was removed (not aliased — AGENTS.md "Pre-release breaking changes"):
`deck.capabilities`, `workbook.capabilities`, `document.capabilities`, and
the `PptxCapabilities`/`XlsxCapabilities`/`DocxCapabilities` types.
Re-introduce it only when a value varies per file or environment, and then
with a consumer that reads it.

Kept, deliberately:

- **`search()`** — a real function someone building their own search UI
  would otherwise have to write against the resolved models. It stays
  headless-only (the facade boundary in §12 is unchanged).
- **`read()`, `media.get()`, `info.sizes`** — load-bearing: `render()`
  consumes the same model `read()` returns, the renderer resolves pictures
  through `media`, and PPTX slide geometry comes from `sizes`.

What stays true about scope: the headless API exists for custom UI. Reading
data from resolved models is a by-product, not a converter feature.

## 16. Search cost, DOCX search coverage, `info.sizes`, and the Worker bundle

Follow-up to §15, prompted by an audit of what the headless API costs.

**Search cost.** PPTX/XLSX keep each unit's searchable text after the first
read (a per-document index on the main thread, cleared on `close()`), and
search reads skip the font-resolver round trip
(`unit(..., { fonts: false })`). Measured on large generated fixtures, a
repeat search fell from ~1.4 s to ~3 ms on XLSX; the cold first search is
unchanged because parse and model transfer dominate it.

Not done, deliberately: a Worker-side text-only `search` message. It would
remove the model clone (~1.4 s of the 3.9 s XLSX cold search) but adds a
protocol message and moves text extraction into core; revisit if real
documents make the cold search a problem. Also fixed: the worker cached the
XLSX workbook read as a *promise* bound to the first request's abort signal,
so a cancelled first read could fail every later sheet read.

**DOCX search now covers tables and footnotes/endnotes.** Matches gain
`kind`, and `table: { index, row, column }` / `note: { id, label }`. Word
stores notes after the body, so a note reports the section and heading of the
paragraph that first references it. Headers/footers and text boxes (not
parsed) stay out.

**`info.sizes`** (removed again in §31): DOCX now reports one page size (inches) per section; its
length is the section count, not `count`. XLSX stays without it: every size
input lives in the worksheet part and `open()` deliberately parses none.

**Bundle**: the Worker (`core/worker.mjs`, ~121 KB minified when bundled as one
browser file, ~40 KB gzip; the ~110 KB first written here predates the DOCX
parser growth, see §23) bundles all
three parsers, so even `aared/pptx` alone downloads the XLSX and DOCX
parsers there. The docs previously implied one-format imports were fully
one-format; that is only true of main-thread code and is corrected. Also
fixed: `sideEffects` named `src/element/index.ts`, which doesn't exist in the
published `dist`, so a consumer's bare `import "aared/element"` bundled to
nothing (verified with a consumer build). Per-format subpaths for
preview/react and a per-format Worker are recorded in the ROADMAP under
"Before first publish" rather than built now.

## 17. PPTX master/layout surface: flatten the small, share the large

Slides inherit from a layout, which inherits from a master. Both contribute a
**background**, non-placeholder **graphics** (logos, bars, header text), a
**colour map** (`bg1`/`tx1` → theme slots), and placeholder **prototypes**.

Decision, by cost:

- **Background: flatten onto the slide** (`SlideModel.background`, effective
  slide → layout → master). It is a few bytes (a colour, a gradient, or a
  `mediaId`), so there is nothing to save by sharing it, and a second source
  of truth would need a compositing rule. A master image repeated on every
  slide is served by fetching the *bytes* once (below).
- **Graphics: store once per layout** (`SlideLayoutModel`, master's shapes
  then the layout's). Layouts are few and slides are many, so this is where
  duplication would be: the Worker parses a layout once, the session fetches
  it once with a dedicated `layout` request (cancellable, one shared
  in-flight fetch not bound to any caller's signal, never caches a failure),
  and every slide on that layout carries the same object.
  `SlideHandle.read()` attaches it; `search()` never fetches it. A `SlideModel`
  stays renderable without the deck.
- **Media bytes: fetched once per document** for layout pictures and an
  inherited picture background (`sharedMedia`), so paging slides does not
  re-read a master logo from the package each time. Slide-owned pictures are
  not cached, which keeps memory bounded on decks of distinct photos.
- **Placeholders on layouts/masters are never painted.** They are prototypes;
  the slide's own placeholder carries content, and they already feed
  geometry/text-style inheritance.
- **Colour map applies to every colour in a slide**, not only backgrounds.
- **Fields**: `slidenum` shows `SlideModel.number` (applied at paint time, so
  shared layout graphics show the right number per slide); `datetime` keeps its
  cached text.
- **No public `deck.masters`/`deck.layouts`.** aared is view-only and the
  collection would widen the API for no rendering benefit.

The layout-model transfer is a new Worker request type, an additive protocol
change. Layouts with no graphics add no `layoutId` and no request.

A theme background reference (`bgRef`) paints the theme's `bgFillStyleLst`
entry (tint and `satMod` applied), not the bare scheme colour; PowerPoint
shows `#3A4961` where the bare colour is `#0E2841`.

Layout → master placeholder matching is by placeholder *type* (title, body,
dt, ftr, sldNum), as PowerPoint does, because masters number them differently
from layouts and slides; slide → layout stays by index.

Not verified against real PowerPoint output yet: the `showMasterSp` rule,
picture backgrounds, and `p:clrMapOvr` (layout/slide colour-map overrides, not
read). They are listed in the ROADMAP fidelity gaps.

## 18. Core performance pass: streaming parse, cached relationships, split worker

`core` was optimized in isolation against generated fixtures. Public API is
unchanged; outputs were diffed field-by-field and are identical except the two
intended changes below. The headline: a 150k-row XLSX first-sheet read went
from 14.3 s and 1.58 GB peak to 2.5 s and 0.47 GB.

**What changed (still valid):**

- **XML** (`core/xml.ts`) — char-code scanner and a streaming mode
  `parseXml(src, { streamChildrenOf, onChild })`. Parsed trees are read-only:
  nothing in `core` mutates them.
- **ZIP/OPC** (`core/zip.ts`, `core/opc.ts`) — inflate writes straight into a
  buffer of the declared size; stored entries are decoded without a copy;
  `.rels` are parsed once per package (cached relationships); media is
  transferred to the main thread without a double copy (`readPart` returns an
  owning buffer; internal decodes use the aliasing `readShared`).
- **XLSX** — `sheetData` and `sharedStrings` are streamed (no full element
  tree); cells live in a sorted per-row grid, so a hyperlink or conditional
  format on `A1:A1048576` touches authored cells, not the area; theme and
  `workbook.xml` are parsed once; cell styles are shared per style id.
- **PPTX** — placeholder inheritance reuses the cached layout/master chain
  (previously re-inflated both parts per placeholder). This supersedes the
  "layout/master memoization reverted" note in §17.
- **DOCX** — shared header/footer parts parsed once. (The worker-side sections
  cache was superseded by §35: the page keeps the sections, the Worker keeps
  none.)
- **Worker/session** — model walks for media ids and font families go through
  `core/collect.ts` (`collectStrings`); abort helpers live in `errors.ts`; the
  canonical `FontResolver` type is the `aared/core` one.

The split-Worker build this entry also introduced (per-format chunks, a
`worker-single` artifact) was reverted by §20; its packaging failures are §19.

**Intended behavior changes (2):**

1. **XLSX charts use the shared `chart.ts`**, so they match PPTX charts.
   Charts without cached values report `unsupportedReason` instead of drawing
   empty axes; doughnut charts draw as pies in both formats; both formats
   paint the authored `c:title`.
2. **Shared read-only objects**: cells using the same style id share one
   style object, and cells covered by one hyperlink share its object.
   Structured clone preserves this. Code outside `core` must not mutate
   `cell.style` / `cell.hyperlink` in place.

## 19. Split worker vs. consumer bundlers: Vite verbatim copy and the abort taxonomy

Follow-up to §18, prompted by a production `vite build` that loaded forever.
The split-worker parts below were reverted by §20; two rules survive.

**Packaging (split part stale, rule kept).** Vite copied the split worker
verbatim, without its sibling chunks, because the Worker URL was built as
`new Worker(workerInit ?? new URL(...))` and Vite only bundles a worker
target when `new URL()` appears *literally* inside `new Worker()`. Rule that
still holds: the default path keeps the literal constructor shape (pinned by
a comment; never fold the branches back together).
`tools/check-consumer-build.ts` guards it. The `worker: { format: "es" }`
Vite config line the split needed went away with §20.

**Error model (still valid): a dead worker reports `"worker-error"`, never
`"aborted"`.** Boot failure used to construct `AaredError("aborted")`, so
every abort check treated a broken worker as a cancellation and `usePreview`
swallowed it, leaving the UI loading forever. `"worker-error"` matches neither
`isAbort()` nor the `"aborted"` code check, and a failed `postMessage` moved
with it (transport failure, never cancellation). Genuine cancellation
(`abortError()`, `closedError()`, signal paths) still reports `"aborted"`,
which unmount cleanup relies on.

## 20. Worker back to one self-contained file: the split is reverted

Overturns the §18 packaging decision (keeps its parse-streaming,
`collect.ts`, and `errors.ts` work, which were never about chunking). The
split worker — `dist/core/worker.mjs` plus per-format sibling chunks, with a
separate `aared/worker-single` artifact for Blob consumers — saved ~55 KB
for single-format apps but bought a packaging failure class: Vite's worker
detection needed the literal `new Worker(new URL(...))` shape (§19),
verbatim copies dropped sibling chunks with zero errors, consumers needed
`worker: { format: "es" }`, and every artifact had to be deployed as a
directory of hashed names. One self-contained `dist/core/worker.mjs`
(`codeSplitting: false`) deploys alone, inlines into a Blob URL, and survives
a verbatim copy. Cost: every consumer ships all three parsers. The contract
is pinned by `test/worker.test.ts` (no relative imports in the artifact, Blob
URL opens every format) and `tools/check-consumer-build.ts` (the emitted
worker references no sibling chunk).

Measured production cost (minified / gzip): worker 117 KB / 39 KB, all-formats
main thread ~100 KB / ~28 KB, against ~500 KB–1.6 MB for single-format
alternatives in `bench/`. Per-format worker variants only on a real consumer
size budget, never speculatively.

Re-checked (Sep 2026, with the download-size method of §33): one flat worker
per format measured 23.5 / 22.0 / 16.8 KB gzip against 43.1 KB for the shared
one — 20–26 KB less for a single-format page, about 19 KB more for a page that
imports all three entries, and four worker artifacts to ship and document.
Rejected again. The size gap it would close is DOCX's (about 8 KB behind
docx-preview + jszip), and that stays open.

## 21. Measuring text for DOCX pagination: measure in the DOM; Pretext considered, not adopted

Context. The layout-mode paginator (§10) fills pages with block heights it
*estimates*, from a per-script average glyph advance (`AVERAGE_CHAR_EM`) and a
rounding tolerance (`WRAP_TOLERANCE_LINES`) in `core/docx/pagination.ts`.
Those constants were fitted to one fixture, so a different document or font
can break them.

Options considered: keep estimating; Pretext (`chenglou/pretext`, which
measures text through canvas `measureText` and computes breaks
arithmetically); or measure in the DOM, reading real block heights in one
batch and feeding them to the paginator. (Pretext's README, read Sep 2026,
says it does not compute bidi levels, is silent on web-font loading, Workers
and inline images, and gives no licence or version; none of that was checked
beyond the README.)

Why Pretext is not adopted.

- **The gaps are Word's inline model, not just bidi.** Tab stops with leaders,
  inline images, superscripts, letter-spacing and justified lines are not
  modelled, so it would become a second layout engine that must still agree
  with what the DOM paints.
- **Canvas measurement is a grey area under §3.** That rule bans canvas
  `fillText` and WASM text shaping; `measureText` is not named. DOM
  measurement needs no new entry.
- **Unknown font contract.** Measuring and painting must use the same faces,
  including anything delivered by `aared-fonts` / `fontResolver`.
- **No Worker need.** `paginate()` is called from the renderer on the main
  thread (`renderLayout` in `docx/index.ts`), where a DOM exists.

Decision. Do not add Pretext. Move DOCX pagination to DOM-measured block
metrics in the renderer (done in §44). `paginate()` keeps its
algorithm and takes an injectable block-metrics function
(`{ topPx, contentPx, bottomPx }`); the renderer supplies DOM measurements and
headless callers and unit tests keep the estimator. Requirements found in the
fidelity pass:

- explicit image heights (block images are `height: auto` until their Blob URL
  loads and would measure 0);
- wait for `document.fonts.ready` and re-paginate when fonts arrive;
- run the tab-stop pass before reading heights;
- one measuring container per page width;
- measure once and move the same nodes into the pages.

Trade-off accepted: break positions then depend on the machine's fonts, as
Word's do, and can only be pinned by WebView visual tests.

The estimator stays: it is the only page layout that runs without a browser,
which fits the "headless thumbnail export" item in [ROADMAP.md](ROADMAP.md).
Revisit Pretext when headless accuracy is scheduled, or if DOM measuring is
too slow on very large documents; then run a spike first (compare line counts
with the DOM for every fixture paragraph) and verify Worker or server support
and the licence before depending on it.

## 22. DOCX layout follows Word's own numbers; supersedes parts of §10 and §11

The bench comparison in §11 judged fixes by eye against screenshots. Every
layout rule is now read from a native Word export of the frozen fixture
(`tools/word-render` dumps every text span's font, size and position), so a
number is read, not guessed. The rule list is in `ARCHITECTURE.md` § Format
rules (DOCX). What it overturned:

- **Headers are never clipped** (§11, third bullet). A header hangs from its
  `w:header` distance and grows down, and the body starts below it
  (`LayoutPage.bodyInsetsPx`, derived from estimated band heights;
  `marginsPx` stays as authored). No fabricated rule is painted.
- **Side borders sit outside the indent** (§11, first bullet), so the
  estimator no longer narrows the text column by border-plus-space, and the
  bar runs through the paragraph's space-after.
- **List geometry** (§11, second bullet): the marker is inline at the start of
  the paragraph with a hanging indent from the start edge, replacing the
  `<li>` grid, so RTL items carry it on the right.
- **Line boxes:** an authored `w:line` multiple scales Word's ~1.15em single
  line, and a line follows the sizes of the runs on it, not the paragraph
  default.
- **Bidi alignment:** `left`/`right` are start/end in a bidi paragraph.
- **§10's "one-block deviation" no longer holds.** After the fixture was
  re-frozen, Word itself keeps Table 1's caption with its table on page 2, and
  aared now matches.

Pagination remains structural, not Word-exact (§10, §21): break positions
depended on the estimator until measured pagination landed (§44).

## 23. Shared-code pass: one home per rule, and the decisions it forced

A read-only audit (after §22) listed code that had been written twice for lack
of context. What landed is one home per rule:
`core/units.ts`, `core/color.ts`, `core/typography.ts`, `core/xlsx/limits.ts`,
the OPC helpers on `OpcPackage` (`readXml`, `readXmlIfPresent`,
`officeDocument`, `mediaContentType`) and, renderer-side, `appearance.ts`,
`media-url.ts`, `chart-svg.ts`, `table-dom.ts`, `text-style.ts`,
`render-handle.ts` and `format-route.ts` (the map, and what stays per-format
on purpose, is `ARCHITECTURE.md` § Shared modules). Where the two copies
disagreed, this is what was decided and why:

- **`lnSpc` does not scale by Word's single-line factor.** DrawingML's spacing
  multiple is the font size itself; Word's "single" is about 1.15em. Each was
  validated against its own Office app, so `applyParagraphBox` takes the factor
  as a parameter (`singleLine`, default 1; DOCX passes 1.15) rather than
  picking one.
- **One palette, published as `--aared-ink`.** The text token is `--aared-ink`
  in every renderer, as the API reference already said; XLSX's `--aared-text`
  collided with `<aared-viewer>`'s consumer-facing stage colour and is gone. The
  dark surface is a neutral `#1a1a1a` with ink `#e8e8e8` (Word's and Excel's own
  dark modes are neutral; a tinted surface would tint authored greys), and each
  muted/accent value is the higher-contrast of the pair the renderers used to
  paint separately. This overturns XLSX's `#1E1F26`/`#E9E9EC`, the one visible
  change; DOCX paint is unchanged. PPTX now writes the resolved theme into
  `data-aared-theme`, not the raw `"system"`.
- **A handle is inert after `destroy()`, in every format.** The API reference
  said so, PPTX and DOCX did it, and XLSX and `PreviewHandle` alone threw a
  typed `aborted` error. One teardown story: they no-op.
- **A DOCX/XLSX template opens, a PowerPoint template does not.** Word and Excel
  templates carry real content; a `.potx` carries layouts but no slides. The
  policy is now stated once, in `officeFormatOf`, instead of by accident in four
  copies.
- **An empty theme complex-script slot resolves to the theme's `Arab` font.**
  Ground truth from a native Word probe (minor `Arab` = Tahoma, major = Segoe
  UI): a `w:rtl` run whose theme `cs` typeface is empty is set in that entry, a
  literal `w:cs` wins, and `cstheme` or `w:lang w:bidi` alone change nothing.
  PPTX already did this; DOCX now shares the rule (`themeFontFamily`). The probe
  also showed Word ignores the complex-script slot for Arabic text in a run
  *without* `w:rtl` (it uses the Latin font and lets glyph fallback pick Arial);
  aared keeps its text-based detection as a harmless superset, because Word
  always writes `w:rtl` on Arabic runs.
- **Pre-release renames, no aliases:** `DocxTextAlign` → `TextAlign` and
  `TextBodyInsetsModel`/`DocxCellMargins` → `InsetsModel` (public, so
  `core/index.ts` and the API reference moved with them), and `isAbort` is now a
  type predicate (`error is Error`).

Bugs found on the way, fixed with tests: the DOCX parser never honoured
`w:outlineLvl` (it read an attribute named `outlineLvl`; Word writes `w:val`),
so `open()`'s titles disagreed with the heading nodes and both now come from
one rule; a non-page `w:fldSimple` dropped its cached text; a run in a content
control ignored its own character style; `setZoom(0)` collapsed a DOCX.

Deliberately not merged: the SheetJS/docx-preview baseline adapters (a baseline
must not depend on the code it measures); XLSX `tint` (sRGB lerp) versus
DrawingML `tint` (linear light, probe-backed), which are different operations;
`TextRunModel.verticalAlign` versus `baselinePct`, one concept in two dialects
that each renderer paints as measured; and the two facades' `render` calls,
which differ on purpose (the element adds sheet tabs and theme). XLSX keeps raw
Excel border names and maps them in the renderer, where DOCX resolves its
borders in core; moving that changes the public `XlsxCellStyleModel`, so it
waits on the model-contract decision (ROADMAP, "Before first publish").

The Worker bundle did not grow (about 121 KB minified before and after); the
renderer-side modules are not in the Worker.

## 24. Dark mode, looked at: automatic text follows its fill; authored text on the dark surface is an open gap

Stale. A contrast audit of the coverage fixtures rendered dark found that
adaptive dark made authored fills with automatic text unreadable, and that
authored text colours on the darkened surface (Excel black body cells, Word
`1B2029` body copy) read at 1 to 3.4:1, so whole sheets and documents looked
blank. A first fix resolved automatic text against its fill (`automaticInk` in
`appearance.ts`, on `contrastRatio` in `core/color.ts`), which stays, and a
recolouring rule was proposed for authored text. Instead the adaptive mode
itself was removed, because any fix recolours authored values and contradicts
the view-only fidelity promise. Superseded by §25.

## 25. Adaptive dark retired: every format is source-faithful

Overturns the adaptive half of the appearance contract (§24) and closes
§24's open gap by removing it rather than fixing it. The §24 bench audit
showed the shape clearly: text on an authored fill was fixed by
`automaticInk`, while authored text directly on the page (Excel `FF000000`
body cells, Word `1B2029` body copy) reads at 1–3.4:1 on the darkened
surface — whole sheets and documents reading as blank. Any fix recolors
authored values, which contradicts the view-only fidelity promise that
motivated PPTX's source-faithful rule in the first place.

Decision: `ContentMode` is gone (pre-release hard remove, no alias).
`Appearance` carries `theme` only; every renderer paints white surfaces
with authored fills/colors verbatim and resolves automatic text against
its fill, in light and dark alike. Dark `theme` drives viewer chrome only
(stage, controls, sheet tabs, selection, tokens) — a dark viewer shows
light sheets and pages on a dark stage, exactly like a slide. This keeps
one rule for all formats instead of two contracts and a recoloring engine.

## 26. Deferred geometry: the worker describes, the main thread evaluates

The full ECMA-376 preset catalog (~187 shapes) and `custGeom` freeforms
cannot live in the worker: it is built flat (`codeSplitting: false`, so it
works as a single file or Blob URL), and any `import()` the parser makes is
inlined into `worker.mjs`. Hand-coding every preset was rejected (months of
formula transcription, still partial), and so was shipping the catalog in the
worker (every deck would carry it).

Decision: the worker keeps a hand-coded fast path for the 11 common presets
and describes everything else on a `deferred` field on `ShapeGeometryModel`:
the encoded `custGeom` program string, or for a preset just the authored
adjustments, since the catalog supplies the definition by name. The session
resolves deferred geometry on the main thread in `unit()`/`layout()` with a
lazily imported evaluator. One interpreter serves presets and `custGeom`
because both are written in the same ECMA guide language
(`avLst`/`gdLst`/`pathLst`); multi-path presets surface as optional
`fillOnly`/`strokeOnly` layers. `deferred` is always stripped before a model
reaches a consumer. The catalog source is ECMA-376 Part 1
`presetShapeDefinitions.xml` (`scripts/`), compiled by
`scripts/compile-presets.ts` to 176 presets with 16 proven aliases, each
checked against LibreOffice's independent copy (`test/drawingml-oracle.json`).

**Addendum (Sep 2026): the lazy family chunks are reversed; one eager table
ships in `aared/pptx`.** Per-family chunks saved bytes only for decks using
few families, and cost a manifest, a loader per family, eight emitted chunks
whose placement under `dist/core/` had to be pinned for the worker URL, and a
failure mode (a chunk failing offline or under CSP left shapes as
fallbacks). The whole table is ~25KB gzip, so it is accepted:
`scripts/compile-presets.ts` emits one `src/core/presets.ts`
(`presetPrograms`) and `drawingml-resolve.ts` takes the table as a parameter.

The load-bearing constraint is where the table lands. `core/session.ts` is
bundled into every format, so importing the table there would make xlsx/docx
apps carry every preset outline. Instead `aared/pptx`'s `open()` imports it
statically and passes it to its session as the internal
`OpenSessionOptions.geometryPrograms`; a session without it resolves no
deferred geometry. Measured: `pptx/index.mjs` grew from 11.8KB to 31.1KB gzip,
`xlsx`/`docx`/`worker.mjs` are unchanged, and `worker.mjs` holds no preset
names. The `deferred`-by-name transport, the lazily imported evaluator and the
`geometry: false` search gate stay.


## 27. Shape style references with `idx="0"` mean "no reference"

A `p:style` `a:fillRef` (or `a:lnRef`) with `idx="0"` carries a vestigial
color child the schema requires, but PowerPoint renders no fill (or line)
for it. Ground truth came from the coverage deck itself: its unfilled
right brace (`fillRef idx="0"` + accent1 child) paints no fill in
PowerPoint, while a newly added left brace with a direct black fill and red
outline paints exactly that — direct formatting always wins over style
references, and a zero index wins over the color child. The parser used to
resolve the child unconditionally, painting a phantom accent fill on the
brace.

This sat in three places with one root cause (shape fill, shape outline,
table-cell fill), fixed once in `resolveStyleRef`. Table cells, text runs,
backgrounds, charts, XLSX indexed colors and DOCX shading were audited in
the same pass and already honor their own none-semantics (`a:noFill`,
`auto`/`nil`/`none`, explicit palettes), so they were deliberately left
alone. The fixture's brace pair pins the behavior.

## 28. DrawingML fills: one model, one painter, and what PowerPoint actually draws

**Why.** A fill the renderer could not paint used to vanish without a
trace: the model tagged it `unsupported`, the renderer set a data attribute
nothing read, and the shape, cell or slide painted empty with no dev marker.
Pattern fills, picture fills on shapes, `path="rect"` gradients, tiled
pictures and every non-solid table-cell fill hit this, and no fixture
exercised them. `pptx-coverage.pptx` gained slides 7-10 (authored through
PowerPoint's COM object model) to carry every fill kind.

**Model (pre-release, breaking).** Shapes, table cells and slide backgrounds
share one `ShapeFillModel` union: solid, linear, centered radial,
`unsupported`, plus new `pattern`, `picture` (`stretch`/`crop` or `tile`) and
`rect-gradient`. `ImageBackgroundModel` is gone (a picture background is a
`PictureFillModel`). `TableModel.cellFills` holds fills, not colour strings;
the cell's own `a:tcPr` fill wins over the table-style band.

**Painter.** `src/pptx/fill-paint.ts` turns every fill into one SVG paint
server, for shapes, cells (an SVG layer under the text) and the background,
so the three cannot disagree. A fill it cannot paint returns nothing and the
caller marks it: a dashed outline in development builds (`MARK_FALLBACKS`)
and a `data-aared-fill-fallback` reason. Paint servers are sized in real CSS
pixels, not viewBox units.

**Ground truth, each from a PowerPoint render.**
- *Pattern bitmaps.* ECMA names the 54 presets but draws none. Each is an
  8x8 bitmap whose cell is one CSS pixel, scaling with the slide and anchored
  to the slide origin. The table is `src/pptx/pattern-presets.ts`. Cells are
  anti-aliased; `crispEdges` banded them (moire) in a small viewer.
- *Rect-path gradients.* `fillToRect` is read against the fill box, not
  against `tileRect` as the spec wording suggests.
- *How stops blend.* Colours mix in **linear light** and the weight follows
  **smoothstep** within each pair of stops. SVG and CSS only interpolate
  linearly in sRGB, so `officeGradientStops` expands each segment into 16
  intermediate stops (worst error 44 levels down to under 8). The same stops
  serve shape, background, cell and WordArt gradients.
- *Picture tiles.* Tile size is the image's natural pixels at 96 DPI times
  `sx`/`sy`.

**Found alongside.** A gradient background SVG was appended after the slide
scene and painted over it (now the viewport's first child). Every table got
a grey outer frame that PowerPoint does not draw. `a:ln w="0"` painted
nothing where PowerPoint draws a hairline (strokes now floor at one CSS
pixel). Tables and plain text boxes ignored `p:defaultTextStyle`'s run size
(they now take it per level); table-style `tcTxStyle/@b` makes header rows
bold.

**Deliberate limits.** COM cannot author an off-center `path="circle"` or a
`path="shape"` gradient, so both stay `unsupported` (dashed in dev) until
probed by hand. Patterns in group children anchor to their child-space
position. Tiles assume 96 DPI. Which of `p:defaultTextStyle` and the
master's `p:otherStyle` PowerPoint applies to text boxes is unprobed; they
agree in both fixtures.

## 29. Real zoom fit for PPTX and DOCX; the DOCX heading/page facade gap closed

**Why.** A bench audit flagged places where the bench compensated for real
product gaps. "Bench-only fixes stay in bench" covers genuine harness
concerns, not a stand-in for fixing the product, so each was traced to its
layer.

**PPTX: `fit-page` measured its own slide's height.** A box that only grows
to fit what it holds measures zero once the slide is hidden to measure the
container, so `fit-page` never shrank anything. It now hides the viewport and
measures the container's own `clientHeight`. Separately, the bench applied
CSS `zoom` to a `width: 100%` wrapper: a percentage width resolves against
the unzoomed containing block, so a zoomed-in slide clipped with nothing to
scroll to. `aared/pptx`'s `setZoom()` resizes the viewport's real box
(`naturalWidth * scale` px) with `transform: scale()` painting inside it, so
the browser's own overflow math works; the bench now drives that.

**DOCX: `fit-width`/`fit-page` were a no-op.** `wireViewer` only pinned a
numeric CSS `zoom`; both fit modes just cleared it. They now measure against
the widest laid-out page (mixed portrait/landscape fit against the wider)
with a `ResizeObserver` on the render target. CSS `zoom` still does the
scaling, but the pages column's width is set to the exact unscaled page width
in pixels first (a `width: 100%` column would not shrink, the same
percentage-under-zoom trap). Flow mode has no fixed-width page, so both fit
modes stay a no-op there by design. Pinned by `docx-zoom.visual.test.ts`.

**Hollow WordArt text: a merged-silhouette outline, not a per-glyph stroke.**
`-webkit-text-stroke` leaves a seam wherever two Arabic glyphs join or
overlap (ground truth: `pptx-coverage.pptx` slide 6; PowerPoint draws one
clean outline). An SVG filter works on the element's whole alpha, so touching
glyphs read as one shape. `dilate(alpha) - erode(alpha)` failed: Chromium's
`feMorphology operator="erode"` silently no-ops under an ancestor CSS
`transform: scale()`, which this renderer's zoom/fit always applies. The fix
avoids `erode`: `dilate(alpha) XOR alpha`, an outside-only band. The outline
therefore sits outside the glyph rather than straddling it, a deliberate
visible difference from a centered stroke.

**`PreviewHandle.docx`: DOCX's facade gap, closed.** The DOCX render handle
already exposed `pages`, `pageCount` and `goToHeading()`, and
`DocxDocument.headings` carried the outline, but the format-agnostic
`PreviewHandle` forwarded none of it, so the bench scraped headings and page
sections from the DOM. A slide or sheet is one navigable unit (`index`/
`goTo()` cover it); DOCX has no equivalent to reuse. Added as
`viewer.docx: DocxPreviewControls | null` (`null` for pptx/xlsx) rather than
new top-level fields.

**Bench-only, no product change.** The XLSX sheet-tab pill read a local
`activeSheet` that never flowed back from the product; the bench now also
pushes `hook.index` into it. The aared/baseline wrapper width classes were
documented: aared's fit modes size against the real container edge, a
baseline library paints at its own size.

## 30. PPTX tables and text lines, matched to PowerPoint

**Why.** A pixel diff of `post-production-review.pptx` against PowerPoint's
export found slide 5's table collapsed (six rows squeezed into a third of
their height) and a steady 2-4px vertical drift in single-line text. Both
traced to how aared sized boxes, not to any one slide.

**Line boxes follow their runs** (`lineFontSizePt`, now in
`core/typography.ts`, shared with the DOCX paginator). Every PPTX `<p>`
carried the deck's default size (18pt) even when its text was 13pt, so the
line box was an 18pt strut and small text sat low in it.

**PowerPoint's 100% line is 1.2 x the font size.** `text-style.ts` assumed
DrawingML's single line was the font size itself. Measured: 28pt table text
gives 54.5px rows in PowerPoint's export (0.1in insets plus a 44.9px line).
Unauthored spacing is the same 100%. Exact point spacing (`a:spcPts`) was
already right.

**Tables grow to their rows.** PowerPoint grows the frame around its rows;
aared shrank the rows to fit the authored frame height. Rows now take their
authored height as a minimum and grow to hold their text; the frame height
applies only to rows authoring none. Cells stating no insets get DrawingML's
defaults (0.1in / 0.05in), resolved in core.

**The whole table-style surface.** The reader resolves every region
(`wholeTbl`, both band pairs, first/last row and column, the four corner
cells) with PowerPoint's precedence: fill, text colour, bold/italic, typeface
and six edges. An edge comes from the region's outer border on its outer
sides and `insideH`/`insideV` inside it, mirrored for RTL. Direct `tcPr`
edges win per side, and an `a:noFill` edge clears a style edge. Diagonal
borders, dashed edges (`prstDash`), rotated cell text (`vert`/`vert270`) and
theme line references (`lnRef`) are read and painted.

**Edges straddle the boundary.** Cell borders were CSS borders on each cell,
so a shared edge painted twice. They now paint as lines centred on the
boundary in an overlay above every cell's fill, as PowerPoint strokes them.

**A face the reader lacks falls back to its kind** (`cssFontFamily`). A bare
family (`Arial`) dropped to the browser's serif default; PPTX, DOCX and XLSX
now emit the face then its nearest generic family.

**Deliberate limits.**
- Built-in table styles referenced by GUID but absent from `tableStyles.xml`
  paint unstyled; PowerPoint regenerates them from the theme.
- `a:lnStyleLst` is not read; `lnRef` widths assume the default theme.
- Body text whose theme font PowerPoint fetches from the cloud (Aptos, the
  default since 2023) paints in a substitute; `pptx-coverage.pptx`'s Latin
  table text measures ~12% larger for it. Font work stays on the roadmap.
  (Closed by §49 for anyone passing a font resolver: Aptos has a calibrated
  substitute.)
- A run tagged `en-US` holding Arabic words around a neutral (`عربي +
  إنجليزي`, slide 5) keeps its words left-to-right in PowerPoint; aared applies
  the Unicode bidi algorithm, which reverses them. Unprobed, so unchanged.
- Charts are untouched (deferred).

## 31. One API grammar across formats and layers

**Why.** An audit of the API reference against the code found about thirty
mismatches, several real bugs (XLSX fit modes did nothing; a closed document
reported `"aborted"`, which `isAbort()` swallows; DOCX heading indexes counted
header and table headings), and a second pass found seventeen places where
the same behaviour had different syntax per format or layer (`SheetInfo.name`
vs `SlideInfo.title`; different event-detail shapes; sync DOCX `search`;
`docx: { pagination }` in the facade but top-level in `render()`; three
per-format appearance types). Pre-release, so best design wins: every caller
moved in the same change, with no aliases.

**The rules** (stated once on the API page, § Design principle):

1. **Units vs places.** `count`/`goTo(i)` switch *units* known at `open()`:
   a slide, a sheet, the whole DOCX document. Moving within the rendered unit
   is `goTo<Place>()`: XLSX `goToCell`, DOCX `goToHeading` and the new
   `goToPage`. This is how Word splits it: headings are structure, pages are
   layout (Print Layout only). So DOCX `count` stays `1` in both pagination
   modes, and `goToPage` rejects `unsupported-feature` in flow mode, where
   `pageCount`/`pages` are `0`/`[]`.
2. **Every `goTo...` returns `Promise<void>`** and rejects `not-found`.
3. **A unit reference is `{ index, title }`** (`UnitRef`), in info objects,
   unit events and search matches.
4. **A cell location is `{ address, row, col }`** (`CellRef`).
5. **Title lookups** are exact, then case-insensitive equality, never a
   substring (`findByTitle` in `core/search.ts`); substring matching is
   `search()`'s job alone.
6. **Events describe changes.** No `indexchange` on load; `zoomchange` only
   when the value changes; `goTo(current)` is a no-op. The ready signal
   carries the starting state.
7. **Options: live chrome in `appearance`, structural options top-level
   under one name** (`pagination`, not `docx: { pagination }`).
   `setAppearance` *replaces* the object everywhere. `sheet.read({ range,
   signal })` and `paintSlide({ model, target, ... })` became object forms.
8. **One error model**, including a new `"closed"` code (use after `close()`
   is a bug to surface, not a cancellation to ignore); the element fails with
   the same `AaredError` codes as `preview()`, fired once and rejected with
   the same instance.
9. **After `destroy()`** every method is a no-op; `destroyed` is on every
   handle. `XlsxViewer` extends the shared `RenderHandleBase` and is no
   longer exported.

**Zoom belongs to the facade and survives navigation.** `render()` and
`preview()` accept a starting `zoom` (no event); the hooks and element carry
it across `goTo()`; every handle reports the painted `scale`. The fit math is
shared (`measureFitRoom`, `fitScale` in `render-handle.ts`): room is measured
with the content set aside, so a fit can't read back its own result (the bug
that made DOCX `fit-page` a no-op in a container with no height). PPTX width
fits are CSS-fluid (`width: 100%` + `aspect-ratio`, only the scene's
transform set from script); sizing the viewport in pixels inside a
`ResizeObserver` callback looped in React hosts.

**XLSX opens at 100%, like Excel, and its fit modes are real.** The grid
scales with CSS `zoom` on the sheet host; fits measure the sheet's used area
(never the viewport-fill cells) against the room outside the host. A 25%
floor keeps a 5,000-row sheet from painting every row.

**DOCX `info` and the outline.** `info.titles` is one entry (the core
properties title, else `"Document"`) and `sizes` is gone: a per-unit array
must be `count` long. The outline is one index space (`docx/outline.ts`):
body headings only, what Word's Navigation pane lists, shared by
`headings.list()`, `goToHeading`, the facade and search matches. PPTX
`titles` stay `"Slide N"`: scanning title placeholders would break `open()`'s
"200 slides open as cheaply as 2" contract.

**Facade controls.** `PreviewHandle` gained `xlsx` controls beside `docx`,
each one stable object per render; `preview()` follows XLSX tab clicks into
`index`/`indexchange`. The element gained read-only `format`/`info`, a `type`
setter, and zoom driven through the render handle.

**`::part(page)` changed for PPTX.** The part stretches to the stage and the
slide inside carries the drop shadow, so the slide can fit the stage height;
a consumer's `--aared-background` / `--aared-page-shadow` win in the dark
theme too. `aared/worker` is now exported (the documented `?raw` import had
no export to resolve).

## 32. Measured numbers on the home page, and how they're measured

**Why.** The home page had no evidence for its claims, and `BENCHMARKS.md` was
a skeleton of `TBD` cells. Rule: no hand-filled numbers. The "By the numbers"
section is drawn from generated results, and every caption is computed from
them (a slower result reads as slower).

**One page, one browser.** `bench/measure.html` is driven by
`tools/bench-measure.ts --run` through `Bun.WebView` (Chromium). The page
times bytes-in to first paint (plus two animation frames) through the same
adapter contract the bench UI uses; aared's adapter is
`bench/src/adapters/aared.ts`, used only here.

**Memory is DevTools heap usage over the page and its Workers, not
`performance.memory`.** The page API is bucketed to 10 MB and blind to
Workers; aared parses in a Worker and silurus in WASM, so a page-only reading
flatters both. Whole-process readings were rejected (a shared WebView2
process tree). Canvas/GPU pixels are not counted, so canvas renderers are
under-reported, and the page says so. (§33 corrected how this was read.)

**Navigation is timed to the committed frame, so 33 ms is its floor** (two
frames at 60 Hz). Nearly every library lands on the floor, so the home page
leaves navigation out; per-move numbers stay in `BENCHMARKS.md`.

**Search is measured in the browser, aared only** (the others have no
`search()`), replacing the earlier Bun-runtime figures.

**Result.** Sizes count aared's format module + shared core + parse Worker:
one format is 108-131 KB gzipped (larger than the lightest single-format
libraries), while all three together (168 KB) are lighter than stitching the
lightest library per format (277 KB). aared opens the 300-slide and 13-sheet
files in ~48 ms (baselines 1.3-24 s) but uses more memory than the smallest
baselines on small files. All of it is on the page.

**Limits.** One machine, one engine, one run; fidelity is out of scope (the
paired Office captures own it). A full pass takes about 30 minutes on an idle
machine, so it runs rarely, before publishing numbers or after a change meant
to move them; the site shows the last full run, dated. `--run
fixture:renderer` re-measures single cells.

## 33. The benchmark rulers were wrong; corrected before any product change

Overturns §32's memory reading and size method (its page, one-browser setup
and computed captions stay).

**Memory never read a Worker.** `bench-browser.ts` sent
`Runtime.getHeapUsage` through `Bun.WebView.cdp()` with the Worker's
`sessionId` inside the command's params. `cdp()` is scoped to the view's page
session and ignores it: the "Worker" reading was byte-identical to the
page's. So every renderer that parses in a Worker (aared, silurus, aiden0z)
had its page counted twice and its Worker never; the flat ~2 MB floor under
aared's small files was the page read twice (an empty Worker measures
0.27 MB).

Fix: `tools/cdp.ts` starts the browser with a DevTools port (`url: false`, so
it never attaches to a browser the user has open), holds one browser-level
WebSocket, attaches to each page and its Workers with flat sessions, and sends
every command with a top-level `sessionId`. Each session proves its realm
(`typeof WorkerGlobalScope`) before its first reading, and `--calibrate` must
read 10 MiB held on the page and 10 MiB in a Worker as 10 MiB each, counted
once, before results are trusted.

**Memory samples were contaminated.** The loop began right after the last
warm run's destroy with one GC, so a sample could include the previous open's
garbage (the negative minimums). Pages now settle (the Worker set holds
still, two GCs) before every reading.

**Memory gains two symmetric readings.** *First open*: a fresh page's first
open, so loading the library's code counts for everyone. *After a full read*:
every unit visited or the document scrolled to its end, so a saving that only
holds for the first screen can't read as a win.

**Size weighed different things for different libraries.** Competitors were
the resolved entry file alone (runtime dependencies left out, barrel entries
scored as stubs, some resolved to npm-latest versions the bench never runs),
while aared was weighed unminified with core chunks a format never loads. Now
size is what the browser fetches to open each real fixture, in a fresh page
with caching off, from the same Vite production build the timings use: one
minifier, dependencies in, versions as installed (the run fails if a package
resolves outside the workspace). Every renderer gets a row. Sizes are decimal
KB.

**Harness work left the competitors' cells.** The SheetJS adapter and the
docx-preview lineage loaded JSZip inside the measured path for harness
probes. Both now read the one part they need with `DecompressionStream`
(`bench/src/lib/zip-entry.ts`) before anything is measured, which can only
help the competitors.

**What the corrected rulers say** (published before any product change, see
`BENCHMARKS.md`): aared is the smallest download for XLSX and for all three
formats, about 8 KB behind docx-preview + jszip for DOCX, and well behind
mdurbar for PPTX (mdurbar draws about half of Office's preset shapes, has no
RTL, and parses on the main thread). On memory the costs are a Worker script
stored two bytes per character (§34), a spreadsheet grid of heavy cells, a
DOCX Worker that keeps parse state nobody reads (§35), and a DOCX layout that
builds every page at once (§37).

## 34. Shipped JS is Latin-1

Nine characters above U+00FF in the Worker's code (five em dashes in error
messages, three default bullets, one non-breaking hyphen) made V8 hold the
whole ~128 KB Worker script two bytes per character, in every Worker. Error
messages now use plain punctuation; the characters aared must emit (bullet,
non-breaking hyphen, LRM/RLM, the footnote return arrow) live in
`src/core/chars.ts`, built at runtime from ASCII escapes with `JSON.parse`.
Neither obvious spelling survives a build: a `\u` escape is printed back raw
by Vite's esbuild (`charset: "utf8"`), and rolldown constant-folds
`String.fromCharCode` into the raw character (the new dist test caught it).
Shipped `.mjs` also drops doc comments (rolldown `comments.jsdoc: false`):
`aared/worker?raw` and no-bundler pages run dist verbatim; the `.d.mts` keep
the docs. Guards: `test/shipped-latin1.test.ts` scans dist for wide
characters and wide `\u` escapes; `tools/check-consumer-build.ts` scans what
a consumer's Vite actually emitted.

Measured (heap snapshot of a PPTX Worker): the script's external string fell
from ~310 KB to ~186 KB — about 125 KB less per Worker. `Runtime.getHeapUsage`
does not count external string data, so the benchmark's memory figure does
not move; the saving is real but is not claimed there.

## 35. The DOCX Worker keeps nothing the page holds

A DOCX session kept the whole `w:body` tree, the footnote and endnote trees,
and a cached copy of the parsed sections that `aared/docx` never asked for
again (it reads unit 0 once at open and searches its own copy). The tree's
strings were slices of the decoded `document.xml`, so on the 40x document the
Worker held the whole part plus the tree plus a duplicate model: 14.9 MB of a
52.7 MB total.

Now the index is light (styles, numbering, theme, relationships, title, the
notes' part names), `parseDocxSections` reads the notes itself and streams
`w:body`'s children through `parseXml`'s stream mode (the body never exists as
one tree), and the Worker parses unit 0 on each request instead of caching
it. Strings that outlive a parse go through `ownString` / `ownTree`
(`core/xml.ts`): engines keep a slice of 13+ characters as a view of its
whole parent, so one kept attribute value keeps a whole part alive. That
covers the styles and numbering the session keeps, the title, and the
module-level tag-name tables (capped, like the intern table). The parsed
sections are byte-identical to before.

A missing `w:body` now fails in `open()`'s unit read instead of its index
step (same message), and `open()` closes the session when that read fails.
§18's "DOCX second read: cached" no longer holds: a second `unit(0)` from a
raw core caller re-parses. Measured: large-40x 52.2 to 38.7 MB; open time
unchanged.

## 36. Renderer defaults are classes in one adopted sheet; a lighter grid

Every XLSX data cell wrote ~20 inline longhands and a positioned, z-indexed
text span even when empty, and the viewport fill was one styled cell per fill
column per row: 616-832 elements at about 3 KB each on the small workbooks,
most of their memory. An inline declaration block is a style attribute plus a
CSSStyleDeclaration per element, and positioned or z-indexed elements each
get a paint layer.

`src/class-sheet.ts` is the shared home: `cssClass()` interns a block of
constant declarations at module scope into one rule, and each render root
(the document, or the shadow root a render lives in, `<aared-viewer>`'s
included) gets one constructed `CSSStyleSheet`, counted by the renders using
it, adopted before a render attaches its DOM, refilled with one `replaceSync`
only when classes were added, and removed when the last render on that root
is destroyed. `follow()` moves it when a target rendered detached is attached
later and restores it if an embedder replaces `adoptedStyleSheets`. No
`<style>` text is written, so a strict `style-src` needs no
`'unsafe-inline'`. Rules are `.aared-<hash>:not(#aared-0#aared-1)`
(specificity 2,1,0), so an embedder's own CSS beats a default only with
`!important` or a selector with two IDs plus a class, close to what inline
styles gave. Inline style keeps only what a node authors, and still wins.
Rejected: `<style>` text (fails under CSP), `insertRule` per page (slow),
custom properties alone (still one style attribute per element),
attribute-keyed rules (weaker specificity, a second mechanism).

The XLSX grid uses it throughout: cells carry classes; an empty cell paints
no text span, `dir` or alignment; only a cell under a data bar or with a
spill overlay is positioned; a row's run of fill columns is one cell, and a
whole fill row one cell, with inner lines drawn as gradients one device pixel
thick at the current zoom (`--aared-hairline`). Hot `dataset` writes became
`setAttribute`. Unchanged on purpose: per-cell `tabindex="-1"`, fill column
headings, unclipped wrapped cells, and a fill row wider than HTML's
1000-column span limit keeping one cell per column. DOCX paragraphs,
headings, list markers and pages take their defaults from classes too; PPTX
is unchanged.

Measured (repeat open, page + Worker): post-production-tracker 3.22 to
2.12 MB; a sheet switch 42 to 33 ms (the frame floor).

## 37. DOCX layout pages paint lazily

Layout mode built all 121 pages of the 40x document at once (23,262
elements, 33.9 MB) and its tab-stop pass forced layout of every page. Pages
now carry `content-visibility: auto` where the browser supports it: every
page stays in the DOM (find-in-page, selection, fragment links, printing),
but a page off screen skips style and layout until scrolled near; its fixed
box already sizes it. Tab stops lay out per page the first time a page is
shown (`contentvisibilityautostatechange`) and again once web fonts land;
without the feature the eager pass runs as before. The `span.paragraph-text`
wrapper is gone from DOCX. Windowing pages out of the DOM entirely was set
aside: it would cost browser find on unseen pages and rework navigation,
captures and tests.

Kept on purpose: the paginator's metrics cache stays keyed weakly on the
read-only model, so a re-render of the same document (theme toggle, mode
switch) does not re-measure every block.

Measured: large-40x 38.7 to 16.9 MB, open 606 to 230 ms. The saving is
first-view: after scrolling to the end the document holds 34.2 MB, because a
page keeps its layout once shown (the "after a full read" column reports it).

## 38. The Worker is reused between opens

Every `open()` started a Worker and every `close()` terminated it, so every
document paid for a new isolate and for compiling the parser again (~0.5 MB),
and `preview()` / `<aared-viewer>` paid twice (once to detect the format,
once to open it). Now a closed document's Worker is parked for the next open
(`core/session.ts`, `acquireBridge`): one idle Worker per worker init (the
default, or a Worker URL), stopped after 15 s idle or when a newer one parks;
two documents open at once still get two Workers. Only a clean Worker is
parked (nothing cancelled, failed to post, or in flight at close), because a
cancelled synchronous parse can only be stopped by terminating. A Worker
built by a caller's `() => Worker` factory is the caller's: used once,
terminated on close. `close(sessionId)` rejects only its own session's reads.
The literal `new Worker(new URL("./worker.mjs", import.meta.url), { type:
"module" })` branch is untouched (§19). State that now outlives a document is
bounded and owns its strings: the XML name tables (§35) and the number-format
memo.

For a single document this mostly moves the Worker's fixed cost from the open
into the page's baseline, like a main-thread library's loaded code. So the
benchmark reports first-open memory next to repeat-open memory, and the idle
Worker a page keeps between opens as its own column. A closed document's
memory is released when its Worker next collects garbage or stops, not the
instant `close()` returns.

## 39. Bounded Worker caches; minified dist

A session cached every slide or sheet it ever parsed, so visiting every
slide of a deck kept every model (and, through sliced strings, every part's
text). Now slides and sheets sit in a small most-recently-used cache
(`core/unit-cache.ts`: 4 slides, 2 sheets; an evicted unit re-parses), and a
read that won't paint (search) is served from the cache but never fills or
refreshes it (`retain: false`), so a first search doesn't push out the unit
on screen. The workbook's shared strings and the theme's background fills,
kept for the session, hold owned copies (`ownString` / `ownTree`).

dist ships minified (worker 70 → 41 KB gzip, with a source map): consumer
bundlers re-minify, but the `aared/worker?raw` Blob recipe, no-bundler and
CDN pages run dist as-is. The fair download size (§33) doesn't move, since it
was already measured from a minified consumer build.

Weighed and left for now, with the reason: one Blob per shared PPTX picture
instead of an ArrayBuffer plus a Blob copy per render (the PPTX page already
holds ~0.15 MB, the lightest of any renderer); moving the PPTX-only geometry
evaluator's `import()` out of the shared session chunk (it is never fetched
for XLSX or DOCX, so the download a page makes doesn't change); dropping the
PPTX layout and master trees for a detached skeleton (a larger change to
placeholder inheritance for a small, per-layout saving).

## 40. No layout jumps when a unit is replaced

Three ways a viewer visibly jumped when the next slide or sheet replaced the
current one:

- **The page scrolled to the top on Next.** A new slide went into its
  container with no height (its content is absolutely positioned), so the
  fit pass shrank a page whose height was the slide and the browser clamped
  its scroll position. The slide now goes in at its full-width box (`width:
  100%` and its aspect ratio), the largest any fit makes it, and the fit sizes
  it. `measureFitRoom` also puts every scrolled ancestor's position back after
  setting the content aside.
- **A short sheet showed without its fill, then grew.** The viewport fill
  needs real sizes and the first window was deferred to the next frame. It
  now measures and paints in the same task once attached; a target not yet in
  a document still paints on the next frame.
- **The fill grew a row per frame in a `max-height` box.** A scroll box as
  tall as its content up to a `max-height` measured the grid already inside it
  as its room, so each pass added rows until the limit stopped it. Its room is
  now the limit.

The playground's demo stages had a 12 rem minimum height that "fit page"
fitted the slide to, so some demos drew smaller slides than the element: the
minimum now applies only while a stage is empty.

## 42. npm name: `@aared/viewer`, one package, reserved with a placeholder

The plain name `aared` cannot be published: npm's similarity check rejects it
against an unrelated, empty `aa_red` (403 "too similar"). Scoped names skip
that check, so `aared` is an npm org and the package is **`@aared/viewer`**
(matches the `<aared-viewer>` element). `@aared/viewer@0.0.0` is a
placeholder holding only a README, published 2026-09-25; the repo's code is
not in it. The repo tag `v0.0.0` marks that point.

**One package, subpaths, stays** (§20): `@aared/viewer/pptx`, `/xlsx`, `/docx`,
`/preview`, `/react`, `/element`, `/worker`. Per-format packages were
rejected again: the single Worker holds all three parsers, so splitting would
need per-format workers (a 20-26 KB saving for one-format pages, at the
packaging cost §20 records) plus a shared core package to release in step.

**Versions.** The placeholder is `0.0.0`, never `1.0.0` (a stub must not look
stable; a published version can never be reused). The first real release is
`0.1.0`: semver allows breaking changes in a `0.x` minor. `1.0.0` comes when
the API is frozen (ROADMAP "Before first publish"). The "best design wins over
backward compatibility" stance in AGENTS.md ends at the first real release,
not at the placeholder.

**Publishing is done by a person.** It is public and permanent, the account
needs 2FA on writes, and agents never hold npm credentials or tokens.

## 43. DOCX right-to-left tables: `w:bidiVisual`, logical borders, physical margins

Found while building the showcase term sheet: every right-to-left table Word
wrote painted its columns left to right. The parser looked for
`w:tblBidiVisual`, an element Word never writes; the real one is
`w:bidiVisual` (ECMA-376 17.4.1). No fixture had a right-to-left Word table,
so no test saw it.

Word keeps writing `w:left`/`w:right` inside such a table, so the sides needed
a ground-truth probe (`tools/probe/docx-bidi-visual/`): the same table saved
left-to-right, then switched to right-to-left in Word, then authored
right-to-left from the start, each exported through Word with span geometry.

- Switching a table to right-to-left changes only `w:bidiVisual`; no side is
  rewritten.
- Borders (`tblBorders`, `tcBorders`) are logical: the `w:left` border paints
  on the right edge of a right-to-left table, and setting the physical left
  border through Word's object model writes `w:right`.
- Cell margins (`tblCellMar`, `tcMar`) stay physical: a 0.5in `w:left` margin
  kept each cell's text 46px from the cell's left edge after the switch.
- `tblInd` is measured from the start side (the right). aared now reads it
  (`DocxTableNode.indent`) and paints it as `margin-inline-start` on the
  table grid, which carries the table's direction.

So the parser reads `w:bidiVisual`, the DOCX painter keeps drawing side
borders on the inline start/end, and cell padding moved from
inline-start/end to physical left/right. PPTX is untouched: it already pads in
physical sides and reads `a:tblPr rtl`. Pinned by section 5 of
`fixtures/docx-coverage.docx`.

## 44. Fixes found by the showcase documents (Sep 2026)

Building polished showcase files for the launch video (PowerPoint, Excel and
Word output compared page by page with aared) surfaced real bugs. Each was
traced to its cause, checked against native Office and pinned in the existing
coverage fixtures (edited through Office COM; no new fixture files).

**PPTX lines and connectors were dropped.** PowerPoint writes every line and
connector as `p:cxnSp`; the shape-tree walk only knew `p:sp`, so timelines and
dividers vanished. A connector is a shape without text, so it now parses
through the shape path; `straightConnector1` joins `line` in the hand-coded
fast path. Shape outlines also paint authored `a:prstDash`, scaled to the line
width. Pinned by `pptx-coverage.pptx` slide 11.

**Headings inherited the browser's `<h1>` bold and 2em.** DOCX headings paint
as `<h1>`-`<h6>` for outline semantics. The shared block class now resets
weight and size; runs decide, as in Word.

**Left-to-right runs in right-to-left paragraphs (probe run-direction).** Word
and PowerPoint keep a run that is not right-to-left (DOCX: no `w:rtl`; PPTX: a
left-to-right `lang`) together left to right: "71%", "2.1%" and "15 Oct 2026"
show as written. Plain Unicode bidi detaches "%" and splits dates. Such runs
(with no Hebrew/Arabic letters) carry `TextRunModel.ltrRun`, and the shared
run painter adds a left-to-right mark at both edges through
`::before`/`::after` (generated content, so never copied or searched).
Isolating each run was rejected: a date PowerPoint splits into two en-US runs
would reverse. Deliberate non-changes: right-to-left runs keep plain bidi
(Word and PowerPoint disagree there), and a left-to-right run holding Arabic
words in a left-to-right paragraph still needs its own probe. The PPTX
renderer now adopts the shared class sheet, as DOCX and XLSX already did.

**Long DOCX tables no longer blank a page.** A table taller than a page used
to close the page it would start on and go whole onto the next, its tail
clipped. Word breaks tables between rows. The paginator now splits a table at
the last row that fits (rows are never split; a boundary inside a vertical
merge is skipped); only an unsplittable table still takes the oversize path.

**A table split must also yield to `w:keepNext`.** Word treats a row whose
paragraphs all carry `w:keepNext` as kept with the next row, so a title row is
never split from the row it introduces. `DocxParagraphLayout.keepNext` is
parsed and carries through headings, list items and table rows; the
paginator's keep-with-next chain (a table's leading run of `keepNext` rows
counts as one piece) moves to a fresh page when the chain fits one but not
what is left of the current one. A generator's "keepNext on every row but the
last" is not licence to relax the row-boundary gate: that let a title row
split from its first row, which Word never does.

**Table rows and other blocks are measured, not only estimated (§21).** The
estimator's average glyph-width guess was worst on dense bilingual tables.
`docx/index.ts` `measureFlow()` lays every section's top-level blocks out once,
off screen, at the section's real body width, and reads back each block's
margins and border-box height and each table row's grid-track height (table
cells lay out as CSS grid rows). `paginate()` takes this as
`PaginateOptions.measured` and prefers it wherever it answers; the estimator
stays as fallback (headless use, blocks not placed standalone). The remaining
height gap on heading-styled rows was the exact-pitch stretch below, and the
cell-inset fix further down. With both, the showcase term sheet measures 10
pages against Word's 10.

**Exact line pitch is the paragraph's, not each run's.** `w:lineRule="exact"`
(and PowerPoint's `a:spcPts`) set a paragraph's line height, but every run
still painted with its own line box centred on its font's ascent and descent,
so a Calibri run beside Sakkal Majalla stretched an exact 24px line. Word and
PowerPoint keep an exact line at its authored pitch whatever it holds. A run
on an exact-pitch line now paints at `line-height:0` (`EXACT_PITCH_RUN`,
`text-style.ts`), so the paragraph's own strut sets the line;
`exactLinePitch(box)` decides which paragraphs qualify (`lineSpacingPt`
without `lineSpacingAtLeast`), and `docx/index.ts` and `pptx/index.ts` apply
it wherever a run paints. An `atLeast` line is a floor, not a pitch, and is
untouched. PowerPoint's `a:baseline` shift now paints in `em` (DrawingML
defines it as a share of the font size) instead of a `vertical-align`
percentage, which CSS takes of the run's own line height (zero on an
exact-pitch run). Pinned by both coverage fixtures: the DOCX tracker update
line is exactly 18pt with Calibri beside Arial Arabic, and the PPTX "Mixed
direction" box on slide 11 is exactly 30pt.

**Line rule kept, and a direct-formatting merge bug fixed alongside it.**
Two faults in the paragraph-property resolver. First, `w:lineRule="atLeast"`
was read but discarded, so a floor and an exact pitch looked identical, and a
style's `auto` multiple over a docDefault's exact pitch left both set instead
of the later level replacing the earlier. Second, `resolveParaProps` parsed
direct formatting into a throwaway object and copied it key by key onto the
merged style chain, which could add or overwrite a key but never clear one
(a styled top border plus a direct left `w:pBdr` lost the top border). Now
`DocxParagraphLayout.lineSpacingAtLeast` is set and cleared by the opposite
rule in the same `w:spacing` branch, direct formatting resolves straight into
the merged object (`readParaProps(directPPr, merged)` last), and `styleId` is
read from `directPPr`'s own `pStyle`. The estimator's `textBlockHeightPx`
treats `atLeast` as a floor: `max(pitch, Word's single line)`. Blast radius:
accumulating properties (tab stops, `w:pBdr` sides) now compose under direct
formatting too.

**Table cell side insets follow Word's own centred-border rule.** Probe
`tools/probe/docx-cell-border/` (side borders at 0.5pt / 3pt / 6pt): a side
border moves the cell's text 0.3 / 1.4 / 2.45pt from the grid line. Word
centres a border on the shared grid line and keeps the text a margin beyond
its own inner half, whichever cell paints it. aared took the whole painted
border out of the padding, double-counting the outer half. Now
`sideInsetPx(margin, edge, painted)` is the margin plus half the edge on that
side, less what this cell itself paints there. A shared inside divider is
painted once, by the later cell as its start border; its neighbour keeps the
full margin plus half. PPTX tables are unaffected: they paint edges as a
centred overlay (`cellEdgeOverlay`) that takes no text width.

**Word 2013 justified space shrinking, approximated conservatively.** Probe
`tools/probe/docx-justify/` (real Word, Courier New): in
`w:compatibilityMode` 15 or later, a justified or distributed paragraph (LTR
or RTL) fits one more word on a line by shrinking each space, by up to about
a quarter and about a third of the word it fits in total; never when
left-aligned, never in mode 14. CSS has no shrink-then-refit pass, so the
parser reads `w:compat/w:compatSetting[@name="compatibilityMode"]`
(`wordLayoutShrinksSpaces`, `DocxIndex.justifyShrinks`, document-wide) and
sets `DocxParagraphLayout.justifyShrink` on justified/distributed paragraphs
only; the renderer adds `word-spacing: -0.01em` (`JUSTIFY_SHRINK`). That
stays under both of Word's limits for a line of up to about twenty spaces, so
it only catches near misses (a justified Arabic cell broke one word to a
fourth line for 0.6px). Known limits: one row still breaks a word earlier than
Word, and English Calibri measures 0.3-0.5% wider in Word (glyph-advance
rounding), deliberately not faked with letter-spacing.

**A `w:lastRenderedPageBreak` mid-paragraph is not a break point.** Word
writes one where its own last layout put a page break, which can fall inside a
run's text. The hint is now read only when no text of the paragraph has been
seen yet (a preferred break before the paragraph); a mid-paragraph occurrence
is dropped, since it records Word's stale layout cache. Inside a table cell it
is dropped too (tables break between rows).

**tools/excel-render printed US Letter, edge to edge.** Excel's "Scale content
for A4 or 8.5 x 11 paper sizes" (`MapPaperSize`, on by default in US-locale
installs) mapped the A4 policy to Letter, and margins were set as inches where
Excel reads points. The exporter turns `MapPaperSize` off for its run
(restoring the user's setting), commits page setup as one batch, and converts
margins with `InchesToPoints`.

**A frozen row's own row number scrolled away from its row.** Excel keeps a
frozen row's number with it. The row-header cell only stuck to the
inline-start edge (`paintFrozen`); it now also gets the row axis's `top`
offset (`frozenRowOffset`) and z-index a frozen data cell gets. Pinned by
`xlsx-coverage.xlsx`'s Shot Log sheet.

**A content-sized sheet let its container scroll it.** `<aared-viewer>` and
the bench give the sheet no height of its own (as in Excel, whose grid scrolls
itself inside the window and is never scrolled by its container), and
`syncViewport` only grew the box to its full content height. So the card
scrolled the sheet: frozen rows, numbers and column letters scrolled away
(`position: sticky` pins to the nearest scrolling ancestor), and the
horizontal scrollbar sat below the fold. `syncViewport` now also caps the
scroller's `max-height` to the clipping ancestor's visible room when that is
at least 120px (`MIN_SCROLL_ROOM_PX`; below that the container scrolls as
before). With no clipping ancestor nothing changes. The cap must be lifted
before measuring (only when set) and the *uncapped* height recorded for the
next comparison, or the change detection reads its own stale cap and misses a
resize; `tools/xlsx-render`'s print capture, which grows the scroller to an
oversized height on purpose, hit exactly that.

## 45. Font fidelity: a delivered substitute is actually used, waited for, and reported (Sep 2026)

Follow-on to §44's font-metrics work, prompted by a direct question: layout
fidelity depended on the calibration machine having every requested Office
font installed. On a machine without them, what happens? Investigating found
three real bugs in `aared-fonts` and the renderer, and a substitution table
that needed correction. (§53 later froze `aared-fonts` and made the browser's
own fonts the default; the alias, wait and warning rules below still hold for
any delivered face.)

**A delivered substitute was registered but never selected.** `createResolver`
built each `FontFace` under the substitute's own name (`"Carlito"`) while
every renderer painted `font-family: Calibri`, a single bare name. On a
machine without Calibri the browser fell to its generic default and the
registered face sat unused. `fontAlias(family)` (`core/typography.ts`,
exported from `aared/core`) is now the one shared naming rule:
`cssFontFamily` stacks the requested name, `fontAlias(requested)`, then the
nearest generic, and `createResolver` registers every delivered face under
`fontAlias(requestedFamily)`. Never under the requested name (it would
shadow a reader who really has the font) and never under the substitute's own
name (nothing would select it). A visual test paints a run in a family that
exists nowhere, registers a `local('Courier New')` face under its alias and
confirms the painted width is Courier's.

**The Latin substitutes were not hosted, and the Arabic ones were not picked by
measurement.** Liberation Sans/Serif/Mono return 400 from Google Fonts;
Arimo/Tinos/Cousine are the hosted builds of the same metric-compatible family
(0 of 95 ASCII advances differ from real Arial/Times New Roman/Courier New,
`tools/font-calibration/metric-clone-check.py`; Calibri/Carlito likewise;
Cambria/Caladea under 0.13%). The four Arabic entries had been mapped by name
association. They were re-picked by measurement: real shaping through
`CanvasRenderingContext2D.measureText` (an unshaped glyph sum is wrong for
Arabic), 59 samples (every Arabic-only run in the fixtures) in the real
Windows font files loaded through `FontFace`, against eleven candidates from
the files the resolver delivers, each load verified (a silent fallback would
look plausible). Ranked by the spread (sd) of per-sample width ratios, because
line breaks depend on the spread while the mean is what `size-adjust`
removes; then by alef and seen heights after adjusting. A candidate had to
beat the incumbent by a quarter. Outcome: Sakkal Majalla uses Markazi Text,
Simplified Arabic uses Lateef, Traditional Arabic and Arabic Typesetting keep
Amiri. No open metric clone of any of the four exists (by search only);
Microsoft Typography describes Arabic Typesetting as a modern Naskh, not
"decorative/Diwani". A large `size-adjust` (Lateef 130%) did not distort the
letters after adjusting. Bold differs from regular by up to 5.6% in width, so
`FontCalibration` has an optional `bold` set and the resolver picks by face
weight. Survey script: `tools/font-calibration/arabic-survey.ts`.

**Calibrating one delivered face against several different real fonts.** Each
requested family has its own `FontFace` descriptors from the measured width
ratio: `size-adjust` scales the substitute's advances, and
`ascent-override`/`descent-override` make the resulting line box match the real
font. Measured directly (`tools/font-calibration/size-adjust-interaction.ts`):
`final_ascent = ascent-override% x size-adjust% x font-size`, so the published
`FONT_CALIBRATION` overrides are divided out from that. Families sharing a
substitute still register as separate `FontFace`s, one per alias, each with
its own calibration; the CSS fetch is cached per substitute, the registered
faces per requested family.

**Most Arabic is set in Arial, and Arial's substitute has no Arabic glyphs.**
Across the fixtures' Arabic runs, Arial carries 183, Times New Roman 12,
Aptos 9, Cambria 1; Arimo and Tinos have no Arabic subset, so that text fell to
the reader's `sans-serif` (on Windows that is Arial's own Arabic, which hid it
and misled the fidelity detector). Word's PDF export
(`tools/font-calibration/word-arabic-fallback.ps1`, span widths read with
PyMuPDF) agrees with canvas measurement within about 3%, which licenses canvas
numbers. Arial's and Times New Roman's Arabic match Simplified Arabic's widths
within 1.3%; Calibri carries its own Arabic; Cambria has none and Word draws it
in Times New Roman. So the resolver adds a second face to the same alias for
Arial, Times New Roman, Cambria, Calibri and Calibri Light, limited by
`unicode-range` to a fixed set of Arabic script blocks (Google's subset range
also claims punctuation Arimo owns) and calibrated per original; composition
was proven in a real browser first (`alias-composition.ts`), with spaces and
punctuation staying on the Latin face. Only upright Arabic faces are
registered: neither real font has an Arabic italic, so the browser slants
synthetically as Word does.
Latin-only documents must not fetch the Arabic face (about 130 KB regular plus
bold). It is not `.load()`ed eagerly, and the font wait and the fidelity probe
(whose canvas `measureText` itself triggers a fetch) only include Arabic
letters when `collectFontUse` reports Arabic text in the model. A test
guards both. Deliberate non-changes: Courier New's Arabic is monospace (every
proportional candidate is off by 45% or more, so none is offered); Tahoma and
Segoe UI are not in the table; Aptos was measured later (§49).

**A registered substitute's glyph data can still be loading when a renderer
measures painted text.** `createResolver` calls `FontFace.load()` without
awaiting it, and the session only awaits the resolver settling, which happens
before the bytes arrive. DOCX layout `measureFlow` read painted geometry
against whatever was already loaded. `awaitDocumentFonts(model)`
(`text-style.ts`, shared by both formats) now walks the model for every
`fontFamily` (`collectFontUse`) and calls `document.fonts.load()` on each
family and its alias, racing a 2-second timeout so an offline CDN never hangs
a render. DOCX layout pagination always awaits it before `measureFlow`; DOCX
flow mode awaits it for the fidelity flag, which reads the font that actually
painted. (PPTX's shrink pass and its wait are gone, §48: PowerPoint never
measures on open.) A real race needs a genuine network `url()` fetch:
`document.fonts.load()` on an unregistered family, and a `FontFace` from an
`ArrayBuffer` or `local()`, resolve synchronously. The regression tests
therefore serve real font bytes from a deliberately delayed same-origin route.

**The fidelity flag.** Every render handle carries `.warnings: readonly
RenderWarning[]` (`{code, message}`, `core/model.ts`, via
`RenderHandleBase.setWarnings()` in `render-handle.ts`). DOCX layout mode
merges the paginator's `oversize-block` and `footnote-overflow` (previously
dropped before reaching the embedder) with font entries; DOCX flow mode and
PPTX carry font entries alone. `detectFontFidelity`/`fontFidelityWarnings`
(`text-style.ts`, DOM-required, so not in `core/session.ts`, Rule 3) decide per
family whether it painted something real (the requested name or a delivered
alias) or fell to the browser default, by comparing a `canvas.measureText`
width with a same-generic control. `document.fonts.check()` is unusable: it
returns `true` for a family that exists nowhere
(`tools/font-calibration/fonts-check-probe.ts`). Windows Chromium's
`sans-serif`/`serif`/`monospace` resolve to Arial/Times New Roman/Courier
New, so width alone cannot tell the real font from the generic default (and
the pixels are right either way); `resolvesLikeGenericDefault`
(`core/typography.ts`) keeps those three from reading as unavailable. A
working substitute is not a warning (`"substituted"`); only `"unavailable"`
produces a `RenderWarning`.

**A table row Word's own saved layout broke before is real evidence
pagination uses.** `w:lastRenderedPageBreak` leading a table cell's first
paragraph was dropped (§44), though it is real signal: on two Word-saved files
the count of raw markers in `document.xml` equals `docProps/app.xml`'s
`<Pages> - 1` exactly. `DocxTableRowNode.pageBreakBefore` now carries a marker
leading the row's first cell's first paragraph (`parseTableCell` returns it
alongside the cell); a marker on any other cell, or after text, stays dropped.
`splitTable` prefers a marked, valid boundary (clean of a merge and a
`w:keepNext` chain, fits the room) over its latest-fitting guess. Honouring
the boundary is not enough: the resulting tail must also start a fresh page. A
whole table with a marked first row breaks for it only once the current page
is more than half full (the gate a top-level hint uses), but a tail
`splitTable` itself carved at a marked boundary always starts fresh, ungated,
tracked by object reference through the one `pending` queue.

**Not done in this pass** (not scheduled): a mid-paragraph
`w:lastRenderedPageBreak` is dropped, not promoted to a paragraph-level hint
(paragraphs are atomic in pagination; the next marker resyncs). Embedded fonts
(DOCX `.odttf`, PPTX `.fntdata`) are not read.

## 46. Arabic calibration measured the way it paints: through the alias, unrounded (Sep 2026)

Supersedes the calibration *numbers* in §45 (the method of choosing each
substitute stands). Found while benchmarking a proposal to replace Arial's
two-face fallback (Arimo + Lateef) with one font, Vazirmatn. The numbers now
live in `aared-fonts` (`FONT_CALIBRATION`, `ARABIC_SCRIPT_FALLBACKS`).

**Vazirmatn, rejected on measurement.** `tools/font-calibration/
vazirmatn-bench.ts` laid out every real paragraph in the fixtures (142) at
seven widths in the real Arial files and in each setup, and counted lines
starting at the same word as Arial's. Vazirmatn, fitted per script, lost on
English (97.1% against 100%; Arimo is a metric clone) and mixed paragraphs
(96.8% against 98.2%) and tied on Arabic. Its width-matched alef is about 25%
shorter than Arial's (refitted Lateef's, about 7%), so its Arabic reads
smaller than Word's, and it saves no download (per-script files, the same
count as Arimo + Lateef).

**Bug 1: the Arabic fallback faces were calibrated alone, but paint beside a
Latin face.** An `ARABIC_SCRIPT_FALLBACKS` face covers `ARABIC_UNICODE_RANGE`
only, so the spaces, Western digits and punctuation inside Arabic text paint
in the alias's Latin face (Arimo, Tinos, Caladea, Carlito) at its own size.
§45 measured Lateef and Markazi Text on their own, so as painted every
fallback was 0.9% to 3.3% too wide. `arabic-fallback-fit.ts` now fits each
through its real alias: the alias's width is linear in the Arabic face's
`size-adjust`, so two probes (100% and 80%) solve for the value that matches
the original's total width over every Arabic-only run in the fixtures (59).
Cambria is fitted against Times New Roman's Arabic (what Word draws) through
Caladea (what paints its spaces), so it no longer shares Times New Roman's
numbers.

**Bug 2: vertical metrics were read rounded.** Canvas `fontBoundingBox*`
comes back in whole pixels and the survey read it at 40px, so
`ascentOverridePercent`/`descentOverridePercent` carried up to 2.5% of the em
of error. Both probes now read at 10000px; the values match each font file's
`usWin` ascent/descent (read with fontTools), which is what a browser on
Windows paints the real font with. All nine calibrations (four main faces,
five fallbacks) were corrected.

**Proof.** Over 77 Arabic paragraphs at seven widths (`arabic-fallback-fit.ts`),
Arabic lines breaking like the original gained about 2 points on average at
both weights (Arial 90.5 to 96.2 regular); mixed paragraphs stayed flat within
noise. Every regular width lands at 100.0% of the original, and the line-box
check (`vertical-check.ts`) puts every entry within 0.1% of the real font bar
the non-changes below. Blast radius: numbers only, the resolver is unchanged.
A calibrated face paints only when the reader lacks the real font, so the
calibration machine (every Office font installed) renders nothing differently
and the visual suite cannot show it; the probes are the evidence.

**Deliberate non-changes.**
- Bold keeps §45's rule: its own numbers only where it differs from regular by
  more than about 1.5%. Arial, Times New Roman and Cambria bold differ by
  0.7-0.9%, so they use the regular numbers.
- No `line-gap-override`: aared never paints text at `line-height: normal`, so
  a face's line gap does not reach the page.
- Whose spaces Word draws inside Cambria's Arabic is unknown; a Word probe
  would settle it.
- The substitutes have not been checked on a machine without the Office fonts
  (done in §53).

## 47. Paragraph spacing in lines: what a "line" is in Word and PowerPoint (Sep 2026)

Proposed: paint Word's `w:beforeLines`/`w:afterLines` with the CSS `lh` unit
(one line = the paragraph's own line height). Rejected on measurement: neither
app means that. Probe `tools/probe/line-spacing-units/` builds paragraph pairs
through each app's object model, exports through the app and reads the gap off
the PDF.

**Word: a line is always 12pt.** At 11pt and 24pt, single and double spacing,
"1 line" added 12.0pt and "0.5 line" 6.0pt (`lh` would give 12.7 to 27.6pt).
Word saves the points beside the lines (`w:beforeLines="100" w:before="240"`),
so its own files were already right. The exception is a line grid
(`w:docGrid w:type="lines"`): Word counts the grid pitch and snaps every line
to it; a stored pitch with no grid type changes nothing. aared has no line
grid (a gap).

Precedence, from hand-edited XML: the line value and the point value inherit
separately, and a line value above zero wins even over the paragraph's own
points. `beforeLines="0"` means "in points", and lines alone need no points.
Word writes a style's line value without points and materialises `w:before` on
each paragraph, so only other producers' files differ. `resolveParaProps`
carries both through the chain and `paragraphSpacePt` applies the rule.

**PowerPoint: a line is 1.2 x the paragraph's largest run size**, in every font
tried and at line spacing 0.8, 1.0 and 2.0 (a 28pt word in an 18pt paragraph
makes the line count 28pt). `lh` matches only at 1.0. aared had dropped
`a:spcPct` spacing; it is now `TextParagraphModel.spaceBeforeLines`/
`spaceAfterLines`, painted as an `em` margin on the paragraph's `<p>`, which
already carries its largest run size.

The same probe found three more rules, fixed here:
- PowerPoint drops space before a text body's first paragraph and after its
  last, in points or lines, in shapes and table cells. aared painted both.
- One paragraph's space after adds to the next one's space before. A table
  cell with no vertical alignment was a block, where CSS collapses the two
  margins; `appendTextBody` now makes every text body a flex column.
- The Word renderer already stacks blocks apart, so Word is unaffected.

Slide 12 of `pptx-coverage.pptx` was extended through PowerPoint for this;
`pptx-paragraph-spacing.visual.test.ts` re-reads the painted gaps.

**Seen in the paired captures, not fixed here** (all in ROADMAP's
fidelity-gap table): PowerPoint lowers a paragraph's first line when line
spacing is above 100% where aared centres the extra leading; a text box that
authors no insets gets none in aared where PowerPoint uses 0.1in/0.05in;
bullets larger than their text make aared's lines taller on slide 5.

## 48. PPTX autofit is what PowerPoint stored; bullets hang; a subtitle takes the master body style (Sep 2026)

The slide-5 screenshot (a bullet list spread out down the slide) and the paired
captures of slide 1 turned up four causes with one theme: aared worked out
something PowerPoint reads from the file.

**Autofit follows the stored values, not a measurement.** Overturns the shrink
half of §45. Probe `line-spacing-units` built nine `normAutofit` boxes and
read PowerPoint's PDF: PowerPoint does not measure on open. A bare
`<a:normAutofit/>` paints at 100%, overflow painted (not shrunk, not clipped).
Only the stored `fontScale` and `lnSpcReduction` change anything. A painted
size is `round(size x fontScale)` in whole points, half up; the line step is
1.2 x that size x (line spacing - `lnSpcReduction`); spacing in points is
unscaled; spacing in lines counts 1.2 x the rounded size and ignores the
reduction. aared's old pass (measure `scrollHeight`, x0.95 steps, clamp at 0.6)
had no basis in Office and shrank slide 1's subtitle to 28.95px where
PowerPoint paints 32px. `parseTextBody` now applies both stored values and the
model carries them (`TextBodyModel.fontScale`, `.lineSpacingReduction`; sizes
and `lineSpacing` are already scaled, so a renderer only paints). Only the
slide's own `bodyPr` counts: a layout or master `normAutofit` never carries a
scale. Removed: `shrinkTextToFit`, `scaleFonts`, `hasAutofitShrinkText`, the
PPTX-only `awaitDocumentFonts` wait (DOCX keeps it) and
`data-aared-autofit-applied`. Deliberate non-changes: `lineSpacingPt` and
`bullet.sizePts` stay unscaled (PowerPoint's effect on both is unprobed).
Limit: `none` and `shrinkText` shapes still paint with `overflow: hidden`,
where PowerPoint paints the overflow.

**Bullets keep PowerPoint's hanging indent.** The marker was a grid column with
a gap, which misplaced the text whenever the marker was wider or narrower than
the hanging gap, and a marker larger than its text made the line taller. Read
off PowerPoint's PDF: the marker starts the first line at the margin plus the
(negative) first-line indent, and the text starts at the margin or right where
a wider marker ends. The marker is now an inline box at least as wide as the
hanging gap, with no line height of its own. Along the way: `marL` counts from
the right in a right-to-left paragraph (`margin-inline-start`); an empty
paragraph is still one line at its own size; a `+mj-lt` bullet font is the
theme's heading font (`cssFontFamily`); the first-line indent is painted for
every paragraph (`applyParagraphBox` lost its `firstLineIndent` option).

**A subtitle placeholder takes the master's body style.** `masterTextStyle`
switched on the raw placeholder type, so `subTitle` (and content, table,
picture and chart placeholders) fell through to `otherStyle`, while the
master-placeholder lookup beside it already sent them to the master body,
which carries the 90% line spacing and 10pt before. Both lookups now share
`masterPlaceholderType`; only date, footer and slide number take
`otherStyle`. Blast radius: every placeholder that is not a title or those
three, in every PPTX; the layout's `lstStyle` and the slide's own properties
still sit above the master style, so nothing that authored its own values
moved.

**Seen, not fixed.** At 90% line spacing aared's first line still sits about
3px lower than PowerPoint's; where PowerPoint puts the leading is unprobed
(same family as the gap in §47). In ROADMAP's fidelity-gap table.

## 49. Aptos: calibrated look-alikes measured against the real cloud fonts (Sep 2026)

Aptos has been Office's default theme font since 2023 (Aptos for body, Aptos
Display for headings, Aptos Narrow as Excel's default). It had no substitute,
because §45 thought there was nothing to measure against. No browser has
Aptos, so its text painted in the generic sans-serif; on Windows that is Arial,
5-17% wider than Aptos (slide 1's "Next review: Thursday 12:00": 384.8px in
Aptos, 407.3px in Arial). The per-font numbers live in `aared-fonts`; §53
later froze that package.

**The real fonts were on the machine all along.** Office downloads cloud fonts
into `%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts` and never installs them
as system fonts. Files are named by hash, so
`tools/font-calibration/cloud-fonts.ts` finds them through their name table.
They are used only for measuring and never shipped or copied.

**Choice.** §45's method, applied to Latin text. `latin-survey.ts` measured
every Latin run in the five fixtures (194 samples) in the real Aptos, Aptos
Display and Aptos Narrow against 61 open candidates, each loaded with the
resolver's own request (`fontCssUrl`) and verified. Width spread alone did not
decide it: the narrowest spreads came from condensed faces whose letters end
up 12-33% taller once `size-adjust` matches widths. So the rule is to keep
Aptos's x-height and cap height within 4% at both weights (3% was too strict
for Aptos Narrow), and exactly one candidate per family passes: Aptos uses
Radio Canada, Aptos Display uses Roboto, Aptos Narrow uses IBM Plex Sans
Condensed. `latin-fit.ts` then laid out real paragraphs at eight widths:
paragraphs breaking exactly like Aptos are 93-96% with the picks, against
38-75% in the generic sans-serif. Aptos Display (Aptos spaced tighter for
large sizes, which a uniform `size-adjust` cannot reproduce) takes its own
face.

**Numbers.** Each family has its own `FONT_CALIBRATION` entry. Bold carries its
own numbers in all three, though weights differ by only 1-3%, departing from
the Arabic entries' 1.5% bar on measured cost (at the regular numbers Aptos
Narrow's bold paints 1.3% wide). The vertical overrides are Aptos's typo
ascent and descent (0.939/0.282 of the em), because Aptos sets
USE_TYPO_METRICS and the browser measured those, not the win metrics the other
calibrations use. The line box now matches within 0.1%; Arial's was 5.8% short.

**Arabic under Aptos.** `aptos-arabic-fallback.ps1` sets Arabic in each Aptos
family and reads the span fonts from the PDF. Excel (one font per cell, no
complex-script slot) and Word (Latin or complex-script slot alike) both draw it
in Arial, at Arial's width. So the three families take Arial's Lateef face in
`ARABIC_SCRIPT_FALLBACKS`, fitted through their own alias (the spaces come from
the size-adjusted Latin face, so Lateef is 3-5% larger than under Arial);
96-100% of Arabic line starts land where Arial's do. PPTX and DOCX already set
Arabic runs in the complex-script font, so this matters mostly for XLSX (Aptos
Narrow is Excel's default); the probe is the only evidence for that route.

**Blast radius.** `aared-fonts`'s table and tests only; no renderer code
changed. For a consumer passing a resolver, Aptos documents now download
fonts (at the time all subset faces loaded eagerly, about 700 KB for the
default theme; §50 made that lazy). XLSX reaches the resolver through style 0
(`collectStrings` visits every distinct cell style; the renderer paints
`cssFontFamily(style.fontFamily)`). XLSX column widths are a fixed 7px per
character (`core/xlsx/sheet.ts`), so substitute digits move no column.
End to end with the real `defaultResolver`, slide 1's line paints 385.7px
against real Aptos's 384.8px. No visual test uses the real resolver; the
probes are the acceptance, as for every calibration.

**Limits.** Italics use the upright numbers of their weight (2-3% narrower
than Aptos italic; 1% in Narrow). Aptos Light, SemiBold, ExtraBold, Black, Mono
and Serif are separate Office families and not covered. Letterforms differ;
widths and letter heights match. The substitutes are Google Fonts' current
files, which can change under any calibration: refitting Arial's and Cambria's
Arabic faces came out 0.3-0.9% off the numbers §46 shipped (noticed, not
changed).

## 50. Delivered fonts load when painted, not up front (Sep 2026)

The resolver used to call `.load()` on every non-Arabic subset face of a
substitute as soon as it registered it, in all four styles, so every Google
subset (Latin Extended, Greek, Cyrillic, Vietnamese...) came down whether or
not the document used it. §45 left that alone; §49 made it cost more.

**Change.** The resolver registers faces and never loads them. The browser
fetches a face when text first needs it, by style, weight and `unicode-range`.
Two things relied on the eager load and now load exactly what the document
paints:

- **The wait.** `collectFontUse` records, per family, the variants its runs
  paint in (their own `bold`/`italic`; an XLSX cell style has the same keys).
  `awaitDocumentFonts` loads those variants with a text sample, which triggers
  the download of those faces and subsets only. It used to wait for regular and
  bold and never italic, so DOCX layout could measure an italic run before its
  face arrived.
- **The fidelity check.** `detectFontFidelity` measures in a variant the
  document paints (a bold-only family's regular face may never be fetched).
- **PPTX `render()`** now waits too, like both DOCX modes: it paints, waits
  (bounded, 2 s) for the slide's fonts, then reads the warnings and resolves.
  Lazily, nothing else would load, and every substituted family would read as
  "font-unavailable". The synchronous `paintSlide` still reads at once and says
  so (API page). XLSX neither measures nor warns, so it needs no wait.

**Measured.** `tools/font-calibration/font-bytes.ts` counts the faces the
browser actually downloaded: `docx-coverage.docx` whole document went from 87
files (2307 KB) to 11 (337 KB); `pptx-coverage.pptx` slide 1 from 58 files
(1620 KB) to 5 (185 KB). Most of what remains is the Lateef Arabic face.

**Cost left.** A reader who has the real font still downloads what the wait and
the check load under the alias. Skipping that would need to know the real font
is installed, which cannot be measured reliably (Arial measures the same as the
generic it stands behind, §45). Kept, bounded by the numbers above.

## 51. The no-Office-fonts review: substitutes painted where they matter (Sep 2026)

Every calibration was measured on a Windows machine with Office's fonts, where
the real fonts always win, so no substitute had ever painted a real document.
A phone, or a Mac or Linux machine without Office, is where they paint.

**Simulated on this machine.** An installed font cannot be hidden from a page
(a registered "Arial" face that fails to load falls back to the installed
Arial). So the review renames every family in the model to one no machine has
(`"Arial absent"`), and
`packages/aared-fonts/test/absent-fonts.browser.ts` registers the real
resolver's faces under the renamed family's alias through its `FontFaceCtor`
hook: the same CDN files, calibrations and aliases a phone would get. A family
with no substitute (Wingdings, Segoe UI) falls to the generic.
`bun run test:visual:artifacts --absent-fonts` adds that render as a third
column of the PPTX and XLSX reviews (it uses the network);
`tools/word-render/capture.ts --absent-fonts` writes the DOCX pages.

**What holds.** Against the native Office renders, all three fixture families
lay out as with the real fonts: DOCX keeps its 4 pages, English line breaks
and table geometry; PPTX keeps all 12 slides in place; XLSX moves no row or
column.

**Fixed here.** Chart text in PPTX and XLSX was painted with the bare family
name, skipping `cssFontFamily`'s alias and generic, so a missing font (Aptos,
every default-theme chart) fell to the SVG default serif even on this machine.
`chart-svg.ts` now paints the same stack as every run.

**Found.** Symbol-font bullets became letters (Wingdings' square and check
painted as "q" and "u-umlaut" on slide 5), because those glyphs live at letter
codes and Android, iOS and Linux have no Wingdings; fixed in §52. The rest are
accepted as good enough (ROADMAP fidelity gaps):
- Lateef's Arabic, set 127-133% to match widths, draws taller than Arial's, so
  in dense rows its strokes cross borders, and under a link the underline cuts
  through letters (slide 6).
- Arabic breaks can shift by a word, within the calibration's measured spread
  (96-98% of line starts).
- A mixed-script cell sitting right at its wrap point can fit on one line where
  PowerPoint wraps it.

**Limits.** Characters no substitute covers fall to this machine's generic; a
phone's generic is Roboto or Noto, so unsubstituted families look different
there. One look on a real device was still worth taking (done in §53).

## 52. Wingdings and Symbol bullets paint their Unicode shape (Sep 2026)

Found by §51. Wingdings and Symbol keep their glyphs at letter codes:
Wingdings' square is "q" and its check mark "ü", and Word's default
first-level bullet is Symbol's private-use U+F0B7. Only machines with Office
have those fonts. Anywhere else the bullet paints as the letter, or as an
empty box.

**Change.** `core/symbol-fonts.ts` maps the characters Office offers as
bullets to their Unicode shapes: 20 Wingdings, 5 Symbol.

- Both the plain and the private-use (U+F000 + code) forms are mapped.
- Shapes a phone may draw as colour emoji (☑, ▪, ◻, ♣, ♦) carry the
  text-style selector U+FE0E.

The PPTX parser applies it to `buChar` and the DOCX parser to list markers,
dropping the symbol font, so the model itself carries the shape and search
and copied text see it too. On a machine with Office, the marker now paints
from a standard font instead of Wingdings; the shapes are near-identical.

**Not covered (accepted):** runs set in a symbol font, `a:sym`/`w:sym`
symbols, and Wingdings 2/3 and Webdings, which still paint letters without
those fonts.

## 53. `aared-fonts` frozen and unpublished: the browser's own fonts are the default (Sep 2026)

Overturns the roadmap's old "`aared-fonts` publish + CDN default" item. The real-device
check §51 left open was done, on an Android phone in Opera, over Wi-Fi to the
bench dev server, once with `?fonts=aared` and once without:

| Fixture | With `aared-fonts` | Browser's own fonts |
|---|---|---|
| `pptx-coverage.pptx`, open to first slide | 7,238 ms | 494 ms |
| `post-production-tracker.xlsx`, open | 1,120 ms | 332 ms |

The resolver fetches Google Fonts stylesheets before the first paint, and on
a phone that network wait dominates. The dev server makes the absolute numbers
larger than a production build would; the ratio is the finding.

The look went the same way. The phone's own Arabic (Noto) is drawn for
screens and read more cleanly than Lateef, which was chosen for Office's
widths, not for reading. Without the package, some Arabic lines break a word
differently from Office. Nothing overlapped, and page counts and tables held.
Under AGENTS.md "Good enough beats perfect", a reader notices slow and hard to
read, not a line break.

**Decision.**

- **The default is unchanged:** no resolver unless the app passes one, the
  browser's own fonts, nothing downloaded (§6, §12).
- **`aared-fonts` stays in the repo, frozen and unpublished.** It gets no
  more work. It leaves the first release (`@aared/fonts` is not reserved or
  published), and the docs present it only as an option for Office-exact
  line breaks.
- **Reopen only for a consumer that needs those.** Start from the network
  cost, not from more calibration.

**Kept, because they help every reader without the package:**

- chart text through the full font stack (§51)
- Wingdings and Symbol bullets as Unicode shapes (§52)
- fonts awaited per variant before measuring (§50)

The review tooling stays for anyone reopening this: `--absent-fonts`, the
bench's `?fonts=aared` switch, and `bench-lan` in `.claude/launch.json`.


## 54. PPTX fidelity pass: where PowerPoint puts text, colour maps, table styles, Arabic around punctuation (Sep 2026)

Sixteen PPTX rows of ROADMAP's fidelity-gap table were worked through under
"good enough beats perfect": fix what a reader notices, probe once when the
rule is unknown, accept the rest in writing. Probes (gitignored
`tools/probe/`): `first-line-leading`, `en-us-arabic-run`, `master-graphics`.

**Line spacing off 100% moves the text inside its box (overturns §47/§48's
"seen, not fixed").** In PowerPoint's own PDF the line step is always
`spacing x 1.2 x size` (aared matched), but above 100% the first baseline sits
`0.9 x spacing x size` from the top of the text area, and below it the natural
baseline plus the removed leading, never above that. CSS centres the leading,
so aared sat too low at 90% and too high above 100%. `leadingShiftEm` paints
the difference as a relative `top` on the `<p>`: box heights, steps and flow
are untouched. `MINIMUM_LINE_HEIGHT` (1.1) still floors 70%-90% lines, where
PowerPoint steps tighter: kept, because Arabic ink clips below it. Worst gap
to PowerPoint across the probe: 1.15pt (was 22.6pt). Limits: the natural
baseline is taken as 0.95em; exact spacing (`spcPts`) is unprobed and
unchanged. Blast radius: every PPTX paragraph with a line spacing other than
100%, notably every title and body on a default master (90%). XLSX and DOCX do
not run the PPTX paragraph painter (DOCX's lines are Word's own model, §44).

**A text body's anchor and wrap cascade one attribute at a time.**
`parseTextBody` read `anchor` and `wrap` only from the first non-empty
`bodyPr` in the slide/layout/master chain, so a slide's own
`<a:bodyPr><a:normAutofit/></a:bodyPr>` hid the master's `anchor="ctr"` and
the title painted at the top of its box. Insets already cascaded per attribute
(§9); anchor and wrap now do too. Blast radius: most PowerPoint-authored
titles. Autofit and `prstTxWarp` still read the first non-empty `bodyPr`.

**Text frames with no authored insets get DrawingML's defaults** (0.1in and
0.05in, resolved in `parseTextBody`, one constant pair shared with table cell
margins). Before, an `AddTextbox` box had no padding. Placeholders unchanged.

**Overflow paints in every autofit mode.** `none` and `shrinkText` no longer
clip (`overflow: visible`); the slide viewport still clips at the slide edge.
Unprobed and unchanged: exact spacing and bullet sizes under a stored
`fontScale`.

**Colour-map overrides.** `p:clrMapOvr/a:overrideClrMapping` on a slide or
layout replaces the map above it (slide beats layout beats master);
`a:masterClrMapping` keeps it. A slide-level override does not reach a
layout's shared graphics (the model is shared per layout). From the spec; no
fixture carries one.

**Table styles.** `a:lnRef` edge widths come from the theme's `a:lnStyleLst`
(`DrawingTheme.lineWidths`). The old constant (0.5, 1, 1.5pt) is the 2013-era
theme; the Aptos-era one states 1, 1.5 and 2pt, so edges were too thin there.
The constant stays only as the fallback. A built-in Medium Style 2 (seven ids)
missing from `tableStyles.xml` (generators that write only the id) is built
from PowerPoint's own definition, which matched the stored one in a probe
(`tools/probe/table-style-guids`); any other missing style paints unstyled.

**Arabic around ASCII punctuation in an `en-US` run.** In PowerPoint's PDF, in
a left-to-right paragraph, a run tagged with a left-to-right language orders
Arabic words in written order around every ASCII punctuation character
(`W1 + W2`, `W1 (W2) W3`); spaces alone, an Arabic `lang` and every
right-to-left paragraph follow plain bidi. The parser cuts such a run at the
punctuation between two Arabic words and flags the pieces `ltrRun` (invisible
marks in generated content, never copied or searched). Neighbours are checked
inside the run only.

**Punctuation stretches in right-to-left paragraphs: built, then removed.** An
`en-US` run in an RTL paragraph lays each stretch of two or more punctuation
marks out in written order (`.»` shows `.` then `»`) with mirrored glyphs, so a
plain left-to-right override is ruled out. It was built (parser split,
`TextRunModel` field, bidi isolates) and matched PowerPoint, then removed: the
only visible effect is which of `.` and `»` sits outermost, against a new
public model field, nested bidi wrappers and a painter branch. Only ASCII
punctuation and guillemets were probed. The LTR-paragraph rule stays: it swaps
whole words, which a reader sees.

**Master and layout graphics verified, no change.** `DisplayMasterShapes =
false` on a layout hides the master's shapes; on a slide it hides layout and
master shapes. aared already resolved this identically.

**Route completeness.** `DrawingTheme` is shared, so XLSX and DOCX also carry
`lineWidths` (unused there). `parseTextBody`, the PPTX text painter,
colour-map overrides and table styles are PPTX-only; DOCX resolves its own
tables, insets and bidi (§43-§44). Whether Word orders Arabic around ASCII
punctuation like PowerPoint is unprobed.

**Accepted, not built (recorded in the fidelity table):** fill approximations
(group pattern anchor, 96 DPI tiles, gradient stops); off-centre `circle` and
`shape` gradients stay marked fallbacks; live fields other than slide number
and date; `"Slide N"` titles (§31); embedded fonts (§45); notes (parsed onto
`SlideModel.notes`, not painted, not searched); `defaultTextStyle` versus
`otherStyle`. **Deferred by the maintainer:** charts, to be done in isolation.
**Found, not scheduled:** WordArt preset warps (`prstTxWarp`) paint unwarped
and marked; the arch and circle presets could bend real text along an SVG
`textPath` (keeps rule 1), the others are free-form envelopes with no DOM form.

## 55. DOCX pictures that text wraps around, placed as Word places them (Sep 2026)

An anchored picture (`wp:anchor`) used to paint as a plain left or right CSS
float whatever its wrap, so a centred picture sat at a column edge, a
top-and-bottom picture had text beside it, and a behind-text watermark pushed
the text away. Probe `docx-float-wrap` (gitignored) authored 13 wrap cases and
a page watermark through Word's COM, compared line by line with aared.

**The model.** `DocxImageNode.anchor` (`DocxImageAnchor`) carries what Word
stores: wrap kind, `behindDoc`, `wrapText`, horizontal and vertical position
(`relativeFrom` with `align` or offset, in inches), the four wrap distances and
the wrap polygon as shares of the picture (Word writes it in a 21600 square).
The picture leaves the runs: it hangs off its paragraph or heading as
`floats`. The old `float` field is gone (pre-release; no alias).

**The placement** is pure numbers (`docx/float-layout.ts`, `planFloat`),
testable without a browser. Four kinds:

```text
 wrap      one side, a CSS float; the wrap distance is padding inside its
           border box and `shape-outside: border-box` (or the wrap polygon
           for tight/through) is what text avoids
 band      top and bottom: a float the width of the column
 overlay   no wrap: absolute, z-index -1 behind the text or 1 in front,
           measured from the paragraph, or from the page for a page or
           margin position (layout mode; the header translates the origin)
 split     text down both sides (below)
```

**Two-sided wrap is built, not approximated.** A CSS line box is one interval
between the left and right floats, so CSS cannot put text on both sides of one
obstacle. An earlier pass pushed centred pictures to one side as "good
enough"; the maintainer rejected it as a difference a reader sees at a glance,
and AGENTS.md now says so. `docx/split-wrap.ts` lays the rows beside the
picture out itself: each row is a grid of (left slot | hole | right slot),
filled in reading order (right slot first in an RTL paragraph), the rest of
the paragraph flowing normally below. Words are measured with canvas
`measureText` (measuring only; the text is DOM spans, selectable and
searchable) through a bounded word cache, then every painted slot is checked
and the paragraph re-laid out with 1, 2 or 4px of slack if shaping disagreed.
Canvas measuring was chosen over DOM `Range` (2.4ms against 20.5ms on an
Arabic mixed-font paragraph, identical breaks). Tight and through take each
row's hole from the polygon's extent over that row (`polygonSpan`). The
paragraph's minimum height reaches the picture's bottom less its spacing
after, so the picture may hang into the gap as in Word but never under the
next line.

**Word applies the default paragraph style.** A paragraph with no `w:pStyle`
takes the style marked `w:default="1"` (usually `Normal`), not the bare
document defaults; aared used the defaults, so text was a pixel or two larger.
`StyleData` records `defaultParagraphStyleId` and `resolveParaProps` falls
back to it. Blast radius: every DOCX paragraph without a style; pixels on the
coverage fixture did not move.

**Fixture.** Seven wrap cases were added to `docx-coverage.docx` through COM;
the fixture is six pages in Word and in aared.

**Review pass** fixed ten memory, speed and dead-code points, among them: the
word cache is bounded, the flow `ResizeObserver` and font listener are
disposed with the render (`ImageWiring.cleanups`), and an anchor inside a
table cell is placed against the cell's text width.

**Route completeness.** PPTX pictures are absolutely placed shapes with no
text wrap; XLSX pictures float over cells with no text: neither format has
anchored text wrap, so both are unchanged. The default paragraph style is
DOCX-only (PowerPoint's masters cascade through `core/pptx.ts`).

**Limits (in ROADMAP):** a wrapped picture positioned from the page or a
margin vertically starts at its paragraph's top (the flow does not know where
the paragraph lands on the page); paragraphs after a split picture do not wrap
around its remainder; a top-and-bottom band right after a paragraph whose
picture still hangs is pushed below that picture (about 14px on the fixture's
page 5), since a full-width float cannot sit beside another float; with a
through wrap Word keeps a wider hole on the lowest rows than the polygon
gives at the line's own height, so it wraps about a word more there; VML
pictures (`w:pict`) have no placement; text boxes are not parsed.

## 56. The Worker URL resolves under webpack and Next.js (Sep 2026)

Settles ROADMAP's pre-publish question "Does the Worker URL resolve under
webpack?". Only Vite was proven (§19); webpack's own guide says webpack 5
bundles `new Worker(new URL("./worker.js", import.meta.url))` with no loader
when the URL is written inline, but says nothing of `{ type: "module" }`, and
Next.js 16 builds with Turbopack by default, a different bundler.

Measured with the package packed as it would publish (`npm pack`: `dist`
only) and installed into throwaway apps that open the three gate fixtures,
list their units and read the first slide and sheet:

```text
 build                                        pptx      xlsx      docx  404s
 webpack 5.111.1 production, no loader        7, read   4, read   2     favicon
 Next.js 16.3.8, Turbopack, static export     7, read   4, read   2     favicon
 Next.js 16.3.8, --webpack, static export     7, read   4, read   2     favicon
 both Next builds, aared imported at the top  same, no server-side errors
   of a client component
```

Webpack emits the Worker as one 133 KB chunk, Latin-1 clean, and passes
`{ type: "module" }` through. Importing aared at module scope of a client
component is safe: Next evaluates it on the server during prerender, and
nothing at import touches `window`, `document` or `Worker`.

**Kept proven:** `tools/check-consumer-build.ts` now builds the same consumer
with webpack as well as Vite (webpack and webpack-cli pinned in `bench/`,
about ten seconds a run), with the Latin-1 check and the open-every-format
run for both. Breaking the URL in `dist` to a plain string fails both passes.
Next.js is not in the script: the same Worker constructor is what both
passes guard, and Next with React would add about 150 MB to `bench/`.

**Deployment shapes, measured the same way:**

```text
 setup                                        webpack   Next/Turbopack  Next/webpack
 dev server (webpack serve, next dev)         ok        ok              ok
 server-rendered (next start, per request)    -         ok              ok
 sub-path (publicPath /app/, basePath /sub)   ok        ok              ok
 scripts on another origin (CDN)              refused   refused         refused
 the same CDN with the Blob recipe            ok        ok              ok
```

A browser never starts a Worker from a script on another origin, CORS or
not, so a CDN on another host fails for every bundler (Vite too, by the same
rule; not run). aared cannot route around it itself: a bundler only follows
the literal `new Worker(new URL(...))`, so aared never holds the URL before
the browser refuses it. Two changes follow:

- The refusal is thrown by the constructor, before any `onerror`, and used to
  reach the caller as the browser's raw `SecurityError`. `WorkerBridge` now
  wraps it: `AaredError("worker-error")` with the browser's reason, the usual
  cause and the remedy, and the original as `cause`
  (`session-lifecycle.test.ts`; seen in a real browser on the webpack CDN
  build).
- The API reference's "Worker override" gives the recipe: a same-origin Blob
  Worker that imports `new URL("aared/worker", import.meta.url)`, which every
  bundler measured emits as a file beside the other assets. It needs the CDN
  to send CORS headers, as any module script does.

## 57. Every document failure is typed; a lost session is `closed`, not `aborted` (Sep 2026)

An audit of the error surface (a read-only pass over every throw) found two
breaks in the one-error-model promise (`core/errors.ts`; the abort taxonomy
is §19), both at the Worker boundary:

**Parse failures arrived untyped.** `serializeError` (`core/worker.ts`) mapped
only `ZipError` to `invalid-document`. Malformed XML (`XmlError`), a missing
required part or element (plain `throw new Error(...)` across `opc.ts`,
`pptx.ts`, `theme.ts`, `xlsx/`, `docx/`) and a corrupt deflate stream crossed
as a bare `{ name, message }`, so `isAaredError()` and `error.code` saw
nothing. All of them mean one thing to a caller, so the Worker boundary now
types them `invalid-document` (the `ZipError` precedent extended) instead of
retyping ~35 throw sites. The original text stays, followed by the one action
a reader can take: re-save the file in Office. (§58 later narrows this.)

**A lost session reported `aborted`.** `sessionClosed()` answered a request
for a session the Worker no longer holds with `code: "aborted"`, which
`isAbort()` swallows, so the element and the hooks dropped a real usage error
silently. It is `closed` now, as the main thread already reports after
`close()`.

**Worker errors read in the caller's entrypoint.** The bridge rewrites the
`aared/core:` prefix to its own scope (for example `aared/pptx:`) as it
receives a failure. Cancellation keeps the one `aared/core: request was
cancelled` wording on both sides. Headless `core` calls still say
`aared/core:`.

**A failing font resolver no longer fails the document; `font-error` is
gone.** A `fontResolver` that threw rejected the whole `read()` and
`render()`, and the rejected promise was cached for the session. A font is
cosmetic: the failure is now caught per family and not cached, the family
paints in the browser's fallback, and the render handle's `warnings` names it
(`font-unavailable`, §45). `font-error`, never produced, is removed from
`AaredErrorCode` (pre-release; no alias). A failing `mediaResolver` was
already handled in all three renderers and is unchanged.

**A damaged sheet says what it skipped.** The XLSX sheet parser dropped an
unreadable cell, hyperlink, merge, conditional-format range or drawing anchor
with a bare `catch { continue; }` and nothing said the view was incomplete.
They now go to `SheetModel.unsupported` under a new kind, `damaged`, counted
per kind ("2 cells could not be read and were skipped; the file may be
damaged"). The kind keeps "the file is broken" apart from "aared does not
model this feature". Not render `warnings`: those are measured from what
painted. An unreadable `dimension` stays silent (the extent comes from the
cells).

The audit's matching PPTX claim, that an unreadable chart paints blank and
unmarked, was wrong: it already paints the marked placeholder
(`data-aared-chart-fallback="unsupported"`). Unchanged.

**Every embed layer reports a failed navigation.** The element fires `error`
and rejects `goTo()`; `preview()` rejects; the hook put only a failed mount in
state. Those idioms stay. The gaps were:

- **The hook lost page-change failures.** `status` stayed `"ready"` when the
  next unit failed. The result gains `navigationError`, cleared by a unit
  change and every mount. Chosen over moving `status` to `"error"`, which
  would replace the whole view for one page that would not show.
- **Reported where the failure happens, not where the call came from.** A
  first cut set it in the hook's own `goTo()` wrapper, but hosts also call
  `viewer.goTo()` on the real handle and bypassed it. Instead `preview()`'s
  handle fires `error` on a failed `goTo()` (still rejecting), with the same
  `{ message, cause }` detail `<aared-viewer>` fires, and the hook listens to
  it. Cancellations and overtaken calls fire nothing.
- **Silent no-ops.** `preview().goTo()` after `destroy()` and the hook's
  `goTo()` before mount now reject with `closed`.
- **A race:** an overtaken `preview().goTo()` still set `index` from its stale
  render. `renderUnit` now reports whether it painted, and only the call that
  did moves `index`.

## 58. Error codes say whose fault it is (Sep 2026)

A design review of the error surface after §57 judged the model sound (one
class, a closed code union, `aborted`/`closed`/`worker-error` kept apart, one
event shape) and found the codes coarser than they look. Overturns two calls
in §57.

**`invalid-document` meant "something failed", not "the file is broken".**
§57 typed every untyped Worker throw `invalid-document` with "re-save it in
Office". That blamed the file for aared's own defects and for valid content
(`pptx.ts` threw on an alignment or anchor value it did not know). Now the
fault decides the code:

```text
 the file is broken (a check that names what) -> damagedDocument(...) where the check is
                                                  invalid-document + repair hint
 not a ZIP / not an Office package              -> invalid-document, no repair hint
 malformed XML (XmlError)                       -> invalid-document + repair hint
 corrupt deflate (the platform's TypeError)     -> ZipError("corrupt") -> invalid-document
 an unknown value in a readable file            -> the default, read on (as Office shows it)
 anything else                                  -> internal, original error as `cause`
```

`damagedDocument` (`core/errors.ts`) replaces about 20 plain throws. The
unknown-value helpers (`align`, `verticalAlign`, `onOff`, a non-numeric colour
modifier) fall back instead of throwing: one odd attribute no longer costs a
slide. `internal` is new: "aared failed on this file", never "re-save it",
with the Worker's original error rebuilt as `cause` for the bug report. The
wire type's `code` is now required.

**Usage errors had borrowed other codes** (`not-found` for a non-whole index
and for an element `open()` with no source, `closed` for the hook's `goTo()`
before mount). New codes split them the way the platform does (`TypeError`
vs `InvalidStateError`): `invalid-argument` (a value the call does not take)
and `invalid-state` (a call the viewer is not ready for). `not-found` is left
for whole numbers and names that do not exist; `closed` for after teardown.

**A crashed Worker read as a closed document afterwards.** The bridge now
keeps why it stopped and rejects later requests with it, so a call after a
crash reports `worker-error`, with the browser's event as `cause`.

**Smaller fixes from the same review:** a download that drops mid-body is
`network-error`; `preview()`'s cancellation uses the shared `abortError()`;
the element and `preview()` share one empty-document check
(`assertHasUnits`); an element `open()` that a newer one replaced rejects
`aborted` (no `error` event: nothing failed); navigation inside a unit
(`viewer.xlsx`, `viewer.docx`) fires the handle's `error` like `goTo()`;
XLSX's render handle now reports font fidelity in `warnings` (it was always
empty).

**`messageFor()` stays:** one generic English sentence per code for a UI, not
localized; a host in another language maps `code` itself. The API reference's
Errors section has the code table, the message policy, `cause`, and the
listen-or-catch rule for the event layers.

**Not changed:** `AaredError.name` stays `"AbortError"` for cancellation so
existing `AbortSignal` handling works; docs say to branch on `code`.
`scrollTo()`'s wrong-shape target stays `invalid-range`.

## 59. Small API faults from a developer-ergonomics review (Oct 2026)

Two read-only reviews, one of the package API and one of the docs site, found
faults a developer hits on first use. Each was checked in the code before it
was fixed. Overturns §58's two hand-written facade index checks.

**DOCX page sizes are inches.** `DocxPageInfo.width/height` were CSS pixels
while the API page said inches, and every other size in the API is inches.
Pre-release, the canonical unit wins: pages report `widthPx / 96`.

**A replaced render destroys itself in every format.** Only PPTX did (a
`MutationObserver` on its target). A DOCX or XLSX render replaced by a newer
one in the same target kept observing and refitting, and calls on it silently
did nothing. `onReplaced` (`render-handle.ts`) is now the rule for all three.
XLSX's `display: flex` on its target is kept once per target
(`targetLayouts`) and restored when the owning render goes. Switching sheets
inside one render is not a replacement.

**One unit-index check for the facades.** The element and `preview()` each
hand-wrote §58's split and disagreed (`invalid-argument` vs `not-found` for
`goTo(-1)`). `assertUnitIndex` (`format-route.ts`) is the only check: a whole
number outside the document, below 0 included, is `not-found`.

**`<aared-viewer>` writes no `<style>`.** Its chrome is a constructed sheet
through `constantSheet`. The API page's "no `'unsafe-inline'`" now holds for
the element too (`csp.visual.test.ts`).

**Every entry imports without a DOM.** The element extended `HTMLElement` and
called `customElements.define` at import, so a server import would throw. It
now extends a stand-in without `HTMLElement` and registers only where there
is a registry (`entry-import.test.ts`). Read, not run in Next.

**Smaller:**
- A source `open()` cannot read is `invalid-argument`, not a raw `TypeError`.
  A File, Blob or ArrayBuffer from another frame still opens.
- `FontResolver` returns `Promise<FontFace | undefined> | undefined`: the
  documented `async` resolver did not compile.
- The hook's comment said `goTo()` before mount rejects `closed`; it rejects
  `invalid-state` (§58).

**Docs:** the home install card and "Start here" say nothing is on npm yet and
give the §56 route (build, `npm pack`, install the tarball); the
`@aared/fonts` reservation claim is gone; the element snippets say the import
belongs in bundled code; the React, DOCX, Worker-override and cancellation
examples were fixed and type-checked under the repo's strict settings.

**Not done:** a test that compiles every block on the page (`docs-parity`
checks names only); each block needs a marker, which belongs with the page's
restructure (done in §63).

## 60. Benchmarks measure aared; competitors are measured once (Oct 2026)

A competitor's numbers change only with its version, so a full pass spent
about 20 of its 30 minutes re-measuring numbers that could not have moved.
`--run` and `--size` now measure aared alone and merge its cells into the
latest result in place of its old ones, so every competitor's cells are
kept. `--merge` is gone, since merging is now the default. `--all` measures
every renderer into a new file: the first run, or after a competitor's
version or the method changes. Each cell and size row records `measured_on`
(date, browser, OS), and the generated tables and the home page's footnote
say which renderers were measured when.

Measured while making the change: on battery under Windows' "Silent" plan,
every small cell read about twice its time. aared's warm open on
`pptx-coverage.pptx` was 64 ms there and 32 ms on mains power, same build,
while the heavy competitor cells did not move. That points at the frame
cadence, not the CPU. BENCHMARKS.md now asks for mains power.

Open: pptx-browser, Extend and Apollo joined the bench after the last `--all`
run. They have size rows but no open, navigation or memory rows.

## 61. The hook's result narrows by `status`; its zoom and index are live (Oct 2026)

From the same API review as §59: the code a React consumer wrote around the
hook was mostly null checks and casts the types forced, plus work-arounds
for options that re-opened the document. Overturns the hook's result shape
and the "starting values" rule (§14's hooks). The live-options part was
overturned in turn by §64; the result shape below stands.

**One shape per state.** `UsePreviewResult` was one interface where
`viewer`, `info` and `format` were nullable whatever `status` said, and
`error` and `navigationError` were `unknown`. It is now a union on `status`:

```text
 "idle" | "loading" -> viewer, index, zoom, error, navigationError: null
 "ready"            -> viewer: V (non-null, typed by format), index, zoom,
                       navigationError: AaredError | null
 "error"            -> error: AaredError
```

Every field is on every member (null when it does not apply), so
destructuring keeps the narrowing. The per-format hooks return
`UsePreviewResult<PptxPreview>` and so on. The duplicates went (the user's
choice, option A): the document's `info` and `format` are read from
`viewer`, and `index` and `zoom` exist only when ready. `scale` is not
mirrored: it changes on resize with no event, so `viewer.scale` is live. A
failure that is not an `AaredError` is wrapped as `internal` with the original
as `cause`. The result object keeps its identity until a field changes, so it
can be an effect dependency.

**One handle type per format.** `PreviewHandle` is now
`PptxPreview | XlsxPreview | DocxPreview`, told apart by `format`; each
carries its own `info` type, and the other formats' controls are `null`.
`viewer.format === "xlsx"` makes `viewer.xlsx` non-null. One class still
serves all three, and `preview()` hands it out as the union.

**`viewer.xlsx` / `viewer.docx` are one object per viewer.** They were
rebuilt for every rendered unit and set to `null` on destroy, so an object
kept from before `goTo()` silently did nothing. Now each call goes to the unit
on screen, and after `destroy()` a move rejects with `closed` while
`scrollTo`/`focus` do nothing.

**Inline resolvers don't re-open.** An inline `mediaResolver` or
`fontResolver` re-opened the document on every render because its identity
was a dependency. A media resolver is now called through the latest one given
(adding or removing one re-opens, since the media route changes);
`fontResolver` and `worker` are read at the next mount. `type` and
`pagination` re-open, compared by value.

## 62. Bytes in as a typed array; every event typed (Oct 2026)

Two more items from the §59 review.

**`Source` takes any `ArrayBufferView`.** Bytes an app already holds (a
`Uint8Array` from `Response.bytes()`, IndexedDB, a Node `Buffer`) needed
`bytes.slice().buffer as ArrayBuffer`: a cast and a second copy. §59 made a
typed array a typed `invalid-argument`; now it is accepted. `sourceBytes`
copies exactly the view's own range, once (a Node `Buffer`'s `.buffer` is a
shared pool, so the whole buffer is never read). The element's `file` takes
the same (`Exclude<Source, string>`).

**Typed events and tag.** No event map existed, so listeners cast
`event as CustomEvent` and `document.createElement("aared-viewer")` was a
plain `HTMLElement`.
- Each handle names its events: `PptxRenderEventMap`, `XlsxRenderEventMap`
  and `DocxRenderEventMap` on the render handles, `PreviewEventMap` on
  `PreviewHandle`, `AaredViewerEventMap` on the element (it overrides `error`,
  `load`, `progress` and `selectionchange`, which it fires itself).
- `addEventListener` and `removeEventListener` are typed per name through
  `EventTargetOf<M>` (`render-handle.ts`, not exported), with a plain-string
  overload kept. Types only.
- `aared/element` adds `"aared-viewer"` to `HTMLElementTagNameMap`.
- `JSX.IntrinsicElements` is left to the app: React is an optional peer, and
  augmenting `react` would break a consumer without its types.

## 63. The API page says less, and its examples are compiled (Oct 2026)

The site review (§59) found that about a quarter of the API page was not
about calling the API (renderer behaviour, decision history, benchmark
timings), and several facts were stated three or four times. The page also
had no test that its code compiles.

**Moved or cut:**
- What PPTX paints moved to ARCHITECTURE § Format rules; the Worker "checked
  with" list to ARCHITECTURE § Packaging.
- The `aared-fonts` benchmark, the deferred decorations API, the `preview()`
  history and the future windowed loader were cut; search timings became a
  link to BENCHMARKS.md.
- Each repeated fact is stated once (theming under Dark mode, `::part()`
  under the element's Styling, `LoadProgress` in the events table).
- The design rules shrank to eight one-liners.

The page went from 1,783 lines to about 1,490 and gained a **Sizing**
section: sheets and documents need a height from the host (the site's CSS had
been supplying it).

**Compiled examples.** A runnable block is fenced ` ```ts check `, and
`api-snippets.test.ts` compiles every such block plus every playground
snippet with `tsc` against the package source under the strict settings.
`file`, `myFile`, `container` and `el` are declared for them. Signature
listings stay unmarked; docs-parity still checks names.

**Site fixes:** same-named headings get per-format anchors; the copy button
is one component (`CopyButtons.astro`) shown at rest, since touch screens have
no hover; the home page's Arabic heading carries `lang="ar"`.

## 64. Options start it, the viewer changes it; the element runs on `preview()`'s viewer (Oct 2026)

Overturns §61's live options and the API page's rule 6 ("options with a
setter are live").

**The problem.** After §61 a React app could change the unit, the zoom and
the appearance two ways (the hook option and the viewer), and the option kept
its own copy of the value, so it could disagree with the screen (a sheet-tab
click left the `index` option naming the old sheet). Only `index` changes
inside the viewer by itself (a sheet tab). Separately, `<aared-viewer>` kept
its own copy of the open, render and newest-wins code beside `preview()` (the
copies drifted once, §59), and XLSX sheet tabs were on in the element and off
in `preview()` and the hooks.

**The rule, in every layer:** options set where the viewer starts, the
viewer's methods change it, events report the change, and the state a layer
shows is a copy of the screen.

**Hooks.**
- `goTo`, `setZoom` and `setAppearance` left the result; `viewer`, non-null
  once `status` is `"ready"`, has them. The result keeps `ref`, `retry` and
  the state.
- `index`, `zoom` and `appearance` are read at mount only; a later change does
  nothing until the next mount (a `source` change or `retry()`). The user
  chose no exceptions for `appearance`: a theme kept in app state is forwarded
  with a small effect calling `viewer.setAppearance({ theme })`.
- No console warning when a start value changes after mount: there is no
  development build (`MARK_FALLBACKS` is compiled out of `dist`), so it would
  print in production. The docs and JSDoc say "start values".
- The rest of §61 stands: inline resolvers don't re-open, the result keeps its
  identity, and `navigationError` hears the viewer's `error`.

**One viewer for `preview()` and the element.**
- `PreviewViewer` moved into `embed-viewer.ts` (internal, not an entry) behind
  `openViewer({ scope, source, target, options, onProgress })`; `preview()` is
  one call to it, and the element calls it with scope `aared/element` and its
  progress callback (reading the bytes first would have copied the file twice
  and named `aared/preview` in the element's errors).
- The element's own copy is gone, and with it `el.goTo()`, `el.format` and
  `el.info`. `el.viewer` replaces them: the handle, typed `AaredViewerHandle`,
  without `setAppearance` and `destroy`, which the element owns (the viewer
  fires no appearance event, so the `theme` attribute could not follow a
  direct call). `el.viewer.xlsx` and `.docx` now reach cells, headings and
  pages from the element.
- The attributes and the viewer are one state: setting `index` or `zoom` calls
  the viewer, and `indexchange`/`zoomchange` write the attribute back, like
  `<input value>`. Edge cases were decided: a `zoomchange` to the format's
  default leaves an absent `zoom` absent; an `index` or `zoom` set while the
  file loads is applied before `load`, with no change event ahead of it; an
  `index` that cannot be shown fires `error` once and the attribute goes back
  to the unit on screen; `type` changed on an open document is compared with
  `viewer.format` with no download (another format reports
  `unsupported-format`, the same type does nothing); of `src` and `file`, the
  one set last opens.

**XLSX sheet tabs.** One default for the embed layers, `embedAppearance`
(`format-route.ts`): tabs on unless the caller says `sheetTabs: false`,
applied under the caller's appearance at every render and `setAppearance()`
(which replaces, never merges, so a theme switch used to drop tabs). The
headless `render()` keeps tabs off by default. The bench passes
`sheetTabs: false`, since it draws its own strip.

## 65. A stale render never paints; the hook owns teardown; the element lays out pages (Oct 2026)

Four follow-ups found while landing §64.

**A stale render could paint over a newer one.** A format's `render()`
awaits (the unit's read, then fonts) and then paints in one synchronous
step, with no way to stop it. Read in the code, not measured:
- A cancelled open's first render kept going; if a newer open painted first,
  the old render replaced it.
- The cancelled viewer was destroyed only after that render settled, and
  `destroy()` clears the whole target, which could wipe a newer paint.
- Inside one viewer, `goTo(1)` then `goTo(0)` while unit 0 was on screen
  returned early without stopping the move to 1, which then landed.

The fix is at the source:
- `render()` takes a `signal` in all three formats, like `open()`, `read()`
  and `search()`. Aborted before the paint, the render reads nothing more,
  leaves the target untouched and rejects `aborted`. After the paint the
  signal is ignored: `destroy()` is its teardown, so no painted handle is
  orphaned. DOCX layout mode builds its media URLs and class sheet after the
  font wait, so a cancel there leaves nothing to release.
- The embed viewer gives every render its own signal. A newer render,
  `destroy()`, and a `goTo()` back to the unit on screen cancel the one in
  flight; an overtaken `goTo()` still resolves having changed nothing.
- `openViewer()` destroys the viewer the moment its open signal aborts,
  instead of after the first render settles, and reports `aborted` even when
  the cut-short render failed with `closed`.

**The hook's `viewer` has no `destroy()`.** Calling it left the result
`"ready"` over a dead viewer. The hook owns teardown (unmount, a new
`source`, `retry()`), so `UsePreviewReady.viewer` is `UsePreviewViewer<V>`:
the handle without `destroy()`, distributed per format so `format` still
narrows. It is a type, like `el.viewer`'s (§64); plain JavaScript can still
call it. The bench stopped wiring `destroy` to the hook's viewer.

**One index check, one shape.** With the element on the shared viewer,
nothing called `assertUnitIndex` without a document's unit count, so the
"no document yet" branch went and `count` is required.

**The element lays out DOCX pages.** `pagination="layout"` (`"flow"` by
default) opens a DOCX in layout mode, so `el.viewer.docx.pages` and
`goToPage()` work from the element; other formats ignore it. Changing it
re-opens a DOCX, as the hooks do; the same mode again re-opens nothing. In
layout mode the element's page box drops its shadow and each page carries
`--aared-page-shadow`, the way a slide does; the element's rule carries the
same two-ID specificity bump as the renderer's class-sheet default, or the
page keeps the DOCX default. Closes the ROADMAP gap "Element renders
flow-only".

## 66. Emptying `<aared-viewer>` is not an error (Oct 2026)

**The problem.** Every change that re-opens the element (`src`, `file`,
`type`, `pagination`) called `open()`. With neither `src` nor `file` set,
`open()` closed the old document and then fired `error` with
`invalid-state` ("provide a src URL or file"). So an app that closed a file
(`el.file = null`) or removed `src` heard a failure, and one that shows "Could
not open this file" on `error` showed it to a user who had only closed one.

**The rule.** A change that leaves no source empties the viewer quietly: the
open in flight is aborted, the viewer is destroyed, the page is cleared,
`viewer` is `null`, and no event fires. The re-open paths go through one
private `#reopen()` that makes this choice. Calling `open()` yourself with
nothing to open is still a mistake: it rejects with `invalid-state` and fires
`error` once.

Deliberate non-change: an invalid `type` set while there is no source is not
reported until a source is set, since nothing is being opened.

## 67. The viewer has toolbar methods and says its default zoom (Oct 2026)

**The problem.** A custom toolbar over `preview()`, the hooks or
`el.viewer` had to rebuild what the element's built-in bar did privately:
paging with bounds, zoom steps that start from the scale on screen in a fit
mode, clamping and rounding, and a reset. The reset needed each format's
opening zoom, which was a private table (`DEFAULT_ZOOM`). For a slide or a
document "100%" is not that default, since `setZoom(1)` pins the natural
size, so the bench kept its own copy (`PAGER_DEFAULT_ZOOM`) and the site's
demo bar a copy of the step rules.

**The decision (the user's choice: both the methods and the value).** The
viewer, and so `preview()`, the hooks' `viewer` and `el.viewer`, gains:
- `defaultZoom`: the zoom the format opens at (`"fit-page"`, `1`,
  `"fit-width"`);
- `next()` and `previous()`: one unit on from where the newest move is
  heading, so quick presses add up (stepping from the unit on screen made
  three quick presses move one unit). At either end they change nothing and
  fire nothing; after `destroy()` they reject `closed`, like `goTo()`;
- `zoomIn(step)` and `zoomOut(step)`: `step` is an amount added or taken
  away, where `1` is 100%, so `zoomIn(0.1)` goes from 100% to 110% (the
  user's reading; a first version took it as a multiplier, 1.25 for 25%
  larger, and was corrected). Optional, default `0.1`, so a bare `zoomIn()`
  works; the element's bar uses the default, so its "+" now adds 10 points
  where it used to multiply by 1.25. A step that is not above 0 throws
  `invalid-argument`. In a fit mode they start at `scale`; they round to
  whole percents and stay within 25%–300%, and a step that would move the
  wrong way (a zoom already past the limit) does nothing;
- `resetZoom()`: `setZoom(defaultZoom)`.

The limits stay fixed (not parameters): they are where the renderers stay
legible and virtualization stays cheap, and no caller asked to move them.

**One copy.** The element's bar now calls these methods, and its own copy of
the step rules is gone. The bench uses `resetZoom()` and dropped its table.
The site's demo bar keeps its own rules because it also drives headless
render handles, which have none of these; it follows the headless paging
decision.

## 68. An aborted render never hands back a handle (Oct 2026)

Overturns one line of §65: "after the paint the signal is ignored".

**The problem.** A headless app pages by rendering one slide per click into
the same target and keeping the newest handle. PPTX and XLSX paint first and
then wait up to 2 s for the fonts they used (to report which painted,
`FONT_LOAD_TIMEOUT_MS`), and only then resolve. An abort during that wait
was ignored, so with two quick clicks the first render could resolve after
the second: the app then kept the first handle, already destroyed because
the second paint replaced it, and its zoom buttons did nothing. §65's signal
only stopped renders that had not painted yet. (The first explanation to
the user said the signal closed the race; it only narrowed it.)

**The options put to the user.**
- A: `goTo(i)` on the PPTX render handle, so one long-lived handle switches
  slides and aared keeps which slide is shown. Cost: a handle's slide would
  change over time, so code keyed by handle (a thumbnail strip) silently
  points at another slide; it needs an `index` and a `slidechange` on the
  handle, which rebuilds the embed viewer inside the render handle; there
  would be two ways to switch (render again, or `goTo`), and mixing them in
  one target brings the race back; and a handle painted from a bare
  `SlideModel` has no deck to fetch from, so `goTo` could only throw (as
  XLSX's `goToSheet()` does today). `preview()`, the hooks and the element
  already are that viewer.
- D: keep one handle per painted unit, and make the cancel rule strict.

The user chose D ("A changes the value, D changes the reference").

**The rule.** `render({ signal })` honours the signal until it resolves, the
way an aborted `fetch()` never resolves. Before the paint nothing is painted
(§65); after it, during the font wait, the render destroys its handle, takes
back out the nodes it put in the target (only those still there, so a newer
render's are left), lets XLSX give the target its own layout back, and
rejects `aborted`. One shared helper does it, `settleRender`
(`render-handle.ts`), used by PPTX and XLSX. DOCX already waited for its
fonts before painting and resolves as it paints, so a cancel always landed
before its paint; its docs now say the same rule. Once resolved, the signal
is ignored and `destroy()` is the teardown, as before. The plain pattern is
then correct with no "is this still the newest call" check; the API page
shows it as a compiled example.

The cost: with very fast clicks the target can be empty for a moment, from
the cancelled slide leaving until the next one paints, where before the dead
slide stayed visible.

**Found while checking it: going back to the unit on screen.** The embed
viewer's `goTo()` to the unit already shown cancelled a move in flight and
returned. But a move that has painted has already replaced that unit, whose
handle destroyed itself, so the viewer was left over a dead or empty target
while `index` named the old unit (read in the code, not measured; it needed
the click to land inside the font wait). Now `goTo()` returns early only
when nothing is on its way; otherwise it renders the unit again (the newest
call wins and cancels the move) and announces nothing, since the unit did
not change. The element's `index` attribute had the same early return of
its own; it now always asks the viewer. The site's headless slide pager
had both races: it now passes a signal per render and re-renders when going
back mid-move. (It was a shared `UnitController`; the playground's bars were
later rewritten as plain app code, and the pager now lives in
`site/src/demos/headless-pptx.ts`.)

Not handled: a newer move that fails (its unit cannot be read) after it
cancelled a painted one leaves the target empty; the failure is reported
through `error` and the promise. Rare: a failing read inside a font wait.

## 69. The viewer searches its own document; building from the contents stays low level (Oct 2026)

**The problem.** A page that shows a file with `preview()`, the hooks or
`<aared-viewer>` and wants a search box over it had to `open()` the file
again from a format entrypoint: a second parse, a second Worker, knowing the
format, and two teardowns.

**The rule, now stated on the API page ("Choosing a layer").** The viewer
levels show a file, move around it and find places in it. Anything built
from the file's contents (thumbnails, exports, pictures, several units at
once) is the low level's job, `open()` and `render()`, where the app owns
every step. Mixing works at the price of a second parse. Item 2's toolbar
methods (§67) and this search are the viewer side of that line.

**The options.** The whole document on the viewer (`viewer.document`), or
only `viewer.search()`:

| | `viewer.document` | `viewer.search()` only |
|---|---|---|
| Search box over the viewer | Yes | Yes |
| One parse, one Worker, one teardown | Yes | Yes |
| Thumbnails, range export, pictures | Yes | No: still `open()`, a second parse |
| What we promise | Everything three document classes offer, in four layers | One method |
| Plain JS can `close()` the document and break the viewer | Yes | No |
| An app can render into the viewer's own box by mistake | Yes | No: no slide handles are handed out |
| Changing course later | Shrinking it after 0.1.0 breaks callers | Growing it into `viewer.document` breaks nobody |

The user chose search only: it covers the case that started this, the data
uses fit the low level, and the choice can only be safely reversed in this
direction.

**The shape.** `search(query, { signal })` on each format's viewer type,
returning that format's matches (`PptxSearchMatch[]`, `XlsxSearchMatch[]`,
`DocxSearchMatch[]`), so checking `format` narrows them. It reaches
`el.viewer` and the hooks' `viewer` through the same types. It delegates to
the document the viewer holds; the search rules are unchanged. Three
defaults, agreed with the user:
- a failure only rejects and fires no `error`: that event reports moves,
  and the hooks show it as `navigationError`, which a failed search is not;
- after `destroy()` it rejects `closed` in the viewer's scope, like `goTo()`;
- the three match types are re-exported as types from `aared/preview`.

Limit, not measured: the first PPTX or XLSX search reads every unit through
the same Worker session the viewer paints from, so a unit change during it
may wait behind it. Later searches reuse the kept text.

## 70. A Start page for newcomers; the API page is the reference (Oct 2026)

**The problem.** A newcomer met the home page, then a "Start here" at the
top of a long reference, with the same choice named three ways (home tabs,
playground tabs, the API page's layers), no recipes and no troubleshooting.

**The change.**
- `/start/` (`site/src/content/docs/start.mdx`): Install (moved from the API
  page), Pick an entry, Recipes (a file the user picks, a file behind a
  login, a search box, your own toolbar, errors for people; headless paging
  links to the API page's pattern), Sizing (moved from the API page),
  Frameworks (Vite and webpack, Next.js) and Troubleshooting.
- "Pick an entry" and the API page's "Choosing a layer" import one partial
  (`site/src/partials/choosing-a-layer.mdx`), the rule of §69 with its
  table, so the two never drift. Outside `content/docs`, it is not a page.
  The Start page maps the playground's tab names onto the same entries.
- The API page opens with one line pointing to Start and the playground;
  its "Start here" and "Sizing" moved, and its links to them point there.
- One route, `pages/[page].astro`, serves every docs page; it replaces the
  two near-identical `api.astro` and `playground.astro`. A page lists its
  headings unless its frontmatter says `sidebar: false` (the playground).
- The header gains "Start", and the home page's install note and its
  next-step cards point to it.
- The Start page's examples are `ts check` blocks: `api-snippets.test.ts`
  now compiles both pages' blocks (AGENTS.md says so).

## 71. Charts: a real chart model, read in the Worker, laid out on the page (Oct 2026)

The fixed chart template produced materially incorrect charts: it lost combo
series, misrepresented stacked/ranged data and axes, and substituted defaults
for labels, colours and legends. It is replaced by a real chart model and layout
engine; `ChartGraphicModel` and `ChartGraphicPrimitive` are superseded.

The engine reads chart specifications in the Worker and lays them out on the
main thread at the frame's actual size. It supports column/bar, line, area,
pie/doughnut, scatter, bubble, radar, and combinations; 3D kinds are flattened.
The plain-data `ChartSpecModel` is shared by PPTX, XLSX, and DOCX; visual
primitives are generated only during page layout.

The PowerPoint oracle informed the axis and layout rules; those rules live in
ARCHITECTURE. Its scores were in-sample, so they are review evidence rather
than a general fidelity guarantee.

Charts are optional content, so the Worker parses and transfers the spec while
the page lazily loads the layout code on first use. This keeps chart drawing
out of the Worker and avoids adding its cost to files without charts. A chart
layout load failure is surfaced as a placeholder, not silently omitted.

The reader enforces chart point limits before allocation, pins SVG chart
direction to LTR while preserving Arabic text direction, and measures Arabic
labels with shaped-width metrics. Unsupported plot types show a reasoned
placeholder rather than a misleading partial chart. Chart fills use the shared
fill painter; theme colours resolve through the common theme module.

All three formats use the shared chart reader and painter. DOCX charts follow
the same inline/anchored placement and wrapping path as drawings. A Word-authored
Arabic inline and anchored chart in the DOCX fixture confirms the route. Charts
on PPTX layouts and masters are still not painted.

**Deliberate non-changes and open work.** Thinning very large series, typed
arrays and the remaining reader-visible gaps (data-label text, painted-font
widths, pie label spreading, the legend cap, fills, date axes, the `a:cs`
slot, 2016 chart types, search and screen readers) are in ROADMAP and
ARCHITECTURE "Known limits".

## 72. Hostile counts are bounded where the file is read; three XLSX and PPTX rules corrected (Oct 2026)

**Root cause and decision.** Hostile OOXML counts and ranges could trigger
unbounded allocation/work or raw exceptions. Bounds now apply in the Worker at
parse time: XLSX column ranges are validated, merges are clipped and capped,
PPTX/DOCX table spans are clipped to table extents, and ZIP directory entries
are checked against the declared directory end. This keeps page-side work
bounded and malformed ZIP input typed as corrupt.

**Three rules corrected on the way.**
- *PPTX merged cells never merged.* `gridSpan`, `rowSpan`, `hMerge` and `vMerge`
  are attributes of `a:tc` (ECMA-376 21.1.3.16); the reader took them from
  `a:tcPr`, and its only test wrote them there. No coverage deck has a merged
  cell, so nothing caught it. A real merged table showed separate cells and kept
  the covered ones. Read from `a:tc` now; the test writes what PowerPoint
  writes. Not yet confirmed by a PowerPoint probe (the ECMA schema and
  python-pptx agree).
- *XLSX columns were 5 px too wide.* A stored width already includes the cell
  padding (ECMA-376 18.3.1.13), so pixels are `trunc((256w + trunc(128/7)) / 256
  * 7)`; the code added 5 px again. Excel read back (COM, on fixture copies):
  width 26 draws 182 px (aared gave 187), 16 gives 112 (117), 8.4258 gives 59
  (64). A sheet with no `defaultColWidth` stays 64 px. A 20-column sheet
  drifted 100 px; anchored charts and overflow moved with it.
- *XLSX tint was an sRGB blend.* Excel moves the lightness in HLS (ECMA-376
  18.3.1.15): `L * (1 + tint)` for a negative tint, `L * (1 - tint) + tint` for a
  positive one. Excel's "Darker 25%" of accent 5B9BD5 is 2E75B6; the blend gave
  4474A0. `applySpreadsheetTint` in `core/theme.ts` sits beside the HSL helpers
  `lumMod` uses; DrawingML's `a:tint` (linear light) is untouched.

**Conditional formats** (`core/xlsx/sheet.ts`). Colour scales use every stop
(red-yellow-green was two-colour) with `min`, `max`, `num`, `percent` and
`percentile` bounds (percentile as Excel's PERCENTILE); `stopIfTrue` keeps lower
priorities off a matched cell; Duplicate and Unique Values compare cached values
(text without regard to case). A rule type the viewer cannot evaluate from
cached values (`expression`, `top10`, `aboveAverage`, `iconSet`, ...) and a
`formula` bound are now reported in the sheet's notices, not skipped silently.

**Route completeness.** Spans: PPTX and DOCX, one helper. Merges: XLSX only
(neither other format has a sheet grid). Tint: XLSX only (DOCX and PPTX read
`a:tint`/`w:themeTint` differently; DOCX does not read `w:themeTint` at all,
unchanged). Zip: shared by all three.

The hostile-input and conditional-format tests cover the bounds and supported
rules. Conditional-format cases use Excel-shaped XML; the column-width and tint
corrections were also checked against Excel. Total inflated ZIP bytes, XML
nesting, repeated entities, and URL input size remain open bounds (ROADMAP).

Typed-column transfer and URL-fragment document formats were not adopted; neither
addressed a demonstrated need in this pass.

## 73. Charts: the reader-visible gaps closed, every type but maps and surface drawn (Oct 2026)

**The problem.** §71 landed the engine with its gaps listed: Arabic measured
as one mean width and set in the Latin font, custom label text shown as the
value, grouped categories and date axes flattened, pie labels overlapping,
legends spilling, chart fills replaced silently, stock, of-pie and every
Office 2016 type a placeholder, and charts invisible to search and screen
readers. All of them are things a reader sees. They were closed in one pass,
six pieces of work in parallel, merged one by one with the full suite after
each.

**What changed, by piece.**
- *Arabic.* The complex-script slot (`a:cs`, the theme's `cs` or Arab font)
  is read into `TextLook.csFamily` and an Arabic run is measured and painted
  in it; a line mixing scripts paints as one tspan per script, because
  Calibri and Arial both hold Arabic glyphs and a font list alone would paint
  Arabic in the Latin face. Widths are per letter and joining form, from
  Arial's, Times New Roman's and Tahoma's own Presentation Forms advances
  (`gen_metrics.py`), lam-alef one glyph, marks zero: the 0.333 em mean of §71
  is gone. A label's lines are one paragraph: its first strong letter sets
  every line's direction, so "58%" under an Arabic category reads "%58", as
  PowerPoint draws it.
- *Painted fonts.* Layout runs on Office's widths; the page now measures each
  face a chart uses once (`chartFontProbes`, a hidden SVG text) and scales
  the layout when the device paints a substitute (within 2.5% counts as the
  same face).
- *Content.* Author-typed label text with its fields (`VALUE`,
  `CATEGORYNAME`, `CELLRANGE`, ...) and Office 2013 cell-range labels; manual
  label positions; the outer levels of grouped categories as Office's extra
  label rows with separators; date axes placed by date on base-unit slots.
- *Overlap and spill.* Pie labels outside the pie spread so none overlaps,
  with leader lines; a best-fit label stays in its slice when it is no wider
  than the slice at the rim and wraps at half the radius (PowerPoint kept a
  62pt Arabic name in a 32 degree slice and wrapped a 77pt one). Side legends
  are capped at a third of the chart; a top or bottom legend that wraps is
  Office's grid, equal cells as few columns as keep the fewest rows (20
  entries: 7, 7, 6). Nothing is drawn outside the frame; the SVG stays
  `overflow: visible`, the layout enforces it. Trendlines (all six types) and
  error bars are drawn.
- *Fills.* `pptx/fill-paint.ts` and `pattern-presets.ts` moved to `src/`, and
  the Worker fill readers to `core/fill.ts`: a chart fill is the real
  pattern, gradient (any path) or picture, painted by the one painter shapes
  use, fetched only for charts that need it. A fill it cannot paint leaves the
  shape unfilled and marked. This overturns §71's stand-ins (grey, a colour
  mix, a linear gradient).
- *Types.* Stock charts (hi-low lines, up-down bars, volume combinations),
  pie of pie and bar of pie, and the Office 2016 types (waterfall, funnel,
  histogram, Pareto, box and whisker, treemap, sunburst; read by
  `chart/read-chartex.ts` and lowered onto the same model where they can be).
  §71 made stock and bar-of-pie placeholders because drawing them as lines and
  one pie said something else; they now draw as themselves. Map charts and
  surface charts stay placeholders, with the reason in all three formats (PPTX
  had dropped it).
- *Search and accessibility.* Chart words (title, axis titles, series,
  categories) are searchable in all three formats; a chart is `role="img"`,
  named by its alt text, else its title, else "Chart", and described by a
  `<desc>` of its type, series and categories.
- *Long series.* A line with more points than the plot has half-point columns
  keeps, per column, its first, lowest, highest and last point (M4): 100,000
  points went from 1,306 KB of path text to 50 KB, every peak kept. Typed
  arrays were not adopted: §72 measured the Worker-to-page clone of a real
  workbook at 1.6 ms, and the series cap bounds the worst case.

Excel chartex charts resolve values through their hidden defined names; 1904
date handling and drawings beyond a worksheet's used range were corrected in
the shared XLSX route. A Vite deadlock in the lazy chart chunk was fixed by
removing imports shared with the entry module. PowerPoint oracle comparisons
showed the corrected rules held on the existing corpus; the rules and remaining
assumptions are recorded in ARCHITECTURE.

**Deliberate non-changes and limits.**
- A busy pie with labels outside: PowerPoint lets them overlap; aared spreads
  them. The one case scores 26 / 30pt off for it. Matching Office would mean
  drawing overlapping text; kept as is, the maintainer accepting the
  difference, and recorded in Known limits.
- Many layout rules from this pass are assumptions not yet measured against
  Office: the side legend's third, the date axis' unit thresholds and tick
  ladder, the leader line's default colour, the trendline's default dash, the
  histogram's bin rounding (two data sets), treemap banner labels, inclusive
  quartiles. ARCHITECTURE "Known limits" lists them.

## 74. Charts: the assumed rules checked against Office, and corrected (Oct 2026)

**The problem.** §73 left a list of chart rules taken on trust: the legend
cap, the date axis' thresholds, steps and fallback format, the looks of
unstyled trendlines, error bars, leader lines and stock bars, the pie label
wrap, how error bars meet the axis. Each was a guess until Office drew the
same chart.

**How it was checked.** Twenty cases were added to the oracle corpus
(`g-*` in `gen_corpus.py`, one or more per rule), exported through
PowerPoint and scored as the others are; lines were compared on the paired
renders. The generator gained raw-XML hooks for what its options did not
cover (series extras such as trendlines and error bars, an extension list,
category XML for grouped categories, numeric date categories, date axes, a
stock chart). Three Excel-made variants were added to `chartex-oracle.xlsx`
(bin count, inclusive quartiles, banner labels), compared with Excel's own
render.

**What held.** The side legend's third, the legend grid (two more cases),
the 28-day boundary (dates 28 days apart go by month), grouped categories,
value-from-cells labels leading the value, error bars' standard deviation
centred on the mean, inclusive quartiles.

**What was wrong, and is now Office's.**
- *Unstyled lines.* With no style in the XML, as third-party generators
  write, Office draws axis lines black, tick marks crossing the axis with
  minor ticks (the schema's default for an absent element), and trendlines,
  error bars, leader and hi-low lines thin and black. The reader had a grey
  axis, no ticks, a dotted series-coloured trendline and grey bars. PowerPoint
  and Excel agree on all of it.
- *Stock defaults.* Unstyled stock series take automatic markers; down bars
  are dark grey (`tx1` at 75%), not black.
- *Pie colours.* A point that only carried a label (a dragged one) lost its
  own colour: the colour pass skipped any point already present. Every slice
  of such a pie painted the series colour.
- *Leader lines.* A label dragged off the pie gets a thin black line with a
  short level end; one still touching the pie gets none. Only labels the
  spreading moved had one.
- *Error bars* widen the value axis to their ends (9.9 runs PowerPoint's axis
  to 12).
- *Date axes.* Office labels every base unit while it can: flat when a line
  height is left between labels, else turned 45 degrees, and only then a
  coarser step; never upright. The engine went straight to coarser steps and
  stayed flat. A date axis without a date format shows the serial numbers;
  the engine put in `m/d/yyyy`.
- *Pie label wrap.* Not half the radius: a label decides inside or outside
  unwrapped against the slice's width at the rim, then wraps to the slice's
  width at 0.85 of the radius inside, or at a fifth of the plot outside (the
  agent's 0.28 was a guess too). Both pies, Latin and Arabic, fit the rule.
- *Trendline labels.* "General" is Office's own rounding, not the cells'
  (1.3214, not 1.321428571); the label is a centred block whose right edge is
  at the line's end, centred on it vertically.
- *Treemap banners* are one grey band (`tx1` at 35%) two label sizes tall,
  the text centred in it, not the branch colour.

**Proof.** The 20 cases, scored against PowerPoint before and after:

| | Before | After |
|---|---|---|
| Mean text offset, across / down | 3.1 / 6.6 pt | 1.7 / 2.2 pt |
| Value axis shows PowerPoint's numbers | 17 / 20 | 20 / 20 |
| Cases with a word more than 10pt off | 8 | 3 |

None of the earlier 74 cases moved. The whole corpus: Latin 87 cases 1.61 /
1.72 pt, axis 84 / 87; Arabic 7 cases 0.70 / 1.30 pt, axis 7 / 7.
`test/chart-oracle-guesses.test.ts` pins each corrected rule on its case;
`chart-chartex.test.ts` the banner and inclusive quartiles.

**Deliberate non-changes and limits.**
- *Date format.* `m/d/yyyy` (Excel's built-in short date) shows as written,
  as XLSX cells do; Office shows the reader's own short date. Two of the
  three cases still over 10pt off are that difference (a wider date turns or
  shifts the labels). Kept: one rendering everywhere, as for cells.
- *Chart border.* PowerPoint draws no border round an unstyled chart, Excel
  a thin grey one; the shared reader keeps Excel's. Word was not probed.
- The third case over 10pt is one outside pie label 25pt off in the
  long-name pie: kept.
- The legend grid and the inside-label wrap rest on two cases each now; the
  histogram's overflow and underflow bins and the box plot's mean line could
  not be set over COM and stay unchecked.
- The bench lists the three oracle workbooks as samples
  (`bench/vite.config.ts` serves them from `tools/chart-oracle/`), so every
  chart type can be looked at, dark mode included; they are not fixtures.

## 75. A whole-repo static review: shared failure classes and deliberate behavior (Oct 2026)

A whole-repository review produced 290 findings; a second pass verified them
against the code. It led to cross-cutting corrections: typed errors for damaged
XML/ZIP input, shared chart-point bounds and enum validation, DOCX relationship,
list and page-break fixes, XLSX range and formatting corrections, consistent
PPTX element identity, and tests that assert outcomes rather than merely execute
code. The pre-release bare `aared` alias was removed and the root build narrowed
to package builds.

Office probes established the missing-drawing-id policy. PPTX shapes without ids
are retained with stable internal ids; missing ids in PPTX, DOCX and XLSX are
reported as damage, while renderable pictures and charts remain visible. Office
may prompt to repair these files but does not remove their visible content. The
id is not used to locate DOCX/XLSX media.

Other intentional choices: chart point budgets are claimed before allocation
and bound both per-chart allocations and total reads; XLSX merges extend stale
declared dimensions; charts are not re-laid out on late font arrival; filled
radar charts use point fills. Site theme storage has no pre-release migration,
and `parseXml` relies on `TextDecoder` to remove a BOM.

Office-dependent behavior was not guessed; the follow-up probes are in §§76–77.
Capabilities and remaining limits belong in ROADMAP and ARCHITECTURE.

## 76. Four Office probes: bold by script, number-format ids, empty paragraphs, first-page headers (Oct 2026)

§75 left four findings that depended on what Office does and not on a guess.
Each was probed in the real application (scripts and documents in the
gitignored `tools/probe/`; each finding is in the tests' comments, so nothing
depends on the folders surviving), then fixed.

- **Bold and italic by script** (Word 16, `docx-cs-bold`). A run in Arabic text,
  or any run marked `w:rtl`, takes `w:bCs`/`w:iCs` and ignores `w:b`/`w:i`; every
  other run takes `b`/`i` and ignores `bCs`/`iCs`. Before, the parser read one
  pair for every run. `readRunProps` now keeps `boldCs`/`italicCs` apart and
  `finalizeRun` picks by script. Blast radius: every DOCX run with a bold
  or italic flag. The coverage fixture's appendix header run changed to bold, and
  Word's own span dump shows it bold (TimesNewRomanPS-BoldMT, 15pt), so the old
  value was the bug. The sections snapshot and two local digests changed by that
  run, the rename in §75, and the new empty paragraphs below, and nothing else.
- **Number-format ids** (Excel 16, `xlsx-builtin-numfmt`). Ids 5-8, 14, 22,
  27-36, 37-44 and 50-58 are regional in Excel, not fixed codes. aared now maps
  5-8 and 41-44 to the ECMA-376 en-US codes, 27-31, 36 and 50-58 to the short
  date, and 32-35 to `h:mm:ss`. Comparing against Excel's own text also found a
  real bug in optional digits: `0.0#` drops a trailing zero, `??` blanks it, and
  only a `0` slot prints an integer zero. Left alone: the `*` fill in accounting
  codes is not painted (Known limits).
- **Empty paragraphs** (Word, `docx-empty-paragraph`). One empty paragraph is one
  full line at the size of its paragraph mark (`w:pPr/w:rPr`) plus its own
  spacing. The parser now emits a paragraph node with no runs for it (never a
  heading, and never for a paragraph that holds only a page break), and the
  renderer paints a `<br>` for any paragraph with no runs. A blank line now
  takes its line in layout too, so pagination follows.
- **First-page and even-page headers** (Word, `docx-first-header`).
  `SectionProperties.headerParts` and `footerParts` carry the default, first and
  even kinds, with `titlePage`, `evenAndOddPages` and `pageNumberStart`, and a
  `DocumentSection` has `firstHeader`, `firstFooter`, `evenHeader` and
  `evenFooter`. Each kind inherits from the previous section on its own, so a
  `w:titlePg` section with no first header of its own shows the earlier section's
  first header, not its default. Layout pagination picks each page's pair and
  gives each page `LayoutPage.number`, the printed number, which restarts at
  `w:pgNumType/@w:start`; `PAGE` fields use it. `HeaderFooterParts` is exported
  from `aared/core` and named on the API page. This replaced the interim
  `headerPart` / `footerPart` of §75.

**Deliberate non-changes.** Pages of a section are still fitted to the default
header band (Known limits). The coverage fixture carries `titlePg` since §77;
even-page headers are tested on synthetic documents pinned to Word's printed
numbers. The two probes this entry left open (list counters, an empty
paragraph in a table cell) are §77.

## 77. List counters per list definition, and the coverage fixtures extended (Oct 2026)

The maintainer's read of the last open probes: Office applies one logic, so a
rule measured once is the rule everywhere. Asking Word itself settled both
quickly (`tools/probe/docx-list-counters`: a hand-written `numbering.xml` opened
in Word, each paragraph's marker read back with `Range.ListFormat.ListString`,
and row positions read back with `Information(6)`).

- **List counters** (Word 16). Numbering ids that share a definition
  (`w:abstractNum`) continue one sequence (1, 2, 3, 4 across two `w:num`s); a
  `w:num` with a `w:startOverride` starts its level over the first time it is
  used (a restart prints 1, 2 again); a deeper level starts over after a
  shallower item; an empty numbered paragraph takes a number; a numbered heading
  takes a number. Before, counters were keyed by `numId`, so two ids sharing a
  definition printed 1, 1, 2, 2, and a heading advanced nothing. `advanceCounters`
  (`core/docx/document.ts`) now keys them by definition and applies a start
  override at a num's first use; a numbered heading advances the counter (the
  empty-paragraph case already worked once §76 emitted empty paragraphs). The
  probe's whole document is a unit test with Word's printed markers. Blast
  radius: every DOCX list; the coverage fixture's lists print as before. A
  separate real-file render probe confirmed the remaining paint gap: Word
  reports `ListString` `1.` for a Heading 1 paragraph and paints
  `1. Project Overview`; aared paints only `Project Overview`. The probe uses
  the updated `tools/word-render/out/numbered-heading-probe/word-source.docx`,
  where the numeral is actual numbering, not heading text; Word's list
  indentation places the marker at the paragraph start and the heading text
  after the list marker tab. Paired page-2 captures are in
  `tools/word-render/out/word-source/`. The earlier coverage capture had a
  typed numeral and did not test this behavior. Painting the computed marker
  remains open (ROADMAP).
- **An empty paragraph in a table cell** (Word 16, the same probe): a row whose
  only content in a cell is an empty paragraph is exactly as tall as a row of
  text (15.05pt at Arial 11), whether the other cell has text or not. That is the
  rule §76 already wrote into ARCHITECTURE.
- **Fixtures extended through Word and PowerPoint** (no new files; the label
  stripped, `strip_label.py` now also removes a marking run that carries
  `w:rtl` and a language, and works on a file it has stripped before).
  `docx-coverage.docx` gains a last portrait section
  (`tools/probe/docx-lists-headers`): a different first page with its own header,
  three lists (one continuing, one restarted) and a table with an empty cell;
  nine pages in Word, and aared paginates to nine with the same header on each
  page. It has no even-page header: `w:evenAndOddHeaders` is document-wide and
  would blank the even pages of every earlier section. `pptx-coverage.pptx`
  gains slide 13 (merged table cells: a `gridSpan` and a `rowSpan`) and slide 14
  (Excel's waterfall chart pasted into PowerPoint, which writes a `chartEx` part
  with a `cx:lvl` cache; PowerPoint's own `AddChart2` refuses the 2016 types)
  (`tools/probe/pptx-merged-chartex`). Their tests are in `docx-parser`,
  `docx-pagination`, `pptx-parser` and a merged-cell visual test; counts that
  named seven pages, twelve slides and four tables moved to nine, fourteen and
  five.

- **A second slide master** (`tools/probe/pptx-second-master`): a logo on the
  existing master would have shown on every slide, so slides 15 to 18 of
  `pptx-coverage.pptx` sit on a second design added through PowerPoint, with its
  own theme (accent 1 dark green), a logo, a picture background and two layouts
  (A shows the master's graphics, B hides them; slide 18 hides them with its
  own `showMasterSp="0"`). It replaces the four `test.todo` of
  `pptx-layout.test.ts` with tests: two slides on one layout share the layout
  object, master graphics come before the layout's and a layout with
  `showMasterSp="0"` drops them, only the hiding slide sets
  `hideBackgroundGraphics`, the master's picture background is inherited as one
  shared image fetched once per session, and a visual test reads the painted
  order and the left-out graphics from the DOM. aared passed all of it as it
  stood. The fifth thing the deck is for failed: slide 15's accent box is green
  in PowerPoint, and aared read the first theme in the package for every slide,
  so it painted blue.
- **Per-master themes.** A slide now resolves colours and theme fonts against the
  theme of the master its layout belongs to. `indexPptx` reads each master's
  theme (`readMasters`: the presentation's master list, each master's `theme`
  relationship, `readThemePart` in `core/theme.ts`) with the table styles
  resolved against it, into `PptxDeck.masters`; `parsePptxSlide` and
  `parsePptxLayout` take theirs through `themeOfMaster`, and `deck.theme` stays
  the first master's (the document info). A deck with one master behaves as
  before. DOCX and XLSX have one theme and still call `readDrawingTheme`. Pinned by
  a test that the second master's slide paints `#1E6B3C` and the first
  master's slides keep `#156082` (it fails with the lookup switched off).
  Checked against PowerPoint: the paired captures of slides 15, 17 and 18 match
  (the green boxes, the sand background, the graphics left out); and slide 2's
  table, moved onto the second master inside PowerPoint (`CustomLayout`) and
  exported, paints a `#1E6B3C` header and `#CCD4CE` bands, exactly what aared
  resolves for the same move made in memory. Not checked: a second master's
  theme fonts (both themes name Aptos, so the deck cannot tell them apart).

**Deliberate non-changes.** A DOCX Office 2016 chart is not in
`docx-coverage.docx` yet. Preset colour names stay unmeasured.

## 78. Search matches the Arabic spellings of one word (Oct 2026)

`search()` was a case-insensitive substring over raw text, so a reader who typed
`الانتاج` found nothing in a deck that wrote `الإنتاج`, and `٢٤` found nothing
in a cell that held `24`. That is the failure the project exists to close, and
it sat open on the ROADMAP as small.

**Root cause.** Not a missing step in three places: the three indexes (PPTX
paragraphs, XLSX cells and chart words, DOCX blocks) already compared through
the one `foldText` in `core/search.ts`, which was only `toLowerCase()`. The
fix is one rule in that home. `searchKey` (needle and haystack alike) is the case
fold, then: tashkeel (U+064B-065F), superscript alef (U+0670) and tatweel
(U+0640) dropped; alef with madda, hamza above or below, and alef wasla to bare
alef; alef maqsura to ya; Arabic-Indic and extended (Persian) digits to 0-9,
with the Arabic decimal and thousands separators to `.` and `,`; and, when the
text has an Arabic presentation form (a PDF paste), NFKC first so the form
takes its base letters.

**Blast radius.** Every `search()` in every format, and `viewer.search()` on
`preview()`, the hooks and the element, since they call the document's. A hit
reports the original text (`clipSnippet` of the cell, block or paragraph), never
the key, and matches are locations, not highlights, so the key may be shorter
than the text and no offset map is needed. The index field `lower` became `key`.
Text with no Arabic-script code unit takes the plain case fold after one scan, so
a large Latin sheet indexes as before. The fold is written as code-point
comparisons and two constants in `core/chars.ts`, because a literal or a `\u`
escape above U+00FF would break the Latin-1 rule (§34). A query that folds to
nothing (a lone fatha) returns `[]` like an empty query, instead of matching
every block.

**Deliberate non-changes.**

- **Title lookups stay case-only** (`sheets.get`, `headings.get`,
  `goToHeading`, §31 rule 5). They share `foldText`, not `searchKey`: a lookup
  names one item, and widening it would make `get("احمد")` answer for a sheet
  called `أحمد`, which the API page promises it does not.
- **Not folded:** ta marbuta against ha (`مدرسة` / `مدرسه`: some writers swap
  them, but it merges a feminine noun with a pronoun suffix), the hamza seats
  waw and ya (`مؤمن` / `مومن`) and a bare hamza, the Quranic annotation marks,
  and Persian ya and kaf against the Arabic letters. Each changes some words or
  is one community's habit; the first real miss reopens it. Recorded in
  ARCHITECTURE Known limits.
- PPTX notes are still not searched, and that is closed, not open: the viewer
  never shows them, so a hit would point at nothing the reader can see (the
  Known limits row for notes stands; it left the ROADMAP).

Pinned by `core-shared.test.ts` (each fold class, the kept distinctions, the
lone-mark query, titles unfolded) and, on the real fixtures, by
`search-index.test.ts` for each format: `الانتاج` and a vowelled, tatweel
spelling find `الإنتاج` in the deck and the workbook, `مُقَابَلَـة` finds the
DOCX cell, and Arabic-Indic digits find Western ones (`٩١٬٢٠٠ egp`, `H.264, ٢٤
Mb/s`) and the reverse (`لقطة 12` finds `لقطة ١٢`). No visual change, so no
Office review applies.

## 79. XLSX search covers one sheet, as Excel's Find does (Oct 2026)

`workbook.search()` scanned every sheet, hidden ones included, and returned
matches on all of them. The maintainer's read: a workbook is searched on the
sheet you are on. Excel's Find and Replace has "Within: Sheet" as its default
(Workbook is the opt-in), so a reader of an Excel file expects the same, and a
hit on a sheet that is not on screen is a jump the reader did not ask for.

**The shape.** `SheetHandle.search(query, { signal })`, returning
`XlsxSearchMatch[]` for that sheet. `XlsxDocument.search()` is removed, not kept
as an alias (pre-release stance, §42): the document has no selected sheet, so
there is nothing for it to default to, and a required sheet argument would be a
second spelling of `sheets.get(x).search()`. The viewer (`preview()`, the hooks,
the element) has a unit on screen, so `viewer.search()` searches that sheet
(`#index`, the one painted, not a move still in flight), and a search box
follows the tab the reader is on. PPTX and DOCX are unchanged: a deck or a
document is one thing to search, and a hit on another slide is the whole point
of the jump.

**Kept.** `XlsxSearchMatch.sheet` stays, so a match kept after the reader
changes tabs still says where it was found, and the type is unchanged. The
index (cell text per sheet, with the search key of §78) still lives on the
document so `close()` drops it with the rest; only the loop over sheets moved
out of it. Hidden sheets are searchable by asking for them
(`sheets.get("Hidden")`), as before.

**Blast radius.** The XLSX document's public surface (`search` leaves it,
`SheetHandle.search` arrives); `viewer.search()` for XLSX on every layer; the
API page (the sheet listing, "Searching a workbook", the search grammar, the
viewer table), the Start page's search box, `docs-parity`'s callable list
(`sheet.search`), and the tests and benchmark callers. The benchmarks
(`bench/src/measure.ts`, `tools/bench-search`) loop over every sheet so their
figures keep timing the whole file and the published numbers (§32) still mean
what they said. Cost: the first search on an XLSX viewer now reads one sheet
instead of all of them, which removes most of the cold cost the ROADMAP item
named for XLSX (it stays for PPTX).

**Deliberate non-changes.** No `workbook.search()` convenience that loops: an
app that wants Excel's "Within: Workbook" writes the loop over
`sheets.list()` and chooses its own order and progress, which is the low
level's job (§69). No option on `viewer.search()` to widen the scope, until a
real consumer wants it.

Pinned by `search-index.test.ts` (one sheet read for one sheet searched, the
index reused, a second sheet is its own read), `xlsx-document.test.ts` and
`chart-fills.test.ts` (a phrase written on two sheets is returned only for the
sheet asked, charts included), and `preview.visual.test.ts` (`viewer.search()`
returns only the sheet on screen, then only the next one after `goTo(1)`).

## 80. Search reads words, not models (Oct 2026)

§16 left the first search reading every unit as a full model and named a
Worker-side text-only message as the fix, against the model clone. Measured
before building it (the 300-slide bench deck, in Bun): the first PPTX search
took about 480 ms, the clone only about 50 ms of it, the full slide parse about
270 ms, and unzipping and parsing the slide XML alone about 100 ms (10-45% of
the full parse on the real decks). So the cost was mostly parsing what search
never shows (theme, colours, placeholder inheritance, geometry, the
background, notes), and the fix had to skip that work, not only the copy.

**What landed.** A `text` request (`TextRequest`, `session.text()`): the Worker
answers with words only. PPTX: `readPptxSlideText` (`core/pptx.ts`) walks the
slide's own shape tree as `parseShapeTree` does (shapes and connectors, groups,
tables, charts, Office 2016 charts) and keeps each paragraph's run text and
each chart's words. It shares what decides the words with the full parser
rather than copying it: `openSlide` (the part, tree, relationships and theme
every slide read starts from, split out of `parsePptxSlide`), `paragraphRuns`
and `runText` (`a:t` of `a:r`/`a:fld`, never `a:br`), `tableOf` and
`contentCells` (which cells hold content: merge continuations out, spans cut to
the grid, a missing `a:txBody` the same `invalid-document`), and
`parseChartFrame` for charts, which need the theme to read. XLSX: `sheetText`
(`core/xlsx/sheet.ts`), the page's cell-text loop moved into core, over the
parsed sheet (the cached one when it is on screen), so only strings cross.
DOCX refuses it (`unsupported-feature`): its page already holds the document.
The page adds the search keys (§78) and keeps the text as before.

**Measured after.** In Bun, same machine and session, before and after: the
first PPTX search 483 → 183 ms, the first XLSX search of the whole 13-sheet
workbook 1302 → 888 ms. In the browser bench (Chromium) against the published
run of 24 Sep, which predates §78 and §79 as well: PPTX 317 → 217 ms, XLSX 997 →
576 ms (one sheet, what a viewer search reads since §79, is about a thirteenth
of that). Repeat searches are unchanged (milliseconds). BENCHMARKS.md's search
table and the site's search data were regenerated from the new run; the size
tables were rendered from the same size files as before, so they did not move.
The Worker is 57.9 KB gzip (56.7 KB at §71/§73, before §78-§80).

**Guarded.** `pptx-text.test.ts` compares the walk with the full model's
paragraph text (each shape and table-cell paragraph, each chart's words, in
element order) on every slide of `pptx-coverage.pptx`, the production deck and
`chart-oracle.pptx` (and the 300-slide deck when it has been generated), and
asserts shapes, tables and charts all gave words, so an empty-against-empty
pass can't hide a gap. No fixture has words inside a group yet; the walk enters
groups as the parser does. `search-index.test.ts` asserts a search calls only
`text`, never `unit`, that a sheet already on screen gives the same matches,
and that DOCX refuses the read.

**Removed, now dead.** `UnitRequestOptions.fonts`, `geometry` and `retain`, and
`UnitRequest.retain`: search was their only user. A `unit` read always resolves
fonts and geometry and fills the cache; the Worker keeps `retain` only inside,
for the XLSX text read. `slideText(model)` left `aared/pptx` (the guard test
keeps its own reference). The package grows by about 65 lines, most of them
comments: the walk is about 40 lines, every helper it shares was moved out of
existing code, and the XLSX cell loop moved rather than grew.

**Deliberate non-changes.** No text read in the background before a search
(the next step if a real deck's first search is still felt); the walk still
parses each slide's XML, the floor measured above. The PPTX layout chain is
read for every slide even with no chart on it (it is cached per layout, and a
slide's theme needs it only for charts).

## 81. One theme rule on every viewer layer: light by default, `system` follows the OS (Oct 2026)

The ROADMAP carried a gap: `<aared-viewer>` listened for OS colour-scheme
flips, while `preview()` and the hooks resolved `system` once and relied on the
host calling `setAppearance()` again. The maintainer asked for one rule, with
two decisions.

- **The default is light, on every layer.** It already was: `resolveTheme()`
  gives light for an unset theme, and the element is fixed `light` unless the
  host opts in. Kept, and now written as the rule. Document content is always
  light (§25), so a dark frame is the exception a host asks for, and an embed
  must not surprise a light site because the reader's OS is dark.
- **A host that takes control is followed, everywhere.** An explicit `light`
  or `dark` is never moved by the OS (true already); `system` now follows the
  OS live on `<aared-viewer>`, `preview()` and the hooks alike.

**Root cause.** The listener lived in the element, the one layer that knew
about the OS, while the logic that applies an appearance lives in the shared
viewer (§65, one viewer for `preview()` and the element). The element resolved
`system` itself and forwarded `light` or `dark`, so the viewer never saw the
word `system` and could not follow it.

**What changed.**
- `watchSystemTheme(onChange)` in `appearance.ts` is the one place anything
  listens for the OS scheme. It returns the stop function; a host without
  `matchMedia` never fires.
- The shared viewer (`PreviewViewer`) listens exactly while its appearance's
  theme is `system`. On a flip it applies the same appearance again and each
  renderer resolves it afresh, so the unit, zoom, selection scroll and Worker
  session stay. `setAppearance()` re-evaluates (a theme other than `system`
  drops the listener), and `destroy()` drops it. Every renderer already
  re-resolves on `setAppearance()`, so none changed. The hooks sit on
  `preview()` and get it with no code of their own.
- The element now forwards its raw `theme` (`system` included) and keeps only
  its own watcher, for its chrome (`data-aared-theme`, the stage and controls
  colours); it no longer re-forwards a resolved theme on a flip. Its resolved
  theme field and its private `matchMedia` code are gone.

**Blast radius.** `appearance.ts`, `embed-viewer.ts`, `element/index.ts`; the
renderers, hooks and `preview()` source are unchanged. The low-level `render()`
still resolves `system` once when it paints (an app that owns the handle calls
`setAppearance()` again), now said on the API page's Dark mode section.

**Deliberate non-changes.** No new event: the viewer still fires none for an
appearance change (§65), and the element reads the OS itself for its chrome. No
change to light as the default, or to what dark changes (chrome only, §25).

Pinned by `appearance.test.ts` (the watcher: calls until stopped, independent
watchers, no `matchMedia`) and `system-theme.visual.test.ts`, which replaces
`matchMedia` with one the test flips: on `preview()` for PPTX, XLSX and DOCX
(content follows both ways with the unit and zoom kept, an explicit theme never
moves, `setAppearance` to a fixed theme drops the listener and back to `system`
resolves from the OS as it is now, `destroy()` leaves none), and on the element
(content and chrome follow, the listener counts of one and two, a fixed theme
stops the viewer's, removal drops all). The test fails with the following
switched off.

## 82. DOCX text boxes and section line grids (Oct 2026)

DOCX text boxes now resolve DrawingML `wps:txbx` and legacy VML `w:txbxContent`
into the existing paragraph/list/table flow model. `mc:AlternateContent` selects
the DrawingML Choice or VML Fallback once, avoiding duplicate text. Inline boxes
paint as fixed-size DOM containers; anchored boxes travel through the existing
float planner. Body search includes their text and records a per-section box
index. Headers and footers remain outside search, as before.

Section `w:docGrid` line pitch is carried in the resolved section model. The
renderer snaps paragraph line boxes up to grid multiples, resolves positive
`beforeLines`/`afterLines` against the pitch, honors `snapToGrid="0"` and exact
line spacing, and applies the grid in tables only when
`w:adjustLineHeightInTable` enables it. A missing line pitch uses Word's 18pt
default measured by the `line-spacing-units` probe. Text remains browser DOM
text; no raster or canvas path was added.

**Root cause and reach.** Text-box paragraphs had no model representation, so
their content was dropped by the parser and absent from rendering and search.
Paragraph spacing had preserved line units but substituted a fixed 12pt value
because section grids were not carried through to paint and measurement. The
changes touch the shared DOCX parser/model, flow and page renderers, pagination
measurement, search, API docs and fixture description. No PPTX or XLSX route
uses these DOCX structures.

**Limits and non-changes.** Text boxes preserve fixed extents, insets, fill,
outline and clipping/overflow, but `normAutofit` font scaling and `spAutoFit`
frame growth are not implemented. VML anchor interpretation remains less
complete than DrawingML. Header/footer search was explicitly left unchanged.
The existing Word-authored coverage fixture gains an Arabic/English anchored
box with a Choice/Fallback pair. See `ARCHITECTURE.md` DOCX rules and
`ROADMAP.md` for the remaining search scope.


## 83. DOCX header and footer search (Oct 2026)

DOCX search now includes resolved header/footer stories, deduplicated by source
part so inherited copies produce one result. The render/preview controls can
highlight and navigate to every displayed occurrence, including repeated
headers, using native CSS Custom Highlights rather than wrappers that alter
layout. PPTX/XLSX were unchanged at this point; hidden header variants remain
searchable but cannot be navigated to when not rendered. Section 84 later
replaced these DOCX-only highlight controls with shared `viewer.search()`.


## 84. Viewer search owns highlighting across formats (Oct 2026)

Section 83's DOCX-specific highlight controls duplicated search state and are
removed. `viewer.search(query, { highlight, signal })` owns highlighting across
PPTX, XLSX and DOCX; highlighting is on by default, `highlight: false` clears
presentation state, and an empty query clears it. Headless document/sheet search
remains data-only.

A lazy main-thread DOM/SVG range controller keeps the query active across unit
changes and XLSX virtualization without another Worker read or model/RPC change.
It preserves prior highlights after cancellation/failure and prevents stale
searches from painting. Preview, React and element viewers share the behavior;
low-level render handles do not expose highlight controls. The bench uses one
app-owned search panel for all formats. Rendered-range browser checks cover the
shared paths. Saved field text remains searchable; PAGE/NUMPAGES highlights use
their displayed values.

## 85. PPTX search results count text occurrences (Oct 2026)

PPTX search now returns every non-overlapping matching occurrence, including
repeats within one paragraph, rather than only the first match per slide. It
retains the `{ slide, text }` result shape and existing text index; no extra
Worker reads or model/RPC changes are needed. XLSX remains per-cell and DOCX per
source block. Fixture and rendered-highlight checks pin result/highlight parity.


## 86. Object-form preview and reactive React appearance (Oct 2026)

`preview(source, target, options)` differed from the renderers' object-form
grammar and hid argument names from completion hints. It is replaced with
`preview({ source, target, ...options })`, typed by `PreviewInput extends
PreviewOptions`. There is no compatibility alias or overload during pre-release.
All callers, demos, checked examples and consumer-build probes use the new form;
headless `open()` and `detect()` keep their existing source/options contracts.

React hooks previously read appearance only at mount, requiring consumers to
write an effect for theme or spreadsheet chrome updates. They now compare
appearance fields by value and apply changes through the existing viewer. Removing
the option resets defaults. Inline object identity and property order are ignored,
so an unrelated rerender does not undo an imperative `setAppearance()` call; the
next option-value change or new viewer takes ownership again. Options changed
during loading apply to the viewer when ready. Index and zoom remain start values.
The bench no longer needs its theme-forwarding effect.

The shared hook covers auto-detect and all three format hooks. Real-fixture browser
tests re-read theme markers, spreadsheet headers/tabs, index, zoom and selection,
check viewer/Worker reuse, property-order equivalence, removing appearance,
StrictMode and appearance changed while loading. Existing preview, element, search
and CSP tests exercise the object-form migration. No Worker or model changes or
new dependencies are needed.

## 87. Active find position and compact navigation (Oct 2026)

The bench's result list did not express the current occurrence or support
previous/next find navigation. It is replaced with an input, active/total counter
and wrapping arrows; Enter advances and Shift+Enter goes back. The panel still
starts closed and uses a native search input. Search UI remains application-owned.

The shared viewer exposes `searchNavigation` (`index`, `count`, `next`,
`previous`) and `searchchange`. Existing `search()` results stay plain data.
PPTX navigates occurrences across slides, XLSX navigates matching cells on the
active sheet through the existing cell navigation (including virtualized rows),
and DOCX navigates painted ranges including every displayed header/footer copy.
DOCX's find count can exceed its deduplicated source-result count. Empty and
data-only searches reset find navigation; newer searches prevent stale navigation
from selecting a range.

All ranges use a soft CSS Custom Highlight. A clipped, pointer-transparent
overlay outside the text tree outlines the active range with a stronger fill.
Native `::highlight` supports colors and text decoration, but not borders
([MDN](https://developer.mozilla.org/en-US/docs/Web/CSS/Reference/Selectors/::highlight));
using range rectangles preserves layout and bidi text. Scroll/resize observation
repositions it and teardown removes it. No Worker, model or RPC change is needed.
Real-fixture browser tests re-read counters, active rectangle geometry, wrapping,
clearing and repeated header copies, alongside existing virtualization coverage.
The bench pager now follows the actual PPTX viewer index when find navigation
switches slides. The lazy highlighter is 5,767 bytes (2,596 gzip); the Worker
remains 177,846 bytes. Production consumer checks exercise the lazy chunk.

## 88. Complete viewers from single-format entries (Oct 2026)

A known-format consumer previously chose between low-level `open()` / `render()`
with manual navigation and two lifetimes, or generic preview whose dynamic
imports made bundlers emit sibling renderer chunks. Each format entry now exports
`preview({ source, target, ...options })`, returning `PptxPreview`, `XlsxPreview`
or `DocxPreview`, with navigation, zoom, appearance, search, events and one
`destroy()` that releases renderer and document.

The shared controller receives an internal open/render route; it does not import
formats. Each local entry supplies its own route. Detection and dynamic format
imports move to `auto-viewer.ts`, used by generic preview and the element; React
continues through generic preview. This preserves one controller and consistent
behavior without making a fixed-format consumer reach the other renderers.
Input types omit `type`; PPTX and XLSX also omit DOCX-only pagination. Low-level
open/render contracts and their separate ownership remain available unchanged.
No compatibility alias, new dependency, Worker, model or RPC change is needed.
The Worker remains self-contained with all parsers; this improves main-thread
renderer isolation, not Worker size.

Real-fixture browser checks cover all three local entries, navigation races,
search and visible highlights, data-only search, theme/zoom, repeated cleanup,
wrong-format input and cancellation. Production Vite builds inspect final module
graphs for sibling renderer and auto-router exclusion, then exercise search and
cleanup. Existing generic Vite/webpack consumer probes and preview, React and
element browser checks validate the shared paths. Checked examples document both
convenient viewer ownership and lower-level document sharing.

## 89. Reference apps use the convenient viewer APIs (Oct 2026)

The homepage and measurement adapters still assembled open/render and separate
cleanup after format-local preview became available. They now use preview with
one destroy; initial-render failure and cancellation cleanup belong to the viewer.
Measurement timing boundaries stay unchanged, XLSX tabs remain off and DOCX layout
remains on. Published measurement tables are not regenerated: old PPTX navigation
repainted the current slide, whereas viewer navigation to that slide is a no-op.

The interactive bench's frame previously required dummy destroy functions and
copied the hook's slide/sheet indices into local state. Its display contract now
omits destroy, while cards retain lifecycle ownership. Slide and sheet controls
read hook state directly and navigate through the viewer. DOCX paging calls
viewer.docx.goToPage instead of inspecting aared's DOM page selectors. Baseline
navigation and DOCX scroll-based page observation remain necessary; the relative
zoom controls keep their existing fit semantics. Search UI and low-level
playground examples are deliberately unchanged.

Typechecks and production bench/site builds cover the consumers. Bun.WebView
checks exercise search, theme changes, measured navigation/cleanup, DOCX paging
and homepage format switching. The site build also caught a literal object-form
brace in playground MDX, now escaped so the displayed syntax remains unchanged.

## 90. Empty section-ending page-break paragraphs do not add pages (Oct 2026)

An empty paragraph carrying both `w:pageBreakBefore` and `w:sectPr` produced an
extra blank page: it forced a break before itself, then the next section started
on another page.

The current fixture was copied byte-for-byte and freshly exported through Word:
it reports seven pages (portrait, portrait, portrait, landscape, portrait,
landscape, portrait). A new aared render of those exact bytes also reports seven.
The seven picture-wrapping examples share page 5, charts remain on page 6, andthe final list/table section remains on page 7. Microsoft documents
`w:pageBreakBefore` as starting a paragraph on a new page and paragraph-level
`w:sectPr` as the properties of the section ending at that paragraph mark;
the standard and Microsoft references do not state how Word treats their
combination on an empty paragraph. The exact behavior is established by this
same-file Word 16 export and the paired aared capture, not inferred from the
definitions alone. See [Microsoft's page-break property](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.paragraphproperties.pagebreakbefore?view=openxml-3.0.1),
[paragraph section properties](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.paragraphproperties.sectionproperties?view=openxml-3.0.1),
and the [ECMA-376 standard](https://ecma-international.org/publications-and-standards/standards/ecma-376/).

The parser now omits the flow node only when that paragraph is empty and has
both properties. Ordinary empty paragraphs and nonempty section-ending
paragraphs retain their current behavior. The change affects the DOCX parse and
layout route; it does not change explicit page breaks, even/odd section filler
pages, or PPTX/XLSX. Parser, pagination and painted-DOM checks assert the seven
pages, all seven wrapping examples on page 5, charts on page 6 and the final
section on page 7. Other line-wrap differences from Word remain open.


## 91. DOCX pagination stays DOM-measured; fidelity work is bounded by visible value (Oct 2026)

The DOM-measured approach in §21 remains the right layout source: text shaping,
wrapping and bidi should come from the browser that paints the document. Word 16
and aared both produced seven pages in the October 2026 review.

The fresh page-3 pair retains approximately 2–5px of vertical drift, with the
delivery table on the same page. Retire that historical roadmap item as
good-enough residue. The anchored-picture page 5 has larger accumulated
spacing/wrapping differences, reaching about 54px at a unique text anchor near
the bottom; those remain with the existing anchored-picture gap.

The automated report is diagnostic evidence requiring review. All 25 reported
page mismatches involve repeated tokens; none of the 435 tokens occurring
exactly once in both exports changed pages. The comparator sorts each token's
occurrences by position and pairs them without context. PDF extraction splits
the spaced header into letters and produces different Arabic strings, while
browser text is split at DOM text-node boundaries. These differences change
occurrence counts and can create false pairings, large offsets and unmatched
text. Missing/extra groups do not establish missing/extra rendered content.
Improve matching and report ambiguity before making this an automatic fidelity
gate; retain paired images and unambiguous anchors as the review method.

Native Word/WebView review requires the approved Windows execution path outside
the sandbox; rasterization and comparison run inside it. This is an automation
environment constraint, not a renderer issue.

The measurement boundary is explicit: layout mode measures eligible top-level
blocks and table row tracks after waiting for fonts. Standalone images/charts
and footnote/header/footer reservations still use estimates; headless layout
uses estimates throughout. Change those paths only where a paired Word review
shows a reader-noticeable difference. Pixel-identical pagination is not the
goal.

Table continuation is a separate visible gap. Current pagination may break a
table between rows, but never fragments an individual row or a paragraph/list
block, and it does not repeat `w:tblHeader` rows. Word permits row splitting by
default unless the row or table style prevents it; an over-page row still flows
across pages. Word also repeats contiguous leading rows marked `w:tblHeader`.
These behaviors should be considered together only when a real fixture shows
the long-table case is worth supporting. The probe must cover row properties
and their style inheritance before implementation; `w:trHeight` and related
row properties belong under `w:trPr`.

Blast radius is DOCX parsing, the model, the pure paginator, the DOM renderer,
the Word review tools, and the Arabic DOCX fixture. Deliberate non-changes:
keep the browser as layout engine; do not add a second text layout engine or
chase sub-pixel parity; do not expand pagination based only on theoretical
coverage. Record any residual that readers would not notice as a known limit;
keep visible clipping, wrong-page content, or missing repeated headers as open
gaps when the fixture demonstrates them. The architecture and roadmap now
distinguish current row-boundary splitting from unsupported within-row
continuation.

## 92. Numbered DOCX headings and refreshed table-of-contents results (Oct 2026)

The real Word-authored probe confirms the marker gap: Word paints
`1. Project Overview`; aared paints only `Project Overview`. The heading text
has no typed numeral, and its actual list marker is positioned at the paragraph
start before the tabbed heading text. Sequence advancement already works (§77);
painting the marker remains open because the heading node has no computed marker.

The same probe exposes a page-1 TOC discrepancy. Word's export refreshes the TOC
to include `5. Delivery tracker / متابعة التسليم` and the table's Heading 2 text
`الإجمالي`; aared displays the older cached result. Word updates fields in a
private read-only export copy; aared recalculates `PAGE`/`NUMPAGES` and otherwise
paints saved field results. Refreshing TOC entries with their switches and range,
or defining the cached-result policy, remains open in ROADMAP.

The probe and captures are in gitignored render output; no checked-in fixture
changed.

## 93. Tools consolidation: one Office export driver, one rasterizer, a standard probe shape (Oct 2026)

`tools/` had grown by accretion: three near-identical `*-render/index.ts`
drivers, two `rasterize.py` copies, some twenty-five copies of the probe
COM/temp-copy boilerplate, six names for three probe jobs, and snapshot
trees in four dialects. Merged to one skeleton with parameters, per the
shared-modules rule: `tools/office-export.ts --format <pptx|xlsx|docx>`
drives all three native exports (each pipeline keeps only its own
`export.ps1`, whose COM object model is genuinely per-app), and
`tools/rasterize.py --mode <sheet|document>` rasterizes both (the sheet
fit-to-page assertion and the Word span/word geometry dump stay behind
their modes). JSON output shapes are unchanged, so the two visual-test
callers only changed spawn lines.

Probes get the same treatment going forward: one directory is one
question with one role script (`probe.py`, `author.py`, `extend.py`), a
README, and shared helpers imported from `tools/probe/_common/` (COM
guards, fixture temp-copy, unzip snapshot, plus `strip_label.py`, moved
out of `pptx-lines/` and documented as the all-three-formats label
strip). The font bench moved to `tools/font-calibration/` — it says
itself it is not an OOXML probe. The three mixed-topic directories were
split (`docx-rtl-table`, `docx-delivery-tracker`, `exact-pitch`,
`ltr-runs`) and same-job scripts renamed to the standard roles; the eleven
`line-spacing-units/make-*.ps1` authors folded into one `author.ps1` (one
function per case, shared COM sessions, same output names); finished
snapshot trees were grandfathered, not renamed. Finished probes are
deleted once comment + fixture + test land (AGENTS.md probe rule, step
6), so the count shrinks instead of growing. `tools/probe/` stays
gitignored scratch; only `_common/` and the probe README are tracked.
