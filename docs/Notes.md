# Engineering notes

> Purpose: preserve the package's accumulated engineering experience: decisions,
> reasons, rejected alternatives, measurements, Office findings, pitfalls and
> deliberate non-changes. Read this before rediscovering a rule the hard way.
> Current behavior, gaps and release readiness belong in
> [ARCHITECTURE](ARCHITECTURE.md); the public contract is
> [the API reference](../site/src/content/docs/api.mdx).
> Changes after the baseline belong in [CHANGELOG](../CHANGELOG.md).

## Reading and maintaining these notes

All **128** inherited entries of the former Decision Log remain grouped by
topic, with their original IDs and titles (there was never a §41). Stable
anchors such as `#note-3` survive future title changes. The baseline first
preserved their bodies verbatim; this editorial pass replaces repeated current
mechanisms with links to their canonical home, without deleting the findings,
measurements, rejected alternatives or reasons that make an entry useful.

**These are experience records, not a second current-state inventory.** Dates,
old API shapes, tool paths, retired filenames and superseded claims remain in
their historical context. Entries often explain what replaced earlier choices;
follow the correcting entry and ARCHITECTURE before treating an old statement
as today's contract. In particular the early DOCX pagination, Worker splitting,
font delivery and model/lifecycle decisions were revised. Historical ROADMAP
mentions describe that retired document, not a promise to recreate it. Do not
erase a rejected approach or an older measurement just because it is no longer
the implementation; annotate new evidence and link the replacement instead.

Write an experience note around the failure/question, how it was investigated,
what evidence changed the decision, alternatives and costs, blast radius and
lessons. Include enough mechanism to explain the result, not another complete
API, lifecycle or renderer specification. Keep unique measurements and probe
traps; replace a duplicated contract with a precise link, not a vague summary.
Historical shapes remain when their rejection or correction is the lesson.

Future changes belong in CHANGELOG; current rules and gap dispositions belong
in ARCHITECTURE; call shapes belong in the API reference. Add stable IDs after
§129 without renumbering inherited entries. No pruning is justified solely by
age, length, supersession or a fact also appearing in source.

The original migration provenance, history-squash details and complete
[decision-ID index](archive/CHANGELOG-pre-baseline.md#decision-id-and-migration-index)
are preserved in the single pre-baseline archive.

### Corrections worth following across topics

- Early DOCX pagination boundaries: [§7](#note-7), [§10](#note-10),
  [§22](#note-22) and [§91](#note-91).
- Split Worker experiments and reversal: [§18](#note-18), [§20](#note-20);
  static parser imports are explained in [§114](#note-114).
- Delivered-font experiments and the default-font decision: [§45](#note-45)
  through [§53](#note-53); retained calibration is evidence, not an active task.
- Borrowed document/content capabilities: [§119](#note-119) is narrowed by
  [§122](#note-122), without losing the rejected lifecycle exposure rationale.
- Earlier delivery/status claims: [§113](#note-113) explicitly corrects them;
  [§129](#note-129) explains why the roadmap became one architecture inventory.

## Topics

- [Design, scope and native evidence](#design-scope-and-native-evidence) — 11 entries
- [API, viewer lifecycle and errors](#api-viewer-lifecycle-and-errors) — 26 entries
- [Worker, packaging and memory](#worker-packaging-and-memory) — 16 entries
- [Shared rendering, paint and host isolation](#shared-rendering-paint-and-host-isolation) — 8 entries
- [Performance, benchmarks and navigation](#performance-benchmarks-and-navigation) — 5 entries
- [PPTX fidelity](#pptx-fidelity) — 4 entries
- [DOCX layout, tables and drawings](#docx-layout-tables-and-drawings) — 22 entries
- [Fonts, script and line measurement](#fonts-script-and-line-measurement) — 8 entries
- [Search and highlighting](#search-and-highlighting) — 6 entries
- [Charts](#charts) — 3 entries
- [XLSX layout and conditional formats](#xlsx-layout-and-conditional-formats) — 10 entries
- [Quality, tooling and documentation](#quality-tooling-and-documentation) — 9 entries

## Design, scope and native evidence

**Current boundaries:** [ARCHITECTURE](ARCHITECTURE.md#purpose-and-boundaries).
The entries below explain their origins, alternatives and native evidence.

<a id="note-1"></a>

### 1. Why this project exists at all

The trigger was concrete, not speculative: evaluating client-side renderers
for an existing chat product's document-delivery panel surfaced two
disqualifying findings, preserved as primary evidence in the archived
[PPTXViewJS observations](archive/CHANGELOG-pre-baseline.md#research-pptxviewjs) and
[SheetJS observations](archive/CHANGELOG-pre-baseline.md#research-sheetjs)
(recovered alongside the original [design](archive/CHANGELOG-pre-baseline.md#research-design)):

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

<a id="note-2"></a>

### 2. PDF was scoped out, deliberately

`pdf.js` already solves client-side PDF rendering well. The gap that
motivated this project doesn't exist for PDF — its layout is already fixed
and its glyphs are already positioned by the time the file exists. Spending
effort here would be solving an already-solved problem instead of the actual
gap. Consumers who need PDF too pair aared with `pdf.js` directly.

<a id="note-3"></a>

### 3. Text is rendered by the browser, never rasterized — the load-bearing decision

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

<a id="note-4"></a>

### 4. View-only, permanently — not a phased-in decision

Considered and rejected: building toward eventual editing. Editing requires
a formula engine, round-tripping (preserving everything not touched, the way
shariha's PPTX editor does), and lossless-write guarantees — an entirely
different, much larger project. Scoping to view-only from the start halves
the problem and keeps the fidelity contract honest: aared promises to show
you what's in the file, not to let you change it.

<a id="note-5"></a>

### 5. Build order: PPTX → XLSX → DOCX, not alphabetical or "easiest first"

DOCX is structurally different from the other two: PPTX and XLSX both have
*bounded* layout problems (absolutely-positioned shapes; a grid). DOCX is a
flow-layout problem — line breaking, pagination, floats — which is, honestly,
building a word processor's layout engine. That's unbounded scope if taken
on first, before the shared `core` layer (OPC, DrawingML, charts, theme) has
even been proven out on two bounded formats. PPTX goes first over XLSX
because Arabic/RTL correctness — the whole reason this project exists — is
most visibly, immediately testable there (a slide either has right-aligned
Arabic text with correct bullet placement, or it doesn't).

<a id="note-6"></a>

### 6. Packaging: one package with lazy subpath exports, not `@scope/*` packages

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

<a id="note-8"></a>

### 8. PowerPoint is the text reference; LibreOffice is a companion artifact

Considered and rejected: asserting pixel-identical output against some
ground truth. No two OOXML renderers agree pixel-for-pixel even when both
are "correct" — shariha's own decision log (§15 in its `DECISION_LOG.md`)
independently reaches the same conclusion for a related problem: SVG/pdf-lib,
resvg, and Chromium-HTML painters are different painters, and demanding
identity across them is the wrong bar. But LibreOffice is not PowerPoint:
its font substitution can change Arabic glyph metrics, line wrapping, and
baseline placement. Using its raster to judge PPTX text would accept the wrong
output or reject the right one.

The resulting review policy is in
[ARCHITECTURE](ARCHITECTURE.md#native-references-and-developer-tools).
The lesson is to choose the oracle for the property being judged: LibreOffice
conversion can help broad inspection, but its substituted text is not evidence
of where PowerPoint places the author's glyphs. Pixel disagreement alone says
less than a native pair plus a DOM measurement of the disputed property.

<a id="note-9"></a>

### 9. PowerPoint probes settle renderer disagreements

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

<a id="note-24"></a>

### 24. Dark mode, looked at: automatic text follows its fill; authored text on the dark surface is an open gap

Stale. A contrast audit of the coverage fixtures rendered dark found that
adaptive dark made authored fills with automatic text unreadable, and that
authored text colours on the darkened surface (Excel black body cells, Word
`1B2029` body copy) read at 1 to 3.4:1, so whole sheets and documents looked
blank. A first fix resolved automatic text against its fill (`automaticInk` in
`appearance.ts`, on `contrastRatio` in `core/color.ts`), which stays, and a
recolouring rule was proposed for authored text. Instead the adaptive mode
itself was removed, because any fix recolours authored values and contradicts
the view-only fidelity promise. Superseded by §25.

<a id="note-25"></a>

### 25. Adaptive dark retired: every format is source-faithful

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

<a id="note-42"></a>

### 42. npm name: `@aared/viewer`, one package, reserved with a placeholder

The plain name `aared` cannot be published: npm's similarity check rejects it
against an unrelated, empty `aa_red` (403 "too similar"). Scoped names skip
that check, so `aared` is an npm org and the package is **`@aared/viewer`**
(matches the `<aared-viewer>` element). `@aared/viewer@0.0.0` is a
placeholder holding only a README, published 2026-09-25; the repo's code is
not in it. The repo tag `v0.0.0` marks that point.

**One package, subpaths, stays** (§20): `@aared/viewer/core`, `/pptx`, `/xlsx`,
`/docx`, `/preview`, `/react`, `/element`, `/worker`. Per-format packages were
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


## API, viewer lifecycle and errors

**Current ownership and behavior:**
[ARCHITECTURE](ARCHITECTURE.md#viewer-behavior-and-ownership).
**Calling contracts:** [the API reference](../site/src/content/docs/api.mdx).
Earlier signatures and policies below are retained only where the change,
rejection or correction explains a design choice; they are not pasteable recipes.

<a id="note-12"></a>

### 12. `aared/preview`: superseding the deferred facade decisions

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

<a id="note-13"></a>

### 13. DOCX `render()` joins the shared async object envelope

DOCX rendered positionally and synchronously — `render(sections,
container, options?)` — while PPTX/XLSX rendered as `await render({ slide |
sheet, target, ... })`. Every consumer (bench, element, `preview()`) had to
branch on call shape, not just on format; the shared-grammar design (same
envelope, format-native nouns) exists to eliminate that.

Choosing the same async object envelope removed a branch every integration had
to write. The former positional form was removed, not aliased, while unpublished.
The [DOCX render contract](../site/src/content/docs/api.mdx#rendering-a-document)
owns the current input union and signature. Two gains beyond symmetry:

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

<a id="note-14"></a>

### 14. `aared/react`: the mount lifecycle as hooks, not docs prose

`preview()` removed the per-format open/render boilerplate, but every React
embed still rewrote the same effect: `AbortController` per mount, capture
the handle for cleanup, swallow aborted rejections, mirror `index`/`zoom`
into state, destroy on unmount. That logic is fixed — it never varies by
app — so it belongs in the package, not in every consumer.

The format wrappers were chosen for readability, not separate lifecycles:
"preview this one PPTX" reads as `usePptx`, without requiring a caller to spell
the format assertion. One implementation prevents format-specific cleanup drift.
The current [hook surface](../site/src/content/docs/api.mdx#react-hooks-aaredreact)
and [ownership](ARCHITECTURE.md#one-controller-several-entry-points) are canonical;
§69 later widened the original facade boundary to include search.

Two packaging choices, both about not taxing non-React consumers:

- `react` (`>=18`) is an *optional* peer behind the `aared/react`
  subpath. Vanilla and element consumers never install it; the entrypoint
  is import-time safe, and `react` itself is never bundled.
- Plain-data options are keyed by value inside the effect so inline
  literals (`appearance={{ theme }}`) don't remount; `source` and the
  function resolvers are identity-keyed and documented as must-be-stable.
  StrictMode mount → cleanup → remount resolves exactly once.

<a id="note-15"></a>

### 15. `capabilities` cut before first publish; `search()` and the headless boundary kept

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

<a id="note-31"></a>

### 31. One API grammar across formats and layers

**Why.** An audit of the API reference against the code found about thirty
mismatches, several real bugs (XLSX fit modes did nothing; a closed document
reported `"aborted"`, which `isAbort()` swallows; DOCX heading indexes counted
header and table headings), and a second pass found seventeen places where
the same behaviour had different syntax per format or layer (`SheetInfo.name`
vs `SlideInfo.title`; different event-detail shapes; sync DOCX `search`;
`docx: { pagination }` in the facade but top-level in `render()`; three
per-format appearance types). Pre-release, so best design wins: every caller
moved in the same change, with no aliases.

**What the grammar clarified.** Units are not places: Word's outline is
structure and its pages are layout, so a DOCX document remains one unit even
when it has many physical pages. Named navigation stays asynchronous; typed
`UnitRef` and `CellRef` replace layer-specific spellings. Exact/case-insensitive
title lookup is deliberately distinct from substring search. Events describe
changes rather than inventing a change on initial load. Named object inputs
and one top-level `pagination` avoid both positional ambiguity and per-layer
nesting. The shared render-handle base removed the exported XLSX-only viewer
type instead of adding a compatibility name.

Several choices in that original grammar were revised, not silently retained:
appearance replacement became merge in §119; the current-unit no-op gained an
in-flight exception in §68; teardown no-ops were narrowed by typed closed-use
checks (§57–§59); and dual error emission/rejection was replaced by action-owned
reporting in §120. Those reversals are experience, not today's contracts.
The [API design principle](../site/src/content/docs/api.mdx#design-principle),
[navigation](ARCHITECTURE.md#navigation-appearance-and-cancellation) and
[errors](ARCHITECTURE.md#errors) own the current rules and data shapes.

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

<a id="note-57"></a>

### 57. Every document failure is typed; a lost session is `closed`, not `aborted` (Sep 2026)

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

<a id="note-58"></a>

### 58. Error codes say whose fault it is (Sep 2026)

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

<a id="note-59"></a>

### 59. Small API faults from a developer-ergonomics review (Oct 2026)

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

<a id="note-61"></a>

### 61. The hook's result narrows by `status`; its zoom and index are live (Oct 2026)

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

<a id="note-62"></a>

### 62. Bytes in as a typed array; every event typed (Oct 2026)

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

<a id="note-63"></a>

### 63. The API page says less, and its examples are compiled (Oct 2026)

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

<a id="note-64"></a>

### 64. Options start it, the viewer changes it; the element runs on `preview()`'s viewer (Oct 2026)

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

<a id="note-65"></a>

### 65. A stale render never paints; the hook owns teardown; the element lays out pages (Oct 2026)

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

<a id="note-66"></a>

### 66. Emptying `<aared-viewer>` is not an error (Oct 2026)

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

<a id="note-67"></a>

### 67. The viewer has toolbar methods and says its default zoom (Oct 2026)

**The problem.** A custom toolbar over `preview()`, the hooks or
`el.viewer` had to rebuild what the element's built-in bar did privately:
paging with bounds, zoom steps that start from the scale on screen in a fit
mode, clamping and rounding, and a reset. The reset needed each format's
opening zoom, which was a private table (`DEFAULT_ZOOM`). For a slide or a
document "100%" is not that default, since `setZoom(1)` pins the natural
size, so the bench kept its own copy (`PAGER_DEFAULT_ZOOM`) and the site's
demo bar a copy of the step rules.

**Why both methods and a default value.** The user chose both so a custom
toolbar need not recreate the algorithm or guess a format's reset target.
The [navigation/zoom API](../site/src/content/docs/api.mdx#navigation) owns
the complete current method contract. Two corrections explain the design:
stepping from the unit on screen made three rapid presses move only one unit,
so steps follow the newest pending destination; the first zoom proposal used
a multiplier, but the user's meaning was additive percentage points
(`zoomIn(0.1)` means 100% to 110%, not a 1.25 multiplier). Fit-mode steps must
start from painted scale, not from an abstract fit label.

The limits stay fixed (not parameters): they are where the renderers stay
legible and virtualization stays cheap, and no caller asked to move them.

**One copy.** The element's bar now calls these methods, and its own copy of
the step rules is gone. The bench uses `resetZoom()` and dropped its table.
The site's demo bar keeps its own rules because it also drives headless
render handles, which have none of these; it follows the headless paging
decision.

<a id="note-68"></a>

### 68. An aborted render never hands back a handle (Oct 2026)

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

<a id="note-70"></a>

### 70. A Start page for newcomers; the API page is the reference (Oct 2026)

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

<a id="note-81"></a>

### 81. One theme rule on every viewer layer: light by default, `system` follows the OS (Oct 2026)

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

**Lesson from the correction.** Forwarding the resolved colour had erased
the host's intent before the shared controller could act on it. Moving the
watcher to the layer that owns appearance, and forwarding the raw `system`
choice, fixed every viewer route without modifying the renderers. The element
still needs its own chrome watcher; that is a different owner, not duplicated
document-state logic. The present watcher/lifetime rules are in
[ARCHITECTURE](ARCHITECTURE.md#navigation-appearance-and-cancellation).

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

<a id="note-86"></a>

### 86. Object-form preview and reactive React appearance (Oct 2026)

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

<a id="note-87"></a>

### 87. Active find position and compact navigation (Oct 2026)

The bench's result list did not express the current occurrence or support
previous/next find navigation. It is replaced with an input, active/total counter
and wrapping arrows; Enter advances and Shift+Enter goes back. The panel still
starts closed and uses a native search input. Search UI remains application-owned.

**Why occurrence state is separate from results.** Plain source results do
not identify the painted occurrence a reader is on: a DOCX header part may
appear on many pages. The find counter therefore can exceed the deduplicated
result count. Reusing XLSX cell navigation also preserves virtualization rather
than inventing a second scroller. The current
[navigation/highlight rules](ARCHITECTURE.md#search-and-highlights) own reset,
stale-query and cross-format activation behavior; the
[viewer API](../site/src/content/docs/api.mdx#auto-detect-preview-aaredpreview)
owns the methods/events.

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

<a id="note-88"></a>

### 88. Complete viewers from single-format entries (Oct 2026)

A known-format consumer previously chose between low-level `open()` / `render()`
with manual navigation and two lifetimes, or generic preview whose dynamic
imports made bundlers emit sibling renderer chunks. Each format entry now exports
`preview({ source, target, ...options })`, returning `PptxPreview`, `XlsxPreview`
or `DocxPreview`, with navigation, zoom, appearance, search, events and one
`destroy()` that releases renderer and document.

The useful separation was dependency direction, not another controller:
format routes supply what varies without making the controller import them.
That gave fixed-format consumers the convenient lifecycle without sibling
renderer costs. The Worker still contains every parser, so this is a
main-thread isolation result, not a Worker-size claim. No alias, dependency,
model or RPC change was needed. The current
[controller map](ARCHITECTURE.md#one-controller-several-entry-points) and
[local-preview inputs](../site/src/content/docs/api.mdx#single-format-preview)
own the routing, option exclusions and lifetimes.

Real-fixture browser checks cover all three local entries, navigation races,
search and visible highlights, data-only search, theme/zoom, repeated cleanup,
wrong-format input and cancellation. Production Vite builds inspect final module
graphs for sibling renderer and auto-router exclusion, then exercise search and
cleanup. Existing generic Vite/webpack consumer probes and preview, React and
element browser checks validate the shared paths. Checked examples document both
convenient viewer ownership and lower-level document sharing.

<a id="note-89"></a>

### 89. Reference apps use the convenient viewer APIs (Oct 2026)

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

<a id="note-119"></a>

### 119. Viewers lend their document, take any search result, and merge appearance updates (October 13, 2026)

An outside review of the consumer-journey audit (`docs/USE-CASES.md`) named the
places where an app outgrowing "show this file" suddenly had to learn more
machinery. Three of its proposals were checked against the source and adopted.
All three are pre-release breaking changes, made without aliases (AGENTS.md
§ Pre-release breaking changes).

**`viewer.document`: one parse for the main view and its thumbnails.** The
viewer kept its parsed document private, so a thumbnail strip, a range read or
a second view meant `open()` again (a second Worker parse) or giving up the
viewer for the low-level route. Each preview handle now exposes the document it
opened: `PptxDocument`, `XlsxDocument` or `DocxDocument`, narrowed by `format`.
The element's and the hooks' viewers inherit it. Ownership did not change. The
viewer still owns the document: `destroy()`, removing the element, unmounting
the hook or a new source closes it, and its reads then reject `closed`.
*Rejected:* a viewer over a caller-owned document (`preview({ document })`). It
adds a second lifetime story to the one-call path for the same result.
*Rejected:* `Omit<…Document, "close">`. The document classes have private
members, so the omitted type no longer goes into `render({ document })` for a
second DOCX view. The full type is exposed, and the docs say that calling
`close()` yourself ends the viewer's session. This overturns §12 and §15 as far
as they kept the viewer from handing out its document; the low-level route
remains for an app that pages itself or keeps a document past any viewer.

**`viewer.goToSearchMatch(match)`: one call for any result.** A compact find
bar was already format-free (`searchNavigation`), but a clickable results list
branched three ways: `goTo(slide.index)`, `xlsx.goToCell(address)` and
`docx.goToSearchMatch(match)`. The method now sits on every preview handle and
takes `PreviewSearchMatch`, the union of the three, declared once on the shared
base, so it can be called on an un-narrowed `PreviewHandle`. A PPTX result
shows its slide. An XLSX result switches to its own sheet (`match.sheet`) when
another is on screen, then selects its cell. A DOCX result goes where the
render handle's `goToSearchMatch` always went: the first displayed
header/footer copy, else the nearest body heading. A PPTX or XLSX result of
the newest highlighted search also becomes the active one, through the same
step `searchNavigation.next()` uses (`#moveToMatch`).

A DOCX result becomes active too, but needs a mapping. DOCX find navigation
counts painted occurrences, while `search()` returns one result per block (a
paragraph, a cell's paragraph, a chart's text) or per header/footer part. A
part repeats on every page that uses it. `#docxOccurrence` (`embed-viewer.ts`)
picks the copy a reader expects:

- **Header or footer result:** the copy of that part on the first page still
  in view (else the first copy).
- **Body result:** the first occurrence in its own block, found by the block's
  text. The n-th result of identical text takes the n-th such block.

The highlighter exposes its occurrences with their blocks for this
(`occurrences`). Two decisions sit inside it:

- **The rule is by reading position, not by the first copy.** Clicking a
  header result while reading page 3 should outline page 3's header, not jump
  to page 1.
- **"In view" is measured by the page box, not the text.** A lazily painted
  page skips its contents, so its text reports a zero rect, and the first
  page's copy then looked in view.

A result that is not among the active highlights (a data-only search, an older
query, or a block split across pages whose painted piece no longer holds its
text) falls back to the render handle's scroll. A sheet switch clears a
workbook's search (it is the sheet on screen's), so a stale-sheet result only
moves. The result's format is read from its shape; another format's result
rejects `invalid-argument`. Like every move, a failure rejects and fires
`error` once. `viewer.docx.goToSearchMatch` was removed as a duplicate.
`DocxRenderHandle.goToSearchMatch` stays for the low-level route.

**Why patches and declarative state cannot share replacement semantics.**
`setAppearance({ theme: "dark" })` silently brought back hidden sheet tabs
when it replaced the whole object (§31/§64); documenting that surprise was
not enough. Imperative updates needed patches, while removing a React option
still needed a real reset. `undefined` is not an unambiguous reset under
`exactOptionalPropertyTypes` or plain-JS omissions. The viewer also must pass
complete resolved state to its renderer, or a dropped knob survives in the
renderer's own merge. These are the reasons for two *intentional* update modes,
not a second public patch method. The exact merge depth, defaulting and React
comparison rules live in [ARCHITECTURE](ARCHITECTURE.md#navigation-appearance-and-cancellation)
and [the appearance API](../site/src/content/docs/api.mdx#appearance).

**Not adopted from the same review.** A single error channel (navigation both
rejects and fires `error`; search only rejects) stays open: dropping the event
brings back unheard failures from a button handler that never awaits. Renaming
the element's `src`/`file` to `source` was rejected, because an HTML attribute
holds only text. The audit's dispositions and its new UX-cost column record
the rest.

**Proof.** `appearance.test.ts` pins the merge. `preview.visual.test.ts`
hides the tabs, switches theme and checks they stay hidden, then brings them
back with `sheetTabs: true`. `xlsx-render.visual.test.ts` checks the same at
the render-handle depth with headers. `react.visual.test.ts` still checks
that removing the hook option restores the defaults. `search-preview.visual.test.ts`
goes to the last result in each format through one call, with no format
branch, checks the active index, the stale-sheet case and the
another-format rejection with one `error`. `preview-document.visual.test.ts`
renders three PPTX thumbnails, reads a sheet range, renders a second sheet and
a second DOCX view from `viewer.document`, counts one Worker `open` each, and
checks that the document's reads reject `closed` after `destroy()`.

**Also fixed.** `search-preview.visual.test.ts` searched an XLSX sheet for
"budget", which the re-authored coverage workbook no longer contains, so its
"highlights survive leaving and returning to a sheet" check had no
highlights to test and failed. It now searches "نمو" on the Summary sheet and
asserts the search finds something, so a future fixture change fails with a
clear cause instead of passing vacuously.

<a id="note-120"></a>

### 120. Failures report through the action's owner, not two channels (October 13, 2026)

The outside UX review's remaining error proposal was initially left as a
documentation problem (§119). That preserves an inconsistent API: navigation
rejected and emitted `error`, search only rejected, and synchronous cell controls
could also throw and emit. Code tracing found a third problem: the XLSX render
handle's tab buttons swallowed every failure, including a failed Worker read.
Explaining that matrix was not good enough for an unpublished API.

**Chosen rule:** public calls throw/reject only. Built-in actions catch their
own failures and emit one `error` event. This overturns §57's dual-reporting
choice and the navigation contract in §64/§119. Search, unit navigation,
search-result activation and format controls all follow the same public rule;
explicit `element.open()` does too. Automatic element loads, attributes and
toolbar buttons catch at their action boundary. XLSX tab buttons report on the
render handle, forwarded unchanged through preview and the element. Shared
`background-error.ts` keeps cancellation quiet and wraps unexpected failures
as `internal`, preserving a typed event cause. Failed/stale actions after
teardown stay quiet. The element restores a failed index attribute before
announcing the failure, so an error listener can safely start a new action.

**React:** remove `navigationError`, not replace it with another operation-specific
field. `error` represents the hook-owned mount or a built-in action. A failed
mount has `status: "error"`; a failed built-in action leaves it ready with the
view intact. Unit changes and mounts clear the error. App-owned calls reject
only to the app; app buttons handle the rejection themselves. The bench, site
toolbar and checked examples were updated, not left with empty catches.

**Rejected:** a configurable error policy (two stories to learn), events for
all calls (consistent duplication is still duplication), and trying to detect
whether a caller awaited a promise (timing-dependent magic). A dropped public
promise is an application error; `void` is not a rejection handler.

**Blast radius and deliberate non-changes:** all preview routes, element, hooks,
XLSX render tabs and consumer integrations. PPTX/DOCX renderers have no built-in
navigation buttons, so their public reject/throw behavior stays unchanged.
Worker crash/request errors, codes, cancellation, newest-wins behavior, view-only
rendering, document ownership and warning fallbacks stay unchanged. No separate
global crash notification or automatic retry is added.

**Proof:** existing preview/element/search/React browser suites pin the new
contract. `error-ownership.visual.test.ts` uses the real Arabic coverage fixtures
and re-reads DOM output: calls reject without events, tab failures report once
at both low-level and preview depth, element toolbar failures report once,
explicit element open only rejects, no unhandled rejection leaks, and the
previous content remains. React's real tab failure updates `error` without
replacing the view. USE-CASES rates the residual app-callback cost Medium,
not friction-free.

<a id="note-121"></a>

### 121. Explicit Worker inputs own special-deployment resource cleanup (October 13, 2026)

The default bundled Worker remains zero-configuration. The former
`string | URL | (() => Worker)` override (§15/§56) made CDN and single-file
consumers create Blob URLs themselves, without the examples releasing them.
A wrapper factory also bypassed the Worker pool. Instead of adding a separate
`crossOriginWorker()` helper, the existing `worker` option now takes one explicit
object-form input: `{ url }`, `{ url, crossOrigin: "blob" }`, `{ source }`, or
`{ create }`. There are no compatibility aliases. Supplying no override still
uses the exact bundler-detectable `new Worker(new URL(...), { type: "module" })`
constructor; changing that expression would repeat the production build failure
in §15.

Validation lives in `core/worker-init.ts`, before construction: exactly one
input, nonempty URL/source, a function factory, and the cross-origin opt-in
only on URL input. Invalid configuration is `invalid-argument`; a browser
refusal or startup failure is `worker-error`. Relative URLs resolve at open
time in the page, so importing remains headless/SSR-safe.

**Ownership lesson:** a recipe that creates a Blob URL has also created a
resource lifetime; giving the caller only its allocation leaks the cleanup
problem into every integration. Startup failure/cancellation must release it
before idle pooling, and an unready Worker cannot be reused safely. Pooling
native URL/source modes by value avoids paying for every inline options object;
factories remain unpooled because their construction semantics are unknown.
The complete current mode, URL-revocation and pooling rules belong in
[ARCHITECTURE](ARCHITECTURE.md#worker-execution-and-deployment).

**Rejected:** automatic cross-origin fallback or relaxed browser permissions.
The explicit CDN mode does not bypass CORS/CSP. A blocked bootstrap/source
rejects; it never retries through an alternative transport. Apps that forbid
Blob Workers host the Worker same-origin. Native factories remain an escape
hatch, not a custom Worker framework.

**CSP ground truth:** the initial two-origin production probe failed before
any CDN request, despite `worker-src 'self' blob:` and the CDN in `script-src`.
Chrome's `Log.entryAdded` named `worker-src` as the blocked directive for the
imported CDN module. Adding the CDN origin to `worker-src` made both bundlers
work; removing it independently reproduced the refusal. The page's own
`securitypolicyviolation` listener did not receive that Worker-scope violation,
so the production probe also captures the browser error log. The API recipe
now names both required Worker sources, not just the Blob permission. There
is no permission relaxation in the library itself.

**Blast radius:** shared session construction covers all three low-level opens,
generic/format-local previews, detect, hooks and the element. Parser/RPC/model
contracts and the self-contained Worker script are unchanged. No public
helper, font change or additional package dependency is introduced.

**Proof:** `worker-pool.test.ts` checks value-based reuse, generated source,
ready-time revocation, construction failure and pre-ready cancellation, plus
invalid configuration without construction. `worker.test.ts` opens and reads
all three actual formats with inline source on one pooled Worker.
`tools/check-consumer-build.ts` extends the existing production probe with
two-origin CDN and single-file cases for Vite/webpack: real reads, chart DOM
output, highlighted search and cleanup. Single-file pages cannot fetch an
external script. Negative controls block Blob CSP, omit CORS or attempt direct
cross-origin construction; all must reject `worker-error`, without an unhandled
promise or hang. Next.js's earlier manual CDN evidence (§56) is retained; it
is not represented as a fresh Next.js verification.

<a id="note-122"></a>

### 122. Viewers lend content capabilities, not their document lifecycle (October 13, 2026)

The maintainer pointed out that §119 contradicted their explicit §69 decision
to expose search only rather than a whole document. That conflict should have
been raised before implementation. Full `viewer.document` exposure let callers
close a viewer's session, and it brought internal methods/ordinary TypeScript
private properties along with the useful reads. An `Omit<...Document, "close">`
would hide a name from types, not remove its runtime access paths.

**Correction:** remove `viewer.document` with no alias. Expose `viewer.content`,
a real restricted read/render facade; keep search on `viewer.search()`. The
maintainer approved this boundary and requested exploration first. Exploration
confirmed that reads plus metadata, not rendering only, serve the existing
thumbnail, cell-reader and picture-gallery journeys. Read-only resolved data
do not close sessions and are not ownership escapes. §69's lifecycle concern
is preserved while its former second-parse cost for combined views is removed.

**Why a runtime facade.** Hiding names in TypeScript would leave the same
object reachable through reflection. Closure-backed facades, private viewer
fields and a renderer-private association instead preserve useful rendering
context without handing out the lifecycle owner. A borrowed view must retain
automatic media, layout media, sheet tabs and heading navigation; a text-only
facade would pass a shallow test but fail those normal-use journeys.

**Why two kinds of lifetime guard.** A closed-use check prevents new operations;
it cannot prevent a render already waiting for bytes/fonts from painting late.
That needs the owner's cancellation signal as well as the caller's. Conversely,
tracking/destroying every completed secondary view would introduce another
ownership registry: callers already own those handles. Data already returned
cannot be un-returned; this boundary restricts lifecycle capabilities, not
ordinary resolved data or arbitrary application behavior. The full current
[borrowed-content contract](ARCHITECTURE.md#borrowed-content) and
[input shapes](../site/src/content/docs/api.mdx#borrowed-content) live once in
their specification homes.

**Alternatives rejected:** rendering-only content would remove useful reads;
`content.renderSlide()`/`renderSheet()` would introduce duplicate render grammar;
returning ordinary handles would expose XLSX's `sheetDocument()` and TS-private
owner/session fields; type-only omission would not restrict the runtime object.
`preview({ document })` was discussed but is not introduced: it adds a second
ownership story and is unnecessary for the approved viewer-plus-thumbnails job.
Low-level `open()` + `render()` remains the path for a caller-owned document
that lives past any viewer.

**Blast radius and deliberate non-changes:** the shared preview viewer, its
generic/fixed-format routes, hooks and element; all three format render input
unions and their source resolution. No parser, Worker/RPC, model data, visual
paint rules, low-level document lifecycle, search, error ownership (§120) or
Worker configuration (§121) changes. Tests that formerly closed an exposed
document now inject an actual Worker failure at the Worker boundary instead.
The optional `messageFor()` helper remains optional, not an error-policy switch.

**Proof:** `preview-content.visual.test.ts` checks real coverage fixtures through
both generic and format-local preview: one Worker open, reads/metadata, media
bytes and loaded images, rendered thumbnails/ranges/secondary DOCX pages,
secondary navigation, runtime surface/reflection, stable borrowed inputs,
closed operations after teardown, and cancellation without late paint.
Type checks reject document/close/search/session/owner access. Element source
changes and React unmount use the same boundary. The existing Vite/webpack
production tool also renders borrowed content at both preview depths, including
single-format/multi-page builds and special Worker deployments. Durable contracts
live in ARCHITECTURE and the API/Start pages; the journey audit reflects the
remaining Low-cost teardown responsibility rather than a completed proposal.


## Worker, packaging and memory

**Current pipeline:** [ARCHITECTURE](ARCHITECTURE.md#performance-model).
The dated costs and failed packaging experiments here explain that pipeline;
old bundle sizes and request shapes are not present-day deployment instructions.

<a id="note-16"></a>

### 16. Search cost, DOCX search coverage, `info.sizes`, and the Worker bundle

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

<a id="note-17"></a>

### 17. PPTX master/layout surface: flatten the small, share the large

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

This originally added a layout request to the protocol; empty layouts avoided
the transfer. The present [scene composition](ARCHITECTURE.md#pptx-scene-and-text-bodies)
and [retention model](ARCHITECTURE.md#opening-and-resolved-data) own the current
representation. The flatten/share distinction above is the decision's cost
argument, not a second inventory of document collections.

A theme background reference (`bgRef`) paints the theme's `bgFillStyleLst`
entry (tint and `satMod` applied), not the bare scheme colour; PowerPoint
shows `#3A4961` where the bare colour is `#0E2841`.

Layout → master placeholder matching is by placeholder *type* (title, body,
dt, ftr, sldNum), as PowerPoint does, because masters number them differently
from layouts and slides; slide → layout stays by index.

Not verified against real PowerPoint output yet: the `showMasterSp` rule,
picture backgrounds, and `p:clrMapOvr` (layout/slide colour-map overrides, not
read). They are listed in the ROADMAP fidelity gaps.

<a id="note-18"></a>

### 18. Core performance pass: streaming parse, cached relationships, split worker

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

<a id="note-19"></a>

### 19. Split worker vs. consumer bundlers: Vite verbatim copy and the abort taxonomy

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

<a id="note-20"></a>

### 20. Worker back to one self-contained file: the split is reverted

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

<a id="note-26"></a>

### 26. Deferred geometry: the worker describes, the main thread evaluates

The full ECMA-376 preset catalog (~187 shapes) and `custGeom` freeforms
cannot live in the worker: it is built flat (`codeSplitting: false`, so it
works as a single file or Blob URL), and any `import()` the parser makes is
inlined into `worker.mjs`. Hand-coding every preset was rejected (months of
formula transcription, still partial), and so was shipping the catalog in the
worker (every deck would carry it).

Original decision (before later fast-path additions): the worker keeps a
hand-coded fast path for the 11 common presets
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

<a id="note-34"></a>

### 34. Shipped JS is Latin-1

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

<a id="note-35"></a>

### 35. The DOCX Worker keeps nothing the page holds

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

<a id="note-38"></a>

### 38. The Worker is reused between opens

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

<a id="note-39"></a>

### 39. Bounded Worker caches; minified dist

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

<a id="note-56"></a>

### 56. The Worker URL resolves under webpack and Next.js (Sep 2026)

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

<a id="note-72"></a>

### 72. Hostile counts are bounded where the file is read; three XLSX and PPTX rules corrected (Oct 2026)

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

<a id="note-80"></a>

### 80. Search reads words, not models (Oct 2026)

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

<a id="note-112"></a>

### 112. One-pass preview opening and real lazy viewer/effect boundaries (October 11, 2026)

**The problem.** Generic preview normalized the source, sent `detect`, imported
the selected route, then normalized/copied again and sent `open`. Detection and
indexing therefore traversed the archive twice. Standalone `detect` also normalized
before a helper that normalized again. Format entries eagerly re-exported their
viewer, so model-only consumers carried the embed controller; XLSX imported the
effect painter just to paint rotation/flips. The alternate implementation in
`packages/aared-optimized` identified these four opportunities; its source graph
counts alone were not evidence of production download savings.

**Why port the opportunities, not the alternate package.** An internal
owned-buffer path can remove duplicate work without weakening the public
caller-buffer/view copy guarantee. Session adoption needs the same format
initialization as ordinary opens, or deferred PPTX geometry and eager DOCX
fonts quietly diverge. Cancellation must own a partially opened session even
before a viewer exists; browser chunk downloads cannot be cancelled, so their
late failures still need observation. The useful split was the cheap transform
helper from optional visual effects, not another package identity. The current
[opening/adoption sequence](ARCHITECTURE.md#opening-and-resolved-data) and
[lazy boundaries](ARCHITECTURE.md#build-and-lazy-boundaries) own the full mechanism.

**Production findings, not source-only guesses.** Removing the eager XLSX effect
import exposed its first-layout race: `settleEffects` drained the targets while
the painter was still downloading, marking them unavailable permanently. Targets
now stay pending until success or actual failure; failed downloads still mark,
and existing disconnected-box requeue/teardown behavior stays intact. Native ESM
tests hold the chunk request open, inspect the picture's DOM before release,
then re-read its computed rotation and shadow after settlement; a 503 case proves
the marked fallback. The authored Notes picture and its native geometry/effect
values are unchanged (§108).

The first rebuilt multi-page bench also failed with `#root element missing`:
Vite folded the controller into the React app's eagerly importing bootstrap,
and the lazy format adapter imported that other page. All embed routes now defer
the controller consistently; the element reads event forwarding metadata from
`format-route.ts` without importing it. No manual chunk policy or consumer-only
workaround. The consumer probe includes three multi-page Vite builds whose other
page bootstrap throws if executed in the preview page, alongside its existing
single-format/generic Vite and webpack gates.

**Built size.** Static-import closure of each emitted tsdown entry, counting each
file once and summing gzip level 6 per file; excludes dynamic imports, source maps,
and the Worker. Baseline is the pre-port build with §111 already present:

| Entry | Raw bytes before → after | Gzip bytes before → after |
|---|---:|---:|
| PPTX | 174,631 → 166,095 | 51,771 → 49,627 |
| XLSX | 90,638 → 76,379 | 32,875 → 28,366 |
| DOCX | 110,904 → 102,167 | 38,864 → 36,424 |
| Generic preview | 28,892 → 13,896 | 10,674 → 5,333 |

These are eager module-graph costs, not whole-viewer download savings or opening
speed claims: using preview still fetches its controller/adapter, and XLSX pictures
with effects still fetch their painter. Worker output is unchanged at 247,866 raw
bytes / 79,436 gzip bytes. No parser splitting, specialized Worker, or dependency
change; §20's packaging tradeoffs stand.

The existing production bench measured all three coverage fixtures successfully
(`bench-measure.ts --run <fixture>:aared`): cold medians 128/148/800 ms for
PPTX/XLSX/DOCX (five samples), warm 32/28/166 ms (eleven samples after warmup).
These are current-build smoke measurements on this Windows machine, not a
controlled before/after comparison; other verification ran concurrently. Its
no-cache download pass (`--size`) measured 134.3 KiB gzip for either PPTX fixture,
113.1 for XLSX coverage / 143.1 for the chart-heavy tracker, and 157.9 for DOCX,
including each opened viewer's fetched code and Worker, not fixture bytes.
The size pass's version reader needed a small correction: aared has no root
export, so it now resolves the adapter's `aared/<format>` before walking to the
manifest, with a test for all three routes. No generated benchmark tables were
rewritten to imply a speed improvement.

**Blast radius and proof.** Preview, React and the element share the new one-open
path for all three formats; fixed-format previews share owned-buffer transfer.
Ordinary opens retain their format checks/initialization, search and close behavior.
`preview-open.test.ts` records requests and transfer lists at the real Worker
boundary, checks caller-buffer preservation and owned-Blob detachment, failed and
cancelled imports/adoptions, and compares each adopted document to ordinary open,
including deferred PPTX geometry and eager DOCX fonts. Preview DOM tests count one
open/no detect for every format. Shipped static-graph/native-ESM tests pin lazy
loading and rendered output; existing embedding, capture and consumer tests cover
the shared lifecycle and paint routes. Public signatures and defaults do not
change. The Notes picture's paired native Excel/aared capture was inspected
(`bun tools/xlsx-capture.ts fixtures/xlsx-coverage.xlsx --part 4:svg-picture=A1:H12 --pair-full`):
rotation and shadow match at normal use, with the already accepted print-grid/header
differences unchanged. Typechecks, package/bench builds, unit tests, the full visual
suite (30-second test deadline) and consumer builds passed. With the maintainer's
approval, the comparison directory/report were removed after porting, resolving
their duplicate workspace identity. Task-only scratch/review files were cleaned;
shared benchmark history containing other renderers' measurements was preserved.

The maintainer subsequently requested publishing those saved numbers to the site
and generated benchmark tables, without a new run. Only three of eight aared
timing cells are fresh; all five download-size rows are fresh, while search stays
older. The caption generator now counts cells when a renderer spans measurement
dates instead of hiding its older cells inside "the other renderers", pinned by
the benchmark contract test. A full rerun is deferred until the next editing pass.

<a id="note-113"></a>

### 113. Corrections to earlier status statements (Oct 11, 2026)

Earlier entries are history and are not rewritten. This entry records the
present-tense statements in them that no longer hold, with the current fact and
where to check it. Documentation audit only; no code changed.

- **§112 (lines 4851-4852):** "all five download-size rows are fresh" is four.
  The size table has four aared rows (pptx, xlsx, docx, and "all three"; lines
  201, 209, 215, 221 of `docs/BENCHMARKS.md`). The timing count, three of eight,
  is right. The same miscount sits in `docs/BENCHMARKS.md` lines 14-15.
- **§12 (lines 205-207) and §14 (lines 256-257):** "search ... stay on the
  headless entrypoints" and "the same boundary as the facade" are superseded by
  §69: `viewer.search()` is on `preview()`, the React hooks and `<aared-viewer>`
  (`packages/aared/src/preview/index.ts` lines 170, 186, 202).
- **§16 (lines 340-342):** the per-format preview and react subpaths are not
  "recorded in the ROADMAP under Before first publish". No per-format react
  subpath exists (`packages/aared/package.json` exports); the react item sits in
  ROADMAP "Not yet scheduled". Per-format preview was built as each format
  entry's own `preview()` (§88: `pptx/index.ts` line 1646, `xlsx/index.ts` line
  2651, `docx/index.ts` line 2183), not as a subpath.
- **§17 (lines 390-392):** "listed in the ROADMAP fidelity gaps" is false, and
  the ROADMAP has no fidelity table. `showMasterSp` was verified in §54 (lines
  2047-2049). `p:clrMapOvr` is read since §54 (lines 2014-2018) and its limit is
  in ARCHITECTURE Known limits, "By design".
- **§47 (lines 1700-1701), §48 (line 1758) and §54 (line 2057):** "ROADMAP's
  fidelity-gap table" no longer exists. Its accepted rows are in ARCHITECTURE
  Known limits (PPTX rows, lines 1277-1284).
- **§71 (lines 2851-2854):** the open chart work is not "in ROADMAP". §73 and §74
  closed the listed gaps; ROADMAP keeps only the chart-oracle and
  generated-table rows.
- **§72 (lines 2900-2901):** "XML nesting, repeated entities ... remain open
  bounds (ROADMAP)" is out of date. Nesting is capped at 512 levels
  (`packages/aared/src/core/xml.ts` lines 59, 285), and DOCTYPE and other
  declarations are refused (line 333), so entity expansion cannot happen. The
  open residue in ROADMAP is total inflated bytes, the high-ratio spread, the raw
  cost of character-reference-heavy parts, the URL download ceiling and
  count-driven loops.
- **§77 (lines 3170-3172) and §92 (lines 3708-3721):** "painting the computed
  marker remains open" and "refreshing TOC entries ... remains open in ROADMAP"
  were closed by §94 (lines 3767-3782). ROADMAP carries neither.
- **§79 (lines 3316-3318):** "(it stays for PPTX)" points at a ROADMAP item that
  is not there. §80 measured the PPTX first search at 483 to 183 ms in Bun
  (lines 3359-3361).
- **§82 (lines 3475-3481):** "spAutoFit frame growth" is not implemented: §95
  resolved it (lines 3812-3814; ARCHITECTURE Known limits, DOCX row, line 1303).
  The "remaining search scope" in ROADMAP does not exist; header and footer
  search landed in §83.
- **§89 (lines 3608-3609):** "Published measurement tables are not regenerated"
  is superseded: the tables were regenerated from saved numbers (§112, lines
  4850-4853).
- **§99 (lines 4029-4030):** the band-size limit is not in the Known limits
  table. It is a bullet in ARCHITECTURE § Format rules (line 1092). The move is
  open for the maintainer.
- **§105 (lines 4327-4329):** "their decisions remain open in ROADMAP" is false.
  `timePeriod` and `iconSet` paint since §107, `expression` rules since §110,
  and the ROADMAP's conditional-formats item was deleted in §111.
- **§107 (lines 4465-4480):** "`expression` rules and complex formula bounds
  still report" and the recommendation to defer the evaluator are superseded:
  §110 built the bounded evaluator.
- **§110 (lines 4687-4688):** "Cross-sheet values, x14 inline `dxf` ... volatile
  functions report" no longer holds. §111 evaluates cross-sheet and x14
  expression rules and `TODAY()`/`NOW()`. Names, table refs, arrays, lookups,
  `TEXT` and unlisted functions still report (§111; ARCHITECTURE Known limits,
  conditional-formula row, line 1295).
- **§87 (lines 3571-3572):** "the Worker remains 177,846 bytes" is the
  pre-formula size. The bounded conditional-formula evaluator grew it to
  247.8 kB raw / 79.4 kB gzip (§111, lines 4737-4739), confirmed at 247,866
  raw / 79,436 gzip bytes (§112, lines 4814-4816). History stands; use the
  §112 numbers for anything current. The lazy-highlighter 5,767 bytes
  (2,596 gzip) in the same sentence was not re-measured here.
- **§109 (lines 4613-4617) and §110 (lines 4669-4671, 4673):** the
  `three-colour-scale` paired part was `A1:E12` in §109 and `A1:F12` in §110;
  it is now `A1:J13` after §111's `G2:J13` columns (§111, lines 4729-4735;
  `fixtures/README.md` lines 157-160). The §110 "Worker growth to be measured
  on build" is the §111 measurement above. History stands.

<a id="note-114"></a>

### 114. Worker parsers move from dynamic to static imports (October 11, 2026)

**The problem.** `core/worker.ts` reached its three format parsers through a
dynamic-`import()` map, but the Worker builds flat (`codeSplitting: false` in
`tsdown.config.ts`), so those imports resolved in place inside the one file:
zero runtime laziness, with namespace wrappers and weaker whole-program
tree-shaking. The laziness the docs promise lives on the main thread
(`auto-viewer.ts`, the format `viewer.ts` adapters), which this does not touch.

**What changed.** Static, symbol-level imports (`indexPptx`,
`parsePptxSlide`, `parsePptxLayout`, `readPptxSlideText`; `indexXlsx`,
`readWorkbook`; `parseXlsxSheet`, `sheetText`, `sliceSheetToRange`;
`indexDocx`, `parseDocxSections`) at all eleven call sites, plus the
`tsdown.config.ts` comment. The Worker stays one self-contained file; §20's
packaging tradeoffs stand. No parser split, no property mangling (model
shapes cross `postMessage`, so mangling could break the protocol while
looking like a saving).

**Measured worker delta** (`dist/core/worker.mjs`, same build flags):

| Measure | Before | After |
|---|---:|---:|
| Raw bytes | 247,866 | 245,414 (−2,452) |
| Gzip level 6 | 79,436 | 78,053 (−1,383) |
| Gzip level 9 | 79,211 | 77,809 (−1,402) |
| Brotli level 11 | 68,694 | 67,621 (−1,073) |

**Proof.** Typecheck, rebuild, full unit suite (867 pass, `shipped-latin1`
included), and `check-consumer-build` (Vite/webpack, single-format and
multi-page) with the worker self-containment assertion. No public API change;
the §112 numbers above are superseded by this table for anything current.

**DOCX open breakdown (same day, scratch harness, deleted after).** Medians
over fresh pages in layout mode with no resolver, 11 pages: module
import/eval 16 ms, Worker `open` 7 ms, unit 0 (full section parse) 34 ms,
font wait 17 ms, render with layout and settlement 211 ms, frame commit
25 ms — about 307 ms low-level, 374 ms through `preview()` (controller
overhead ~65 ms). Layout/settle is two thirds of aared's own path; worker
parse plus fonts are ~60 ms combined, so worker bytes cannot move this
number. The bench's 800 ms cold includes harness and environment overhead
outside aared's stages. No code change follows: layout cost stands unless a
real consumer budget demands it.


## Shared rendering, paint and host isolation

**Current shared homes:** [the module map](ARCHITECTURE.md#shared-modules),
[paint rules](ARCHITECTURE.md#geometry-fills-pictures-and-effects) and
[CSS/resource ownership](ARCHITECTURE.md#css-and-resource-ownership).
Keep the native discoveries and failed approaches below, not a second module catalog.

<a id="note-23"></a>

### 23. Shared-code pass: one home per rule, and the decisions it forced

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

<a id="note-27"></a>

### 27. Shape style references with `idx="0"` mean "no reference"

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

<a id="note-28"></a>

### 28. DrawingML fills: one model, one painter, and what PowerPoint actually draws

**Why.** A fill the renderer could not paint used to vanish without a
trace: the model tagged it `unsupported`, the renderer set a data attribute
nothing read, and the shape, cell or slide painted empty with no dev marker.
Pattern fills, picture fills on shapes, `path="rect"` gradients, tiled
pictures and every non-solid table-cell fill hit this, and no fixture
exercised them. `pptx-coverage.pptx` gained slides 7-10 (authored through
PowerPoint's COM object model) to carry every fill kind.

**Why one fill vocabulary.** Colour strings on table cells and a separate
`ImageBackgroundModel` could not describe the same fills shapes already needed.
Moving cells/backgrounds onto the same union removed that structural mismatch;
a picture background became an ordinary picture fill rather than a new paint
route. One SVG painter made those surfaces agree, including explicit unsupported
results instead of a silent empty shape. The current model kinds, precedence,
diagnostics and paint sizing are in
[ARCHITECTURE](ARCHITECTURE.md#geometry-fills-pictures-and-effects) and the
[resolved-model reference](../site/src/content/docs/api.mdx#resolved-models).

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

<a id="note-36"></a>

### 36. Renderer defaults are classes in one adopted sheet; a lighter grid

Every XLSX data cell wrote ~20 inline longhands and a positioned, z-indexed
text span even when empty, and the viewport fill was one styled cell per fill
column per row: 616-832 elements at about 3 KB each on the small workbooks,
most of their memory. An inline declaration block is a style attribute plus a
CSSStyleDeclaration per element, and positioned or z-indexed elements each
get a paint layer.

**Why classes in a constructed sheet.** Interning a constant block lets cells
share declarations instead of allocating the same longhands on every node.
Constructed sheets meet strict CSP without `<style>` text; explicit specificity
preserves the strength formerly supplied by inline defaults while authored
inline values still win. A shared root needs reference-counted lifetime, not
one render removing a sheet another still uses. Detached attachment and hosts
replacing adopted sheets explain why adoption must be recoverable. The current
[sheet contract](ARCHITECTURE.md#css-and-resource-ownership) owns those mechanisms.

Rejected: `<style>` text (fails under CSP), `insertRule` per page (slow),
custom properties alone (still one style attribute per element), and
attribute-keyed rules (weaker specificity, a second mechanism). This was not
permission to put authored geometry into classes or expose internal class names
as styling API.

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

<a id="note-101"></a>

### 101. Group fill is inherited paint; group effects are composite paint (Oct 2026)

The PPTX group parser kept identity and coordinate transforms but discarded
group appearance. `grpFill` was recognised as unsupported, not resolved;
the earlier roadmap wording claiming there was no parser arm was already stale.
The correction passes resolved group paint through recursive parsing, and emits
2D effects on the element that actually authors them.

**Native evidence and fixture provenance.** Python drove PowerPoint's grouping,
solid fills, gradient fills and effects on a new bilingual coverage slide.
Setting a group fill writes that fill on `grpSpPr` and `grpFill` on child shapes
and nested groups. A child's direct fill overrides it; a native `noFill` child
stays transparent. COM group shadow/glow/reflection operations instead write
their effects onto each child; a soft-edge setter on a group is rejected.
To exercise actual group-level storage, the probe moved PowerPoint-authored
effect payloads and a native gradient into `grpSpPr`, replaced the participating
child fills with `grpFill`, then reopened, saved and exported with PowerPoint.
The final fixture therefore includes native-authored values in both layouts,
with a native render of the group-level interpretation, not a screenshot oracle
inferred from the XML. Its added slide has no slide-number placeholder.

**What the native pair required.** Missing parent context is not an explicitly
unfilled group, and inherited gradient colour must continue across children
rather than restart in each child box. A rectangular backing on the group
would invent paint; overriding a child's direct no-fill would discard authored
intent. Carrying the ancestor paint box solved those distinctions without a
second fill painter. The resolved propagation/coordinate rules live in
[ARCHITECTURE](ARCHITECTURE.md#pptx-scene-and-text-bodies).

**Text is not shape inheritance.** A native probe put `grpFill` on a red
child's text inside a blue group. PowerPoint's saved XML keeps it, but its PNG
paints no glyphs. PDF extraction misleadingly reports black text: the paint trace
has opacity zero. The same happens on a dark slide. The parser therefore emits
`noFill` for the run and keeps its source flag, rather than painting blue or
trusting the PDF's black RGB metadata. The real fixture and DOM test pin this.

**Effects.** `DrawingEffectsModel` and `readDrawingEffects` preserve outer and
inner shadows, glow, soft edges and conventional reflections. Group effects
operate over its combined DOM/SVG silhouette; child-authored effects stay on
their child, matching COM's distributed form. An ordinary outer shadow uses the
browser's native CSS `drop-shadow`, not a custom SVG graph. Browser filters keep text
as text. Native comparison caught the reversed direction of the first inner
shadow and a reflection mask fading from the wrong end; both were corrected.
Reflections use native CSS where available and a standard masked, inert,
aria-hidden DOM mirror otherwise, with unique definition IDs and teardown-owned
media/text observers. The fallback and shadow-root route are exercised in the
rendered-output test.

**Blast radius and stopping point.** Slide and layout/master groups share the
same recursive parser and renderer. Standalone PPTX shapes and pictures use
the same effect painter; DOCX/XLSX drawings do not acquire this PPTX capability.
Their shared fill and colour paths are regression-tested. Search still reads
the original words once, including nested groups; reflections add no hits.
Blur/glow/soft-edge softness and tiny reflection blur differences are accepted cosmetic
residue. Effect DAGs and perspective/skewed effects
remain explicit limitations, not a claim that every DrawingML effect is drawn.
The normal group-appearance gap leaves the roadmap; ARCHITECTURE owns these
limits, and the real fixtures plus parser/DOM/native comparison own the evidence.

**Performance informed the stopping point.** The first production-bundle
comparison (`bun tools/bench-group-effects.ts`) showed ordinary deck zoom frames
around 16.6ms, the decorated fixture around 21–22ms, and eight overlapping copies
of its real authored groups around 159ms. Isolating effects identified morphology:
eight glows took 74ms, eight soft edges 58ms; outer/inner shadows and reflections
stayed around 17ms. Native outer shadows now avoid a custom graph, and glow/soft
edges use Gaussian blur plus alpha adjustment rather than dilation/erosion.
The paired native appearance stays usable; exact softness is deliberately not
chased.

Final two-repeat comparison (`--optimized`), Windows, HeadlessChrome 155,
960x540 viewport; these are warm display/zoom measurements, not download/parse
benchmarks or GPU-memory measurements:

| Case | Effects off: median zoom frame | Effects on | Forced DOM reflection fallback |
|---|---|---|---|
| Real 7-slide production deck | 16.6–16.7ms | 16.8–16.9ms | 16.7–16.8ms |
| Single authored effects slide | 16.5–16.6ms | 16.6–16.8ms | 16.6–16.7ms |
| 8x overlapping group stress | 16.5ms | 55–61ms | 43.9–47.2ms |

The stress case improves roughly 63% versus the initial implementation but is
still expensive; it is a stress ceiling, not evidence that consumers author
such stacks. Every teardown returns the DOM to its six-node harness baseline,
with zero filters/mirrors. The fallback adds 18 nodes per reflected pair in this
fixture, and was not the primary bottleneck, so it stays. Results are written to
the gitignored PowerPoint review output; `--isolate` reproduces single-effect
comparisons. We retain the working normal-use capability, record dense-stack
limits in ARCHITECTURE, and stop here under the existing good-enough guidance;
no new blanket rule or rendering option was added.

**What to remember from this work.** After the targeted changes, effects had
little measured impact in the ordinary production deck and the bounded effects
fixture: their zoom medians remained around 17ms, comparable to effects off.
That is the reason to retain the implemented support, not the effort already
spent and not an assumption that browser effects are always cheap. The expensive
primitive mattered more than the feature label: morphology made glow/soft edges
costly; replacing it preserved a usable appearance and recovered normal-use
responsiveness. Dense overlapping stacks remain costly. These results describe
this PPTX implementation on one measured environment, not all Office files or
devices.

**Word applicability was separate from delivery, and is now closed** - see §102.

<a id="note-102"></a>

### 102. Word drawings use the shared effect painter, and a picture has no text to justify (Oct 2026)

§101 deliberately stopped at PPTX and said so. Word authors the same effect
families on its own drawing objects (Microsoft's
[Word Shape reference](https://learn.microsoft.com/en-us/office/vba/api/word.shape)
lists Shadow, Glow, Reflection and SoftEdge), so leaving DOCX out meant two
readers and two painters for one DrawingML vocabulary. Native Word COM authoring
(the gitignored probe `docx-drawing-effects`, deleted once its findings became
the fixture, the tests and this entry) confirmed what Word writes: the same
`a:effectLst` children the PPTX reader already understands, on a text box's
`wps:spPr` and on a picture's `pic:spPr`, beside Word's own `wp:effectExtent`
in EMU. Nothing about the payload was new; only the route was missing.

**Why paint settlement must follow layout.** The shared effect vocabulary did
not justify a new DOCX painter. It did require ownership by the ink surface,
not a measuring frame: applying filters during off-screen measurement would
feed their paint bounds into pagination. Word's `effectExtent` reserves ink
without moving an inline picture off its baseline, a distinction ordinary
padding can accidentally erase. The current effect route, settlement order and
lazy loading live in [ARCHITECTURE](ARCHITECTURE.md#geometry-fills-pictures-and-effects)
and its [DOCX drawing rules](ARCHITECTURE.md#docx-flow-pages-and-drawings).

**Glow was calibrated against Word, not guessed.** Pixel sampling of the native
page showed Word holds a glow near full strength for several points outside the
shape and fades to nothing by the authored radius; aared's first blur-and-gain
mask reached half that strength in a much narrower band. The shared mask now
blurs at 0.6 of the radius, gains, and bends the falloff with a transfer table
into that plateau, which is what §101's measurement-driven rule asks for. The
change is in the shared painter, so PPTX glows match better too.

**A picture sits on the paragraph's start side.** The native comparison caught
something unrelated to effects: the Word-authored picture-only paragraph is
`w:bidi` with `w:jc="both"`, Word paints its picture at the right edge, and
aared painted it left, because a block picture inherited the paragraph's
alignment verbatim and the renderer only acted on `center` and `right`. There
is no text in a picture to justify, so Word leaves it on the start side. The
parser now resolves an unauthored or justified block picture (and chart) to the
paragraph's start side, using the direction it already resolved. Anchored
pictures are untouched: their anchor, not the paragraph, places them.

**Blast radius and stopping point.** The shared fill, colour and search paths
are regression-tested; PPTX shape, picture and group effects still paint through
the same painter. A plain DOCX with no effect never fetches the painter. The
rendered-output test proves the reflection fallback is inert, aria-hidden,
sized from layout (never a transformed rect), torn down with the render, and
that removing every effect from the model leaves page count and page heights
identical. XLSX remains unchanged: its charts have no DrawingML effect list.
Accepted residue, recorded in ARCHITECTURE: a picture-only paragraph followed by
a text line can sit about 15px from Word's line box (ordinary DOCX line rhythm,
not the effect painter), and the unsupported effect DAG and 3D families stay
unpainted and marked as before.

**Performance follows §101's method.** The PPTX production-bundle benchmark was
re-run after the glow change (`bun tools/bench-group-effects.ts --optimized`,
same machine): the real seven-slide deck and the bounded authored-effects slide
still measure about 16.6-16.9ms median zoom frames with effects on, off and in
the forced fallback - no measurable cost at normal use. The eight-times
overlapping group stress moved from 55-61ms to 62-65ms for the wider glow, and
effects off remains 16.4-16.7ms; that stress case is a ceiling, not a consumer
claim, and no further tuning was done.

**What to remember.** Word needed no new effect vocabulary, only the shared
route, and the native pair caught two things no unit test would have: the glow
profile and the justified-picture side. Both were fixed in the shared or shared
reading code rather than worked around in the renderer.

<a id="note-108"></a>

### 108. SVG-first pictures, shared render settlement, lazy-painter boundary, XLSX picture effects (October 10, 2026)

A reader-supplied article pointed at `asvg:svgBlip`: Office stores an
inserted SVG beside the PNG it rasterized for compatibility, and the main
`a:blip` names the PNG. Verified against MS-ODRAWXML ("Pictures":
`svgBlip` plus a rasterized PNG the main blip names) and by inserting SVGs
through real Office 16.0 COM (probe `svg-pictures`): PowerPoint writes a
lone `svgBlip` for a COM-inserted picture (the same shape as its picture
bullets), Excel the full pair (`<a:blip r:embed="png"><a:extLst>...
<asvg:svgBlip r:embed="svg"/>`). The XLSX conditional-format icon catalogue
(§107) is untouched by this: those rules name built-in families and carry no
media, and Fluent icons stay a viewer-chrome option only, never a document
substitute.

**Why resolution precedes paint.** A first-embed subtree search selected
Office's compatibility raster, not necessarily the image the author inserted;
picture bullets had that same defect. One image-source reader avoids fixing
only ordinary pictures. The change existence-checks the SVG candidate alone,
so a dangling SVG preserves the old raster/missing-part behavior instead of
inventing a new validation policy. The current source/fallback route is in
[ARCHITECTURE](ARCHITECTURE.md#geometry-fills-pictures-and-effects); optional
fallback fields belong to the [model reference](../site/src/content/docs/api.mdx#resolved-models).

**DOCX joins the shared settlement.** Its staged post-paint waits ran past
the handle with no `signal`, and its comments still claimed a synchronous
paint. Both modes now go through `settleRender` like PPTX/XLSX (no model or
signature change), covered by extending the replaced-render visual test to
abort-during-settlement in flow and layout modes.

**The lazy-painter boundary is real again.** `svgElement()` moved to its
own `svg-element.ts` so the effect painter and WordArt outlines build SVG
without fetching 10 kB of fill code, and `chartLayoutReady()` preloads only
the chart layout: a solid-only chart no longer fetches the fill painter (the
render's `chartsSettled()` wait covers a fill the layout turns out to
need). The preload test's synchronous-fill expectation was updated
deliberately. PPTX still imports fills eagerly, so it claims no new saving.

**Spreadsheet pictures gain effects.** The Worker resolves
`pic/spPr/effectLst` through the shared reader (default colour map, as for
charts) plus the picture's own transform; the sheet paints through the
shared effect painter with cleanups owned by the paint session (switches
and appearance repaints redispose) and an `effectsReady` the render's
settlement includes. A box not yet attached re-queues instead of dropping
its effects — found because the pre-loaded settlement pass runs before the
host is placed — and the resize observer settles it once layout exists.
Rotation/flips paint through one shared helper (`paintTransform`,
`effect-paint.ts`), which the five PPTX call sites now use too instead of
their local copy.

**Corrections alongside.** The shared `star5` text rectangle put the radius
inside the trig angle and clamped `hf`/`vf` below their defaults; it now
follows the ECMA `a:rect` guides (`sx1 = hc − iwd2·cos342°`), pinned by
regression numbers. Invalid standalone x14 icon sets were silently dropped;
they keep their rule so the unsupported path reports them (`NoIcons` still
paints nothing, by design, in both base and x14 shapes).

**Cost.** Worker `core/worker.mjs` 206.7 kB → 208.8 kB (+0.6 kB gzip);
XLSX chunk 43.6 kB → 44.5 kB (+0.3 kB gzip); `effect-paint` +90 B;
`svg-element` a 275 B shared chunk. No new dependency, no runtime fetch.

**Blast radius.** Image resolution runs in the Worker for every picture in
every format; files without SVGs resolve identically (pinned by the
raster-only DOCX assertions and the unchanged PNG captures). Renderer
changes touch all three picture routes plus fills and bullets; the PPTX
transform unification is string-identical CSS. Public surface grows only by
optional model fields (`fallback*`, `SheetPictureModel.effects/transform`)
and one `XlsxViewer.effectsReady` getter, all documented on the API
reference page.

**Deliberate non-changes.** Linked (external) SVG is never followed. DOCX
has no Word-authored SVG picture yet (its route shares the proven resolver;
synthetic unit tests plus unchanged raster captures cover it) — a
Word-pasted SVG case is ROADMAP open verification, since appending to that
fixture moves pagination and drawing counts. XLSX shapes/text boxes, cell
fill redesign and the formula engine stay out (see §107 and the audit).

**Proof.** `image-source.test.ts` (seven synthetic edge cases: pair,
lone-SVG, dangling, raster-named, external/mistyped, absent), fixture tests
for the COM-authored slide 20 (SVG-only) and Notes drawing (SVG + PNG
fallback, shadow values, rotation, anchor), the extended replaced-render
abort coverage, and `svg-pictures.visual.test.ts` re-reading decoded vectors,
the corrupt→fallback swap, the corrupt→marked-fallback end, and the painted
`drop-shadow` + `rotate(12deg)`. Native PowerPoint/Excel exports pair slide
20 and the Notes drawing (rotation, offset shadow). Full unit, visual
artifacts, typecheck, build and consumer-build runs gate the change.

<a id="note-128"></a>

### 128. Rendered content owns inherited text traits and paint-resource names (October 14, 2026)

**Root causes.** The text painters deliberately write only resolved/authored
emphasis and tracking. Their roots did not reset all inherited traits, so an
application's bold/italic, font variant/stretch, tracking, casing or indentation
could change text and wrapping even when the file's font resolved correctly.
DOCX already reset paragraph weight, page family/pitch and surface ink; it still
inherited other traits in both flow and layout. PPTX already supplied a 1.2
default paragraph pitch, XLSX cells 10pt and line-height 1/1.1. Missing literal
OOXML typefaces usually resolve through the theme or font zero. This was not a
reason to replace those format rules with `line-height:normal` or a new shared
font/black-ink policy.

PPTX's fill names hashed part paths and source shape IDs. Those repeat between
files and between concurrent renders of one slide, so an SVG `url(#id)` could
resolve to another render's definition. Shared filters, chart fills/descriptions,
Word text-box fills and highlight names used module counters, which also repeat
across independently bundled copies. The hollow-text filter cache retained every
encountered colour/width combination merely to reuse those names.

**Why two shared corrections, not a host reset.** Resetting demo ancestors
would hide inherited typography defects from library tests. Paint and pagination
measurement need the same content-owned baseline, while each Office dialect's
family, size, ink and pitch remain resolved as before. A blanket `all:initial`
or `revert` would erase useful format defaults and still would not protect
against direct descendant selectors. Shadow DOM remains the encapsulation choice.

For resources, source IDs and module counters were both insufficient across
independent bundled copies. A private bundle namespace avoids global registries
and their retention costs; an owned name also removes the need to cache every
hollow-filter colour/width merely for identity reuse. Search's sheet has a
render lifetime, not the constant-sheet document lifetime. Restoring adopted
sheets is not a race fix: no task race in synchronous CSSOM writes was shown.
The exact current baselines, identifier construction and sheet lifetimes are
specified in [ARCHITECTURE](ARCHITECTURE.md#css-and-resource-ownership).

**Default-style consolidation.** Extending §36's class-default policy to repeated
PPTX paint layers and fallback boxes removes its old PPTX exception for these
blocks. DOCX chart/text-box layers reuse the same frame-paint class. Authored
geometry and measured sizes remain inline. The element imports the specificity
suffix rather than copying its literal. Class interning encodes pseudo-elements
and declarations separately and fails explicitly on a detected local hash
collision; no registration-order suffix pretends to solve cross-copy collisions.
The specificity remains (2,1,0). XLSX's unused spacer rule is removed, and its
range/reference washes no longer use `!important`: inset shadows paint above
backgrounds, so §115's inline-fill rationale was not a cascade conflict.

**Site/bench boundary.** One shared Shiki sheet owns theme-token selection and
the dark comment contrast override. Error text uses one light/dark token;
CopyButtons follows the existing accent focus-ring convention. Verified unused
tokens and a never-written active-state selector are removed; unstyled classes
remain legitimate hooks. Docs h1/code-wrapper/continuation-line prose selectors
exclude demo descendants like their sibling rules. Bench declares its native
control colour scheme. Paper remains paper in dark chrome. Existing component
focus/control layout rules stay local rather than adding a global reset or
retokenizing every repeated number.

**Blast radius and deliberate non-changes.** Renderer CSS/resources and docs/site
chrome only: no OOXML/Worker/model contract, public API, fixture bytes, Office
spacing/colour rules, font calibration, embedding teardown or benchmark numbers
change. Minimal native capture pages never loaded bench/site CSS: they were not
contaminated evidence, but alone could miss embedding-specific defects. The
measurement page intentionally omits app chrome/preflight; changing benchmark
methodology is not part of this CSS correction. No ROADMAP item is converted
into a duplicate completion narrative.

**Proof.** Real Arabic/table fixtures compare computed text traits, family,
size, pitch and geometry under clean/hostile ancestors with and without the
bench's installed Tailwind preflight, covering PPTX, XLSX and DOCX flow/layout.
A negative control still inherits the hostile traits. Concurrent real PPTX
gradient, table-fill, WordArt, group-effect and chart renders check unique DOM
definitions, own-subtree references and surviving paint after sibling teardown,
both within one package evaluation and across two copies. Search checks independent
registries, sheet restoration and final cleanup. Pure tests pin CSS interning,
collision rejection and resource-name uniqueness. Existing range/reference,
effect, text-box, CSP, exact-pitch and packaging tests cover the shared routes.

**Verification.** Package, site and bench production builds; strict workspace/
tools typecheck; 906 unit tests (92 browser-gated skips); all 91 browser tests
with retained native PPTX/XLSX pairs and absent-font columns; Vite/webpack
production probes including fixed-format, multi-page, CDN and inline builds.
Word's native/normal/absent-font review also regenerated successfully (11 pages;
median matched-token dy 1.7px, with existing pagination/TOC mismatches still
visible, not a new fidelity acceptance). Sample paired Arabic tables, gradient
paint and Word text were inspected. Live home/playground light/dark code colours,
low-level Word embedding and the Tailwind bench's slide/dark chrome were checked
with browser and server consoles; paper stays white. Benchmark measurements
were not rerun or published.


## Performance, benchmarks and navigation

**Current mechanisms:** [ARCHITECTURE](ARCHITECTURE.md#performance-model).
**Published numbers and measurement policy:** [BENCHMARKS](BENCHMARKS.md).
The before/after values below remain dated experiments, including corrections
to earlier rulers; they must not be copied into a current-size or speed promise.

<a id="note-32"></a>

### 32. Measured numbers on the home page, and how they're measured

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

<a id="note-33"></a>

### 33. The benchmark rulers were wrong; corrected before any product change

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

<a id="note-37"></a>

### 37. DOCX layout pages paint lazily

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

<a id="note-40"></a>

### 40. No layout jumps when a unit is replaced

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

<a id="note-60"></a>

### 60. Benchmarks measure aared; competitors are measured once (Oct 2026)

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


## PPTX fidelity

**Current scene/text rules:** [ARCHITECTURE](ARCHITECTURE.md#pptx-scene-and-text-bodies),
with [shared table rules](ARCHITECTURE.md#tables). Native measurements and
rejected approximations below explain those rules, including later reversals.

<a id="note-30"></a>

### 30. PPTX tables and text lines, matched to PowerPoint

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

<a id="note-48"></a>

### 48. PPTX autofit is what PowerPoint stored; bullets hang; a subtitle takes the master body style (Sep 2026)

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

<a id="note-54"></a>

### 54. PPTX fidelity pass: where PowerPoint puts text, colour maps, table styles, Arabic around punctuation (Sep 2026)

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

<a id="note-103"></a>

### 103. PPTX visual review accepted (October 10, 2026)

The maintainer explicitly accepted the existing native PowerPoint/aared visual
comparisons as acceptable for normal reading, closing the PPTX visual review
sign-off gate. This records acceptance of the review already performed, not a
new capture run or a claim of pixel-identical rendering.

The open sign-off section leaves ROADMAP. The accepted rendering differences
remain documented in ARCHITECTURE's Known limits. The resolved-model contract
decision and the release checklist remain separate open work.


## DOCX layout, tables and drawings

**Current layout/painting:** [ARCHITECTURE](ARCHITECTURE.md#docx-flow-pages-and-drawings).
The probes below are the evidence for that design, not an alternative list of
what remains open. [The inventory](ARCHITECTURE.md#docx-pagination-wrapping-and-text-boxes)
owns today's dispositions.

<a id="note-7"></a>

### 7. DOCX pagination is explicitly not promised

Faithful pagination — matching Word's exact page breaks — needs the same
layout engine that made DOCX the hardest format (§5). Promising it would
either delay DOCX indefinitely or produce a renderer that's "mostly right,"
which is worse than being upfront: aared renders DOCX as flow-layout HTML
(matching the `docx-preview` tier of fidelity) and states plainly that page
count won't match Word. This is recorded as part of the fidelity contract in
[ARCHITECTURE.md](ARCHITECTURE.md#fidelity-contract) so it's a documented
boundary, not a discovered one.

<a id="note-10"></a>

### 10. DOCX layout mode: structural pagination without Word parity

Stale. Opted DOCX into physical pages (`pagination: { mode: "layout" }`)
beside flow rendering (§7), because page markers alone cannot paginate a real
document: Word printed four pages where the markers implied two. The parser
resolves `DocumentSection[]` and a headless paginator fills section boxes with
estimated block heights, so break positions match structurally, not exactly.
Its layout rules were later re-derived from native Word's own numbers.
Superseded by §22.

<a id="note-11"></a>

### 11. DOCX layout fidelity corrections from bench comparison

Stale. A bench comparison against Word fixed three defects, each traced to its
OOXML source: paragraph borders (`w:pBdr`) were dropped, list indents painted
twice, and headers/footers stacked inside the page consumed room the paginator
never reserved. It also made inline images inherit their paragraph's
alignment. The header clipping, border/indent geometry and list geometry
were later re-derived from native Word and changed. Superseded by §22.

<a id="note-21"></a>

### 21. Measuring text for DOCX pagination: measure in the DOM; Pretext considered, not adopted

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

The estimator stayed because it was the only page layout that ran without a
browser, motivating the then-proposed "headless thumbnail export". That proposal
now lives in [ARCHITECTURE](ARCHITECTURE.md#unscheduled-proposals), not a delivered
browser-free renderer or the retired ROADMAP.
Revisit Pretext when headless accuracy is scheduled, or if DOM measuring is
too slow on very large documents; then run a spike first (compare line counts
with the DOM for every fixture paragraph) and verify Worker or server support
and the licence before depending on it.

<a id="note-22"></a>

### 22. DOCX layout follows Word's own numbers; supersedes parts of §10 and §11

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

<a id="note-29"></a>

### 29. Real zoom fit for PPTX and DOCX; the DOCX heading/page facade gap closed

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

<a id="note-43"></a>

### 43. DOCX right-to-left tables: `w:bidiVisual`, logical borders, physical margins

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

<a id="note-44"></a>

### 44. Fixes found by the showcase documents (Sep 2026)

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

<a id="note-55"></a>

### 55. DOCX pictures that text wraps around, placed as Word places them (Sep 2026)

An anchored picture (`wp:anchor`) used to paint as a plain left or right CSS
float whatever its wrap, so a centred picture sat at a column edge, a
top-and-bottom picture had text beside it, and a behind-text watermark pushed
the text away. Probe `docx-float-wrap` (gitignored) authored 13 wrap cases and
a page watermark through Word's COM, compared line by line with aared.

**The model and planner separated two problems.** The old `float` field could
name a side, not Word's wrap/position vocabulary. Retaining the anchor and a
paragraph-owned `floats` collection enabled a numeric, browser-independent
planner instead of another CSS guess. Word writes wrap polygons in a 21600
square; the model normalizes those coordinates to picture shares. The present
wrap/band/overlay/split placement rules are in
[ARCHITECTURE](ARCHITECTURE.md#docx-flow-pages-and-drawings), and the field
shape belongs to the [model reference](../site/src/content/docs/api.mdx#resolved-models).

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

<a id="note-76"></a>

### 76. Four Office probes: bold by script, number-format ids, empty paragraphs, first-page headers (Oct 2026)

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

<a id="note-77"></a>

### 77. List counters per list definition, and the coverage fixtures extended (Oct 2026)

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

<a id="note-82"></a>

### 82. DOCX text boxes and section line grids (Oct 2026)

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

<a id="note-90"></a>

### 90. Empty section-ending page-break paragraphs do not add pages (Oct 2026)

An empty paragraph carrying both `w:pageBreakBefore` and `w:sectPr` produced an
extra blank page: it forced a break before itself, then the next section started
on another page.

The current fixture was copied byte-for-byte and freshly exported through Word:
it reports seven pages (portrait, portrait, portrait, landscape, portrait,
landscape, portrait). A new aared render of those exact bytes also reports seven.
The seven picture-wrapping examples share page 5, charts remain on page 6, and
the final list/table section remains on page 7. Microsoft documents
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

<a id="note-91"></a>

### 91. DOCX pagination stays DOM-measured; fidelity work is bounded by visible value (Oct 2026)

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

<a id="note-92"></a>

### 92. Numbered DOCX headings and refreshed table-of-contents results (Oct 2026)

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

<a id="note-94"></a>

### 94. A DOCX correctness pass: two limits ruled, autofit and fills probed, numbered headings with a live TOC, and a zoomed container (Oct 2026)

Five DOCX defects closed and two differences ruled, each against Word 16 ground truth. Every one is a rule read from the wrong place or owned in two places rather than a wrong rule, which is why the fixes are shared rather than local.

**Two differences carried two verdicts.** An unstyled chart's area border and DOCX text-box `normAutofit` were each listed twice with opposite verdicts: ARCHITECTURE's Known limits settled both while ROADMAP's Open gaps called both work. Two lists claiming one row makes neither trustworthy, so each is ruled once and now lives only in Known limits. An unstyled chart draws Excel's thin grey border in all three formats and PowerPoint draws none; it stays, being a hairline only Office's own default supplies, and the "per-format default" clause that read as open work is gone. `normAutofit` stays unread, for the reason the probe below gives.

**Word ignores a stored `normAutofit`, honours `spAutoFit` in some files, and clips the overflow aared spills.** Probe `autofit` (Word 16.0): five inline boxes at 200x30pt, 11pt Arial needing seven lines, each `wps:bodyPr` matched to its case by the token in the same `wps:wsp` block and never by document order, because Word writes shapes in the reverse of creation order, which made an earlier positional run of this probe report that `noAutofit` grows a frame. Sizes were read from Word's own PDF (`word-spans.json`), not judged from a screenshot: 11pt is 14.72px at 96dpi.

A bare element, `fontScale="90000"` (would be 13.2px) and `fontScale="62500" lnSpcReduction="20000"` (would be 9.17px) all painted at 14.72px, and Word wrote every `bodyPr` back byte-identically: neither consumed nor discarded. aared already paints the sizes in the file, so **reading `fontScale` would have moved aared away from Word and toward PowerPoint** - the opposite of the fix this item assumed. `spAutoFit` is the half Word does honour: in one probe file the box grew from its authored 30pt to 91.6pt (seven lines at Word's 12.65pt pitch) and the PDF painted all seven, while a second file differing in box size (320x40pt) and carrying Arabic as well as English stayed at exactly 40pt, like the boxes beside it. So "Word grows the frame" is not yet a fact and "aared never grows it" is; since `Resize shape to fit text` is the one autofit option Word's UI offers the gap is real either way, and what the condition is stays open rather than guessed. Separately, Word paints only the lines that fit the frame, clipping overflow even with `vertOverflow="overflow"` in the file, where aared sets `overflow: visible` and spills text Word cuts - a visible difference in the opposite direction, now a ROADMAP item in its own right. Whether Excel honours a carried `normAutofit` inside a chart's `c:txPr/a:bodyPr` is settled in §95: it does, and both `fontScale` and `lnSpcReduction` are now applied.

*Method note.* `TextFrame.AutoSize` proved useless as evidence **in that run**: setting a scratch shape to all three `msoAutoSize*` values read back `0`. That conclusion is superseded by §95, which found `-1` (`msoAutoSizeShapeToFitText`) does write `<a:spAutoFit/>` - the run's `TextFrame2` accessors are unavailable on this Word build, and its helper swallowed the COM failure, so nothing was ever assigned. The claim that Word has no shrink-to-fit option rests on Microsoft's Text Box pane reference (Word's AutoFit group lists only `Resize shape to fit text`; the three-option group is PowerPoint's). Recorded so nobody re-runs the enum without checking the assignment reached the shape.

**A numbered heading carries its marker; a supported TOC refreshes.** The reproduction is the Word-authored `tools/word-render/out/word-source/word-source.docx`, which reports `ListString` `1.` for its Heading 1 paragraph and paints `1. Project Overview`; aared now paints the same line, and its page-1 table of contents lists all nine entries, including the two the cached result had missed (`5. Delivery tracker / متابعة التسليم` and the delivery table's Heading 2 `الإجمالي`).

The parser already advanced a heading's counter and then threw the computed marker away, leaving the node unable to paint it. `DocxHeadingNode` now carries `marker`, `markerRun` and `markerSuffix` beside `DocxListItemNode`'s three, resolved through the same `advanceCounters`/`markerFor` path, so one sequence runs through both (a heading between two items gives `1.`, `2.`, `3.`), and the renderer paints it through `paintLine`'s existing marker path rather than a second one. One semantic call: `text` leads with the marker (`1. Project Overview`) while `runs` keep the authored text alone - `text` is what a reader sees and what Word puts in an outline and a TOC entry, so outline titles, `section.title` and refreshed entries agree, while `runs` stay what search matches. The marker is rendered text, never folded into the runs, so the browser still lays it out (rule 1).

A TOC field refreshes, or it keeps its whole cached result, never half. Field tracking was per paragraph, which cannot see a TOC: Word's field spans several paragraphs (instruction and first entry in one, later entries in their own, the `end` in a paragraph of its own) and each entry nests its own `PAGEREF`. Tracking is now context-level and enabled only while streaming the body; tables, text boxes, headers, footers and notes parse untracked, since a TOC never lives there. Cached paragraphs are held in draft markers and either replaced or unwrapped in one place after the body is parsed, so the public model never sees them and a half-read field is never refreshed. One subtlety cost the first run entirely: the opener's own `separate` arrives before its session exists, so the flag is recorded per paragraph and handed to the session - without that, a nested `PAGEREF` instruction leaks into the TOC's own switches and every field silently falls back to its cache.

Supported switches are heading outline levels (`\o`, and `\u`, which resolves to the outline levels the parser already computes), `\h`, `\n` with or without a level range, `\p`, and bookmark ranges (`\b`, honoured through the body's bookmark index — see below); `\w`, `\x` and `\z` are accepted and change nothing (`\z` hides page numbers in Web layout, which flow mode does not have). Everything else keeps the cache - caption and sequence (`\a`, `\c`, `\s`, `\d`), TC entries (`\f`, `\l`) and anything unrecognized - because a partial refresh would silently drop entries the switches select, which is worse than a stale table whose structure is still the reader's own.

Selection is its own document-order walk, not `document.headings`: Word's refreshed export includes `الإجمالي`, a Heading 2 inside the delivery table, and the public outline deliberately excludes table-cell headings. So entries are collected from the parsed drafts and given `tocAnchor` fragments to link at when they sit outside the outline. Page references are a live `PAGEREF` run resolved after pagination against the pages aared laid out (section restarts honoured, map render-local so resolved models stay read-only); flow mode has no pages, so an entry matching a cached one keeps its saved number and an entry added since the TOC was saved shows none rather than an invented one.

**The first attempt at refreshed entries was wrong in a way only a capture caught.** Synthesizing entries from the headings bypassed `resolveParaProps`, so entries ran tighter than Word's, level-2 entries sat flush, and an RTL heading's entry was mirrored because direction was copied from the heading where Word takes it from the `TOC{n}` style. Three visible differences, one cause. Resolving the entry style the way `parseParagraph` does cleared all three; the entry is also *set* in that style, because copying the heading's resolved runs printed every entry bold, since `Heading 1`'s bold is style-derived and Word drops it. A test then caught a fourth: `finalizeRun` drops empty text, which deleted the page-number run for an entry with no cached number - exactly the newly-added entries - so the run is now pushed directly and kept even when empty.

**The refresh rule is one invariant, stated once.** A field is refreshed only when every switch that could change its result is understood; anything else keeps the whole saved result. A refresh that silently drops what an unread switch selects is worse for a reader than a table that is merely stale, because the stale one is at least the reader's own structure. Two places enforce it, and a third now exists for switches whose data lives in the document rather than the instruction: `parseTocInstruction` returns `undefined` for a field it cannot fully evaluate, `resolveTocRanges` unwraps a session whose instruction went undefined mid-stream (how a late switch in a split instruction falls back), and `tocFieldSupported` rejects a field whose data the body does not carry. Adding a switch therefore means saying what it does; one that cannot be honoured belongs with the others rather than parsed and ignored.

**`\b` is honoured, and it needed a bookmark index rather than a TOC special case.** A bookmark is a named label around content - `w:bookmarkStart` with a name, `w:bookmarkEnd` with the same `w:id` - and `{ TOC \b Part1 }` collects entries only from that range. The parser discarded both elements, so honouring the switch meant first learning where content sits. Rather than a map that exists only for TOCs, the body records a `DocxBookmarkIndex` while `w:body` streams: named ranges in **block-index** space, half-open, because a heading is exactly one block and that is what a `\b` collects. Boundaries are read both between paragraphs and inside one (a boundary inside a paragraph takes that paragraph's blocks, which is why the range end is the index after it, not the index of it - the first version got that wrong and listed nothing). Word merges same-named ranges into one, so a name maps to a list and a heading inside any of them is listed, in document order. A range that would pair a start and an end across a section break is dropped rather than joined across two index spaces, so such a field resolves honestly or not at all: `tocFieldSupported` treats a `\b` naming a bookmark this document does not hold as unsupported, and the field keeps its cache instead of listing the whole document into something that asked for a chapter. Word's own `_Toc...` bookmarks stay unread - aared mints its own fragments for those - so the index carries only the labels a person made.

**One residue is settled rather than queued: direct character formatting inside a heading.** Worth separating two things that read alike and behave oppositely. What works is the loud case - an entry is *set* in its `TOC{n}` style, so a heading's **style-derived** appearance never carries, `Heading 1`'s bold prints in no entry, and no TOC line goes bold because a heading is bold. What aared drops is formatting applied **directly** to a run inside the heading (a bolded or coloured word, authored as `w:rPr` on the run itself), which Word keeps. The obstacle is that the distinction no longer exists by refresh time: `readRun` applies the style chain first and the run's own `w:rPr` on top, and `TextRunModel` carries no record of which came from where. Recovering it means either carrying direct-only formatting alongside each run - widening the public model, on every heading in every document - or, at refresh time, carrying only the attributes that differ from the heading style's own run defaults, which is roughly thirty lines and no model change. Deferred past `0.1.0` on release order rather than effort: the case is rare and invisible until someone bolds a word mid-heading, it needs a Word probe first because the behaviour is asserted here and proven by nothing in the repository, and carrying a run's direct attributes wholesale can make one entry's weight disagree with the rest of its own table of contents, which reads worse than the flat entry does. Faithful is still the target; the timing is what waits. One case survives any of it: a run that sets bold *explicitly* where the heading style **already** sets bold is indistinguishable from the style alone, so it would still be dropped - redundant authoring, unfixable without the model change. Recorded as a Known-limits row rather than a ROADMAP row, since it is settled as accepted and the row carries its own reopen condition (probe, then fixture).

**A shape's fill is the fill it authors itself.** Probe `docx-drawing-fill`: aared painted DOCX text boxes solid black, and in their outline colour, where Word paints no fill. One expression caused it - `drawingFill` looked for a `solidFill` anywhere under `wps:spPr`, and Word gives every inserted DrawingML shape a black outline, so `a:ln/a:solidFill/a:prstClr val="black"` sits under almost every text box; turning the outline off leaves `a:ln/a:noFill` plus `a:extLst/a:ext/a14:hiddenLine` carrying the same black so the formatting can come back, and a subtree search finds either one. The blast radius was wider than one shape: a box with the fill off and a red outline painted solid red, a box with both off painted black, and a **gradient** and a **pattern** fill each painted black, because the outline beside them was the only `solidFill`. Every text box whose fill was not a direct `solidFill` was wrong, and a direct fill was the only input on which the deep and scoped searches agree - which is why the fixture never showed it. Effect lists are not part of this (`a:outerShdw` and `a:glow` carry their colour as a direct `EG_ColorChoice`), and PPTX is immune because it reads fills through `resolveDirectFill`.

One home for the rule: `core/fill.ts` gained `directFillProperty`, the one fill child a container authors itself, `resolveDirectFill` switches on it, `drawingFill` returns a colour only for a direct `solidFill`, and `drawingBorder` uses it for its `noFill` check. The DOCX route deliberately keeps a string fill and a plain `border` rather than a `ShapeFillModel`: a text box is a flow box painted with `backgroundColor`, so there is no SVG paint server to route through `fill-paint.ts`, which is innocent. Fixed in the same read: `wps:bodyPr` follows `wps:txbx`, so the search that found it returned a **nested** box's insets and `vertOverflow` while walking `spPr` twice for one answer; both are scoped to the first `wps:wsp` in document order. Gradient, pattern and picture fills are now cleanly dropped rather than miscoloured but still dropped and still unmarked; `prstGeom`, `a:xfrm/@rot`, `bodyPr/@anchor`, `bodyPr/@wrap` and `prstDash` stay unread - separate capabilities, not defects in this one. `fixtures/docx-coverage.docx` gained a Word-authored box with fill and outline both off, written through Word's COM object model and label-stripped, with five cases in `test/docx-text-box.test.ts` reading it back (§95 adds seventeen more on three native pages).

**Pagination measured painted pixels against layout pixels, so a zoomed container invented pages.** Reproducing a consumer report - Chrome showing ten pages for `docx-coverage.docx` where Word showed seven - reproduced it in a clean profile: with `zoom: 1.5` on the render container, the fixture paginated to ten. Root cause: two coordinate spaces in one measurement pass. `measureFlow` read each block's height from `getBoundingClientRect()` (painted pixels) while its margins came from `getComputedStyle()` and the page box from the model (layout pixels); a CSS `zoom` on any ancestor scales only the first, and Chromium rounds the two differently per factor, so they drifted 2.3x apart at 1.5 - a paragraph whose `clientHeight` was 36px reported a rect of 82.75px. The ratio is not monotonic in zoom, which is why headless Chrome and Edge agreed on seven, Firefox agreed on seven, and one windowed Chrome did not. Blast radius: the DOCX pagination probe, and the split-wrap pass beside it, whose slot-overflow test compared a `Range`'s painted width against a layout-pixel slot width - reachable with no consumer action at all, since aared puts `zoom` on the render root itself for `fit-width` and `fit-page`. `tabs.ts` already divided by a measured ratio and was left alone: it is calibrated, correct, and takes its ratio from its own root's width. The fix has one home: `src/docx/measure.ts` gained `paintedToLayoutZoom(root)`, a 100px ruler read once whose painted width over 100 is the ratio; `measureFlow` divides block heights by it and `split-wrap.ts` takes it as a parameter. At zoom 1 the ratio is exactly 1 and the paired capture is unchanged token for token, so this is inert in the normal path; the visual test renders with the zoom in place *before* `render()` - the case that matters, since a consumer's browser zoom lands before the first paint - and fails 10 against 7 with the division removed.

**The report that sent us looking was not this defect, and it is worth saying so plainly.** A consumer saw ten pages in one Chrome where Word, Firefox and Edge all showed seven. It survived a reboot and was not reproducible in any clean profile this repository could build: window sizes 900 to 1600, display scales 1 and 0.9, CSS zooms 0.75 to 1.5, six consecutive loads, repeatedly with `fonts.status=loaded`. The zoom defect was found on the way, because `zoom: 1.5` produced the same number, and was worth fixing on its own merits - but it was not the cause. The cause was that reader's own Chrome profile, `webkit.webprefs.minimum_font_size = 16`, which floors every run below 16px, and the fixture is full of 9-12pt text; set to 0, that browser returned to seven pages. No engine, no aared code, and no amount of bisecting a profile would have found it as a defect. What it costs: the browser lays out and paints the enlarged text, so aared measures it honestly and paginates on it - no overlap, no clipping, simply more pages than Word. aared does not override the setting, being an accessibility choice the reader made, so this is a Known-limits row and no warning is emitted: it would need a fixture and a test per format, and "your browser enlarged the text" is not actionable enough to earn that.

**Deliberate non-changes across this pass.** No exported symbol changed except the DOCX model additions named above. Numbered-heading markers are not folded into runs, and a TOC field is never refreshed on a partial reading.

<a id="note-95"></a>

### 95. DOCX text boxes: resize, clipping, geometry, rotation, fills (Oct 2026)

The correctness pass in §94 left five DOCX text-box gaps a paired Word/aared
review showed a reader notices: a box saved with `Resize shape to fit text` grew
its text outside a fixed background, gradient/pattern/picture fills painted
nothing, `prstGeom`/`@rot`/`@anchor`/`@wrap`/`prstDash` were unread, and a fixed
box spilled text Word clips. This entry closes them as one capability - the
model, one shared read path, one renderer structure - and settles the chart
`c:txPr` autofit question §94 left open.

**Word's own `AutoSize` setter works after all.** §94 recorded that
`TextFrame.AutoSize` "read back 0" for all three `msoAutoSize*` values, from a
probe whose COM calls were swallowed by a helper, so nothing was ever assigned.
Probe `word-textbox-native` found the truth: `TextFrame.AutoSize = -1`
(msoAutoSizeShapeToFitText) writes `<a:spAutoFit/>`, and Word recomputes the
frame on reopen; the `TextFrame2` accessors that probe tried are unavailable on
this Word build. Every resize case now in the fixture is authored by Word
itself, not patched XML. A Word-saved `spAutoFit` box keeps the *authored*
extent in `wp:extent/@cy` and recomputes only on reopen, which is why the
committed fixture cannot carry a final height - the test drives both directions
from one authored box instead.

**The frame's fit mode, not `vertOverflow`, decides clipping.** Word-authored
DrawingML text boxes carry `vertOverflow="overflow"` universally and still paint
only the lines that fit, so the attribute carries no signal; §94's paint rule
(`clip` -> `hidden`, else `visible`) was reading the wrong input. A `resizeShape`
box lets the browser size the frame to its content - Word recomputes in both
directions - and a fixed box keeps its authored extent and clips. Clipped words
stay in the model, so search still reaches text Word itself hides.

**Why the model had to change.** The former `verticalOverflow` flag encoded
the wrong signal. An autofit-driven frame and shared shape fill/geometry can
express what Word actually paints, rather than adding exceptions to that flag.
The model fields belong to [the reference](../site/src/content/docs/api.mdx#resolved-models);
the text-rectangle/inset formulas belong to
[ARCHITECTURE](ARCHITECTURE.md#geometry-fills-pictures-and-effects).

Preset named colours also moved through `resolveDrawingColor` before modifiers,
so the private Word colour lookup no longer bypassed the shared pipeline.
This is a **cross-format** gain, not a DOCX one: `resolveDrawingColor` is
the one colour reader for PPTX shapes, table cells, backgrounds and
WordArt and for charts in all three formats, so an
`<a:solidFill><a:prstClr val="dkGreen"/></a:solidFill>` that resolved to
`undefined` before - leaving the shape's style fallback, or no fill, in its
place - now paints the named colour in every format. DOCX simply held a
private twelve-name table for the twelve names it had seen; the other ~128
ECMA `ST_PresetColorVal` tokens vanished there. One table, one pipeline.

**Why frame, surface and content must be separate.** Rotation can change a
bounding rectangle without changing the flow space Word reserves; §94's
painted/layout-pixel error therefore reappears if pagination measures the
rotated surface. Aligning individual paragraphs would also lose their collapsing
margins. Those are the reasons for an untransformed frame and one browser-text
flex item, not three layers chosen merely for naming symmetry. The complete
[text-box structure and settlement](ARCHITECTURE.md#text-boxes) live in
ARCHITECTURE, including batched measurement and lazily skipped pages.

**§94's fill reasoning was incomplete, and this overturns it.** §94 kept a
string fill and a plain `border` for the DOCX route because "a flow box has no
SVG paint server to route through `fill-paint.ts`". That described the
then-current implementation, not a constraint: a PPTX table cell is a flow box
and is painted through the same shared painter. Text-box fills and outlines now
paint through `fillPaint` exactly like a PPTX shape, cell or chart, with
`data-aared-fill-fallback` marking what cannot paint. The preset geometry comes
from `drawingml-shape.ts`'s fast path, so no preset catalog reaches the DOCX
route, and a resized round rectangle keeps circular corners instead of
stretching its authored radius.

**Chart `c:txPr` autofit is honoured by Excel, and now by aared.** Probe
`chart-txpr-autofit` (Excel 16.0 build 20430, PDF span sizes at fixed print
scaling) measured a stored `normAutofit/@fontScale` halving both an explicit
24pt title (22.08 -> 11.04 PDF pt) and an inherited default (16.56 -> 8.28), and
`@lnSpcReduction="20000"` shortening the baseline gap from 25.39pt to 20.30pt
without shrinking the glyphs. `readText` applies both after run/default
inheritance; `lnSpcReduction` became `TextLook.lineSpacingReduction`, which the
shared `lineHeight` honours, because folding leading into font size would have
shrunk the glyphs too. Saveback keeps a non-default `fontScale` and drops the
default `100000`. No real chart part in `fixtures/` or `tools/chart-oracle/`
carries an autofit child - which is why this was never visible in the corpus,
and why corpus absence says nothing about whether Office honours one, only how
often it writes one. The roadmap had predicted "most likely document nothing to
read"; the measurement said the opposite.

**Fixture, tests and the paired review.** `fixtures/docx-coverage.docx` grew
three native Word pages (fills and geometry, placement, sizing) holding
seventeen Word-authored text boxes, label-stripped with
`tools/probe/_common/strip_label.py`; the seven original pages are unchanged and
the document is ten pages in both engines.
`test/docx-text-box.test.ts` reads the model for each case;
`test/docx-text-box.visual.test.ts` re-reads the painted DOM - gradient, pattern
and picture paint servers, preset paths with arcs, a rotation matrix, computed
whitespace, vertical gaps - and proves resize in both directions, that a fixed
frame clips while keeping its text, and that a render never mutates the resolved
model. The paired capture of the three new pages
(`tools/word-render/out/docx-coverage/review.html`) was inspected: fills,
geometry, rotation, middle and bottom anchoring, dashed outlines, resize and
no-wrap all match; the clipped box differs only in how many whole lines fit
before the cut.

*Deliberate non-changes.* Word's stored DOCX `normAutofit` font scale stays
unread (§94: Word ignores it, so reading it would move aared away from Word and
toward PowerPoint). The PPTX route's `resizeShape` handling is unchanged - it
grows only the text layer, which is right for PowerPoint, so it was not the
reuse the roadmap suggested. A chart pattern anchors at its box's own origin,
which is recorded in Known limits as an approximation. VML fallback text boxes
keep their reduced fill and anchor reading.

<a id="note-97"></a>

### 97. A row's height is read from where Word writes it (Oct 2026)

An audit of the roadmap against the code found one silent drop and one stale
row. This entry records the defect.

**`w:trHeight` was read off the wrong parent.** `parseTable` took
`firstChild(row, "trHeight")`, but ECMA-376 17.4.81 puts `w:trHeight` inside
`w:trPr`, a child of `w:tr` — so the lookup could never match a real file and
every row height Word authored was dropped. The rule is a plain reading
mistake, which is why nothing caught it: the model field, the paginator and the
renderer were all consistent about a `height` that never arrived. Two
independent reasons the suite never showed it — no test asserted `row.height`,
and `fixtures/docx-coverage.docx` happens to carry **zero** `trHeight`
elements (its three `<w:trPr>` blocks hold only `w:tblHeader`), so the
fixture could not have shown it either. The regression test is therefore
synthetic, as the parser's other edge cases are (see the revision-tracking
test above it), and pins the nesting rule rather than a reading.

`w:hRule` is still not read, deliberately, because the consumer is already
right without it: `foldTableHeightPx` takes the larger of the cell content's
height and the authored one, which is Word's default `atLeast` semantics. So a
`hRule="exact"` row paints taller here than Word clips it to — the same
divergence as everywhere else in DOCX pagination, where nothing clips a row
(§91). `DocxTableRowNode.height` now says so, because a reader of the model
would otherwise assume a ceiling. `w:tblHeader` sits in the same `w:trPr` and
is still unread, which is the queued ROADMAP gap; `parseTable` is now its home
too.

<a id="note-98"></a>

### 98. A preset's text rectangle, and a fallback that keeps its fill (Oct 2026)

`ShapeGeometryModel.textRect` had been on the shared model since the DOCX text
boxes landed, but only DOCX read it. A PPTX shape laid its text across the whole
painted box, so PowerPoint's narrower column in a rounded rectangle, an ellipse
or a triangle was not reproduced - visible at a glance, since text that sits
outside a shape's narrow waist reads as broken wrapping rather than as a detail.

**Measuring it.** `tools/probe/pptx-text-rectangle/probe.py` authors one shape
per inset preset at two aspect ratios, left-aligned (an AutoShape centres by
default, and a centred line's `BoundLeft` is ink, not the column's edge), and
reads `TextFrame2.TextRange.BoundLeft` back. Two findings the spec text does not
give you: Office **omits** `lIns` when it is the 0.1in default, so comparing
against the raw guide put every row out by exactly 7.2pt; and the hexagon's
`a:rect` guides are conditional on `a = x1 - ss/2`, which is a real branch, not
a simplification - an early formula got `q5` wrong and the probe caught it. The
probe now raises rather than falling back when a shape is missing from the
package read-back, because it once compared only slide 1 while reporting all
twenty rows. Worst disagreement after the corrections: **0.000pt**.

**The fix** composes `textRect` with `a:bodyPr`'s insets in `appendShapeText`.
One trap: a CSS percentage `padding-top` resolves against the element's
**width**, while `a:bodyPr`'s vertical insets are a share of the box's height, so
the vertical sides are resolved in JS and applied in pixels. Using percentages
there is wrong for every non-square shape.

**A geometry aared cannot draw now paints its fill.** The pentagon has no fast
path, and the fallback drew an empty dashed rectangle - while PowerPoint fills it
teal and sets white text. The result was white text on a white slide inside a
dashed outline, which reads as a missing shape, not a missing outline: this is
the same class of silent `fill="none"` that §28 fixed for fills, and the honest
stand-in for an undrawable outline is an undrawable outline. A fill the painter
still cannot draw is left undrawn, so the stand-in never claims a shape that fill
was never in. The test that pinned the old behaviour
(`pptx-render.visual.test.ts`) was updated to the new contract; the ink colour was
never wrong - the parser already resolves automatic text against the shape's own
fill, so only the fill was missing.

Both behaviours are pinned by negative check: reverting either one fails a test.
The probe directory is deleted; the comment, the formula tests and the renderer
tests are the durable record.

<a id="note-99"></a>

### 99. A table's style is drawn from the style, not from the cells (Oct 2026)

A DOCX table naming a `w:tblStyle` painted plain and unbordered: `styles.xml`
resolved paragraph and character styles only, and `parseTable` read direct
formatting alone. The fixture could not show it - all five of its tables carried
their formatting directly, and it defined only `TableNormal`, which authors no
regions.

**What Word actually writes.** The probe (`tools/probe/docx-table-style/`) built
six 7x3 tables in Word and read both the written XML and the vector fills in
Word's own PDF export. The decisive fact: Word authors **no per-cell shading at
all** for a style-driven table (`ownShd=0` on four banded tables). The whole
look is `w:tblStylePr` regions in `styles.xml`, which is why reading only direct
formatting cannot reproduce it.

**Band counting excludes the header row.** With `w:tblLook/@firstRow="1"` the
painted band rows were 2, 4, 6; with `firstRow="0"` they were 1, 3, 5, 7. So the
first *body* row is band 1 when a header row is on. That is
`(row - (firstRow ? 1 : 0)) % 2`, the rule `core/pptx.ts` already implemented
for `a:tblStyle` - so the resolver moved to `core/table-style.ts` and both
formats call it, rather than a second copy. A vertical merge does not shift the
count; the merged cell takes its origin row's band.

**Precedence, strongest first**, all measured: direct `w:tcPr/w:shd` beats a
region's fill (a header cell painted its own `#00FF00`, not the region's
`#156082`); direct `w:b val="0"` beats a region's bold; a *paragraph* style's
colour beats a region's colour (a `Heading 2` paragraph in a header row painted
`#0F4761`, not the region's `#FFFFFF`) while the region's weight still applied
where the paragraph style was silent. So a region layers below the paragraph and
character styles, and `ParseContext.tableStyleText` carries the region's `w:rPr`
and `w:pPr` into the cell's own resolvers at exactly that point.

**`w:tblLook` carries both forms.** Word writes the explicit attributes *and* the
legacy `w:val` bitmask, and they agree (`0420` = firstRow + noVBand). The
fixture's tables write only `val="0000"`, so `tableLookFromVal` reads the
bitmask for files that have no attributes.

Two probe traps worth recording. Borders are **filled** 1px rects in the PDF,
not strokes, so a probe counting strokes sees no borders at all. And fills must
be searched **per page**: every page's tables start at the same y, so one
accumulated list reported bands on a `Table Grid` table that authors none, which
would have been codified as a false finding.

**Deliberately not read:** `w:tblStyleRowBandSize` / `w:tblStyleColBandSize`
(a band wider than one row or column), listed in *Known limits*. The probe's
style sets both to 1, which is what the resolver assumes. `w:tblHeader` and the
rest of `w:trPr` remain the queued row, unchanged.

<a id="note-100"></a>

### 100. Word theme ink includes its shade or tint (Oct 2026)

`readRunProps` found `w:themeColor`, emitted the raw theme colour, and skipped
`w:themeShade`/`w:themeTint`. A Heading 2 using accent1 with shade BF therefore
painted `156082` instead of Word's `0F4761`. The browser was faithfully painting
the wrong resolved value; no renderer correction could recover the discarded
adjustment.

**Native evidence.** Python drove Word 16's `Font.TextColor.ObjectThemeColor`
and `TintAndShade` on bilingual paragraphs in the coverage document. Word wrote
BF for -25%, 99 for +40%, 7F for -50%, and 33 for +80%; its own PDF painted
`0F4761`, `45B0E1`, `0A2F40`, and `C1E4F5` from base `156082`. These are HSL
lightness adjustments, not DrawingML's linear-light tint or a direct RGB blend.
The resulting channels truncate. The probe also caught two authoring traps:
Word's theme enum numbers differ from PowerPoint's, and `TextColor.Brightness`
does not encode this operation like `TintAndShade`; neither guessed COM call was
used as evidence.

`resolveWordRunColor` lives with the existing colour arithmetic in `theme.ts`.
It resolves the theme slot and byte adjustment, with literal/cached `w:val`
only as a fallback. A test changes the theme while keeping the cached colour
stale, proving the fix is resolution rather than copying the cache. Invalid
bytes leave the base intact; zero/full bytes, aliases and absent themes are
covered. Native RGB values are asserted on both parsed runs and actual DOM ink.

**Blast radius.** One `readRunProps` path serves direct runs and inherited
paragraph, character, table and numbering styles, including notes, headers,
footers and text boxes. DrawingML and SpreadsheetML transforms retain their own
semantics and tests. This change is text ink, not Word's separate shading/border
theme attributes. No model or renderer change is needed for Word.

<a id="note-104"></a>

### 104. Table-level borders and padding inherit through the full style chain (October 10, 2026)

The DOCX table reader already resolved cell regions and text properties
base-first, but table-level borders and default cell margins selected only
`tableStyleAncestry(...).at(-1)?.whole`. That discarded the base styles whenever
the final style omitted a property, including a child that supplied no table
properties at all.

`parseTable` now folds every style's whole-table `tblPr` base-first, merging
border edges and padding sides individually, then applies direct table
formatting last. Explicit `nil`/`none` borders and zero margins remain overrides.
The existing cycle-safe ancestry and border reader/merger are reused; resolved
models and public signatures do not change.

Fixture-derived tests vary only the style ancestry and direct table properties
of the existing Word-authored coverage table. They cover empty descendants,
three style levels, partial overrides, and direct formatting. Rendered-output
coverage re-reads actual borders and padding in flow and layout modes, including
the RTL table's logical borders and physical margins. Every DOCX table parsed
through this path benefits, including tables in nested content and stories.
Cell-region/text inheritance and PPTX/XLSX paths are unchanged. This closes the
previously flagged table-level inheritance bug; existing table-style limits
such as multi-row/column band sizes remain in ARCHITECTURE.


## Fonts, script and line measurement

### Frozen calibration outcome

These are the retained choices from the calibration work, not the default
renderer's fonts or a promise of identical Office line breaks. §§45–§53 below
explain the measurements, rejected candidates and eventual freeze. The current
[font delivery contract](ARCHITECTURE.md#font-delivery-and-aliases) belongs in
ARCHITECTURE; pinned descriptors and rerun procedures belong in
[the calibration catalog](../tools/font-calibration/README.md).

| Requested | Substitute | Approach |
|---|---|---|
| Calibri / Calibri Light | Carlito | Latin metric clone |
| Cambria | Caladea | Latin metric clone |
| Arial | Arimo | Latin metric clone |
| Times New Roman | Tinos | Latin metric clone |
| Courier New | Cousine | Latin metric clone; Arabic not substituted |
| Aptos / Aptos Display / Aptos Narrow | Radio Canada / Roboto / IBM Plex Sans Condensed | Calibrated per family/weight |
| Traditional Arabic / Arabic Typesetting | Amiri | Calibrated per requested font |
| Simplified Arabic | Lateef | Calibrated |
| Sakkal Majalla | Markazi Text | Calibrated |
| Arabic under Arial / Times New Roman / Cambria / Aptos | Lateef on the Latin alias | Arabic-only unicode-range face |
| Arabic under Calibri / Calibri Light | Markazi Text on the Latin alias | Arabic-only unicode-range face |

<a id="note-45"></a>

### 45. Font fidelity: a delivered substitute is actually used, waited for, and reported (Sep 2026)

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

<a id="note-46"></a>

### 46. Arabic calibration measured the way it paints: through the alias, unrounded (Sep 2026)

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

<a id="note-47"></a>

### 47. Paragraph spacing in lines: what a "line" is in Word and PowerPoint (Sep 2026)

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

<a id="note-49"></a>

### 49. Aptos: calibrated look-alikes measured against the real cloud fonts (Sep 2026)

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

<a id="note-50"></a>

### 50. Delivered fonts load when painted, not up front (Sep 2026)

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

<a id="note-51"></a>

### 51. The no-Office-fonts review: substitutes painted where they matter (Sep 2026)

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

<a id="note-52"></a>

### 52. Wingdings and Symbol bullets paint their Unicode shape (Sep 2026)

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

<a id="note-53"></a>

### 53. `aared-fonts` frozen and unpublished: the browser's own fonts are the default (Sep 2026)

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


## Search and highlighting

**Current search scope, navigation and paint:**
[ARCHITECTURE](ARCHITECTURE.md#search-and-highlights).
Source results, painted occurrences and title lookup are different concepts;
the entries below explain why their contracts diverged.

<a id="note-69"></a>

### 69. The viewer searches its own document; building from the contents stays low level (Oct 2026)

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

<a id="note-78"></a>

### 78. Search matches the Arabic spellings of one word (Oct 2026)

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

<a id="note-79"></a>

### 79. XLSX search covers one sheet, as Excel's Find does (Oct 2026)

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

<a id="note-83"></a>

### 83. DOCX header and footer search (Oct 2026)

DOCX search now includes resolved header/footer stories, deduplicated by source
part so inherited copies produce one result. The render/preview controls can
highlight and navigate to every displayed occurrence, including repeated
headers, using native CSS Custom Highlights rather than wrappers that alter
layout. PPTX/XLSX were unchanged at this point; hidden header variants remain
searchable but cannot be navigated to when not rendered. Section 84 later
replaced these DOCX-only highlight controls with shared `viewer.search()`.

<a id="note-84"></a>

### 84. Viewer search owns highlighting across formats (Oct 2026)

Section 83's DOCX-specific highlight controls duplicated search state and are
removed. `viewer.search(query, { highlight, signal })` owns highlighting across
PPTX, XLSX and DOCX; highlighting is on by default, `highlight: false` clears
presentation state, and an empty query clears it. Headless document/sheet search
remains data-only.

The important boundary was shared viewer state rather than DOCX-only controls:
one app-owned search panel should not lose its query at a slide or virtualized
row replacement, nor should a failed replacement destroy useful prior paint.
This required no new Worker/model protocol. Rendered-range checks covered the
shared routes, including displayed page fields rather than just saved text.
The current controller/lifetime/clear rules live in the topic's architecture link.

<a id="note-85"></a>

### 85. PPTX search results count text occurrences (Oct 2026)

PPTX search now returns every non-overlapping matching occurrence, including
repeats within one paragraph, rather than only the first match per slide. It
retains the `{ slide, text }` result shape and existing text index; no extra
Worker reads or model/RPC changes are needed. XLSX remains per-cell and DOCX per
source block. Fixture and rendered-highlight checks pin result/highlight parity.


## Charts

**Current algorithms and type boundaries:** [ARCHITECTURE](ARCHITECTURE.md#charts).
These entries retain which rules were guessed, which native cases tested them,
what failed and why an intentional difference was kept. That evidence is not
a second algorithm table or a guarantee for an unscored host/font route.

<a id="note-71"></a>

### 71. Charts: a real chart model, read in the Worker, laid out on the page (Oct 2026)

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

<a id="note-73"></a>

### 73. Charts: the reader-visible gaps closed, every type but maps and surface drawn (Oct 2026)

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

<a id="note-74"></a>

### 74. Charts: the assumed rules checked against Office, and corrected (Oct 2026)

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


## XLSX layout and conditional formats

**Current behavior and bounded evaluator semantics:**
[ARCHITECTURE](ARCHITECTURE.md#xlsx-grid-selection-and-conditional-formatting).
COM/PDF values, unsupported experiments and cost studies below remain evidence;
their once-open status is historical, not today's backlog.

<a id="note-105"></a>

### 105. XLSX conditional formats: top/bottom and above-average evaluated; Excel's behaviour measured first (October 10, 2026)

**The problem.** `top10` and `aboveAverage` rules were reported in
`SheetModel.unsupported` and never painted, and a formula bound was reported
even when it named one cell. The ROADMAP's XLSX conditional-formats item asked
for Excel's behaviour to be measured before any matcher was written, because
the spec and Office diverge on tie-breaking and rounding.

**What Excel does (measured through COM on a scratch workbook, 16.0).** Each rule
was applied to its own column, painted cells were read back from
`DisplayFormat`, and the XML Excel wrote was read from the saved package.

- *top10, count:* the N-th largest (or, with `bottom`, smallest) number is the
  threshold, and every cell tied with it is painted, so a rule can paint more
  than N cells. Ten numbers `10, 9, 8, 8, 8, 7, ...` with rank 3 paint 10, 9,
  8, 8, 8.
- *top10, percent:* the count is `floor(numeric count x percent / 100)`, never
  below one: 1.5, 2.5, 2.9 and 3.5 of ten give 1, 2, 2 and 3; 0.5 gives 1.
  Blanks and text are not counted in the denominator (50% of ten numbers plus a
  blank and a text cell paints five).
- *aboveAverage:* the mean is over numbers only (text and blanks do not count as
  zero). `aboveAverage="0"` means below; the default is above. The comparison is
  strict, and `equalAverage="1"` admits cells equal to the mean.
- *Rank and percent in the XML:* `rank` carries the N, `percent="1"` says it is a
  share, and `bottom="1"` takes the smallest. No formula is written for `top10`.
- *Not measured:* a standard deviation. The `AboveAverage` object in this Excel
  build has no `StdDev` property, so the variant cannot be authored through COM.
  It stays reported, and its XML is not pinned.
- *Not implemented, measured for later:* `timePeriod` follows the clock (its
  formula floors to the date, so time of day is ignored; weeks start on Sunday;
  `last7Days` includes today; months are calendar months). `iconSet` families
  could not be selected through this COM binding, so only the default
  three-symbol shape was seen; the families are not measured.

**What changed.** `top10` (count and percent, top and bottom) and `aboveAverage`
(above, below, and their equal variants) are evaluated from cached numbers, in
`core/xlsx/sheet.ts`, beside the existing cell-value rules. Both paint through
the existing differential fill and font path, and both respect `stopIfTrue` and
priority order. A `formula` colour-scale or data-bar bound now resolves when it
is a number or one same-sheet cell holding a number (`$F$5`). Anything that needs
a formula engine, such as `PERCENTILE(...)`, a sheet-qualified reference or an
empty cell, is still reported and left unpainted. A standard-deviation
`aboveAverage` rule is reported, not guessed.

**Blast radius.** Conditional formats are read in the Worker for every XLSX
sheet, so the changed parse path touches every workbook that carries them. No
resolved-model field changed, the renderer is unchanged, and no public API
moved. The formula-bound resolver is shared by colour scales and data bars,
so both gain the single-cell case together. `numericValuesIn` and the
differential path are reused, not copied.

**Deliberate non-changes and unmeasured cases.** A rank below one paints
nothing: its meaning is not measured. A rule whose range is a union of several
references counts the union; only single-range rules were measured. A blank
referenced cell does not resolve a formula bound; Excel's treatment of it was
not measured. `timePeriod`, `iconSet`, `expression` rules and formula-engine
bounds stay reported, and their decisions remain open in ROADMAP.

**Proof.** `xlsx-conditional.test.ts` re-reads parsed cells from the Shot Log
(93 numbers) for top and bottom rules, the percent rounding and minimum,
aboveAverage with its four variants, the single-cell bound, the reported
standard-deviation rule, and the formula-engine bound that stays unpainted. The
new tests were run before the matchers existed and failed for the expected
reason. The rules are Excel-shaped XML; the coverage fixture is not yet
COM-authored with these rules, which remains open in ROADMAP.

<a id="note-106"></a>

### 106. XLSX data bars paint Excel's x14 gradient and default bar lengths (October 10, 2026)

**Root cause.** The base `<dataBar>` parser read only its bounds and colour.
Excel stores the Office 2010+ appearance in a worksheet-level x14 conditional
format block, linked back to the base rule by `x14:id` / `cfRule/@id`. The old
parser never followed that link. The renderer then painted one flat fill with
`opacity: 0.45`, which was neither the Office gradient nor its leading-edge
colour.

**What Excel does (measured).** A read-only COM probe of `fixtures/xlsx-coverage.xlsx`
with Excel 16.0 read the Shot Log's `H5:H16` rule as
`BarFillType=xlDataBarFillGradient`, `PercentMin=10`, `PercentMax=90`,
`ShowValue=True`, `AxisPosition=xlDataBarAxisNone`, and `Direction=xlContext`.
The linked worksheet x14 rule sets `axisPosition="none"`, bounds 0 and 20, and
a negative fill colour; it omits `gradient`, lengths, `showValue`, border and
direction. Excel's COM values establish the omitted defaults for this real
rule: gradient, 10%/90% length, values shown, context direction and no border.
The Microsoft [MS-XLSX `CT_DataBar` definition](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/88c5ca0a-3a34-4fad-8c91-fdb4550d3fe1)
and [Excel `BarFillType` documentation](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.barfilltype)
agree that gradient is Excel's default. The paired native `H4:I16` capture
shows an opaque teal leading edge fading to transparent over the cell fill.
For value 6.5 of 20, x14's 10/90 lengths give a 36% bar, not the old 32.5%.

**What changed.** The Worker indexes worksheet-level x14 data-bar rules by
their `cfRule/@id` and resolves the base rule's nested `x14:id` to that entry.
The cell overlay now includes gradient, value visibility, direction, border,
axis and negative-colour metadata. Missing x14 lengths use Excel's observed
10/90 defaults; non-x14 base rules retain their standard 0/100 lengths. The
bar ratio includes those lengths. The DOM renderer paints the default gradient
as the authored colour fading to transparent, removes the artificial 45%
opacity, honors solid/gradient, `showValue`, border and direction settings,
and keeps the existing cell-relative positioning and text-over-bar stacking.

**Related routes and shared logic.** This is XLSX conditional-format data, so
there is no equivalent conditional data-bar route in DOCX or PPTX. PPTX, DOCX
text boxes, and charts use the shared `fill-paint.ts` SVG painter for
DrawingML-authored fills. A data bar is instead a per-cell HTML overlay with a
value-dependent width and a fade to transparency; routing it through that SVG
shape-fill contract would conflate different model semantics. The existing
shared renderer behavior reused here is the grid's direction, cell text
layering, and conditional-rule order. No DOCX/PPTX code changed.

**Deliberate limit.** The coverage range has no negative values and
`axisPosition="none"`. The parser preserves x14 axis/negative-colour metadata,
but negative-value bars and an explicit middle axis need a zero-axis layout
that this change does not implement. To avoid painting them as positive bars,
the reader reports the rule and leaves it unpainted when the range contains a
negative value or requests a middle axis. The official x14 schema also exposes
same-as-positive flags for negative fill and border; those matter when negative
bars are implemented. DOCX/PPTX have no parallel feature to cover here.

**Proof.** A fixture-backed core test re-reads `H5` and the zero-valued cell's
resolved overlays, including the linked x14 defaults; an extension mutation
proves an explicit solid-fill setting overrides the default. The rendered DOM
test re-reads bar width, gradient, and opacity from Excel's authored tracker
fixture. The negative-range test verifies unsupported negative bars are
reported rather than mispainted. The paired capture of `H4:I16` was reviewed
against Excel's native page.

<a id="note-107"></a>

### 107. XLSX conditional formats: icons, date periods and standard deviations painted; native behaviour measured first (October 10, 2026)

**The problem.** `iconSet`, `timePeriod` and standard-deviation
`aboveAverage` rules were reported in `SheetModel.unsupported` and never
painted. The ROADMAP asked for Excel's behaviour to be measured before any
matcher was written, and for the icon cost to be measured before committing
to a full catalogue.

**What Excel does (measured through COM, 16.0).** Probes live in
`tools/probe/xlsx-conditional-path/` (deleted on completion; the findings
below are the record).

- *Standard deviation is population, strict.* Values
  `10,9,8,8,8,7,6,5,4,3` with one deviation above paint 10,9 only, and one
  below paint 4,3 only. That is mean +/- population stdev (8.94), not sample
  stdev (9.05). `AboveBelow` 4/5 select above/below deviation with
  `NumStdDev`; the XML is `<cfRule type="aboveAverage" stdDev="1"/>` (above)
  with `aboveAverage="0"` for below. `equalAverage` admits equality.
- *Date periods need the date operator.* `FormatConditions.Add` creates them
  as type 11 with the period as the 7th argument (`DateOperator` 0-9), not
  the 2nd: `Add(11,0,"","","","",op)`. Periods 0-9 are today, yesterday,
  tomorrow, last7Days, thisMonth, lastMonth, nextMonth, thisWeek, lastWeek,
  nextWeek. There is no `next7Days`: `ST_TimePeriod` holds ten values, and
  Excel writes a `<formula>` child (`FLOOR(A1,1)=TODAY()` for today), so the
  ROADMAP's "no formula child" and "eleven values" were wrong.
- *Icon families are all reachable.* `Workbook.IconSets` holds 20 entries:
  IDs 1-17 are the base families (ID 4 omits `iconSet`, defaulting to
  `3TrafficLights1`), IDs 18-20 are the x14-only `3Stars`, `3Triangles`,
  `5Boxes` (written as standalone `<x14:conditionalFormatting>` with
  `<x14:cfvo><xm:f>`, not linked by ID). `NoIcons` is not selectable. Every
  `IconCriteria` reads `Operator 7` (GreaterEqual), so thresholds are
  inclusive and the first `cfvo` (always 0) is ignored. `showValue` defaults
  true, `reverse` false, and `reverse` is ignored when `custom="1"`.
- *Corrections to the earlier record.* `serialToDate()` already existed with
  both date systems; `NumStdDev` and `Workbook.IconSets` are documented and
  reachable (the earlier COM failure was the wrong `Add` signature, not a
  missing capability); a `timePeriod` rule does carry a formula.

**What changed.** `core/xlsx/sheet.ts` parses base `<iconSet>` and
standalone x14 icon rules, `timePeriod` attributes and `stdDev`; `parseXlsxSheet`
takes an optional pinned `XlsxToday` (defaulting to the reader's local date,
one value per parse so sheets agree). Thresholds reuse `scaleBound`;
timePeriod matches only date-formatted numeric cells via the existing
serial helpers. The cell overlay gains `conditional.icon` (`family`, tier
`icon`, `showValue`); per-tier `NoIcons` paints nothing. The XLSX renderer
(`xlsx/icons.ts`, `xlsx/index.ts`) paints a 16px SVG at the inline start
with shared geometry, honouring `showValue` and keeping arrow directions
unmirrored in RTL. `fixtures/xlsx-coverage.xlsx` gains a COM-authored
`CF Rules` sheet (icons, today, 1-deviation); the manifest and
`fixtures/README.md` name it.

**Icon cost (measured, not estimated).** Full 20-family catalogue vs the
pre-change build: Worker `worker.mjs` 202.31 kB to 206.68 kB (+4.37 kB,
gzip +1.32 kB); XLSX chunk 40.00 kB to 43.63 kB (+3.63 kB, gzip +1.30 kB).
`xlsx/icons.ts` is 6.9 kB source sharing arrows, circles, flags, diamonds,
checks, stars, quarters and boxes. The glyphs are drawn in the package, not
extracted from Excel (a conditional-format rule names a family; the file
carries no artwork): the arrow tiers are single filled triangles, while
Excel paints shafted arrows — directionally and chromatically right, visibly
different side-by-side, and recorded as accepted in ARCHITECTURE Known
limits. A 3-family canonical set would save
roughly 2 kB minified (~0.7 kB gzip) but needs a substitution policy,
mapping documentation and its own review; preserving the authored families
keeps directional meaning (an arrow is not a rating) and reuses the same
catalogue for custom sets. Runtime cost is one 16px SVG per painted cell;
virtualization bounds how many paint. Recommendation: keep all families.

**Conditional-formula evaluator (cost study, no engine built).** `expression`
rules and complex `formula` bounds still report rather than guess. A bounded
evaluator over cached values (comparisons, arithmetic, boolean/text helpers,
relative/absolute and sheet-qualified cached references, blank/error
coercion, `PERCENTILE`-style aggregations, no named ranges/tables/volatile
functions, clock injected like `timePeriod`) would be a second expression
language beside the number-format parser, carried by every consumer in the
flat Worker bundle with its own parse surface, fuzz surface and maintenance.
Estimated weight is single-digit kB minified for a tight subset, growing
with each function, plus per-cell evaluation over whole ranges on every
parse. It would earn its keep only against real `expression` rules and
function bounds in the wild, which are not yet surveyed. Two constraints any
such evaluator must respect: it must never be *wrong* (aared does not guess —
a wrong colour is worse than none), and a bounded subset is what makes it
tractable. Recommendation: defer the build until such a survey exists; the
notices stay explicit rather than guessing.

**Blast radius.** Conditional parsing runs in the Worker for every XLSX
sheet. No public API moved except one optional `today` parameter and one
optional `conditional.icon` field; the renderer change is XLSX-only (one new
`xlsx/icons.ts` module, icon-aware emptiness/spill/positioning). DOCX/PPTX
are untouched. The fixture gains one sheet; existing sheets, extents and
captures are unchanged.

**Deliberate non-changes.** Negative data bars and middle axes stay reported
(§106). `NoIcons` as a whole-set family paints nothing. Malformed custom
sets and engine-needing bounds report rather than guess. Icon artwork is
recognisable normal-use SVG, not pixel-traced glyphs (ARCHITECTURE Known
limits). `timePeriod` follows the reader's clock by design, like Excel.

**Proof.** `xlsx-conditional.test.ts` pins stdDev thresholds, pinned-today
date matching (and that Hours numbers never match), icon tiers/inclusivity,
`reverse`/`showValue` and the reported formula bound, plus an Excel-authored
`CF Rules` sheet test (icons across tiers, stdDev painting 9 and 10, date
rule evaluated but missing today). The visual DOM test re-reads the icon
overlay, SVG and text padding. Typecheck, build (with the Latin-1 dist
test) and the XLSX suites gate the change; paired Excel captures cover the
new sheet.

<a id="note-109"></a>

### 109. Three-colour scales have Excel-authored, paired A4 proof (October 10, 2026)

**The gap.** The three-stop colour-scale matcher had synthetic XML coverage
over the real Shot Log numbers, but no native-authored case or paired Excel
capture. This was a verification gap, not a missing renderer capability.

**Native evidence.** Excel 16.0 COM `FormatConditions.AddColorScale(3)`
authored `CF Rules!E2:E12` in the existing `xlsx-coverage.xlsx`: eleven values
0 through 100, min / percentile 50 / max, red `F8696B`, yellow `FFEB84`,
green `63BE7B`. Excel saved those exact three `cfvo` bounds and `color`
children in the worksheet. After recalculation, `DisplayFormat.Interior.Color`
gave `F8696B`, `F98370`, `FA9D75`, `FCB77A`, `FDD17F`, `FFEB84`, `E0E383`,
`C1DA81`, `A2D07F`, `83C77D`, `63BE7B`. Aared's three stops match exactly;
interpolated channels differ by at most one level, accepted as visually
negligible rounding in ARCHITECTURE Known limits. Both full-sheet and curated
part captures show the same red-to-yellow-to-green progression. The native
print header and icon artwork differences remain the already documented
limits, unrelated to colour scales.

**Repeatable proof.** `xlsx-conditional.test.ts` compares every resolved
numeric cell against the measured native colours, with exact assertions for
the stops. `xlsx-capture.visual.test.ts` re-reads all eleven cells' text and
computed background colours from the DOM and includes the `three-colour-scale`
part (`A1:E12`; curated parts start at column A). Reproduce the review with:

```bash
bun tools/xlsx-capture.ts fixtures/xlsx-coverage.xlsx --part 5:three-colour-scale=A1:E12 --pair-full
```

The review is `tools/excel-render/out/xlsx-coverage/review.html`; the readable
pair is `pair-sheet-5-part-three-colour-scale.png`. These are generated,
gitignored evidence; the native fixture, measured colour assertions and
capture plan are the durable record. The authoring scratch directory
`tools/probe/xlsx-three-colour-scale/` was deleted after codification.

**Blast radius.** Fixture coverage, the conditional parser test and the
existing paired-capture plan only. No parser, renderer, public API or package
weight changed. Earlier columns and their rules are retained; this
SpreadsheetML-only capability has no DOCX/PPTX route. Typecheck, the full
unit suite and the targeted native/DOM capture suite passed. The closed
verification item was removed from ROADMAP; formula-engine work remains open.

<a id="note-110"></a>

### 110. Bounded conditional-formula evaluation over cached values (October 11, 2026)

**The gap.** `expression` rules and formula bounds beyond a bare number or
single-cell reference reported instead of painting. Excel stores only the
condition, so there is nothing to read without evaluating the formula.

**Native evidence (Excel 16.0 COM, `tools/probe/cf-formulas/`, deleted after
codification).** Relative references shift per target from the first range's
start (`$B2` on `A2:A6` paints per-row; `A2>5` on `B2:B4` paints only B2;
`cellIs >A2` on `6,1,6` vs `10,0,5` paints the lower two). Blank equals `""`,
`0` and `FALSE` but not `TRUE`; `""` equals `""` but not `0`. `-2^2` is `4`,
`10%` is `0.1`, `MOD(-3,2)` is `1`. `AND("abc")`, `IF("",1,2)` and
`SUM("a")` are `#VALUE!` while `"TRUE"`/`"FALSE"` coerce in `AND`/`OR`/`SUM`;
`IF(FALSE,1)` is `FALSE`; `SUM(TRUE)` and `SUM("1")` are `1`; out-of-range
`PERCENTILE` is `#NUM!`. Cross-sheet `expression` lives in x14 with
`<xm:f>`/`<xm:sqref>` and inline `dxf`, not base `<formula>`/sqref.

**What changed.** New Worker-only `core/xlsx/formula.ts` (hand-rolled,
Latin-1, zero DOM): tokenizer/parser for numbers, strings, booleans, error
literals, A1 refs/ranges with `$` axes and quoted sheets, the 22-function
allowlist and `+ - * / ^ & = <> < > <= >= %` with Excel precedence
(unary tighter than `^`, so `-2^2` is `4`). Evaluation reads cached `<v>`
only; a formula cell with no cache is unknown, never blank or zero.
`sheet.ts` compiles once per distinct text (bounds, operands, expressions),
evaluates `expression` per authored cell plus matching blanks in small ranges
(at most 50k cells per rule), evaluates `cellIs` operands per cell with the
same shift, and resolves `formula` bounds once at the anchor. Errors never
match for `expression`; `IFERROR` never conceals unknown/unsupported/budget.
Whole-rule failure (unsupported syntax, cross-sheet without a reader,
missing cache, budget) leaves the rule unpainted and reported, never partial.
x14 non-icon, non-dataBar rules now report instead of silently dropping;
linked x14 dataBar extensions stay silent as before.

**Fixture.** `CF Rules!F2:F12` gains COM-authored `=MOD($F2,2)=0` over `1..11`,
painting `2,4,6,8,10` (native `DisplayFormat` readback `[3,5,7,9,11]`);
`fixtures/README.md` and the `A1:F12` paired part record it.

**Cost.** No dependency. Worker growth to be measured on build; the subset is
deliberately small (parser plus typed evaluator, no recalculation graph).
Runtime compiles once per distinct formula and memoizes nothing across parses;
aggregates scan authored numbers only, and the 5M-visit budget with 2048-node
and 64-depth caps bounds hostile input. Uncached search still parses the full
sheet including CF; small `read(range)` still evaluates before slicing.

**Blast radius.** XLSX Worker parse path only; PPTX/DOCX untouched. No public
API change (optional `conditional` overlays as before, plus `expression`
now painting). Renderer unchanged (re-reads fills as before). Existing
literal `cellIs`/single-cell bounds keep their behavior; `PERCENTILE` bounds
and `F5>10`-style expressions now paint. Typecheck, full unit suite and the
targeted native/DOM capture suite passed.

**Deliberate non-changes.** Cross-sheet values, x14 inline `dxf`, names,
tables, arrays/spills, lookups, UDFs and volatile functions report.
Large blank-sensitive ranges stay authored-only. Full recalculation stays
permanently out of scope (view-only).

<a id="note-111"></a>

### 111. Closing the conditional-formula gap: dates, text search, conditional aggregates, x14 and cross-sheet (October 11, 2026)

**The gap.** §110 left the normal-use middle reporting: `TODAY()` overdue
rules, `COUNTIF` duplicate highlighting, `ISNUMBER(SEARCH(...))` text rules,
and every cross-sheet formula (stored as x14, never base XML).

**Native evidence (Excel 16.0 COM, `tools/probe/cf-engine-close/`, deleted
after codification).** `TODAY()` is the date serial with
`TODAY()=INT(NOW())`; `DATE` rolls month/day overflow, adds 1900 to years
0-1899, keeps fictitious 1900-02-29 (serial 60), and is `#NUM!` below serial
0; `YEAR(-5)` is `#NUM!` while `WEEKDAY(0)` is 7 and `DAY(0)` is 0;
`WEEKDAY(x,99)` is `#NUM!`. `SEARCH` is case-insensitive with `*`/`?`/`~`
wildcards, `FIND` is case-sensitive and literal, a miss is `#VALUE!`;
`MID` is 1-based (start 0 `#VALUE!`, length 0 and past-the-end `""`).
`CONCAT` joins ranges, `CONCATENATE` with a range spills. `COUNTIF` is
case-insensitive with wildcards: `"5"` matches 5 and text `"5"`, `"TRUE"`
matches `TRUE`, `"#DIV/0!"` matches that error, `""` matches blanks and
zero-length strings while `"="` matches only blanks and `"<>"` everything
but blanks; a blank criteria matches nothing; `SUMIF` aligns a short sum
range top-left and `AVERAGEIF` with no match is `#DIV/0!`. A COM-added
cross-sheet rule is accepted, paints, and lands only in x14 (`<xm:f>`,
`<xm:sqref>`, inline `<dxf>` with a bare `<bgColor>` fill) — base XML never
carries it. Same-sheet `expression` rules stay in base `<formula>`/sqref.

**What changed.** `formula.ts` gains the date/text/aggregate subset
(`TODAY`, `NOW` as the parse-time snapshot, `DATE`/`YEAR`/`MONTH`/`DAY`,
`WEEKDAY` kinds 1-3, `SEARCH`/`FIND`/`MID`/`CONCAT`/`CONCATENATE`
scalar-only, `COUNTIF`/`COUNTIFS`/`SUMIF`/`SUMIFS`/`AVERAGEIF`/`AVERAGEIFS`)
with a criteria compiler (operators, wildcards with star/length caps,
Excel's blank rules) and positional multi-range alignment
(`COUNTIFS`/`AVERAGEIFS` size mismatch is `#VALUE!`, `SUMIF` overlays).
`sheet.ts` parses x14 `expression` rules (inline style via
`XlsxStyles.inlineDxf`, priority-merged with base rules) and preloads every
referenced sheet's values-only grid before evaluation, so `evaluate` stays
sync and CF can never recurse. `COUNTIF`-family blanks stay arithmetic;
expression/cellIs blank materialization keeps its 50k cap.

**Fixture.** `CF Rules!G2:G13` frozen 2020/2099 serials under `=$G2<TODAY()`
(paints rows 2-7); `H2:H13` fruit under
`=COUNTIF($H$2:$H$13,H2)>1` (apple x4, banana x3); `I2:I9` text under
`=ISNUMBER(SEARCH("apple",$I2))` (5 hits); `J2:J13` hundreds under the x14
`=$J2>Notes!$B$6` rate rule (500 and up red). Native `DisplayFormat`
readbacks `[2-7]`, `[2,3,4,6,8,11,13]`, `[2,3,5,7,9]`, `[6-13]`; the paired
A4 part is now `A1:J13`.

**Cost.** No dependency. Worker `core/worker.mjs` 208.8 kB / 70.3 kB gzip
(§110) to 247.8 kB / 79.4 kB gzip (+39.0 kB raw, +9.2 kB gzip): the
criteria/date/text surface. Flat bundle, so PPTX/DOCX download it too, as
before. Runtime compiles once per distinct formula; range scans touch
authored cells only with blanks as arithmetic, and scans charge the 5M-visit
budget. Search parses keep their full-sheet behavior; sheets with
cross-sheet CF preload each referenced sheet's values once per parse.

**Blast radius.** XLSX Worker parse path only; PPTX/DOCX untouched. No public
API change (`XlsxStyles.inlineDxf` is internal). Renderer unchanged.
Existing literal `cellIs`, single-cell bounds and §110 expressions keep
their behavior. Typecheck, full unit suite (46 conditional tests), the
targeted native/DOM capture suite, `shipped-latin1` and
`check-consumer-build` passed.

**Deliberate non-changes.** Names, table refs, arrays/spills, lookups,
`TEXT`, unlisted functions, exotic `WEEKDAY` kinds and cross-workbook refs
report (Known limits). Full recalculation stays permanently out of scope.
The ROADMAP conditional-formats item is deleted: the subset above is the
whole contract.

<a id="note-115"></a>

### 115. Spreadsheet viewer chrome: formula bar, range select and inline aggregates (October 11, 2026)

**The gap.** Aared's XLSX grid had a single-cell cursor with keyboard moves
but no range, and no Excel-like chrome: nothing named the cursor or showed the
cell's formula. A competitor recording showed a name box, an `fx` bar with the
raw formula, and live aggregates over a drag-selected range.

**What changed.** All viewer-side, off by default, no Worker or model change:
`appearance.spreadsheet.formulaBar` paints a read-only name box plus the
cursor cell's formula (else its formatted value), and a Sum/Average/Min/Max/
Count dropdown inside the same bar for a selected range. Aggregates inspect
cached cell values only; formulas are never evaluated, and Count counts
nonempty cells. There is no bottom status/footer bar. `Shift`+arrows extends
from the anchor and dragging extends from the press point; plain moves, clicks
and sheet switches re-anchor or collapse. The wash is one class-sheet rule
over a per-theme accent variable with `!important` (authored fills are inline,
so a plain rule would lose), the cursor inside a range paints paper-white, and
both ride the existing selection-restore funnel so scrolling repaints keep them.
When a formula cell is active, visible same-sheet A1 precedents (including
ranges) receive a reference wash after quoted string literals are removed
from the scan; explicitly different-sheet references are left alone because
their cells are not on screen. The reference wash is reapplied with range
selection after windowed scroll repaints.
The `xlsx-coverage.xlsx` `Summary` sheet was reauthored through Excel COM to
match the bilingual sales data and formulas in `formula-bar-sheet.html`; its
`E` and `F` formulas give the real Office-authored fixture same-sheet
precedents for the visual test, while all displayed results remain Excel's
cached values. Its light palette is authored in Excel too: warm gray grid
lines and HTML-matching green/red conditional font colors on positive/negative
change values. This is source worksheet styling, not a viewer theme change.
`selectionchange` gains an optional `range: { start, end } | null`, which the
facade forwards untouched; `render()` gains a `range` getter. Display-only
throughout: the name box and formula text are text, not inputs, and there is
no fill handle, no auto-scroll past the window edge, and no editing.

**Cost.** XLSX entry +5.2 kB raw / +1.3 kB gzip (chrome lives in the entry;
splitting it out would need viewer internals, so it stays). No other entry
moves; defaults keep every existing render byte-identical.

**Proof.** `xlsx-chrome.visual.test.ts` re-reads the DOM on the real tracker
Dashboard and coverage workbook: cached formula source, same-sheet precedent
wash, no foreign-sheet wash, `B12`→`C12:C14` extension, aggregate menu values
and metric switching, range wash membership, collapse on click, drag with
release-click retention, menu fit and stacking over sticky headings, and
absent chrome by default. Typecheck, full unit suite, snippets/parity and the
API reference (appearance, both `selectionchange` rows) gate it.

<a id="note-116"></a>

### 116. XLSX: two Excel differences kept by choice, missing fonts measured, centred overflow fixed, the Summary redesigned (October 11, 2026)

**What the review found.** The coverage workbook's Summary sheet, printed by
Excel next to aared's render, showed three differences.

1. **General-aligned text.** Excel places a General string by its own first
   strong letter: in the right-to-left Summary, "SKU", "Q2", "Status",
   "HP-204" and "Wireless Charger" sat left while Arabic in the same columns
   sat right, and "Q3" ran into "التغيير" across the D|E border. aared puts
   every General string at the sheet's start edge (`textAlign`): all right.
   Only the right-to-left case was read off Excel; Arabic-first General text
   in a left-to-right sheet is unprobed. A fix would compute the side from
   the first strong character in script, not CSS `start`, because the spill
   overlay is `dir="ltr"` and reads the resolved side.
2. **Shared borders.** Each cell paints its own four CSS borders with
   `border-collapse: separate`, so an edge both neighbours author (Excel's
   All Borders writes it on each) draws as two adjacent 1px lines, 2px
   against Excel's one, while the range's outer edge is 1px. Measured in the
   DOM (adjacent cells' borders at x=733 and x=734); it applies to every
   bordered XLSX range. PPTX tables had the same class of fault and stroke
   each shared edge once (§30).

**Decision on 1 and 2: kept, by the maintainer's choice.** Both are recorded
in ARCHITECTURE § Known limits and marked at `textAlign` and at the border
painting in `paintCell`, so neither is "fixed" by accident. A sheet that
states its alignment and draws single-sided rules is unaffected, which is how
the redesigned Summary below is built.

3. **"Arabic looks smaller."** Not a renderer fault. The sheet's font, IBM
   Plex Sans Arabic, is not installed on the machine. Excel's PDF export of a
   cell set in it, and of one set in a made-up name, drew every glyph, Latin
   and Arabic, in ArialMT at Arial's widths. The browser's `sans-serif` on
   Windows is Arial as well: the same strings measured within 1% of Excel's
   widths (`الطلب ثابت تقريبًا` 5.58 em in Excel, 5.62 em in aared; "Wireless
   Charger" 7.74 em against 7.72 em). Arial's Arabic is drawn small next to
   its Latin in both. Enlarging it would move aared away from Excel, so the
   renderer is unchanged; the workbook's font is what changes (below).

**Centred overflow, fixed.** The redesign exposed it: the total row's status,
centred in F9 beside an empty G9 and a filled E9, painted inside G. Unwrapped
text that may run into blank neighbours gets an overlay across the run, and
centred text was centred on the run, not on its cell, so any centred text
with unequal blank space on its two sides moved, even text that fits. Excel's
PDF export of a centred cell settled the rule: a word that fits stays centred
in its own cell beside an empty neighbour; an overflowing line keeps the same
position whether its left or its right neighbour holds text, and is only cut
at that neighbour. The overlay now pads its shorter side by the difference
(`TextSpill.after`), so the centring box's middle is the cell's middle and the
run still clips. Left- and right-aligned spills are unchanged.

**The Summary redesigned through Excel COM** (one `extend` pass, label
stripped). Three rows were inserted above the table, so Excel shifted every
formula, conditional range and cached value itself; the other four sheets
parse to identical models before and after. Now: a merged two-row teal title
band (Arabic title, bilingual subtitle), the table at A4:G9 frozen under its
header, Segoe UI throughout (its Arabic is drawn at its Latin's size; SKUs in
Consolas), every cell explicitly aligned, rules as single top or bottom edges,
banded rows, signed change percentages in the existing green and red
conditional fonts, the status column as three `cellIs`-equal chips, a light
total row, a merged footnote at A11, gridlines off, and the window scrolled to
the top (it had been saved at `topLeftCell="A3"`, hiding row 2 under the
frozen header). Excel now writes the Summary's five differential styles ahead
of the others, so `xlsx-conditional.test.ts` finds the red and tan fills its
synthetic rules use by colour rather than by a fixed offset.

**Proof.** The paired Excel/aared full-sheet capture matches the redesign.
`xlsx-parser.test.ts` pins the layout, merges, explicit alignment, single
edges, chips and formulas; `xlsx-chrome.visual.test.ts` the moved precedents
(E5 reads C5 and D5, C9 reads C5:C8); `xlsx-render.visual.test.ts` re-reads
the merged title, the hidden gridlines and F9's text centred on F9 within a
pixel. Typecheck, unit and visual suites pass.

<a id="note-117"></a>

### 117. XLSX fit: the file's sizes, the browser's layout, fonts as residue (October 12, 2026)

**The question.** Several XLSX differences come down to "does this text fit?":
text running into empty cells, the gridlines it crosses, headers that wrap or
are cut. aared never draws text at Excel's exact width (a substitute or
generic face, and the browser's own rounding), so a rule built on computed
widths is wrong whenever the text nearly fills its cell, and must be redone
when a font loads or the zoom changes. A viewer does not need Excel's
widths to the pixel; it needs no gross mistakes.

**The rule.** Keep the file's sizes exactly (column widths, row heights,
merges, frozen panes), let the browser's layout of the painted text decide
fit, and record what fonts leave over as residue rather than fixing it by
resizing or measuring. Size changes would go through the sheet geometry
(charts, frozen offsets, virtualization and scroll-into-view all read it),
never through a cell growing in CSS.

**The inventory behind it.** Excel's print beside aared for all nine sheets of
the two XLSX fixtures, with and without Office's fonts, each difference
confirmed in the DOM. Row heights matched everywhere, so no geometry work was
needed. What a reader would see, and what was done:

1. **A negative percent read "(81.5)%"** where Excel shows "(81.5%)", under
   `0.0%;\(0.0%\);\-` (the tracker's budget sheet). Not a fit problem: the
   percent plan stripped `%` and appended it after the whole section. `%` now
   becomes a literal where the code writes it, so a fixed section keeps it in
   its prefix or suffix; scientific and fraction sections, which have no
   affixes, still append it. A unit test had pinned the wrong output.
2. **Cell padding.** aared set text 5px inside the cell box (a 1px border and
   4px of padding). Excel's PDF export, read with filled cells as the exact
   box, puts left- and right-aligned text 2.4-2.5px in for Arial 10 and
   Calibri 11 once its print scale is taken out (it prints an 89px column
   95.8px wide). The padding is now 1.5px (`CELL_PAD_PX`), shared by the cell,
   the spill overlay, the centring shift and the indent.
3. **Gridlines beside fitting text.** The spill pass hid the gridlines of
   every empty neighbour text might run into, measuring nothing, so the
   budget sheet's row 11 lost two lines beside a 48px word in a 110px cell.
   Which lines go is now read off the painted text: after each window paints,
   `settleSpillLines` reads every spilled text and its run's cells once, then
   hides only the edges the text reaches, leaving authored right borders
   alone, and runs again on `document.fonts`' `loadingdone`. A paper-coloured
   mask behind the text was weighed first and rejected: it would paint over
   the selection wash, a neighbour's own fill and authored borders.
4. **Font residue, accepted.** Aptos and Aptos Narrow text paints in the wider
   generic face, so CF Rules headers and Notes rows 18 and 20 wrap or cut
   earlier. A number too wide for its cell stays cut rather than Excel's `####`, which
   a wider face would show where Excel shows the number.

Two smaller findings stay open without a row: Excel shows "٥,١" where the file
holds "٥.١" (tracker budget sheet, column E), unexplained; and the capture
tool's part pairs can come out blank, with full-page pairs of a long sheet
unreadable.

**Proof.** `xlsx-numberfmt.test.ts` and `xlsx-parser.test.ts` pin "(81.5%)";
`xlsx-render.visual.test.ts` re-reads the budget row's lines (all kept beside
the fitting word) and the Dashboard's A16, whose long label hides A16's edge
and keeps B16's. The paired Excel captures were re-run.

<a id="note-118"></a>

### 118. XLSX selection past the data: a click, a drag, the arrow keys and Excel's plus pointer (October 12, 2026)

**The gap.** The empty area past the data is painted as inert cells, so a
click on it did nothing, the arrow keys stopped at the last row and column,
and the pointer over it was the default arrow. Excel selects any cell and
shows the plus pointer over the grid.

**The decision.** Empty cells select like data cells, numbered on from the
sheet's extent, and the cursor over every selectable cell is Excel's plus
(`cursor: cell`). Headings keep the default pointer. Cells past the data stay
out of the model: they hold nothing, so the viewer never needs their content.

**How it works.** Fill cells are not in the model, so a pointer event on one
is placed by geometry (`fillCellAt`). A fill row's track cell reads its own
column from the cell's start edge (the right edge in a right-to-left sheet),
and a whole-row fill cell walks the visible columns and then the fill columns
in the same order the row paints them. The selection past the data is drawn
as boxes in a layer over the canvas, on the offsets the drawings already use
(`paintSelection`), so it moves with the sheet's geometry and zoom. The range
wash is one box per row run, not one per cell, so a drag over the empty area
stays cheap. The cursor past the data is a white box with Excel's frame; the
cursor inside the extent keeps its own outline.

**Why not per-track cells.** Splitting the fill cells into one element per
column would make selection exact, but a viewport can hold hundreds of fill
columns on every row, so the painted area would multiply for a selection the
user may never make. The geometry hit test costs nothing when unused.

**Keyboard.** The arrow keys walk into the painted fill. Home, End and the page
keys stay on the data: Excel's Ctrl+End goes to the last used cell, and the fill
has no last cell of its own, so those keys keep to the rows and columns the
file has.

**The range wash no longer replaces fills.** Reviewing a selection over a
header, the wash (a solid `background-color` with `!important`) painted over
the header's dark fill and made its white text unreadable. The wash is now a
translucent inset shadow, which sits above the cell's fill and below its text:
the fill and the text keep their own colours, and only the tint changes. The
cursor cell inside a range keeps its own look and shows its frame, instead of
the paper-white it used to paint over its content. The formula-precedent wash
follows the same rule. A cursor cell with no fill of its own turns paper-white
again, as Excel does; a cell with a fill keeps it. While a range is selected,
the formula-reference outline is not shown: the range is the highlight, and an
outline on its cells read as a second range.

**Known limit.** Geometry-based selection overlays use natural sheet positions,
so the dashed perimeter and fill-area cursor/wash do not follow sticky frozen
rows or columns while scrolling; the cells and their tint still stick
(ARCHITECTURE § Known limits). Revisit with a range crossing a frozen pane.

**Proof.** `xlsx-selection.visual.test.ts` re-reads the DOM: the plus pointer on
data and empty cells and not on headings, a click naming the first empty
column (`I1`) and the next (`J1`), a drag from a data cell across the empty
columns (`B11:I13`) with its wash boxes, the arrow key into the fill column
(`H1` to `I1`), a click on an empty row past the extent (`A12` on the budget
sheet), and `ArrowDown` from the last data row into it.

**One dashed range boundary.** The per-cell wash remains translucent so
authored fills and text stay visible, but an extended selection now gets one
absolute dashed frame around the full selected rectangle, including the
viewport-fill cells. The cursor cell no longer adds its own solid outline inside
that range; a single-cell selection still has its original solid frame. The
visual test checks one dashed boundary both for a range within the sheet and
for one extending into empty columns, and confirms that selected data cells do
not keep competing outlines.

**Formula-reference ranges use the same treatment.** A formula's same-sheet
references keep their distinct per-range tint and color, but no longer put a
solid outline on each referenced cell. Each reference range gets one dashed
perimeter in its assigned color, so `SUM(C4:C8)` is shown as one rectangle
around C4:C8. The visual test checks the single-range geometry and the separate
colors for the two references in `D5/C5`.

**Selection appearance restored by request.** The dashed perimeter applies to
formula-reference ranges only. Extended cell selections again use the earlier
translucent per-cell wash with a solid outline on the active cell; the one-frame
selection experiment was removed. Formula references retain their one dashed
per-range perimeter and colors.


## Quality, tooling and documentation

**Current test/review structure:** [ARCHITECTURE](ARCHITECTURE.md#quality-strategy).
**Commands and artifacts:** [the tools catalog](../tools/README.md).
The entries here retain failure modes, safeguards that earned their keep and
dated proof boundaries, not a second command manual or current test-count ledger.

<a id="note-75"></a>

### 75. A whole-repo static review: shared failure classes and deliberate behavior (Oct 2026)

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

<a id="note-93"></a>

### 93. Tools consolidation: one Office export driver, one rasterizer, a standard probe shape (Oct 2026)

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

<a id="note-96"></a>

### 96. Probe cleanup and the label-strip record (Oct 2026)

The five probes behind §95 (`autofit`, `word-textbox-native`,
`word-textbox-fixture`, `word-textbox-layout`, `chart-txpr-autofit`) are
finished: their findings are codified in `core/docx/document.ts`,
`core/chart/read.ts`, `test/docx-text-box.test.ts`, `test/chart-autofit.test.ts`
and the ARCHITECTURE text-box and chart rows, so the probe directories are
deleted rather than left as dead weight (AGENTS.md, probe step 6). What replaced
the ignored scratch evidence is the permanent fixture, the semantic tests and
the paired native review.

<a id="note-123"></a>

### 123. One developer-tools entry point, shared review backends and checked tooling (October 13, 2026)

**Root cause:** tools had shared native export/rasterization already, but the
user-facing workflow still depended on knowing individual script names. DOCX
required four commands and repeated output directories, PPTX's aared capture
was embedded in a fixture-specific test, and the XLSX command/test copied
export types and disagreed on part filenames. Unchecked tooling also retained
36 broken imports after the font-calibration directory moved, plus two
machine-specific repository paths.

**Why a dispatcher over the existing backends.** The missing capability was a
usable workflow, not another Office/browser harness. Extracting PPTX capture,
making XLSX callable and composing DOCX phases let arbitrary files follow the
same evidence path as fixtures. Shared review chrome needed only report framing,
not a renderer framework. Side-effect-free validation matters because a bad
argument must not launch Office or replace an existing review. Caller-relative
input paths avoid surprising a command run outside the repository. Current
commands, options and prerequisites live in [tools/README.md](../tools/README.md).

**Evidence-safety findings.**
XLSX part slugging is common to native arguments, CLI captures and tests;
same-sheet filename collisions (including Windows case folding) fail early.
Filename-derived export folders reject `.`/`..` stems so cleanup can never
target the whole artifact root or its parent; missing input files also fail
before replacing previous captures.
Visible-sheet ordering agrees with Excel. Absent-font columns include XLSX
parts; DOCX's absent token file cannot overwrite the normal token evidence.
Page-count differences are evidence to display, not a reason to suppress a
Word review.

**Native finding:** the first real XLSX part review produced an all-white
aared capture while Excel showed the Arabic title band. The existing helper
expanded the scroller to 100,000px in both dimensions. The renderer correctly
anchors an RTL grid at that scroller's right edge, outside the A4 capture.
The old tests checked scale and PNG dimensions, not where the content painted.
Expand horizontally to the actual grid width instead; keep the tall viewport
for complete row materialization. The shared helper now checks the painted
grid's viewport intersection, and a real-fixture browser test reads visible
rows, direction, text and bounding rectangles for RTL and deep LTR bands.
This corrects capture chrome only; no spreadsheet renderer rule changes.

**Verification boundary:** tools now have strict no-emit checking (including
retained research, excluding ignored probes). Package tests can import shared
tool helpers across the package root; a separate source-only declaration
build config prevents tools/tests from emitting stray `.d.ts` files beside
their source. The checked in-process inspector uses the actual RPC result
types and includes picture/fallback media IDs. Failed/empty benchmark passes
retain diagnostic JSON but return failure; only compatible same-machine
results are merged. The one-shot group-effects benchmark writes under
`bench-results/`, not a native-review output root.

**Blast radius and deliberate non-changes:** developer tooling, review tests,
typecheck/declaration-build boundaries and their docs. No package public API,
OOXML parsing, Worker construction, renderer text/fill/effect/font rules,
calibration numbers, native COM adapters, rasterizer or benchmark methodology
changes. Chart-oracle vector scoring, page/Worker CDP routing, font research
and live probe helpers remain separate. No probe/evidence deletion or blanket
directory move; output roots and same-stem/per-export replacement semantics
are preserved. Remaining capture limits are documented in the catalog:
curated merge-clean A-column XLSX row bands, built DOCX renderer, explicit
network for absent fonts, and human rather than pixel-equality review.

**Proof:** CLI/report/benchmark unit tests (including unknown selectors,
outside-repo help and conflicting modes); responsive review-shell and real
RTL/deep-LTR DOM tests; native reviews of PPTX/XLSX/DOCX coverage fixtures,
XLSX cropped/full pairs and a spaced part name; absent-font PPTX/XLSX/DOCX
capture paths; focused native capture suites; strict workspace/tools
typecheck, package build, unit suite and Vite/webpack production probes.
Frozen calibration scripts were typechecked, not rerun. Performance runs and
published benchmark tables were not regenerated by this workflow cleanup.

<a id="note-124"></a>

### 124. Test cleanup preserves distinct routes and named fixture behavior (October 13, 2026)

**Root cause:** several assertions could pass without the claimed behavior:
DOCX search loops had no nonempty/context guard, mutation checks only rejected
sentinel strings, long-result checks allowed an empty or wrong prefix, and the
element's zoom check only required a positive value. Whole-document page and
drawing totals also made unrelated fixture additions break many tests. Some
proposed duplicates were not duplicates at all: public format entrypoints,
shadow-root styling, default versus explicit XLSX headers, and physical-right
RTL freeze offsets need their own proof.

**Decision:** strengthen the weak assertions, and delete or merge only proofs
with a verified surviving home. React hook exports and DOM-free import checks
share one test. Unread event fields and an error-identity assertion implied by
zero error events are removed. The element's explicit zoom step is checked
against both its synchronized attribute and painted scale; sheet tabs must
cover the loaded collection. Automatic source supersession holds a different
real document's response until the newer one paints and checks cancellation,
final format, content and load events instead of sampling the first load of
two identical files. The superseded automatic open is awaited to settlement,
and all nonprogress event types are checked, including absence of stale errors.

**DOCX fixture durability:** named body labels distinguish the real headings
from their TOC entries. The known cases retain consecutive placement, section
orientation and break causes, A4/Letter geometry, first-page versus continued
headers, missing first-page footers, PAGE/NUMPAGES values and footnote ownership.
An independent fixture-only XML oracle counts the explicitly authored section
and page breaks, preserving rejection of unexplained trailing pages without a
remembered total. A real-fixture mutation adds an explicit page and proves the
oracle and pagination both grow by one; it is not a general pagination oracle.
Anchored pictures are selected by paragraph labels/alt text rather than global
slices. Missing-id counts come from the XML mutation actually performed; the
unchanged resolved document and warning aggregation remain asserted. Zoom
invariance compares page counts across real scales, not a fixed fixture total.

**Deliberate non-changes:** native XLSX spacing/sticky/gutter constants and the
complete PPTX table-band proof stay. The scoped effect-family counts stay too:
they detect missing or duplicate examples, not the document's total drawings.
Public preview mismatch, closed-use, navigation and per-format route tests
stay; an internal opener is not an adequate replacement for a public entry.
Unrelated hostile-input and DOCX zoom cases are not collapsed just to reduce
the test count. Chart viewport dimensions, mixed-font exact pitch and bidi
percent placement still read actual SVG/DOM output, without pinning float
serialization, pseudo-element content or per-run line-height implementation.

**Blast radius:** tests and quality documentation only; no parser, renderer,
public API, fixture bytes or package weight changes. The earlier restored
translucent-background, non-1 numbering restart and axis-title proofs remain.
Worker-cache observability, positive ISBLANK materialization, further split-wrap
and font/borrowed-content verification, the PPTX capture aspect cross-check and
visual harness consolidation remain separate open work in ROADMAP.

**Safety review:** an independent reviewer checked removed assertions against
surviving proofs, found the initial trailing-page loophole and late-response
settlement weakness, and rechecked both corrections against the actual XML and
automatic-open path. Temporary, process-local fault injection (outside the
repository; no source edits) made the improved DOCX tests fail for missing
heading context, wrong truncation prefix, reordered repeat results and a
phantom trailing page.

**Proof:** package build and strict workspace/tools typecheck passed. The full
unit run passed (887 pass, 83 visual-gated skips, no failures); the visual-enabled
full run passed (969 pass, one skip, no failures), including native PPTX/XLSX
capture routes. No new Word fidelity claim or native Word comparison was made:
the renderer and fixture bytes are unchanged by this test cleanup.

<a id="note-125"></a>

### 125. Tools have task-owned source folders and one generated-output home (October 13, 2026)

**Root cause:** §123 improved the command experience but deliberately retained
the old physical layout. The maintainer still saw many scripts in `tools/`'s
root, benchmark code split across two locations, and generated data mixed into
native/chart source folders. A front door alone did not resolve that clutter.

**Why task-owned folders.** A command front door did not make the code's
responsibilities discoverable. Leaving only the entry point, catalog and
configuration at the root, and grouping implementation by task, addressed that
separate problem. Centralizing stable roots avoided a move-dependent cwd/path
bug. Imports, consumers and docs moved together rather than leaving old-path
wrappers to prolong two layouts. The current
[layout catalog](../tools/README.md#layout) owns the file/directory map; old
paths in dated entries remain provenance, not runnable instructions.

**Outputs:** ignored `tools/out/` holds `pptx/`, `xlsx/`, `docx/`, `charts/`,
`benchmarks/` and `search-fixtures/`. The six old local output roots were moved
without overwriting anything. All 208 files were SHA-256 verified at their
final destinations, including three older group-effects JSONs relocated from
native output into benchmarks. An ignored local migration manifest records
source/destination/hash. Only empty retired source folders were removed.
Native per-file replacement and same-sanitized-stem behavior are unchanged.
Reviews written with relative asset URLs remain readable after the move.

**Research:** the chart oracle stays a distinct vector/word-scoring pipeline;
its tracked decks/workbooks/cases are test inputs, not trash. Capture/scoring
output is centralized. Corpus/chartex generators now default to scratch under
`out/charts/` rather than quietly writing tracked files or cwd; promotion of a
reviewed corpus is explicit, as the types generator already required.
Calibration scripts and their pinned PDFs/JSON/TSV stay frozen and unmoved.
No corpus, fixture, calibration or published performance data was regenerated.

**Probe cleanup:** seven topics were verified closed against their own README,
standalone source rules, real fixtures and passing regression tests:
`group-id`, `title-insets`, `line-spacing-units`, `table-style-colors`,
`svg-pictures`, `xlsx-cf-rules`, `xlsx-summary-redesign`. Inset/tint comments
were distilled into standalone Office findings before removing their scratch
trees; fixture authoring notes no longer point to deleted scripts. Twenty-five
topics remain because closure is unverified, including nineteen with no
README and one contradictory bidi record. The tracked probe README records
the surviving evidence and retention reasons without pretending the ignored
topics ship in a fresh clone. No blanket cleanup, age-based deletion or shared
helper deletion; one affected retained probe import was repaired.

**Blast radius and deliberate non-changes:** tool locations/default output
paths and references across TypeScript, Python, PowerShell, tests, bench/site
and docs. Native COM/export semantics, rasterization, browser capture layout,
RPC, parser/renderer behavior, public API and benchmark methodology are
unchanged. Source changes outside tools are comments and import/path repairs
only. This is physical organization, not a new rendering or Office harness.

**Proof:** layout/root/output-ignore guards, strict tooling checks, existing
CLI/report/benchmark tests and the 209-test closure verification; moved CLI
help/doctor/inspector and page/Worker memory calibration; strict workspace/tools
typecheck, package/docs builds, unit suite, renderer review suites and all
three native review routes at their new locations. The 94-case chart renderer
and browser word measurer write/read the centralized chart output; Python and
PowerShell syntax checks do not execute corpus generators. Vite/webpack
production probes also pass through the moved checks command. No performance
tables were regenerated and no frozen calibration was rerun.

<a id="note-126"></a>

### 126. IFS range shapes and the deferred coverage proofs (October 14, 2026)

**Root causes.** `SUMIFS` checked criteria ranges against one another but not
against its sum range; `AVERAGEIFS` had the same omission. Separately, the
shared font-use walker only recognized Arabic in run `text`, while XLSX cells
carry it in `value`/`formatted`. The documented visual command filtered titles
by "visual", omitting nine browser tests including the borrowed-content
lifetime proof. The remaining gaps were observability: screenshots with fixed
dimensions did not independently prove slide proportions, and equal sheet
read results did not prove Worker cache hits.

**Native Excel finding.** Read-only evaluation against Shot Log gave
`ERROR.TYPE = 3` (`#VALUE!`) for SUMIFS and AVERAGEIFS with longer/shorter value
ranges, different widths, and equal-area 3x1 versus 1x3 ranges. Matching
`F5:F7`/`E5:E7` with `"Locked"` gave 1.95 and 0.975; SUMIF with the shorter
criteria range still gave 1.95. The shared IFS pair validator now also checks
the value range's width and height. This corrects the omission in §111's
range-size account; SUMIF/AVERAGEIF top-left overlays are deliberately unchanged.
Errors remain scalar Excel values, not unsupported-rule reports: a direct
erroneous condition is unpainted, and ISERROR can match it. Pure values,
real-fixture conditional output, the shipped Worker route and rendered cell
paint pin the behavior. No workbook recalculation or editing path is added.

**Arabic font sampling.** A loaded Courier alias was falsely called unavailable
on the Arabic sheet: its bold Latin sample measured 134.42px against 134.29px
for the generic, inside the 0.5px heuristic tolerance, while the Arabic sample
distinguished them. The missing cell-text detection also omitted Arabic-only
subset waits. The shared walker now recognizes the three text representations;
pure guards and the actual XLSX substitute render pin it. PPTX/DOCX run-text
behavior remains unchanged and keeps its browser proofs. All formats now have
registered-alias success coverage; PPTX also re-reads painted text width and
withholds a font's network response until render resolves, proving the bounded
offline wait and success after delivery on a subsequent render.

**Worker cache proof.** A test-only real Worker wraps streaming worksheet XML
reads at the OPC boundary. It runs the actual request dispatcher and parser,
not a fake cache or main-thread request count. Tests pin repeated/different
range reads from one parse, search reuse without filling/refreshing retention,
LRU eviction and rendered-read refresh, close/reopen on the same Worker, and
an aborted first parse leaving neither a rejected nor partial retained sheet.
No production instrumentation, public seam or RPC message is added. Live
Worker failures are awaited through the existing rejection helper before
asserting; Bun 1.4.3's promise-rejection matcher stalled Worker message dispatch
in the direct closed-session test, not the production dispatcher.

**Blank and lifetime proof.** Positive ISBLANK rules materialize only matching
missing cells and one deliberately targeted missing row, with null values and
empty formatted text; existing cells and unrelated rows are unchanged. A
50,001-recipient range remains authored-only without extra notices or extent
growth, and the blank fill is re-read from DOM. Borrowed-content helpers gain
headless source/signal/freeze/memoization/media-guard tests. The existing browser
test adds closed list/get operations, same-format element replacement for all
three formats, and retains the original PPTX-to-XLSX stale-input case.

**Split-wrap and capture proof.** The real centred-picture paragraph is varied
only in direction/run styles: flow/layout tests read the first slot's physical
side for RTL/LTR, every visible word, bold/colour, justification and absence of
picture overlap. Narrow/wide reflow stays covered; post-destroy resize and a
late font-ready notification leave the row nodes unchanged. Native fixture
bytes and split-wrap algorithms are not changed. PPTX's shared capture backend
compares the actual scene rectangle with each public info.sizes entry, including
absent-font captures; screenshot dimensions remain a separate check.

**Verification policy and blast radius.** Visual commands select visual files,
not titles, so content, search, themes and lifecycle tests always execute.
Browser-only proof remains valid; plain bun test deliberately skips it. Code
changes are limited to the IFS validator, the shared font text scanner and the
capture/selection tools. Public APIs, fixtures, cache policy, text painting,
split geometry and SUMIF/AVERAGEIF semantics stay intact. No screenshot pixel
equality or new Office harness is introduced. Completed coverage rows leave
ROADMAP; visual harness/build memoization remains open and was not implemented.

**Proof.** Native Excel read-only evaluation confirmed the error/value matrix;
the scratch probe was removed after codification. Build and strict
workspace/tools typecheck pass. The normal suite passes (900 pass, 84 skips,
no failures); the documented file-selected visual command passes all 83 tests
across 42 files, both normally and with `--absent-fonts` (native PPTX/XLSX
capture paths included). The restored cross-format stale-content case passes
alongside the three same-format cases. The explicit `dist/core/worker.mjs`
integration case verifies both IFS functions in the shipped artifact, not Bun's
source-extension fallback. Independent read-only review found no remaining
code or documentation blockers. Native Word comparison was not rerun: DOCX's
run-text font sampling and split geometry are unchanged; its actual browser
font/layout/lifetime cases were rerun.

<a id="note-127"></a>

### 127. One visual harness: a shared page runner and per-run browser builds (October 14, 2026)

**Root causes.** Every browser test repeated the same three pieces: the
`AARED_VISUAL` gate as a local `const visual = ...` line, a `Bun.serve` that
re-implemented static module/file/page routing, and a `try`/`finally` wrapping
`new Bun.WebView` plus `navigate`. Thirteen files also called `Bun.build`
directly or through a local copy, so each test re-bundled an identical entry
point. The duplication made the setup in each test look bespoke when it was not,
and it was the reason one file could ship its own MIME types and 404s that no
other file shared.

**Scope of the consolidation.** At this migration, 42 visual files moved to
one gate, 71 page runs to one runner, and direct `Bun.build` calls disappeared.
`docx-harness.ts` and `xlsx-review.ts` became callers too, without changing the
DOCX media routes or A4 host page. Custom request/header escape hatches were
necessary to preserve genuine font/MIME/404 behavior, not a reason to retain
the boilerplate. The current helper map and execution/lifetime rules are in
[ARCHITECTURE](ARCHITECTURE.md#quality-strategy) and AGENTS' testing rule.

**Isolation is preserved, not relaxed.** Each `withVisualPage` call gets its own
server on its own port and its own WebView, so per-test isolation is exactly what
it was. The only shared state is the memoized module text, which carries none.
The runner stops the server it drove on the way out, *including* one the test
supplied through `server`, because the page it navigated to lives exactly as
long as that call and the caller has no longer any scope in which to stop it.
A test that needs one server across several calls passes `keepServer` and owns
the lifetime itself; the lazy-chunk suite does this, because its request log and
effect-painter gate are state the server's closure owns across its pages.
`visual-harness.visual.test.ts` pins this lifecycle, including the throw path:
removing the stop makes three of its four cases fail.

**Cache key normalization.** The memo keys on the resolved path. Callers spell
one entry two ways (`join(dir, "../src/xlsx/index.ts")` and a template literal),
which are different strings on Windows and built the same bundle twice per run.

**Build memoization is bounded and self-healing.** `browserModule` caches by
entry path for the life of the run and drops a failed build so a later retry
rebuilds instead of replaying a stale rejection. `visual-harness.test.ts` pins
that a repeated call returns the identical string, that a second entry builds
separately, that a failed build does not poison the cache, and that two servers
serve the same content on different ports without interfering.

**Deliberate non-changes.** `xlsx-capture.visual.test.ts` keeps its own server
and direct WebView: it wraps `navigate` and every `screenshot` in `within(...)`
timeouts, and the runner would have dropped those guards. The PPTX/XLSX/DOCX
render suites keep their bespoke servers (font gating, per-format media, shared
state across tests) and pass them to the runner as `server`. No assertion, page
script, MIME type, or timeout changed anywhere.

**Measured gain.** Memoization removes about 580ms of bundling per visual run
(707ms across the 63 call sites down to 123ms for the 9 distinct entries).
That is real but small next to the suite's ~85s, which is dominated by WebView
launch and the native Office export; end-to-end wall-clock is unchanged within
run-to-run variance, and no larger speedup is claimed.

**Verification.** Full suite 904 pass / 88 skip / 0 fail; `bun run test:visual`
87 pass across 43 files, and the same with `--absent-fonts`; workspace and tools
typecheck clean. Native Office captures were exercised through the unchanged
`xlsx-capture`/`pptx-capture` suites, so the paired review output is untouched.
An independent read-only review found the server-lifetime leak this section now
records; it was fixed before the numbers above were taken.

<a id="note-129"></a>

### 129. One gap inventory: the roadmap retires into ARCHITECTURE (October 14, 2026)

**The problem.** Open work lived in two places: `docs/ROADMAP.md` carried the
status, the pre-publish gates, the open gaps and the release checklist, while
ARCHITECTURE carried Known limits. §94 had already shown the failure mode this
produces — one difference (an unstyled chart's area border, DOCX
text-box `normAutofit`) listed twice with opposite verdicts, so neither list
could be trusted — and entries through §128 kept adding rows and corrections
across both files. Anyone reading a row first had to know which document owned
it, and the file's own banner ("finished work leaves this file") had drifted
from what it actually held.

**Why move content, not keep a shorter competing roadmap.** The inventory,
release gates and undecided proposals acquired one present-state owner in
ARCHITECTURE; ROADMAP retired. Evidence gaps needed their own identity: a
Word-pasted SVG, a DOCX Office 2016 chart and unchecked generated chart tables
are questions about proof, not established reader-visible failures. Named row
dispositions prevent a deferred item or unscheduled idea from masquerading as
an implementation commitment. The exact current inventory definitions and
release/proposal sections are in [ARCHITECTURE](ARCHITECTURE.md#gaps-limits-and-fidelity-gaps).

AGENTS.md, README.md, CONTRIBUTING.md, docs/USE-CASES.md, the API reference
page and two source comments follow the rename, so no active link resolves to a
deleted file.

**Why one inventory rather than a shorter roadmap file.** Grouping by topic is
what lets one list hold what the old shape could not: an evidence gap is not a
difference a reader ever sees, and it was being argued about in the same list as
a reader-visible layout bug. Naming the verdict per row (OPEN, DEFERRED,
ACCEPTED, BY DESIGN) also stops a deferred item from being read as owed work,
which is what §94's duplicate rows were really costing.

**Why this is not closure by decree.** Relocating a row cannot settle its
meaning or erase its evidence. Acceptance still depends on the specific
difference and may be reopened; completed work leaves a durable rule and proof,
not a rewritten backlog narrative. Those maintenance rules are canonical in
AGENTS' Root-cause discipline. The move did not alter the pre-release policy or
settle the model contract by changing where its release gate was written.

**Current-state corrections.** The rewrite uses the post-§125 tool paths and
§126–§128 font, harness, CSS and resource-lifetime rules. It removes superseded
bundle/page-count claims instead of presenting them as current guarantees.
The source check also corrects the inherited XLSX "headings off" statement:
headings default on everywhere; only sheet tabs have a layer-specific default.
PPTX datetime fields retain cached text; slide numbers are live. Durable rules
such as per-level default text sizes, inclusive quartiles and searchable clipped
text remain in their topic sections. The DOCX band-size limit is recorded with
§99's deliberately-unread scope, not promoted into new work.

**Blast radius and deliberate non-changes.** Documentation and two comments
only: no runtime code, public API, model contract, fixture or test logic changed.
ARCHITECTURE's retained headings keep their anchors (#package-layout,
#shared-modules, #performance-model, #format-rules, #fonts, #fidelity-contract,
#known-limits, #quality-strategy), so existing links into it still resolve, and
`docs/archive/` is untouched. Historical entries keep their "in ROADMAP" wording
  on purpose: they record prior states and claims, not a live index. Stale
  AGENTS.md facts were corrected in the same pass rather than
left to rot — the Worker's parser imports are static and symbol-level (§114),
  and `MARK_FALLBACKS` controls drawn diagnostic chrome, not all metadata.
  WordArt's `data-aared-text-warp` and other fallback attributes can ship without
  a dashed outline or notice bar (no code change in this pass).

**Verification.** Docs-only, so no typecheck, test or build run (AGENTS.md §
Testing rule exempts prose and comment changes). The diff was read end to end
and every remaining live reference to the deleted file was checked. Local
architecture paths/anchors resolve; historical entries 1–128 and the checked
API examples are unchanged. Read-only review compared every former gap/limit,
release decision and unscheduled proposal with its destination and identified
the corrections above before closure.

<a id="note-130"></a>

### 130. XLSX gradient and pattern cell fills, painted as Excel draws them (October 14, 2026)

**The problem.** A cell with a gradient or a pattern background painted blank.
`resolveFill` (`core/xlsx/styles.ts`) kept `patternFill type="solid"` only and
never read `gradientFill`, so the fill was gone before the renderer saw it. A
probe workbook put it side by side with Excel: the solid cell matched, and the
gradient and checker cells were white. This was the one silent fill drop left,
and it was held behind the model-contract decision because the cell style's
`fill` was a CSS colour string.

**Root cause and fix.** The parser, not the painter. The model had no way to
say "gradient" or "pattern" for a cell, while the shared fill model and painter
(`fill-paint.ts`) already drew both for slides, charts and Word text boxes. The
cell style's `fill` becomes a `ShapeFillModel`, and Excel's fills map onto it
(`gradientFill` to `linear-gradient` or a `rect` `path-gradient`, `patternFill`
to `pattern`). The cell painter puts a gradient or pattern behind the text as an
SVG layer, which is how PPTX table cells already paint non-solid fills. A solid
fill stays the CSS background, so the common case costs nothing new. The
painter is fetched before the grid paints, and only for a sheet that has such a
cell (`needsFillPainter`, in `render()` and on a sheet switch). A conditional
format's differential still applies its colour only; Excel's dialog offers
patterns there too, but no fixture has one.

**What Excel does, from its own screen bitmaps** (probe xlsx-cell-fills: each
case filled through Excel's COM object model, black on white, then copied with
`Range.CopyPicture` at 100% and read pixel by pixel):

- **Gradients blend straight in sRGB.** Black to white across a cell read 4, 28,
  53, 78 … 251, even steps. PowerPoint's two-stop gradients ease in linear light
  (§28), so the gradient models gain `blend: "srgb"`, which the painter honours
  in every gradient kind.
- **`degree` is measured in the cell's own proportions.** 0 runs left to right,
  90 top to bottom, 180 right to left. 45 runs exactly corner to corner in a
  216 x 80 px cell: the middle row reads from 25% to 75%. That is DrawingML's
  `scaled` angle. It is not mirrored in a right-to-left sheet (paired capture).
- **`type="path"`** is a rectangular gradient from a box whose `left`, `right`,
  `top` and `bottom` are positions from the cell's left and top edges. All 0.5:
  the first stop at the centre, reaching the last at every edge, the diagonal
  reading like the middle row. All absent (Excel's "from corner"): from the
  top-left corner, the middle row flat at 50% until halfway across.
- **The 17 patterns** each repeat within 8 x 4 px, one bitmap cell per CSS
  pixel. Their tiles are in `pattern-presets.ts` under Excel's names (`darkGray`,
  `gray125`, `darkGrid` …); `darkGray`, `mediumGray` and `lightTrellis` equal
  PowerPoint's `pct70`, `pct50` and `pct30`.
- **Patterns anchor to each cell.** Vertical and horizontal line patterns in
  cells starting off the 4 px grid (column B at 70 px, row 6 at 105 px) showed
  the same phase from every cell's own corner, so the tile starts at the cell,
  where PowerPoint's starts at the slide origin.
- **Pattern colours** are `fgColor` on `bgColor`; `indexed` 64 and 65 are the
  system foreground and background (black and white).

**Rejected.** Approximating a gradient by its middle colour, or a pattern by
its background: it hides a visible difference rather than drawing it, and the
painter for both already existed. Reusing PowerPoint's preset bitmaps by
nearest look: only three are identical, and the probe gave Excel's exact tiles.

**Format 47, found while closing the number-format probe.** ECMA-376's table
prints built-in id 47 as `mmss.0`; Excel resolves it to `mm:ss.0` and shows
45123.756 as `08:38.4`, so the spec's missing colon is a misprint. Showing
`08:38.4` also needed fractional seconds, which the date formatter dropped:
`ss.0`, `ss.00` and `ss.000` now show tenths to thousandths, rounded at that
precision first so 59.96 s under `ss.0` carries into the next minute.

**Blast radius.** The cell style's `fill` changes type (public in this tree;
the model-contract decision may hide it later). Its readers were the style
reader, the cell painter and five parser assertions. Conditional fills,
borders, charts and PPTX/DOCX fills are unchanged; the painter's default blend
is unchanged for every existing caller.

**Proof.** `xlsx-coverage.xlsx` gains a "Fills" sheet authored through Excel's
COM object model (label stripped): an RTL Arabic table with 90° and 0°
gradients, a from-centre gradient, and `darkGrid`, `gray125` and `lightDown`.
`xlsx-parser.test.ts` pins each fill as Excel wrote it.
`xlsx-cell-fills.visual.test.ts` opens the workbook through the Worker, renders
the sheet, and re-reads each cell: the layer, the paint server, two
un-interpolated stops for the sRGB gradients, a cell-anchored 8 x 8 pattern
tile, no fallback marker, and the value text on top. A paired capture against
Excel's print shows the six cells matching. `xlsx-numberfmt.test.ts` covers
format 47 and fractional seconds.
