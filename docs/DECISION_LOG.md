# Decision Log

> Purpose: a chronological record of what was decided and why, with the
> reasoning behind each call. Read this when you ask "why is it shaped this
> way" months from now. For the current architecture, see
> [ARCHITECTURE.md](ARCHITECTURE.md). For the public API, see the API
> reference page ([`site/src/content/docs/api.mdx`](../site/src/content/docs/api.mdx)).
> For what's queued next, see [ROADMAP.md](ROADMAP.md).

This is a distilled record from a single design session (2026-09-03), not a
rediscovered history — the raw decision table it's drawn from, with every
alternative considered, was `docs/archive/ms-previewrs/DESIGN.md`. That
archive has left the tree; read it from history with
`git show 5a32bf0^:docs/archive/ms-previewrs/DESIGN.md`.

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

Supersedes the blanket reading of §7 without rewriting it: flow rendering
stays exactly as promised there, and `pagination: { mode: "layout" }` opts
into physical pages with an honest, narrower promise.

Root cause: a marker-only page model cannot paginate real documents. The
gate fixture (`docx-coverage.docx`) carries one explicit `w:br w:type="page"`,
zero `w:lastRenderedPageBreak` hints, zero `w:pageBreakBefore` flags, and one
`nextPage` section boundary — yet Word prints four pages. The missing page is
layout overflow inside the first section, so page markers are evidence, not
pagination. These counts come from unpacking the fixture's OOXML and from
Word's own print of it; re-derive them the same way (the audit notes were
never committed).

What changed: the parser now resolves `DocumentSection[]` (per-`sectPr`
geometry, break type, header/footer references) instead of flattening
sections into one flow; a headless heuristic paginator (`core/docx/`)
fills section content boxes with estimated block heights, honoring explicit
breaks, section transitions, and a footnote fixed-point; the renderer paints
one `section[data-aared-element="page"]` per laid-out page with repeating
headers/footers, per-page footnotes, and post-layout `PAGE`/`NUMPAGES`.

Deliberate limits: estimates model the renderer's own CSS (not Word's text
engine), so break positions match structurally rather than exactly — the
gate fixture's Table 1 caption lands with its table on page 2 instead of
opening page 3 as in Word, a documented one-block deviation that keeps
table and caption together. Tables never split across pages yet;
`nextColumn` degrades to continuous; layout pages paint tighter
Word-like metrics (1.15 base line box) than the flow article's 1.5.

## 11. DOCX layout fidelity corrections from bench comparison

*Partly superseded by §22: header clipping, border/indent geometry and list
geometry below were later re-derived from native Word and changed.*

Side-by-side bench review against the Word reference exposed three defects
in one pass, each traced to its OOXML source rather than patched visually:

- **Quote border dropped.** The pull quote's `w:pBdr` left border (gold,
  2.25pt, 14pt space) never reached the model — `readParaProps` had no
  `pBdr` case. Borders now resolve per side (merging through the style
  chain, `nil`/`none` skipped) into `DocxParagraphLayout.borders` and paint
  as CSS border plus `space` padding; the estimator narrows the text column
  by indent plus border-plus-space.
- **Double list indent.** Numbering-derived indents painted twice: once as
  the `<li>` offset and again as the inner paragraph's own margin, pushing
  item text ~84px in instead of ~47px. The item now starts at the numbering
  indent exactly once while the inner paragraph drops its start indents
  (model untouched); the estimator wraps items in the same
  indent-minus-marker column.
- **Table clipped by stacked furniture.** Headers/footers painted inside the
  fixed-height page consumed ~100px the paginator never reserved, so page
  2's table overflowed into `overflow: hidden` and clipped mid-row.
  Layout sheets now follow Word's box model — furniture absolutely
  positioned in the margin zones, body owning the full page-minus-margins
  box the paginator filled. An oversized header clips in its zone rather
  than overlapping body text: the safe side of Word's overlap behavior.

  Inline images inherit their host paragraph's alignment onto the `<img>`
  itself (auto margins need a fixed-width block — the `<figure>` fills its
  container, collapsing them to zero), so centered source paragraphs center
  their figures; floats ignore alignment, as in Word.

## 12. `aared/preview`: superseding the deferred facade decisions

The bench's aared adapter (~205 lines) re-implemented what the library
already knew: try each format's `open()` in turn on `unsupported-format`,
then hand-manage generation guards, `destroy()`/`close()`, and DOM cleanup
per format — while `<aared-viewer>` already detected, opened, and guarded
the same lifecycle internally. Two earlier deferrals ("no `preview()`
convenience API", "no all-format auto-detect entrypoint") blocked promoting
that logic into the package; both rested on one rationale — one-format
consumers must not pay for all formats.

`aared/preview` supersedes both deferrals without breaking that rationale:
the facade lives behind its own subpath (root `.` stays `aared/core`),
so `import { open } from "aared/xlsx"` never includes it, and the format
entrypoints load via runtime `import()` — exactly one format chunk loads,
the detected one. Detection reuses `WorkerSession.detectFormat` (Worker-side
probe, no main-thread ZIP/XML parsing), and the facade is extraction, not a
fork: it shares the element's detect → lazy-open → generation-guard shape.

Deliberate boundaries, held from review against the settled rows:

- `preview()` covers open → render → navigate → destroy only. Search, range
  reads, and media stay on the headless entrypoints.
- One teardown verb (`destroy()` does render + session + painted DOM)
  instead of the headless `destroy()`/`close()` split — the facade owns the
  whole session, so the split would be ceremony. The split stays where
  callers own the session.
- Fonts stay opt-in (`fontResolver` passthrough, never a CDN default):
  zero-network-by-default stands (see §6).
- `core` stays DOM-free: the facade is its own subpath importing the
  internal session plumbing, never part of `aared/core`.

## 13. DOCX `render()` joins the shared async object envelope

DOCX rendered positionally and synchronously — `render(sections,
container, options?)` — while PPTX/XLSX rendered as `await render({ slide |
sheet, target, ... })`. The positional form dated from when sections were
"already in memory, so sync is honest," but every consumer (bench, element,
`preview()`) had to branch on call shape instead of just on format — the
branch the shared-grammar design (§-level API principle: same envelope,
format-native nouns) exists to eliminate.

Since the package is pre-release (see AGENTS.md "Pre-release breaking
changes"), the old form was removed, not aliased: `await render({
document, target, ... })` everywhere. Two real gains beyond symmetry:

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
  consumer could not branch on it usefully. The only booleans that could vary
  never did.
- **No consumer.** Nothing in the package (element, `preview()`, hooks,
  renderers) read it. Its only readers were tests asserting the exact
  literal, and the API page.
- **Already wrong once.** `presentation.notes: true` claimed something
  nothing rendered (see the ROADMAP gap row). A flag that must be kept
  honest by hand, with no reader to notice when it drifts, is a liability.
- **Cheap to add, costly to remove.** A new field is non-breaking; removing
  a published one is not. The package is pre-release, so this is the
  cheapest moment to cut.

It was removed (not aliased — AGENTS.md "Pre-release breaking changes"):
`deck.capabilities`, `workbook.capabilities`, `document.capabilities`, and
the `PptxCapabilities`/`XlsxCapabilities`/`DocxCapabilities` types. Dark-mode
support is not gated by it either: PPTX rejects `contentMode: "adaptive"`
with an `unsupported-feature` `AaredError`. Re-introduce it only when a value
varies per file or environment, and then with a consumer that reads it.

Kept, deliberately:

- **`search()`** — a real function someone building their own search UI
  would otherwise have to write against the resolved models. It stays
  headless-only (the facade boundary in §12 is unchanged) and is documented
  with its costs: no index (PPTX/XLSX parse each unit in turn), text only,
  120-character results, and DOCX coverage that excludes table and footnote
  text.
- **`read()`, `media.get()`, `info.sizes`** — load-bearing: `render()`
  consumes the same model `read()` returns, the renderer resolves pictures
  through `media`, and PPTX slide geometry comes from `sizes`.

What stays true about scope: the headless API exists for custom UI. Reading
data from resolved models is a by-product, not a converter feature.

## 16. Search cost, DOCX search coverage, `info.sizes`, and the Worker bundle

Follow-up to §15, prompted by an audit of what the headless API costs.

**Measured first** (`tools/bench-search/`: generated large fixtures, 300
slides / 13 sheets ≈ 52k rows / DOCX ×40; Bun, one machine, so read the
ratios, not the absolutes). Before → after:

| | cold first search | repeat search |
|---|---|---|
| PPTX | ~495 ms → ~500 ms | ~495 ms → 0.2 ms |
| XLSX | ~3.9 s → ~3.9 s | ~1.4 s → ~3 ms (9 ms for a 22k-hit query) |
| DOCX | (in-memory; index build 7 ms) | 0.2 ms |

What changed: PPTX/XLSX keep each unit's searchable text after the first
read (a per-document index on the main thread, cleared on `close()`); the
PPTX worker caches parsed slides like XLSX already did; search reads skip the
font-resolver round trip (`unit(..., { fonts: false })`); the worker skips
re-walking a cached unit for media ids; DOCX flattens its blocks once. The
first search is unchanged because parse and model transfer still dominate it.

Tried and reverted: memoizing layout/master XML per package. Measured effect
was ~2% (inside run-to-run noise) for shared read-only XML trees, so it was
not kept.

Not done, deliberately: a Worker-side text-only `search` message. It would
remove the model clone (~1.4 s of the 3.9 s XLSX cold search) but adds a
protocol message and moves text extraction into core; revisit if real
documents make the cold search a problem. Also found and fixed: the worker
cached the XLSX workbook read as a *promise* bound to the first request's
abort signal, so a cancelled first read could fail every later sheet read.

**DOCX search now covers tables and footnotes/endnotes.** Matches gain
`kind`, and `table: { index, row, column }` / `note: { id, label }`. Word
stores notes after the body, so a note reports the section and heading of the
paragraph that first references it. Headers/footers and text boxes (not
parsed) stay out.

**`info.sizes`**: DOCX now reports one page size (inches) per section; its
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
Before this, aared painted every slide white, ignored layout/master graphics,
resolved `bg1`/`tx1` to nothing, and dropped `a:fld` fields (slide numbers,
dates) on every slide.

Decision, by cost:

- **Background: flatten onto the slide** (`SlideModel.background`, effective
  slide → layout → master). It is a few bytes (a colour, a gradient, or a
  `mediaId`), so there is nothing to save by moving it to a shared object, and
  a second source of truth would need a compositing rule. The user's point that
  a master image is repeated on every slide is served by fetching the *bytes*
  once (below), not by restructuring the model.
- **Graphics: store once per layout** (`SlideLayoutModel`, master's shapes
  then the layout's). Layouts are few and slides are many, so this is where
  duplication would be: the Worker parses a layout once, the session fetches
  it once with a dedicated `layout` request (explicit, cancellable, one shared
  in-flight fetch that is not bound to any caller's signal and never caches a
  failure), and every slide on that layout carries the same object.
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
change. Layouts with no graphics (many decks) add no `layoutId` and no
request.

Verified against real PowerPoint output: a master "Background Style" on
`pptx-coverage.pptx` (`bgRef idx="1002"` on `bg2`, dark colour map)
renders `#3A4961` in PowerPoint. An early version painted the bare scheme
colour (`#0E2841`); the theme's `bgFillStyleLst` entry (`tint 95000` +
`satMod 170000`) is now applied, which needed `satMod` support in
`resolveDrawingColor` (and `lumMod` above 100% no longer clamps).

Also found by that fixture: date, footer, and slide-number placeholders were
drawn at the slide's top-left. The master numbers them 2/3/4 while layouts and
slides use 10/11/12, and every level was matched by index, so the master was
never found. Layout → master matching is now by placeholder *type* (title,
body, dt, ftr, sldNum), as PowerPoint does; slide → layout stays by index. This
also gives non-first content placeholders the master body's text style.

Not verified against real PowerPoint output yet (the fixtures don't exercise
it): the `showMasterSp` rule, picture backgrounds, and `p:clrMapOvr`
(layout/slide colour-map overrides, not read). They are listed in the ROADMAP
fidelity gaps; verifying them needs an edited fixture, which is the
maintainer's to author.

Tried and reverted along the way: memoizing layout/master XML for placeholder
inheritance (measured ~2%, within noise).

## 18. Core performance pass: streaming parse, cached relationships, split worker

`core` was optimized in isolation against generated PPTX/XLSX/DOCX fixtures
(incl. a stored/uncompressed ZIP and a hand-corrupted edge-case sheet) and
landed as a whole-directory replacement. Public API and every existing export
are unchanged; outputs were diffed field-by-field against the original and are
identical except the two intended changes below.

**Measured** (Node 22, worker message loop driven directly by
`tools/core-harness.ts` — one machine, so read the ratios, not the
absolutes):

| Case | Before | After |
|---|---|---|
| XLSX, 150k rows × 10 cols, first sheet read | 14.3 s, 1.58 GB peak RSS | 2.5 s, 0.47 GB |
| XLSX with whole-sheet hyperlink / CF range | crashes (`Map maximum size exceeded`) or 4.5–9 s | ~2 ms |
| PPTX, 80 slides with placeholders, all slides | 453 ms | 100 ms |
| DOCX second read of same document | full re-parse | cached (0 ms) — superseded by §35: the page keeps the sections, the Worker re-parses |
| Worker bundle, split build, per-format load | 115 KB min (all formats) | ~60 KB (shared + one format) |

**What changed:**

- **XML** (`core/xml.ts`) — char-code scanner; no entity decoding without
  `&`; shared frozen empty `children`/`attrs`; interned tag names;
  allocation-free `attr()` misses; new streaming mode
  `parseXml(src, { streamChildrenOf, onChild })`. Parsed trees are read-only:
  nothing in `core` mutates them.
- **ZIP/OPC** (`core/zip.ts`, `core/opc.ts`) — inflate writes straight into a
  buffer of the declared size (capped at deflate's max ratio, chunk fallback
  on lying headers); stored entries are decoded without a copy; one shared
  `TextDecoder`; `.rels` parsed once per package; media is transferred to the
  main thread without the old double copy (`readPart` returns an owning
  buffer; internal decodes use the aliasing `readShared`).
- **XLSX** — `sheetData` and `sharedStrings` are streamed (no full element
  tree); cells live in a sorted per-row grid instead of `"row:col"` string
  maps, so hyperlinks/conditional formats visit existing cells only (a rule
  on `A1:A1048576` touches authored cells, not the area); no
  `Math.min(...big)`; no extent-sized row-offset array; corrupt `<col max>`
  clamped; theme and `workbook.xml` no longer parsed three times
  (`readWorkbook(pkg, signal, known?)` reuses the index); cell styles shared
  per style id; number formats memoized/precompiled; faster `General` path.
- **PPTX** — placeholder inheritance reuses the cached layout/master chain
  (previously re-inflated both parts per placeholder); link maps memoized.
- **DOCX** — parsed sections cached in the worker (superseded by §35: the
  page keeps them, the Worker keeps none); shared header/footer parts parsed
  once.
- **Worker/session** (`core/worker.ts`, `core/session.ts`, new
  `core/collect.ts`) — format parsers load via dynamic `import()`, so a
  splitting bundler ships each format separately. Single-file consumers
  cannot use that split (a Blob worker has no siblings for the relative
  chunk imports to resolve against), so the build also emits a
  self-contained `dist/core/worker-single.mjs` (same source, imports
  flattened; exported as `aared/worker-single` for the Blob recipe) —
  proven by `test/worker-single.test.ts`, which opens every format through
  a Blob-URL worker. Model walks
  for media ids and font families go through `collectStrings` (no per-object
  allocation, each shared cell style visited once); abort helpers moved to
  `errors.ts` (still re-exported from `zip.ts`, and outside-`core` importers
  now point at `errors.js`); duplicate `FontResolver` type removed in favour
  of the canonical `aared/core` one.

**Intended behavior changes (2):**

1. **XLSX charts now use the shared `chart.ts`** (as the old adapter's
   comment planned), so they match PPTX charts. Charts without cached values
   now report `unsupportedReason` instead of drawing empty axes; doughnut
   charts are drawn as pies in both formats. The old XLSX-only adapter also
   painted the authored `c:title`; that moved into the shared module too, so
   both formats now paint chart titles (centered, bold, in the band above
   the plot). `adaptChartGraphic` now takes the theme, and
   `chartUnsupportedReason` explains an undrawn chart.
2. **Shared read-only objects**: cells using the same style id share one
   style object, and cells covered by one hyperlink share its object.
   Structured clone preserves this, so the main thread receives fewer
   objects. Code outside `core` must not mutate `cell.style` /
   `cell.hyperlink` in place.

Also: XML tag/attribute names with non-ASCII characters are now accepted
(valid XML; previously a parse error). Single-file (non-split) worker bundles
are ~3 KB larger because of the lazy-import wrappers.

This supersedes the "layout/master memoization reverted" note in §17 (the
chain-cache reuse above it now wins well past noise) and closes the
"per-format Worker" item in the ROADMAP bundle-size section.

## 19. Split worker vs. consumer bundlers: Vite verbatim copy and the abort taxonomy

Follow-up to §18, prompted by a broken production build: `bench/` dev worked
but `vite build` + `preview` loaded forever. Two compounding defects, one in
packaging and one in the error model.

**Packaging: Vite copied the worker instead of bundling it.** The consumer
build emitted `worker-*.mjs` byte-identical to `dist/core/worker.mjs` — raw
`import ... from "./errors-HASH.mjs"` lines intact — with none of the sibling
chunks. `WorkerSession` built the worker URL as
`new Worker(workerInit ?? new URL("./worker.mjs", import.meta.url))`, and
Vite only bundles a worker target when `new URL()` appears *literally*
inside `new Worker()`; the `??` indirection defeated detection, so Vite fell
back to a static-asset copy. The pre-split worker survived that accident
because it had no siblings to miss. Fix: branch on the override kind first so
the default path is the literal documented shape (pinned by a comment —
never fold the branches back together). Verified with
`tools/check-consumer-build.ts`, which production-builds a minimal Vite
consumer, asserts every chunk the worker names is emitted, serves the result,
and opens one real fixture per format headlessly.

**Vite consumers need one config line.** The split worker is a multi-chunk
graph and Vite's `iife` worker default cannot emit it — the build fails
loudly (`Invalid value "iife" for option "worker.format"`). `worker: {
format: "es" }` in `vite.config` is now the documented setup (dev needs
nothing); `bench/vite.config.ts` dogfoods it. Single-file consumers are
unaffected — they use `aared/worker-single` with the Blob recipe.

**Error model: a dead worker reported `"aborted"`.** The boot-failure path
(`worker.onerror` → bridge `fail()`) constructed `AaredError("aborted")`,
so every abort check treated a broken worker as a cancellation —
`usePreview` swallowed it and the UI loaded forever instead of erroring.
Worker death (boot or mid-session) now reports the new `"worker-error"`
code, which matches neither `isAbort()` nor the `"aborted"` code check, and
a failed `postMessage` moved with it (a failed post is transport failure,
never cancellation). Genuine cancellation is untouched: `abortError()`,
`closedError()`, and signal paths still report `"aborted"`, which is also
what unmount cleanup relies on. Pinned by a lifecycle test asserting boot
failure rejects with `"worker-error"` and `isAbort() === false`.

## 20. Worker back to one self-contained file: the split is reverted

Overturns the §18 packaging decision (keeps its parse-streaming,
`collect.ts`, and `errors.ts` work, which were never about chunking). The
split worker — `dist/core/worker.mjs` plus per-format sibling chunks, with a
separate `aared/worker-single` artifact for Blob consumers — saved ~55 KB
for single-format apps but bought a packaging failure class: Vite's worker
detection needed the literal `new Worker(new URL(...))` shape (§19),
verbatim copies dropped sibling chunks with zero errors, consumers needed
`worker: { format: "es" }`, and every artifact had to be deployed as a
directory of hashed names that change per build. One self-contained
`dist/core/worker.mjs` (same `src/core/worker.ts` source, `codeSplitting: false`)
has none of that: it deploys alone, inlines into a Blob URL, and survives a
verbatim copy. Cost is the reverted saving — every consumer ships all three
parsers (~115 KB min instead of ~60 KB for one format). The `xml.ts`
`codeSplitting` group and the `worker-chunks` test existed only to pin the
split graph, so both are gone; `test/worker.test.ts` pins the new contract
instead (no relative imports in the artifact, Blob URL opens every format),
and `tools/check-consumer-build.ts` asserts the emitted worker references no
sibling chunk.

Follow-up: measured production cost (minified / gzip) is worker 117 KB /
39 KB, all-formats main thread ~100 KB / ~28 KB — ~67 KB gzip total for a
three-format renderer, against ~500 KB–1.6 MB for the single-format
alternatives in `bench/`. Chunking would save ~15–20 KB gzip for
single-format apps only. Decision confirmed: keep the single file; per-format
worker variants only on a real consumer size budget, never speculatively.

Re-checked (Sep 2026, with the download-size method of §33): one flat worker
per format (`dist/{pptx,xlsx,docx}/worker.mjs`, each built alone and started by
a literal constructor in its own entry) measured 23.5 / 22.0 / 16.8 KB gzip
against 43.1 KB for the shared one — 20–26 KB less for a single-format page,
about 19 KB more for a page that imports all three entries directly, and four
worker artifacts to ship and document. Per-format workers have been tried
several times before and the packaging overhead around them was never worth
it; rejected again. The size gap it would close is DOCX's (about 8 KB behind
docx-preview + jszip), and that stays open.

## 21. Measuring text for DOCX pagination: measure in the DOM; Pretext considered, not adopted

Context. The layout-mode paginator (§10) fills pages with block heights it
*estimates*. Line counts come from a per-script average glyph advance
(`AVERAGE_CHAR_EM`: 0.46em Latin, 0.40em Arabic) with a 0.1-line rounding
tolerance (`WRAP_TOLERANCE_LINES`), in `core/docx/pagination.ts`. Those
constants were fitted to one fixture (Arial and Arabic body text on Windows).
They are the weakest part of the engine: they land on Word's page breaks for
this file only, a different document or font can break them, and while they
were being fitted the spec table spilled to the wrong page more than once.

Options considered.

1. Keep estimating and tune the constants. The status quo; fragile.
2. Pretext (`chenglou/pretext`): measures text once through canvas
   `measureText`, then computes breaks arithmetically with no DOM reads.
3. Measure in the DOM: the renderer builds the blocks off-screen, reads their
   real heights in one batch, and feeds them to the paginator.

What Pretext's README says (fetched Sep 2026; the README only — its licence and
version were not checked in the repository or on npm, and the speed and size
figures in third-party write-ups were not verified):

- Web Workers / `OffscreenCanvas`: not stated.
- Bidi: "doesn't compute bidi levels"; segment widths are for line breaking
  and "aren't enough to position individual characters correctly in Arabic or
  mixed bidi text"; it advises rendering each paragraph as one DOM element with
  its direction set.
- Mixed fonts in one paragraph: supported through a separate `rich-inline`
  entry point.
- Tabs: "follow the default browser-style `tab-size: 8`". Inline images and
  other non-text inline objects: not mentioned.
- Font loading (web fonts, waiting before measuring): no guidance.
- Version, licence, API stability: not in the README.
- Caveats: only `white-space: normal` / `pre-wrap`; `system-ui` is unsafe;
  needs `Intl.Segmenter` and Unicode property escapes.

Why it is not adopted.

- **The gaps are Word's inline model, not just bidi.** The bidi caveat is about
  positioning glyphs; counting lines needs only breaking widths. What DOCX
  paragraphs need beyond that — tab stops with leaders, inline images,
  superscripts, letter-spacing, justified lines — is not modelled, so it would
  sit behind adapters until it is a second layout engine that must still agree
  with what the DOM paints.
- **Canvas measurement is a grey area under §3.** That rule bans canvas
  `fillText` and WASM text shaping; `measureText` is not named. Using it would
  need its own entry here. DOM measurement needs none.
- **Unknown font contract.** Measuring and painting must use the same faces,
  including anything delivered by `aared-fonts` / `fontResolver`; Pretext says
  nothing about loading them.
- **Unstated maturity** on the pagination path.
- **No Worker need.** `core/docx/pagination.ts` is DOM-free, but `paginate()` is
  *called from the renderer on the main thread* (`renderLayout` in
  `docx/index.ts`), where a DOM exists, so there is no measurement to run in the
  Worker.

Decision. Do not add Pretext. Move DOCX pagination to DOM-measured block
metrics in the renderer — a follow-up, **not done yet**. `paginate()` keeps its
algorithm and takes an injectable block-metrics function
(`{ topPx, contentPx, bottomPx }`); the renderer supplies DOM measurements and
headless callers and unit tests keep the estimator. Requirements found in the
fidelity pass:

- explicit image heights — block images are `height: auto` until their Blob URL
  loads and would measure 0;
- wait for `document.fonts.ready` and re-paginate when fonts arrive, or pages
  jump after first paint;
- run the tab-stop pass before reading heights, since tabs change wrapping;
- one measuring container per page width (portrait and landscape sections);
- measure once and move the same nodes into the pages, rather than rendering
  the document twice;
- the paginator's fold and margin-collapsing logic stays as it is.

Trade-off accepted: break positions then depend on the machine's fonts, as
Word's do, and can only be pinned by WebView visual tests, not unit tests.

The estimator stays. It is the only page layout that runs without a browser,
which makes it the fit for the "headless thumbnail export" item in
[ROADMAP.md](ROADMAP.md) (a browser-free, agent-oriented render): a headless
render cannot measure the DOM, so it takes its page breaks from the estimator.
That raises the bar on the estimator rather than removing the need for it — its
constants are fitted to one fixture today — so headless accuracy is where a
measurement approach can still earn its place (measured font metrics injected
into the estimator, or a library such as Pretext, whose README says
server-side rendering is coming "soon").

Revisit Pretext when headless accuracy is scheduled (the thumbnail-export item
above, or an accurate page count in the Worker before render), or if DOM
measuring proves too slow on very large documents. Then run a spike first:
compare Pretext's line counts with the DOM's for every paragraph of the
fixtures, adopt it only if they agree, and verify Worker or server support and
the licence before depending on it.

## 22. DOCX layout follows Word's own numbers; supersedes parts of §10 and §11

The bench comparison in §11 judged fixes by eye against screenshots. Since
then every layout rule has been read from a native Word export of the frozen
fixture instead — `tools/word-render` exports through the installed Word,
rasterizes with PyMuPDF, and dumps every text span's font, size and position —
so a number is read, not guessed. The full rule list was kept in the
handoff document, since folded into `ARCHITECTURE.md` § Format rules (DOCX).
What it overturned:

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

Pagination remains structural, not Word-exact (§10, §21): break positions still
depend on the estimator until measured pagination lands.

## 23. Shared-code pass: one home per rule, and the decisions it forced

A read-only audit (after §22) listed code that had been written twice for lack
of context. Every finding was checked against the code first; most held, a few
did not, and two were larger than reported. What landed is one home per rule:
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
which is also why `open()`'s titles disagreed with the heading nodes, so both
now come from one rule; a non-page `w:fldSimple` (DATE, TITLE, REF…) dropped its
cached text; a run in a content control ignored its own character style; DOCX
passed a raw numeric zoom, so `setZoom(0)` collapsed the document; the bench
swallowed every typed `AaredError` from a sheet render.

Deliberately not merged: the SheetJS/docx-preview baseline adapters (a baseline
must not depend on the code it measures); XLSX `tint` (sRGB lerp) versus
DrawingML `tint` (linear light, probe-backed), which are different operations;
`TextRunModel.verticalAlign` versus `baselinePct`, one concept in two dialects
that each renderer paints as measured; and the two facades' `render` calls,
which differ on purpose (the element adds sheet tabs and theme).

Tests and tools got the same treatment (`test/harness.ts`, `test/zip.ts`,
`test/docx-harness.ts`, `tools/shared.ts`), with the same care about what not to
merge: `writeReview` has different signatures per format, the two
`browserRenderer`s build different renderers, the inline `Bun.build` +
`Bun.serve` harnesses serve different routes, the three fixture lists serve
different jobs (timing, consumer build, bench samples), `tools/core-harness.ts`
walks for a different purpose than `core/collect.ts`, and fixture paths and
manifest lookups are mechanical churn across fifteen files with no drift risk.
XLSX keeps raw Excel border names and maps them in the renderer, where DOCX
resolves its borders in core; moving that changes the public
`XlsxCellStyleModel`, so it waits on the model-contract decision (ROADMAP,
"Before first publish").

Size: the Worker bundle did not grow. Bundled the same way from the last commit,
the pre-pass tree and now: 124,419 → 121,005 → 121,000 bytes minified (40,583 →
40,033 → 40,029 gzip); the shipped `dist/core/worker.mjs` is 253,897 bytes
unminified (64,450 gzip). The renderer-side modules are not in the Worker.

## 24. Dark mode, looked at: automatic text follows its fill; authored text on the dark surface is an open gap

Adaptive dark had only ever been tested for the chrome and for the surface
colours; nobody had rendered the coverage fixtures dark and *looked*. Doing so
in the bench, with a contrast audit over the live DOM (each text element's
computed colour against the composited background behind it), found two
distinct defects in XLSX and DOCX. PPTX has none: a slide is source-faithful on
purpose, and only its stage is themed.

**Fixed — automatic text now resolves against the fill it sits on.** Adaptive
dark paints automatic text in the light ink (`#e8e8e8`), and a cell or paragraph
the author filled with a light colour kept that ink, so the Shot Log's cream
zebra rows, its total row and its conditional colour scales read at 1.02–1.23:1
(150-odd cells, invisible). The contract already said adaptive dark re-themes
*automatic* text; it never said automatic text is one constant. It is now
whichever palette ink reads better on the effective fill (`automaticInk` in
`appearance.ts`, on the new `contrastRatio` in `core/color.ts`) — the rule
Word's and Excel's Automatic colour follows, and PowerPoint's here since the
slide-title work. XLSX resolves it per cell from the conditional overlay's fill,
else the style's; DOCX marks each paragraph, run and table cell that has an
authored fill (`data-aared-shade`) and `paintAutomaticInk` re-resolves them, so
a theme change repaints in place with no re-render. Outside adaptive dark nothing
changes: source mode still paints Excel's black on any fill and inherits the
page ink in Word. This needed no contract change. After it, the audit's only
sub-4.5 text on a filled cell is an authored pair that reads the same in light
mode (amber on cream 3.28, teal on pale teal 3.98).

**Open — authored text colours on the surface aared darkens.** What remains is
text with an authored colour and no fill of its own: Excel `FF000000` body cells
(1.21:1), a `12171F` title (1.03), `0000FF` inputs (2.03), the `1155CC` link
(2.65), `5B6470` notes (2.9), `008000` (3.39), `2E7D74` (3.57); Word's `1B2029`
body text (1.07, 129 spans in the DOCX fixture alone) and the same greys and
blues. Those colours are legible on the white the author saw and unreadable on
the `#1a1a1a` aared substituted, so a whole sheet or document reads as blank.
It was left for a decision because the API reference says adaptive "never
touches explicit authored styling" and explicit colours "paint verbatim", and
fixing it overturns that line. Not built; the next step is to decide and build
it. The candidates, with what the evidence says:

- **Re-theme only where aared's own surface broke it (recommended).** An
  authored colour stays verbatim if it reads on the surface it sits on
  (contrast ≥ 4.5:1); if it does not, lift its lightness with hue and saturation
  kept until it does, and let a near-black, low-saturation ink (about HSL
  saturation ≤ 0.3 and lightness ≤ 0.2: `000000`, `12171F`, `1B2029`) become the
  page ink outright, so "black text" is one grey rather than a dimmer one beside
  the automatic text. Text over an authored fill or shading is never touched:
  the author chose that pair, and it reads as it does in light mode. One
  sentence covers it: adaptive dark re-themes an authored value only where the
  surface change is what broke it. It needs a small cache (colours repeat by the
  thousand), the effective fill for DOCX from the same `data-aared-shade`
  markers, and an amended Dark mode section in the API reference.
- **Keep the surface light where the text on it is authored dark.** Rejected:
  the DOCX fixture's body text is authored, so adaptive dark would be a no-op for
  exactly the documents that need it.
- **Re-theme only near-neutral dark inks (black and greys).** Smaller, but blue
  inputs (2.03) and links (2.65) stay hard to read.
- **Leave it.** Keeps the contract, and keeps the sheets blank; an embedder can
  pick `contentMode: "source"` for such files.

The test that would have caught it, and should gate whichever is chosen, is a
visual invariant rather than pinned colours: render each coverage fixture in
light and in dark, pair the text elements, and assert that dark contrast is at
least `min(light contrast, 4.5)` for every pair. It needs `getComputedStyle`, so
it is a `Bun.WebView` test (like `element-viewer.visual.test.ts`), and the audit
used here is short enough to become its helper.

Not covered by this pass, so still to look at: the XLSX data-bar overlays and
authored border colours on the dark surface (black borders), and the DOCX
header/footer rules. Charts stay white cards by design - a chart is source-
faithful content, like a PPTX slide.

## 26. Deferred geometry: the worker describes, the main thread evaluates

The full ECMA-376 preset catalog (~187 shapes) and `custGeom` freeforms
cannot live in the worker: the worker is built flat (`codeSplitting: false`,
so it works as a single file or Blob URL), and any `import()` the parser
makes is inlined into `worker.mjs` - a "lazy" catalog inside the parser
would ship whole. Three options were on the table: hand-coding presets
(rejected: months of formula transcription, and the catalog would still be
partial), shipping the catalog table in the worker (rejected: the flat
bundle would carry every deck's catalog whether used or not), and a
build-time-compiled catalog behind lazy main-thread chunks.

Decision: the third. The worker keeps a hand-coded fast path for the 11
common presets and describes everything else on a `deferred` field on
`ShapeGeometryModel` - the encoded `custGeom` program string, or (for a
preset) just the authored adjustments, since the catalog supplies the
definition by name. The session resolves deferred geometry on the main
thread in `unit()`/`layout()`, lazily importing an evaluator plus one small
catalog chunk per shape family - the same pattern as font resolution
(DECISION_LOG on the font resolver). One interpreter serves presets and
`custGeom` because they are written in the same ECMA guide language
(`avLst`/`gdLst`/`pathLst`); multi-path presets surface as optional
`fillOnly`/`strokeOnly` layers beneath/above the main path. A chunk that
fails to load (offline, CSP) leaves the visible dashed fallback in place;
`deferred` is always stripped before a model reaches a consumer.

Verification is independent by construction: LibreOffice's shipped copy of
the catalog was evaluated with a separate implementation and compared
endpoint by endpoint against the engine (seeded into
`test/drawingml-oracle.json`, which runs without LibreOffice), the compiler
*proves* mirror/transpose aliases by evaluating both sides through the
engine itself, and the flat-worker rule was checked from the built output:
no evaluator or catalog bytes appear in `worker.mjs`.

The catalog source is the ECMA-376 Part 1 `presetShapeDefinitions.xml` (`scripts/presetShapeDefinitions.xml`, ECMA license in `scripts/NOTICE-ECMA-376-geometries`); `scripts/compile-presets.ts` compiled the full 176-preset catalog with 16 proven aliases, and every entry is oracle-checked against LibreOffice's independent conversion. The derived copy the XML was reconstructed from omits plain `upArrow`; that one entry came from LibreOffice's git-tracked copy of the same ECMA file instead (noted in the XML header). The 11 hand-coded fast-path presets stay worker-side and are never catalogued.

**Addendum (Sep 2026): the lazy family chunks are reversed; one eager table
ships in `aared/pptx`.** The per-family chunks saved bytes only for decks
that use few families, and paid for it with a family manifest, one loader
per family, eight emitted chunk files whose placement under `dist/core/` had
to be pinned for the worker URL to resolve, a failure mode (a chunk that
fails to load offline or under CSP) that left shapes as fallbacks, and a
standing question of whether a search read could trigger a download. The
whole table is ~25KB gzip. That is accepted: `scripts/compile-presets.ts`
now emits one `src/core/presets.ts` (`presetPrograms`, 176 presets, 16
aliases, byte-identical programs), `drawingml-resolve.ts` takes the table as
a parameter, and the manifest, loaders and family chunks are gone.

The load-bearing constraint is where the table lands. `core/session.ts` is
bundled into every format's chunks, so importing the table there would make
xlsx/docx apps carry every preset outline. Instead `aared/pptx`'s `open()`
imports it statically and threads it into its session through the internal
`OpenSessionOptions.geometryPrograms`; a session without it resolves no
deferred geometry. Checked from `dist/`: `pptx/index.mjs` grew from 11.8KB
to 31.1KB gzip, `xlsx`/`docx`/`worker.mjs` are byte-identical to before, no
family chunks are emitted, and `worker.mjs` contains no preset names. The
engine, program format, compiler, oracle, the `deferred`-by-name transport,
the lazily imported evaluator, and the `geometry: false` search gate all
stay.

## 25. Adaptive dark retired: every format is source-faithful

Overturns the adaptive half of §12/§15's appearance contract and closes
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
trace. The model tagged it `unsupported`, the renderer set a data attribute
nothing read, and the shape, cell or slide painted empty: no fill, no dev
marker. Only unsupported *geometry* drew the dashed outline. Pattern fills,
picture fills on shapes, `path="rect"` gradients, and tiled pictures all hit
this. So did every non-solid table-cell fill, because cells only carried a
colour string. None of the fixtures exercised them, so nothing failed.
`pptx-coverage.pptx` (renamed from `arabic-pptx-rtl-coverage.pptx`, together
with the DOCX and XLSX coverage files) gained slides 7–10, authored through
PowerPoint's COM object model, to carry every fill kind.

**Model (pre-release, breaking).** Shapes, table cells and slide
backgrounds share one `ShapeFillModel` union. It gains `pattern`, `picture`
(stretched with `stretch`/`crop`, or tiled with `tile`) and `rect-gradient`,
beside solid, linear, centered radial and `unsupported`.
`ImageBackgroundModel` (`kind: "image"`) is gone: a picture background is a
`PictureFillModel`. `TableModel.cellFills` holds fills, not colour strings,
with the cell's own `a:tcPr` fill winning over the table-style band.

**Painter.** `src/pptx/fill-paint.ts` turns every fill into one SVG paint
server. The same code paints shapes, cells (an SVG layer under the text)
and the background, so the three cannot disagree. A fill it cannot paint
returns nothing. The caller then marks it: a dashed outline in development
builds (`MARK_FALLBACKS`) and a `data-aared-fill-fallback` reason. A shape
keeps its authored stroke in shipped builds. Paint servers are sized in real
CSS pixels, not viewBox units.

**Ground truth, each from a PowerPoint render.**
- *Pattern bitmaps.* ECMA names the 54 presets but draws none. Every preset
  was rendered black on white and exported at 1280 and 2560 px. Each is an
  8×8 bitmap whose cell is one CSS pixel (the tile is 1/12 inch). The bitmap
  scales with the slide and is anchored to the slide origin: shapes placed
  off the 8 px grid show the same bitmap shifted by that offset. The table
  is `src/pptx/pattern-presets.ts`. Cells are anti-aliased: PowerPoint blends
  them when the slide is shrunk, and `crispEdges` banded them (moiré) in a
  small viewer.
- *Rect-path gradients.* `fillToRect` is read against the fill box, not
  against `tileRect` as the spec wording suggests. PowerPoint puts the first
  stop at the box's top-left corner for `tileRect l=t=-100%` +
  `fillToRect r=b=100%`. The fill is the focus rectangle plus four sides,
  each graded from the focus edge to the extent edge.
- *How stops blend.* Pixel fits against PowerPoint's own exports show two
  things. Colours mix in **linear light**, so blue-to-orange passes through
  a warm grey, not the olive of an sRGB blend. The weight follows
  **smoothstep** of the position within each pair of stops, measured on
  black-to-white 2- and 3-stop gradients. SVG and CSS only interpolate
  linearly in sRGB, so `officeGradientStops` expands each segment into 16
  intermediate stops. Slide 7's worst error fell from 44 colour levels to
  under 8. The same stops serve shape, background, cell and WordArt
  gradients.
- *Picture tiles.* Tile size is the image's natural pixels at 96 DPI times
  `sx`/`sy`, matching PowerPoint's 25 % and 30 % tiles on slides 7 and 9.

**Found alongside.** These surfaced while matching slides 2, 4 and 10:
- A gradient background SVG was *appended* after the slide scene and
  painted over it. It is now the viewport's first child.
- Every table got a grey outer frame, even styled ones that carry their own
  cell borders. PowerPoint draws none.
- `a:ln w="0"` painted nothing, where PowerPoint draws a one-pixel hairline.
  Strokes now floor at one CSS pixel.
- Tables and plain text boxes ignored the deck's default run size, so a
  size-less run fell to the browser's 16 px. They now take
  `p:defaultTextStyle`'s size per level; colour stays with automatic
  contrast (see `parseRun`). Table-style `tcTxStyle/@b` now makes header
  rows bold.

**Deliberate limits.** COM cannot author an off-center `path="circle"` or a
`path="shape"` gradient, so both stay `unsupported` (dashed in dev) until
they are added by hand on a new slide at the end of the deck and probed. Patterns in group children
anchor to their child-space position, an approximation. Tiles assume
96 DPI; an image carrying other DPI metadata would tile at a different size
in PowerPoint. Where `p:defaultTextStyle` and the master's `p:otherStyle`
disagree on size, which one PowerPoint applies to text boxes is unprobed.
They agree in both fixtures.

## 29. Real zoom fit for PPTX and DOCX; the DOCX heading/page facade gap closed

**Why.** A bench audit flagged several places compensating in the bench for
what turned out to be real product gaps — the corrected reading of "bench-only
fixes stay in bench" (that rule covers genuine harness concerns, not a stand-in
for fixing the product) meant tracing each one to its actual layer.

**PPTX: `fit-page` measured its own slide's height, and a consumer's CSS
`zoom` never overflows.** `fit-page`'s available-height measurement read the
slide's own painted height back (a box that only grows to fit what it holds
measures zero once the slide is hidden to measure the *container*), so it
never shrank anything. Fixed by hiding the viewport and measuring the
container's own `clientHeight` instead. Separately, the bench applied CSS
`zoom` to a `width: 100%` wrapper for its zoom-in/out buttons: a percentage
width resolves against the *unzoomed* containing block, so the box never
grows past its container and a zoomed-in slide clipped with nothing to
scroll to. `aared/pptx` already exposes a real `setZoom()` that resizes the
viewport's actual box (`naturalWidth * scale` px) with `transform: scale()`
painting the visual result inside it — real DOM size, so the browser's own
overflow/scrollbar math works. The bench now drives that instead of its own
CSS trick for aared PPTX cards.

**DOCX: `fit-width`/`fit-page` were a no-op.** `aared/docx`'s `wireViewer`
only ever pinned a numeric zoom via CSS `zoom`; both fit modes just cleared
it, leaving a layout-mode page at its authored pixel size regardless of the
host — the audit's "DOCX layout pages don't fit their container". Fixed with
real measurement against the widest laid-out page (mixed portrait/landscape
sections fit against the wider of the two) and a `ResizeObserver` on the
render target, mirroring PPTX's own reason for observing its container. CSS
`zoom` (not `transform`) still does the scaling, since it already correctly
resizes the layout box — the trap was pinning it on the `width: 100%` pages
column, which — for the identical percentage-under-zoom reason as the PPTX
bench bug above — doesn't shrink; the column's width is now set to the exact
unscaled page width in pixels before `zoom` is applied, matching what pages
already carry as fixed pixel sizes. Flow mode has no fixed-width page to fit
(the article already flows to the container's width like ordinary content),
so both fit modes stay a no-op there by design. Pinned by
`docx-zoom.visual.test.ts`.

**Hollow WordArt text: a merged-silhouette outline, not a per-glyph stroke.**
`-webkit-text-stroke` strokes each glyph's own vector outline before they
composite, leaving a visible seam wherever two Arabic glyphs join or overlap
(ground truth: `pptx-coverage.pptx` slide 6, "تعبئة" — PowerPoint draws one
clean outline). An SVG filter fixes it, since filters operate on the
element's whole rendered alpha: touching glyphs already read as one shape
before the outline is traced. The first attempt used `dilate(alpha) -
erode(alpha)` for a centered ring; Chromium's `feMorphology
operator="erode"` silently no-ops whenever the filtered element sits under
an ancestor CSS `transform: scale()` — which this renderer's own zoom/fit
always applies — collapsing the ring into a solid blob at any real radius.
`dilate` alone has no such problem at any scale. The fix (found by searching
for how others solve this outside a `w3c/fxtf-drafts` `#117` "outside stroke"
discussion, rather than continuing to invent a workaround) avoids `erode`
entirely: `dilate(alpha) XOR alpha`, an outside-only band from the dilated
edge back to the original edge — a deliberate, visible difference from a
centered stroke (the outline sits outside the glyph rather than straddling
it), not a bug. Verified at three ancestor scale factors, including the
live bench's own `transform: scale()`.

**`PreviewHandle.docx`: DOCX's own facade gap, closed.** `aared/docx`'s
render handle already exposed `pages`, `pageCount`, and `goToHeading()`, and
`DocxDocument.headings` already carried the outline (text, level, section
index) — but `aared/preview`/`aared/react`'s format-agnostic `PreviewHandle`
never forwarded any of it, so the bench scraped `h1`–`h6` elements and
`section[data-aared-element="page"]` inline styles from the painted DOM
instead. A slide or sheet is one navigable unit, so `index`/`goTo()` already
cover it; DOCX renders as one unit internally laid out into physical pages,
with no other-format equivalent to reuse. Added as `viewer.docx:
DocxPreviewControls | null` (`null` for pptx/xlsx) rather than new top-level
fields on `PreviewHandle`, sourced from the render handle's own `pages`/
`pageCount`/`goToHeading` plus `document.headings.list()` (independent of
pagination mode). The bench's outline rail and pager now read this instead
of the DOM.

**Bench-only, no product change.** Two more audit findings were genuinely
bench-side: the XLSX sheet-tab pill could read a locally-owned `activeSheet`
state that only ever flowed *into* the product (`renderSheet(activeSheet) →
viewer.goTo()`), never back from it — the bench now also pushes the real
`hook.index` (already mirrored from `indexchange`) into that state, so the
two can't drift even if the index ever changes some other way. And the
aared/baseline wrapper width classes (`w-full` vs `w-fit max-w-full`)
were undocumented — aared's fit modes measure and size against the real
container edge, so its box has to reach that edge; a baseline library paints
at its own natural or CSS-`zoom`ed size with no such measurement, so
stretching it to fill unclaimed space would be wrong. Documented in place,
no behavior change.

## 30. PPTX tables and text lines, matched to PowerPoint

**Why.** A pixel diff of `post-production-review.pptx` against PowerPoint's
own export found slide 5's table collapsed (six rows squeezed into a third of
their height, text clipped and unpadded) and a steady 2–4px vertical drift in
single-line text on every slide. Both traced to how aared sized boxes, not to
any one slide.

**Line boxes follow their runs** (`lineFontSizePt`, now in
`core/typography.ts`). Every PPTX `<p>` carried the deck's default size (18pt)
even when its text was 13pt, so the browser's line box was an 18pt strut and
small text sat low in it — slide 2's badge numbers, every footer (4px), every
title (2px). Word already had this rule in the DOCX paginator; it moved to
the shared module and both renderers use it.

**PowerPoint's 100% line is 1.2 × the font size.** `text-style.ts` assumed
DrawingML's single line was the font size itself. Measured: 28pt table text at
100% spacing in `pptx-coverage.pptx` gives 54.5px rows in PowerPoint's export —
the 0.1in of insets plus a 44.9px line, 1.20 × 37.3px. Unauthored spacing is
the same 100%. Exact point spacing (`a:spcPts`) was already right, which is
why slide 3's 22pt body lines always matched. The residual text offset is now
under a pixel.

**Tables grow to their rows, the way PowerPoint does.** Slide 5's frame is
authored 1in tall around six 0.52in rows; PowerPoint grows the frame, aared
shrank the rows to fit it (percentage row tracks inside a fixed height with
`overflow: hidden`). Rows now take their authored height as a minimum and grow
to hold their text; the frame height only applies to rows authoring none.
Cells stating no insets get DrawingML's defaults (0.1in / 0.05in), resolved in
core, replacing a CSS `0.2em 0.35em` guess.

**The whole table-style surface.** The reader now resolves every region
(`wholeTbl`, both band pairs, first/last row and column, the four corner
cells) with PowerPoint's precedence, each with its fill, text colour,
bold/italic, typeface (`tcTxStyle` font or `fontRef`) and six edges; an edge
comes from the region's outer border on its outer sides and `insideH`/
`insideV` inside it, mirrored for RTL tables. Direct `tcPr` edges win per side,
and an `a:noFill` edge now clears a style edge instead of being ignored.
Diagonal borders, dashed/dotted edges (`prstDash`), rotated cell text
(`vert`/`vert270`) and theme line references (`lnRef`, at the default theme's
0.5/1/1.5pt) are read and painted.

**Edges straddle the boundary.** Cell borders used to be CSS borders on each
cell, so a shared edge painted twice and a thick one ate into its own cell.
They now paint as lines centred on the boundary in an overlay above every
cell's fill, as PowerPoint strokes them.

**A face the reader lacks falls back to its kind** (`cssFontFamily`). A bare
family (`Arial`) has no fallback, so a missing face dropped to the browser's
serif default. PPTX, DOCX and XLSX now emit the face then its nearest generic
family.

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
mismatches, several of them real bugs (XLSX fit modes did nothing; a closed
document reported `"aborted"`, which `isAbort()` swallows; DOCX heading
indexes counted header and table headings, so nested tables produced
duplicates), and a second pass found seventeen places where the same
behaviour had different syntax per format or per layer (`SheetInfo.name` vs
`SlideInfo.title`; `{ slide: number }` vs `{ sheet: name ?? index }` event
details; sync DOCX `search`; `docx: { pagination }` in the facade but
top-level `pagination` in `render()`; three per-format appearance types for
one `Appearance`). The API page's own principle — same grammar,
format-native vocabulary — was not what the code did. Pre-release, so best
design wins: every caller (package, element, bench, site, tests) moved in the
same change, with no aliases.

**The rules** (stated once on the API page, § Design principle):

1. **Units vs places.** `count`/`goTo(i)` switch *units* known at `open()`:
   a slide, a sheet, the whole DOCX document. Moving within the rendered unit
   is `goTo<Place>()`: XLSX `goToCell` (was `goTo(cell)`, which collided with
   the unit verb), DOCX `goToHeading` and the new `goToPage`. This is how
   Word itself splits it: headings are structure (Navigation pane, every
   view); pages are layout (Print Layout only). So DOCX `count` stays `1` in
   both pagination modes — pages are a render result, not a unit — and
   `goToPage` rejects `unsupported-feature` in flow mode, where
   `pageCount`/`pages` are now `0`/`[]` (flow used to report `pageCount: 1`
   beside an empty `pages`).
2. **Every `goTo…` returns `Promise<void>`** and rejects `not-found`. The
   element's `goTo` and DOCX `goToHeading` were `void`.
3. **A unit reference is `{ index, title }`** (`UnitRef`), in info objects,
   unit events and search matches. `title` everywhere: it already was the
   PPTX field, and a sheet's tab name and a heading's text are both labels.
4. **A cell location is `{ address, row, col }`** (`CellRef`) — `col`, not
   `column`, in XLSX events and DOCX table matches alike.
5. **Title lookups** are exact, then case-insensitive equality, never a
   substring (`findByTitle` in `core/search.ts`); substring matching is
   `search()`'s job alone.
6. **Events describe changes.** No `indexchange` on load (the element used
   to fire one); `zoomchange` only when the value changes; `goTo(current)` is
   a no-op. The ready signal carries the starting state.
7. **Options: live chrome in `appearance`, structural options top-level
   under one name.** `preview(file, el, { pagination })` and the hooks'
   `pagination` replace `docx: { pagination }`. `setAppearance` *replaces*
   the object everywhere (a caller must be able to reset a knob).
   `sheet.read({ range, signal })` and `paintSlide({ model, target, … })`
   became object forms like every other read and render.
8. **One error model**, including a new `"closed"` code (use after
   `close()` is a bug to surface, not a cancellation to ignore) and the
   element failing with the same `AaredError` codes as `preview()`, fired
   once and rejected with the same instance.
9. **After `destroy()`** every method is a no-op; `destroyed` is on every
   handle. `XlsxViewer` now extends the shared `RenderHandleBase` instead of
   re-implementing it, and is no longer exported.

**Zoom belongs to the facade and survives navigation.** `render()` and
`preview()` accept a starting `zoom` (applied on first paint, no event);
`preview()`, the hooks and the element carry it across `goTo()`; every handle
reports the painted `scale`. The fit math is shared (`measureFitRoom`,
`fitScale` in `render-handle.ts`): room is measured with the content set
aside, so a fit can't read back its own result — the bug that made DOCX
`fit-page` a no-op in a container with no height. PPTX width fits are now
CSS-fluid (`width: 100%` + `aspect-ratio`, only the scene's transform set
from script); sizing the viewport in pixels inside a `ResizeObserver`
callback looped in React hosts.

**XLSX opens at 100%, like Excel, and its fit modes are real.** Before, the
element labelled a sheet "Fit width" while it sat at 100% and scrolled.
The grid now scales with CSS `zoom` on the sheet host; fits measure the
sheet's used area (never the viewport-fill cells) against the room outside
the host, with one scrollbar-correction pass and, for `fit-page`, one pass
for rows that paint taller than modeled below 100% (borders snap to device
pixels). A 25% floor keeps a 5,000-row sheet from painting every row. The
bench stops applying its own CSS `zoom` to aared XLSX.

**DOCX `info` and the outline.** `info.titles` is one entry — the core
properties title, else `"Document"` — and `sizes` is gone: a per-unit array
must be `count` long, and a document has no single size. The outline is one
index space (`docx/outline.ts`): body headings only, what Word's Navigation
pane lists, shared by `headings.list()`, `goToHeading`, the facade and
search matches. PPTX `titles` stay `"Slide N"`: scanning title placeholders
would break `open()`'s "200 slides open as cheaply as 2" contract.

**Facade controls.** `PreviewHandle` gained `xlsx` controls (cells) beside
`docx` (headings, pages), each one stable object per render; `preview()`
follows XLSX tab clicks into `index`/`indexchange`. The element gained
read-only `format`/`info`, a `type` setter, and zoom driven through the
render handle (so each format's own fit applies).

**`::part(page)` changed for PPTX.** The part now stretches to the stage and
the slide inside carries the drop shadow, so the slide can fit the stage
height; theme defaults moved to private properties so a consumer's
`--aared-background` / `--aared-page-shadow` win in the dark theme too.
Styling the shadow belongs to `--aared-page-shadow`.

**Packaging.** `aared/worker` is exported (the Worker script shipped but the
documented `?raw` import had no export to resolve).

## 32. Measured numbers on the home page, and how they're measured

**Why.** The home page had no evidence for its claims, and `BENCHMARKS.md` was
a skeleton of `TBD` cells. Rule: no hand-filled numbers — so the page's "By the
numbers" section is drawn from generated results, and every caption is
computed from them (a slower result reads as slower).

**One page, one browser.** `bench/measure.html` (built with the bench) is
driven by `tools/bench-measure.ts --run` through `Bun.WebView` (Chromium). The
page times bytes-in → first paint (plus two animation frames) through the same
adapter contract the bench UI uses; the harness samples, takes statistics and
reads memory. aared gets a bench adapter on its per-format entrypoints
(`bench/src/adapters/aared.ts`), used only here.

**Memory is DevTools heap usage over the page and its Workers, not
`performance.memory`.** The page API is bucketed to 10 MB and blind to Workers;
aared parses in a Worker and silurus in WASM, so a page-only reading flatters
both. `Runtime.getHeapUsage` (JS heap + DOM wrappers + buffers, WASM memory
included) is read for the page and every Worker target after a forced GC.
Whole-process readings were rejected: the browser here is a shared
WebView2 process tree. Canvas/GPU pixels are not counted, so canvas renderers
are under-reported, and the page says so.

**Navigation is timed to the committed frame, so 33 ms is its floor** (two
frames at 60 Hz). Nearly every library lands on the floor, so a chart says
nothing; the home page leaves navigation out and the per-move and per-pass
numbers stay in `BENCHMARKS.md`. (omni-doc-viewer has no sheet navigation: its
bench adapter's `renderSheet` is a no-op, so its rows there time nothing.)

**Search is measured in the browser, aared only** (the others have no
`search()`), replacing the earlier Bun-runtime figures quoted on the API page
(now 0.6 s for 300 slides, 1.8 s for the 13-sheet workbook).

**Honest exclusions and results.** The size chart leaves out silurus (1.6–1.8
MB WASM behind a stub entry) and omni/xlscalc (thin re-export entries). Sizes
count aared's format module + shared core + parse Worker, which puts one
format at 108–131 KB gzipped — larger than the lightest single-format
libraries (mdurbar, docx-preview) — while all three formats together (168 KB)
are lighter than stitching the lightest library per format (277 KB). aared
opens the 300-slide and 13-sheet files in ~48 ms (the baselines take 1.3–24 s)
but uses more memory than the smallest baselines on small files, and its
Excel sheet switch is ~42 ms against ~33 ms. All of it is on the page.

**Limits.** One machine, one engine, one run; fidelity is out of scope (the
paired Office captures own it). A full pass takes ~25 minutes (about 30 now,
with first-open and full-read readings) and must run on an idle machine, so it
runs rarely — before publishing numbers or after a change meant to move them —
and the site shows the last full run, dated; `--run fixture:renderer`
re-measures single cells.

## 33. The benchmark rulers were wrong; corrected before any product change

Overturns §32's memory reading and size method (its page, one-browser setup
and computed captions stay).

**Memory never read a Worker.** `bench-browser.ts` attached to each Worker
target and then sent `Runtime.getHeapUsage` through `Bun.WebView.cdp()` with
the Worker's `sessionId` inside the command's params. `cdp()` is scoped to the
view's page session and ignores it: a probe showed the "Worker" reading
byte-identical to the page's, `typeof WorkerGlobalScope` answering
`"undefined"`, and a Worker holding 20 MB uncounted. So every renderer that
parses in a Worker (aared, silurus, aiden0z) had its page counted twice and
its Worker never; the flat ~2 MB floor under aared's small files was the page
read twice, not a Worker's cost (an empty Worker measures 0.27 MB). The
forced GCs never reached Workers either.

Fix: `tools/cdp.ts` starts the browser with a DevTools port (`url: false`,
so it never attaches to a browser the user has open), holds one browser-level
WebSocket, attaches to each page with flat sessions, auto-attaches its
Workers, and sends every command with a top-level `sessionId`. Each session
proves its realm (`typeof WorkerGlobalScope`) before its first reading, and
`--calibrate` must read 10 MiB held on the page and 10 MiB held in a Worker as
10 MiB each, counted once, before results are trusted.

**Memory samples were contaminated.** The memory loop began right after the
last warm run's destroy with one GC, so a sample could include the previous
open's garbage or a Worker still shutting down — the negative minimums. Pages
now settle (the Worker set holds still, two GCs) before every reading, and
every sample is kept.

**Memory gains two readings, both symmetric.** *First open* — a fresh page's
first open, so loading the library's code counts for everyone (every adapter
imports its library on open); the repeat-open number alone keeps competitors'
compiled code in the baseline while a per-open Worker pays it each time.
*After a full read* — every unit visited or the document scrolled to its end,
so a saving that only holds for the first screen can't read as a win. The page
also records how many elements each open painted.

**Size weighed different things for different libraries.** Competitors were
the resolved entry file alone: dependencies they import at runtime (jszip for
docx-preview, jszip + konva + lodash-es for docx-renderer, echarts for
aiden0z) were left out, barrel entries (omni, xlscalc, silurus) scored as
stubs, and resolution fell through to Bun's global auto-install cache, so
sheetjs, docx-preview and silurus were weighed at npm-latest versions the
bench never runs (xlsx 0.18.5 for 0.20.3). aared was weighed unminified,
comments included, with core chunks a format never loads. Now size is what
the browser fetches to open each real fixture, in a fresh page with caching
off, from the same Vite production build the timings use — one minifier,
dependencies in, versions as installed (recorded per row; the run fails if a
package resolves outside the workspace). Every renderer gets a row; the
left-out footnote is gone. The site's sizes are decimal KB (they were KiB
labelled KB).

**Harness work left the competitors' cells.** The SheetJS adapter loaded
JSZip on every open and kept it for the session (its sheet RTL probe), and
the docx-preview lineage parsed `document.xml` with JSZip during the timed
open (its widest-section probe). Both now read the one part they need with
`DecompressionStream` (`bench/src/lib/zip-entry.ts`) and run before anything
is measured, so a column carries only its library — which can only help the
competitors.

**What the corrected rulers say** (published before any product change, so
the part of the gap that was the ruler stays visible): see `BENCHMARKS.md`.
In short, aared is the smallest download for XLSX and for all three formats,
about 8 KB behind docx-preview + jszip for DOCX, and well behind mdurbar for
PPTX (mdurbar draws about half of Office's preset shapes, has no RTL, and
parses on the main thread — the page says so next to the bar). On memory the
page side of a small PPTX is the lightest of all; the costs are a Worker
script stored two bytes per character, a spreadsheet grid of heavy cells, a
DOCX Worker that keeps parse state nobody reads, and a DOCX layout that builds
every page at once.

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
strings were slices of the decoded `document.xml`, so on the 40× document
the Worker held the 2.5M-character part (5 MB, two-byte) plus ~8 MB of tree
plus a duplicate model: 14.9 MB of a 52.7 MB total.

Now the index is light (styles, numbering, theme, relationships, title, and
the notes' part names), `parseDocxSections` reads the notes itself and
streams `w:body`'s children through `parseXml`'s existing stream mode (the
whole body never exists as one tree), and the Worker parses unit 0 on each
request instead of caching it. Strings that outlive a parse go through
`ownString` / `ownTree` (`core/xml.ts`): engines keep a slice of 13+
characters as a view of its whole parent, so one kept attribute value keeps
a whole part alive. That covers the styles and numbering the session keeps,
the title, and the module-level tag-name tables (capped, like the intern
table), which used to pin whichever part first introduced a long tag name.
The parsed sections are byte-identical (snapshot of the committed fixtures,
digests of the local Word exports and the 40× document,
`test/docx-sections-stream.test.ts`).

A missing `w:body` now fails in `open()`'s unit read instead of its index
step — same message — and `open()` closes the session when that read fails
(it used to leave it behind). §18's "DOCX second read: cached" no longer
holds: a second `unit(0)` from a raw core caller re-parses.

Measured (repeat open, page + Worker): docx-coverage 2.96 → 2.50 MB,
large-40x 52.2 → 38.7 MB; open time unchanged.

## 36. Renderer defaults are classes in one adopted sheet; a lighter grid

Every XLSX data cell wrote ~20 inline longhands (padding, four borders, font
size, line height, colour, alignment, white-space, overflow) and a
positioned, z-indexed text span even when empty, and the viewport fill was
one styled cell per fill column per row: 616–832 elements at about 3 KB each
on the small workbooks, most of their memory. An inline declaration block is
a style attribute plus a CSSStyleDeclaration per element, and positioned or
z-indexed elements each get a paint layer.

`src/class-sheet.ts` is the shared home now: `cssClass()` interns a block of
constant declarations at module scope into one rule, and each render root
(the document, or the shadow root a render lives in — `<aared-viewer>`'s
included) gets one constructed `CSSStyleSheet`, counted by the renders using
it, adopted before a render attaches its DOM, refilled with one `replaceSync`
only when classes were added, removed when the last render on that root is
destroyed. `follow()`, which renderers call whenever they paint, render a
window, re-zoom or resize, moves it when a target rendered detached is
attached later and restores it if an embedder replaces `adoptedStyleSheets`
— on the next of those, not instantly. No
`<style>` text is written, so a strict `style-src` needs no
`'unsafe-inline'`. Rules are `.aared-<hash>:not(#aared-0#aared-1)` —
specificity (2,1,0) — so an embedder's own CSS beats a default only with
`!important` or a selector with two IDs plus a class, close to what inline
styles gave. Inline style keeps only what a node authors, and still wins.
Options weighed: `<style>` text (fails under CSP), `insertRule` per page
(slow, fragments the sheet), custom properties alone (still one style
attribute per element), attribute-keyed rules (weaker specificity, a second
mechanism).

The XLSX grid uses it throughout: data, fill and heading cells carry
classes; an empty cell paints no text span, no `dir` and no alignment; only
a cell under a data bar or with a spill overlay is positioned; a row's run
of fill columns is one cell, and a whole fill row one cell, with the inner
lines drawn as gradients one device pixel thick at the current zoom
(`--aared-hairline`; the real columns' non-uniform edges come from per-sheet
custom properties). Hot `dataset` writes became `setAttribute` (a
`dataset` access creates a DOMStringMap wrapper per element). Deliberately
unchanged: per-cell `tabindex="-1"` (keyboard focus), fill column headings
(one per track), wrapped cells staying unclipped, and a fill row wider than
HTML's 1000-column span limit keeping one cell per column. DOCX paragraphs,
headings, list markers and pages take their defaults from classes too; PPTX
already had the lightest page of any renderer and is unchanged.

Measured (repeat open, page + Worker): post-production-tracker 3.22 →
2.12 MB (page 2.28 → 1.16), xlsx-coverage 2.82 → 1.49 MB (page 1.94 →
0.61), large-13-sheets 4.53 → 3.20 MB; first open on the tracker 3.75 →
2.70 MB; a sheet switch 42 → 33 ms (the frame floor).

## 37. DOCX layout pages paint lazily

Layout mode built all 121 pages of the 40× document at once — 23,262
elements, 33.9 MB of DOM, style and layout — and its tab-stop pass forced
layout of every page. Pages now carry `content-visibility: auto` where the
browser supports it: every page stays in the DOM (find-in-page, selection,
fragment links, printing), but a page off screen skips style and layout
until it is scrolled near. Its fixed box already sizes it while skipped. Tab
stops lay out per page the first time a page is shown
(`contentvisibilityautostatechange`, plus one frame for pages relevant from
the start), and again for shown pages once web fonts land; without the
feature the eager pass runs as before. The `span.paragraph-text` wrapper is
gone from DOCX (nothing read it; one element per paragraph). Windowing pages
out of the DOM entirely was set aside: it would cost browser find on unseen
pages and rework navigation, captures and tests.

Kept on purpose: the paginator's metrics cache stays keyed weakly on the
(read-only) model, so a re-render of the same document — `<aared-viewer>`'s
theme toggle, a mode switch — does not re-measure every block.

Measured (repeat open, page + Worker): large-40x 38.7 → 16.9 MB (docx-preview
24.9), open 606 → 230 ms; docx-coverage 2.50 → 2.05 MB. The saving is
first-view: after scrolling to the end the 40× document holds 34.2 MB,
because a page keeps its layout once shown — the "after a full read" column
reports it.

## 38. The Worker is reused between opens

Every `open()` started a Worker and every `close()` terminated it, so every
document paid for a new isolate and for compiling the parser again (~0.5 MB
and some milliseconds), and `preview()` / `<aared-viewer>` paid twice: once
to detect the format, once to open it. Now a closed document's Worker is
parked for the next open (`core/session.ts`, `acquireBridge`): one idle
Worker per worker init (the default, or a Worker URL), stopped after 15 s
idle or when a newer one parks; two documents open at once still get two
Workers. Only a clean Worker is parked — nothing was cancelled, failed to
post, or was still in flight at close — because a cancelled synchronous
parse can only be stopped by terminating, and a Worker built by a caller's
`() => Worker` factory is the caller's: used once, terminated on close.
`close(sessionId)` rejects only its own session's reads. The literal
`new Worker(new URL("./worker.mjs", import.meta.url), { type: "module" })`
branch is untouched (§19), and a dist test now checks it with Vite's own
detection pattern. State that now outlives a document is bounded and owns
its strings: the XML name tables (§35) and the number-format memo.

Honesty: for a single document this mostly moves the Worker's fixed cost
from the open into the page's baseline, like a main-thread library's loaded
code. So the benchmark reports first-open memory (a fresh page, the Worker
and code included) next to repeat-open memory, and the idle Worker a page
keeps between opens as its own column.

Consequence: a closed document's memory is released when its Worker next
collects garbage or stops, not the instant `close()` returns.

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
current one, found in the playground:

- **The page scrolled to the top on Next.** A new slide went into its
  container with no height (its content is absolutely positioned), so the
  layout the shrink-to-fit pass forced shrank a page whose height was the
  slide, and the browser clamped the page's scroll position. The slide now
  goes in at its full-width box (`width: 100%` and its aspect ratio), which
  is the largest any fit makes it, and the fit sizes it. Separately,
  `measureFitRoom` sets the content aside to measure the room and puts every
  scrolled ancestor's position back afterwards.
- **A short sheet showed without its fill, then grew.** The viewport fill
  (the empty rows and columns that continue the grid) needs real sizes, and
  the first window was deferred to the next frame, so the grid painted at its
  data height and then filled the viewport. It now measures and paints in the
  same task once attached; a target that isn't in a document yet still
  paints on the next frame.
- **The fill grew a row per frame in a `max-height` box.** A scroll box that
  is as tall as its content up to a `max-height` measured, as its room, the
  grid already inside it, so each pass added rows until the limit stopped it.
  Its room is now the limit.

The playground's demo stages had a 12 rem minimum height that "fit page"
fitted the slide to, so React, JS and JS-full-control drew smaller slides
than the element: the minimum now applies only while a stage is empty. Those
three demos also get the same Previous / Next / zoom bar the element draws,
built from the public API only (`site/src/lib/demo-controls.ts`).

## 42. npm name: `@aared/viewer`, one package, reserved with a placeholder

The plain name `aared` cannot be published: npm's similarity check rejects it
against an unrelated, empty `aa_red` (403 "too similar"). Scoped names skip
that check, so the project uses an npm org. `aared` is the org and the
package is **`@aared/viewer`** (matches the `<aared-viewer>` element;
`@aared/aared/pptx` would repeat itself). `@aared/viewer@0.0.0` is a
placeholder holding only a README, published 2026-09-25; the repo's code is
not in it. The repo tag `v0.0.0` marks that point.

**One package, subpaths, stays** (§20): `@aared/viewer/pptx`, `/xlsx`, `/docx`,
`/preview`, `/react`, `/element`, `/worker`. Per-format packages were
rejected again: the single Worker holds all three parsers, so splitting would
need per-format workers (measured 20-26 KB saving for one-format pages, at the
packaging cost §20 records) plus a shared core package to release in step.
The org scope leaves room for `@aared/fonts` and thin re-export packages later
without breaking anyone.

**Versions.** The placeholder is `0.0.0`, never `1.0.0` (a stub must not look
stable; a published version can never be reused). The first real release is
`0.1.0`: semver allows breaking changes in a `0.x` minor, which fits an API that
is still settling. `1.0.0` comes when the API is frozen (see ROADMAP "Before
first publish"). The "best design wins over backward compatibility" stance in
AGENTS.md ends at the first real release, not at the placeholder.

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
physical sides and reads `a:tblPr rtl`. Pinned by section 5 of `fixtures/docx-coverage.docx`
(Word-authored, every side different) in the parser test and a DOM read-back
test.

## 44. Fixes found by the showcase documents (Sep 2026)

Building polished showcase files for the launch video (PowerPoint, Excel and
Word output compared page by page with aared) surfaced real bugs. Each was
traced to its cause, checked against native Office, pinned in the existing
coverage fixtures (edited through Office COM; no new fixture files) and
tested by reading the DOM back.

**PPTX lines and connectors were dropped.** PowerPoint writes every line and
connector as `p:cxnSp`; the shape-tree walk only knew `p:sp`, so timelines,
dividers and today-markers vanished with no fallback marker. A connector is a
shape without text (same `spPr` geometry and line), so it now parses through
the shape path; `straightConnector1` joins `line` in the hand-coded fast path
(same path; the catalog test now checks fast-path presets against the oracle
too). Shape outlines also paint authored `a:prstDash` (dash/dot families,
scaled to the line width) — the model carried it for table edges only.
Pinned by `pptx-coverage.pptx` slide 11 ("خطوط وموصلات").

**Headings inherited the browser's `<h1>` bold and 2em.** DOCX headings paint
as `<h1>`–`<h6>` for outline semantics; runs only set `font-weight: 700` when
bold, so a Heading 1 without `w:b` still painted bold (and a sizeless run at
2em). The shared block class now resets weight and size; runs decide, as in
Word.

**Left-to-right runs in right-to-left paragraphs (probe run-direction).** Word
and PowerPoint keep a run that is not right-to-left (DOCX: no `w:rtl`; PPTX:
a left-to-right `lang`) together left to right: "71%", "2.1%" and
"15 Oct 2026" show as written. The Unicode bidi algorithm alone makes a
number after Arabic an Arabic number, detaching "%" ("%2.1") and splitting
dates ("Oct 2026 15"). Such runs (with no Hebrew/Arabic letters) now carry
`TextRunModel.ltrRun`, and the shared run painter adds a left-to-right mark
at both edges through `::before`/`::after` — generated content, so it is never
copied, searched or counted. Isolating each run instead was rejected: a date
PowerPoint splits into two en-US runs ("15 " + "Oct 2026") would reverse.
Deliberate non-changes: right-to-left runs keep plain bidi (Word and
PowerPoint disagree there: Word "71%", PowerPoint "%71"), and a left-to-right
run holding Arabic words in a left-to-right paragraph (the existing
`عربي + إنجليزي` gap) still needs its own probe. The PPTX renderer now adopts
the shared class sheet, as DOCX and XLSX already did.

**Long DOCX tables no longer blank a page.** In layout mode a table taller than
a page took the oversize path: the page it would start on was closed (in the
showcase term sheet, holding only a `pageBreakBefore` paragraph — a blank page)
and the whole table went on the next page, its tail clipped. Word breaks
tables between rows. The paginator now splits a table at the last row that
fits (rows are never split; a boundary inside a vertical merge is skipped, so
each part stands alone); only an unsplittable table (one row, or one merge
spanning it) still takes the oversize path. The term sheet went from 13 pages
with clipped tables to 12 complete ones against Word's 10.

**A vertical-merge boundary must also yield to `w:keepNext`.** The table split
above only skipped a merge; a row whose paragraphs all carry `w:keepNext`
(Word treats the whole row that way — never splitting a title row from the
row it introduces, for instance) can still be split from the row after it.
`DocxParagraphLayout.keepNext` is now parsed and carries through headings,
list items and table rows; the paginator's keep-with-next chain (any block
that keeps with the next, plus the first placeable piece of the one after
it — a table's own leading run of `keepNext` rows counts as one piece here,
same as `splitTable` already treats it) moves to a fresh page when the chain
fits one but not what is left of the current one. A first attempt read a
generator's "keepNext on every row but the last" (a common way to say "try
not to orphan a row") as license to relax the row-boundary gate itself,
reasoning that OOXML has no literal "keep this row with the next row"
property — that surfaced correctly-fitting content (an 11-row milestone
table Word placed whole) but also let a section's title row split from its
first row, which Word never does; reverted (`aared_launch` review caught
both, side by side against native Word).

**Table rows and other blocks are now measured, not only estimated
(DECISION_LOG §21).** The 12-vs-10-page gap above was the estimator's average
glyph-width guess, worst on the term sheet's dense two-column bilingual table.
`docx/index.ts` gained `measureFlow()`: it lays every section's top-level
blocks out once, off screen, at that section's real body width and the
page's own text defaults, and reads back each block's margins and
border-box height and each table row's grid-track height. `paginate()` takes
this as `PaginateOptions.measured` and prefers it over the estimator wherever
it answers; the estimator remains the fallback (headless use, and any block a
render did not place standalone). One forced layout for the whole document.
Reading back a table row's height needed table cells to lay out as CSS grid
rows, which the DOM shell already was. Combined with the two fixes above, the
term sheet's section breaks now land exactly where Word's do on every page —
no more numbers split from their titles anywhere in the document (the bug
that prompted this work). Section 4's own two-column table (a bold title row
plus four `Heading2` clause rows) did measure about 13% taller here than in
Word at the time, concentrated in its two heading-styled rows (about 21-24%
over; its two plain-paragraph rows were within 3%) — just enough, on this one
table, to push an otherwise-fitting 11-row milestone table onto a page of its
own (aared: 11 pages; Word: 10). That gap is now closed (see "Exact line
pitch" and "Table cell side insets," below): it was never a line-height
difference in the CSS the paragraph itself set, but every run on an
exact-pitch line stretching that line past its own font's ascent and descent
— the heading rows pair a Calibri number with Sakkal Majalla, the worst
mismatch in the table. With both fixed, the term sheet measures 10 pages
against Word's 10.

**Exact line pitch is the paragraph's, not each run's.** `w:lineRule="exact"`
(and PowerPoint's points spacing, `a:spcPts`) set a paragraph's line height,
but every run still painted at the browser's own rule: a span's line box is
centred on its own font's ascent and descent and aligned on the shared
baseline. A Calibri digit run inside a Simplified Arabic line stretched an
exact 24px line to about 25px; a Calibri section number beside a Sakkal
Majalla heading stretched it further still — the "heading-styled rows
measure 21-24% taller" residual recorded above was this same stretch,
wherever a differently-metriced run shared an otherwise-exact line. Word and
PowerPoint keep an exact line at its authored pitch whatever it holds — the
pitch belongs to the paragraph, never to a run on it. Fix: a run on an
exact-pitch line now paints at `line-height:0` (`EXACT_PITCH_RUN`,
`text-style.ts`), contributing no line-box height of its own, so the line is
set entirely by the paragraph's own strut; `exactLinePitch(box)` decides
which paragraphs qualify (a set `lineSpacingPt` without `lineSpacingAtLeast`
— see "Line rule kept," below), and both `docx/index.ts` (paragraphs,
headings, list markers) and `pptx/index.ts` (runs) apply it wherever a run
paints. A line whose rule is `atLeast` — a floor, not a pitch — is untouched:
its tallest run still grows the line past the floor, matching Word. In the
term sheet's justified Arabic cell, four lines measured 99px and now measure
exactly 96px (4×24). PowerPoint's `a:baseline` shift was painted as a
`vertical-align` percentage, which CSS takes of the run's own line height:
zero on an exact-pitch run, and 1.2 times too much elsewhere, since DrawingML
defines the shift as a share of the font size. It now paints in `em`, so
DrawingML's font-size meaning holds whatever the line. Pinned by both
coverage fixtures, extended through Office for it (`tools/probe/docx-justify/
extend.py`): the DOCX delivery-tracker update line is exactly 18pt with its
date and percentage in Calibri beside Arial Arabic (it measured 25px under
the old rule, 24px now), and the PPTX "Mixed direction" box on slide 11 is
exactly 30pt with its "2.1%" in Calibri; both visual tests read the line
height and the runs' `line-height: 0` back from the DOM. Word's save put the
org's label marking back into the DOCX headers and recreated the even and
first-page headers to carry it. `tools/probe/pptx-lines/strip_label.py` now
also strips a Word package: it removes the marking runs, then drops any
even/first header or footer left with no run.

**Line rule kept, and a direct-formatting merge bug fixed alongside it.** Two
faults in the paragraph-property resolver, found while chasing the line-pitch
fix above. First, `w:lineRule="atLeast"` was read but discarded —
`resolveParaProps` only ever produced `lineSpacingPt`, so an `atLeast` floor
and an `exact` pitch looked identical to the renderer and the pagination
estimator, and a style's `auto` multiple layered over a docDefault's exact
pitch left both `lineSpacing` and `lineSpacingPt` set instead of the later
level replacing the earlier, as intended. Second, `resolveParaProps` parsed
direct formatting into its own throwaway object and copied it onto the merged
style-chain result key by key — a copy that could only add or overwrite a
key, never clear one, because a `delete` inside `readParaProps`'s own
accumulation logic (a later `w:spacing` rule replacing an earlier one, a
`w:pBdr` composing per side) only ever touched the object passed to that one
call, not the merged object copied from afterward. Concretely: a paragraph
with a styled top border that then set only a direct left `w:pBdr` painted
with the top border gone — the `pBdr` case's own promise ("a direct left
border composes with a styled top border rather than replacing it") held
only inside a single `readParaProps` call, not across the docDefault → style
→ direct chain. Fixed by adding a model field,
`DocxParagraphLayout.lineSpacingAtLeast`, set (and cleared by the opposite
rule) in the same `w:spacing` branch that already resolves `lineSpacingPt`;
and by resolving direct formatting straight into the same merged object the
style chain already built (`resolveParaProps` now calls
`readParaProps(directPPr, merged)` last), instead of merging a separate parse
result afterward. `styleId` — previously read off a full throwaway parse of
the direct properties, only to get one field out of it — is now read
directly from `directPPr`'s own `pStyle`, since the ancestry lookup needs it
before the merge exists. The pagination estimator's `textBlockHeightPx`
treats `atLeast` as a floor: `max(pitch, Word's single line)`, never the
pitch itself. Blast radius: every paragraph property that accumulates rather
than replaces (tab stops, `w:pBdr` sides) now composes correctly under direct
formatting too, not only across the style chain; `numPr`/`ilvl`,
`pageBreakBefore` and `keepNext` are plain presence-wins-if-set fields,
already correct either way. Pinned by parser tests over small synthetic
packages: exact and at-least rules through docDefaults, style and direct
levels (each one replacing the one it inherits), and a direct left border
composing with a styled top one.

**Table cell side insets follow Word's own centred-border rule.**
Ground-truth probe `tools/probe/docx-cell-border/` (real Word, side borders
at 0.5pt / 3pt / 6pt): a side border moves the cell's text 0.3 / 1.4 / 2.45pt
from the grid line — Word centres a border on the shared grid line and keeps
the text a margin beyond its own inner half, whichever cell paints it. aared
took the whole painted border (at least one device pixel — the same rule
`borderSnapSlackPx` uses top and bottom) out of the padding on the side that
painted it, double-counting the outer half no cell needs to give up. Fix:
`sideInsetPx(margin, edge, painted)` computes padding as the margin plus half
the edge on that side, less what this cell itself paints there. A shared
inside divider is painted once, by the later cell as its start border: that
cell subtracts the painted width, while its neighbour paints nothing on that
side and keeps the full margin plus half. PPTX
tables are unaffected: they already paint edges as a centred overlay
(`cellEdgeOverlay`), which takes no text width. Pinned by two visual tests.
In the right-to-left table, the first column's right padding beside its 3pt
start border is 4.8 + 2 − 4 = 2.8px (it was 4.8px before), and its left
padding, with no border on that side, stays at the 0.3in margin (28.8px). In a
left-to-right table with 0.5pt outer edges and a 0.25pt inside divider, both
columns' paddings are checked, including the divider side of the column that
does not paint it.

**Word 2013 justified space shrinking, approximated conservatively.**
Ground-truth probe `tools/probe/docx-justify/` (real Word, Courier New,
binary search on text width): in `w:compatibilityMode` 15 or later, a
justified or distributed paragraph — left-to-right or right-to-left — fits
one more word on a line by shrinking each space, up to about a quarter
(24.6-24.8% whenever that bound applied) and about a third of the word it
fits, in total; never when left-aligned, never in Word 2010 mode (14). CSS
breaks lines at fixed widths and has no equivalent shrink-then-refit pass, so
aared approximates: the parser reads `word/settings.xml`'s
`w:compat/w:compatSetting[@name="compatibilityMode"]`
(`wordLayoutShrinksSpaces`, `DocxIndex.justifyShrinks` — computed once per
document and shared by every part, body, headers, footers and footnotes
alike, since compatibility mode is document-wide) and sets
`DocxParagraphLayout.justifyShrink` on a justified/distributed paragraph
only; the renderer adds a small negative `word-spacing` (`JUSTIFY_SHRINK`,
`-0.01em`) to such a line. The amount stays under both of Word's limits for a
line of up to about twenty spaces, so it can never pull down a word Word
would have pushed to the next line — it only catches near misses, the case
that mattered: the term sheet's justified Arabic cell broke "تيار" to a
fourth line for 0.6px. Known remaining limits: one row still breaks one word
earlier than Word because Word shrinks more there than this conservative
bound allows, and English Calibri text measures 0.3-0.5% wider in Word than
in the browser regardless of shrinking (glyph-advance rounding), so a few
English lines wrap differently — deliberately not faked with letter-spacing.
Snapshot digests in `docx-sections-stream.test.ts` were re-pinned after
proving, by hashing with the new field stripped, that `justifyShrink` is the
only model change. A parser test covers the gating: modes 15, 14 and absent;
justified, distributed and left-aligned; and the coverage fixture's justified
paragraphs. A visual test reads the negative `word-spacing` back from a
justified paragraph, and `0px` from a left-aligned one.

**A `w:lastRenderedPageBreak` mid-paragraph is not a break point.** Word writes
one at the position its own last layout put a page break, which can land
inside a run's text as easily as before a paragraph — the parser treated any
occurrence as a preferred break, splitting the paragraph itself in half (a
heading's number landed alone on the previous page, its title top-of-page on
the next: the exact bug pages 3/4/6 of the showcase term sheet showed). The
hint is now a signal read only when no text of the paragraph has been seen yet
(a preferred break *before* the paragraph); a mid-paragraph occurrence is
consumed and dropped, matching what it actually records (Word's stale layout
cache, never a place to split live content). The same hint inside a table
cell has nowhere to go (tables break between rows) and is dropped there too.

**tools/excel-render printed US Letter, edge to edge.** Excel's "Scale content
for A4 or 8.5 x 11 paper sizes" (`MapPaperSize`, on by default in US-locale
installs) mapped the hardcoded A4 policy to Letter on export, and margins were
set as inches where Excel reads points. The exporter now turns `MapPaperSize`
off for its run (restoring the user's setting after — it persists), commits
page setup as one batch, and converts margins with `InchesToPoints`; pages are
A4 (1241×1754 at 150 dpi) with content at 0.7in/0.75in, matching the aared A4
fit the capture test already assumed.

**A frozen row's own row number scrolled away from its row.** Scrolling a
sheet with frozen rows inside the launch-video's bench card left the frozen
rows pinned at the top, correctly, beside whatever row numbers had scrolled
underneath them — Excel keeps a frozen row's number with it. The row-header
cell only ever stuck to the inline-start edge (`paintFrozen` handles the
column axis); it never got the row axis's own `top` offset and z-index a
frozen data cell gets. It now does, at the same `frozenRowOffset` a frozen
row's cells already use, above the scrolling numbers passing beneath it.
Pinned by `xlsx-coverage.xlsx`'s Shot Log sheet (4 frozen rows, ~100 total) in
the render test, scrolled and read back from the DOM.

**A content-sized sheet let its container scroll it, so frozen chrome and the
horizontal scrollbar disappeared.** The same bench card exposed a second bug
behind the first: `<aared-viewer>` and the bench give the sheet no height of
its own (content-sized, matching Excel's own grid, which never imposes a
scrollbar until the window does), and aared's `syncViewport` only ever grew
that box to its full content height, adopting a clipping ancestor's visible
room as a `min-height` floor and no more. So the *card* scrolled the whole
sheet: frozen rows, their numbers and the column letters all scrolled away
with it (position:sticky pins to the nearest scrolling ancestor, which was
the card, not the sheet), and the sheet's own horizontal scrollbar sat below
the fold at the sheet's full height, reachable only by trackpad. Excel's grid
scrolls itself inside the window's own visible area; the container never
scrolls it.

Fix: `syncViewport` now also caps the scroller's own `max-height` to the
clipping ancestor's visible room (a ceiling, alongside the existing floor),
whenever that room is at least 120px (`MIN_SCROLL_ROOM_PX` — below that, a
box too short to show anything useful, so the container scrolls it as
before). The sheet then scrolls inside its own bounded box: frozen rows, row
numbers and column letters all stay pinned in view, and a normal horizontal
scrollbar sits right under the visible grid. With no clipping ancestor
(a full-length embed), nothing changes — the page still scrolls the sheet.

A first version of this cap broke `tools/xlsx-render`'s print-capture
harness, which deliberately grows the scroller to an oversized explicit
height to force every row into the DOM for measurement, undisturbed by
virtualization or windowing. `syncViewport`'s own cap from an *earlier* pass
was still in place when that resize happened, so the change-detection that
decides whether to recompute measured the scroller's height through its own
stale cap and never saw the resize. Fixed by lifting a previously-set cap
before measuring (only when one is actually set, so the common bounded-viewer
path — which never sets one — pays no extra layout cost), and always
recording the *uncapped* height for the next comparison, never the capped one
the pass just applied. Pinned by the same render test: it renders the sheet
inside a bounded card, scrolls it, and checks the card's own `scrollTop`
stayed at 0 while the sheet's box scrolled, with the column headers and the
frozen row's number both still in the visible area — and by the print-capture
suite continuing to pass, which now exercises exactly the resize-after-cap
sequence that broke the first attempt.

## 45. Font fidelity: a delivered substitute is actually used, waited for, and reported (Sep 2026)

Follow-on to §44's font-metrics work, prompted by a direct question: aared's
layout fidelity today depends on the calibration machine having every
requested Office font actually installed (Calibri, Simplified Arabic, Sakkal
Majalla — all real here). On a machine without them, what happens?
Investigating the answer surfaced three real bugs in `aared-fonts` and the
renderer, on top of the substitution table itself needing correction.

**A delivered substitute was registered but never selected.** `createResolver`
built each `FontFace` under the *substitute's own* Google Fonts name (`new
FontFace("Carlito", …)`), while every renderer painted `font-family: Calibri`
— a single bare name, no fallback stack of its own. Nothing connected the
two: on a machine without Calibri, the browser fell straight to its own
generic default, and the registered Carlito face sat in `document.fonts`
unused. This was an open question recorded in ARCHITECTURE.md and never
verified until now. Fix: `fontAlias(family)` (`core/typography.ts`, exported
from `aared/core`) is the one shared naming rule — `cssFontFamily` now stacks
the requested name, `fontAlias(requested)`, then the nearest generic, and
`createResolver` registers every delivered face under `fontAlias(requestedFamily)`,
never the requested name (would shadow a reader who really has it installed)
and never the substitute's own name (nothing would select it). Verified
end-to-end, not just at the unit level: a visual test renames a run to a
family guaranteed to exist nowhere, registers a face under its alias sourced
from a dramatically different local font (`local('Courier New')`, monospace),
and confirms the *painted* width matches Courier's, not the generic
sans-serif's — this is what actually closes the open question, a unit test
against a fake `FontFace` could not.

**The substitution table's Latin entries pointed at fonts Google Fonts
doesn't host, and the Arabic entries weren't picked by measurement.**
Fetching each `FONT_SUBSTITUTIONS` value from the CDN directly (a one-off
script, not a test — no network in the suite) found Liberation Sans/Serif/
Mono return 400: never hosted there. Arimo/Tinos/Cousine are — the actual
Google Fonts builds of the same Ascender-designed metric-compatible family
Liberation ships locally (confirmed: 0 of 95 ASCII glyphs' advance widths
differ from the real Arial/Times New Roman/Courier New installed here,
`tools/probe/font-metrics/metric-clone-check.py`; Calibri/Carlito likewise 0
differ; Cambria/Caladea differ only by sub-pixel units-per-em quantization,
under 0.13%). So every substitute Arial/Times New Roman/Courier New document
was silently falling back to the generic default the whole time, not to
Liberation as documented — the table said "metric-compatible" for a mapping
that had never actually been exercised. Separately, the four Arabic entries
mapped by name association (Traditional Arabic/Sakkal Majalla → "sounds like
Naskh" Noto Naskh Arabic, Simplified Arabic → Noto Sans Arabic — a different
letterform style entirely, not even the same design family as the other
three). Measured instead, in two passes. The first (three candidates, three real
paragraphs from `fixtures/docx-coverage.docx`, real shaping via
`CanvasRenderingContext2D.measureText`, not a per-glyph table — Arabic's
contextual forms make an unshaped sum wrong) picked Amiri for all four fonts:
it beat Noto Naskh Arabic despite the "Naskh" name and Noto Kufi Arabic (a
geometric, angular style — wrong shape family, confirmed by the numbers, not
assumed from the name). That pass was too narrow, and it had never been checked
against outside sources. The second pass (web research, then a survey:
`tools/probe/font-metrics/arabic-survey.ts`) is what the table now rests on:

- **Sources.** Searches found no open metric clone of any of the four fonts;
  Microsoft Typography describes Sakkal Majalla as Naskh with compact
  proportions and Arabic Typesetting as a *modern reading of Naskh* — the
  earlier code comment calling it "decorative/Diwani" was wrong — and
  Wikipedia describes Simplified Arabic as a typewriter-derived "simplified"
  design, not a Naskh revival. The one page that lists metric clones (the
  ArchWiki) was unreachable, so "none exists" rests on the searches alone.
- **Method.** 59 samples (every Arabic-only run in the three fixtures) measured
  in the real Windows font *files* (served locally and loaded through
  `FontFace`, so the measurement machine's font choices cannot leak in)
  against eleven candidates loaded from the same Google Fonts files the
  resolver delivers, each load verified — a candidate that silently fell back
  would measure a generic font and look plausible. Ranked by the spread (sd) of
  per-sample width ratios, because line breaks depend on the spread while the
  mean is exactly what `size-adjust` removes; then by the height of alef and of
  the seen body after adjusting. A candidate had to beat the incumbent by a
  quarter to displace it.
- **Result.** Sakkal Majalla moves to Markazi Text (sd 0.036 against Amiri's
  0.060: 41% less). Simplified Arabic moves to Lateef (0.023 against 0.044:
  47%). Traditional Arabic and Arabic Typesetting keep Amiri (Lateef would be
  12% better and 10% worse). An early worry — that a large `size-adjust` (Lateef
  needs 130%, Amiri 67% for Arabic Typesetting) would distort the letters —
  did not survive measurement: the alef and seen heights after adjusting land
  at the original's proportions for both.
- **Bold.** Measured per weight: bold differs from regular by up to 5.6% in
  width (Traditional Arabic 80.0% against 84.5%), enough to matter for
  headings, so `FontCalibration` gained an optional `bold` set and the resolver
  picks by each face's weight.

**Calibrating one delivered face against several different real fonts.**
Each requested family has its own `FontFace` descriptors computed from the
measured width ratio: `size-adjust` scales the substitute's advances to match,
and `ascent-override`/`descent-override` are set so the *resulting* line box
matches each real font's measured ascent/descent. Getting the override numbers
right needed one more empirical check, not a spec reading: does `size-adjust`
scale an override percentage, or does the override replace the metric outright
at the unscaled em? Tested directly (`tools/probe/font-metrics/
size-adjust-interaction.ts`, four descriptor combinations on the same local
file): `final_ascent = ascent-override% x size-adjust% x font-size` — the
override is multiplied by size-adjust's own factor, confirmed by measuring
30 = 0.60 x 0.50 x 100 exactly. The published `FONT_CALIBRATION` numbers are
divided out from that formula, not assumed from the spec's own wording (which
reads ambiguously either way). Requested families sharing a substitute (Arial,
Times New Roman, Cambria and Simplified Arabic all deliver Arabic through
Lateef) still register as separate `FontFace`s, one per alias, each carrying
its own calibration — `createResolver`'s CSS fetch is cached per *substitute*
(one round trip serves all of them), but the constructed-and-registered faces
are cached per *requested family*, deliberately not collapsed to one.

**Most Arabic in real documents is set in Arial, not in an Arabic-named
font, and the substitute for Arial has no Arabic glyphs.** Resolving the
families of every Arabic-containing run in the fixtures showed the four
calibrated fonts carry none of them: Arial carries 183 runs, Times New Roman
12, Aptos 9, Cambria 1. Arimo and Tinos have no Arabic subset (checked against
Google Fonts' own metadata), so that text fell through to whatever `sans-serif`
the reader's system has. It never showed on the development machine because
on Windows the generic sans-serif *is* Arial's own Arabic — the same
coincidence that misled the fidelity detector. Ground truth for what to match:
Word's PDF export (`tools/probe/font-metrics/word-arabic-fallback.ps1`, one
sample per page, real span widths read with PyMuPDF; an earlier version read
`Range.Information()` caret positions and was discarded — they flip between
the two ends of a bidi run and gave garbage) agrees with the canvas
measurement of the same installed fonts to within about 3% (Simplified Arabic
0.983 of Arial in Word, 0.988 measured), which is what licenses using canvas
numbers. Findings: Arial's and Times New Roman's Arabic match Simplified
Arabic's widths within 1.3% (Wikipedia's account of Simplified Arabic as the
Arabic supplement of Arial and Times New Roman is right); Calibri 6 carries
its own Arabic; Cambria has none and Word draws that text in Times New Roman.
So the resolver adds a *second face to the same alias* for Arial, Times New
Roman, Cambria, Calibri and Calibri Light, limited by `unicode-range` to the
Arabic script blocks and calibrated per original. Composition was proven in a
real browser first, not assumed (`alias-composition.ts`): Latin under the alias
measures exactly as Arimo alone; Arabic exactly as calibrated Lateef alone;
spaces and punctuation stay with the Latin face (as in real Arial); the vertical
metrics come out at Arial's (36/8 at 40px, against 45/25 uncalibrated); and add
order is irrelevant. The face uses its own fixed Arabic range rather than
Google's subset range, which also claims punctuation Arimo owns. Only upright
Arabic faces are registered: Arial and Times New Roman have no Arabic italic,
so Word slants Arabic synthetically and so does the browser once handed an
upright face.
Two costs are kept off documents that do not need them. The Arabic face is
not `.load()`ed eagerly (the Latin faces keep the earlier behaviour), and the
font wait and the fidelity probe used to put Arabic letters in their sample
unconditionally, which made even a Latin-only Arial document fetch about 130 KB
(regular plus bold) — caught by a test that asserts a Latin-only document
never requests the face. It failed even after the wait was made conditional,
because the fidelity probe measures a sample through canvas `measureText`,
which itself makes the browser fetch whichever face covers the text.
`collectFontUse` now reports whether the model has Arabic text in the same
walk that collects the families, and both the wait and the probe use it. With
the wait's Arabic letters removed the test fails too, so it guards both.
Deliberate non-changes: Courier New's Arabic is monospace (every proportional
candidate is off by 45% or more on individual paragraphs, so no calibration
helps and none is offered); Tahoma and Segoe UI are not in the table at all;
Aptos has no Windows font to measure against (superseded by §49: Office's
cloud-font cache has the real files); and the explicit `.load()` the
resolver calls on every Latin subset face (which downloads all of Arimo's
Cyrillic, Greek and Hebrew subsets whether or not they are painted) predates
this and was left alone (removed in §50: faces now load when painted).

**A registered substitute's glyph data can still be loading when a renderer
measures painted text.** `createResolver` calls `FontFace.load()` without
awaiting it (by design — painting should not block on every font file), and
the session (`core/session.ts`) only awaits the resolver's own promise
*settling*, which happens once the `FontFace` object exists and is added to
`document.fonts`, well before its bytes finish loading. DOCX layout mode's
`measureFlow` and PPTX's `normAutofit` shrink pass both read painted
`getBoundingClientRect`/`scrollHeight` synchronously; before this fix, that
measurement ran with whatever was already painted — an installed real font,
or the browser's own unrelated fallback — regardless of a substitute
mid-flight. Fix: `awaitDocumentFonts(model)` (`text-style.ts`, shared by both
formats) walks the model for every `fontFamily` it uses (`collectFontUse`, a
sibling of the `collectStrings` walk `core/session.ts` uses to find which
families to resolve, which also reports whether any run has Arabic text) and calls `document.fonts.load()` on each family and its alias,
racing the result against a 2-second timeout so an offline or blocked CDN
never hangs a render. DOCX layout pagination always awaits it before
`measureFlow`; DOCX flow mode awaits it too, even though flow mode's own
CSS text reflows automatically once a font loads with no help needed —
`awaitDocumentFonts` there is only for the fidelity flag below, which reads
whichever font actually ends up painting. PPTX only pays this cost for a
slide that actually has `normAutofit="shrinkText"` text
(`hasAutofitShrinkText`, checked against the raw model before painting, so
the common slide with no autofit text never waits); the synchronous
`paintSlide` entry point and a no-autofit `render()` compute font fidelity
from whatever is loaded at that moment, which can occasionally read stale
for a font mid-flight — a documented, narrower scope for PPTX than DOCX.
(Superseded for PPTX by §48: PowerPoint never measures on open, so the shrink
pass and this wait are gone.)
Verified with a genuine race, not a single-tick approximation: `document.
fonts.load()` on an unregistered family (installed or nowhere) resolves in
under 1ms, and even a `FontFace` from an in-memory `ArrayBuffer` reports
`"loaded"` synchronously — a `local()` or binary-data source has no
asynchronous window to test against at all. Only a genuine network `url()`
fetch (what the real resolver actually uses) does, so the regression tests
serve real font bytes from a deliberately delayed same-origin route and
compare a render that registers-and-immediately-triggers (mirroring the
resolver's fire-and-forget exactly) against one that awaits `.load()` first;
both a DOCX page-count test and a PPTX shrink-scale test fail without the
fix and pass with it (confirmed by temporarily reverting each fix and
re-running).

**The fidelity flag.** Every render handle now carries `.warnings: readonly
RenderWarning[]` (`{code, message}` — `core/model.ts`, shared by every
format's handle via `RenderHandleBase.setWarnings()`, `render-handle.ts`,
never one DOCX-only shape). DOCX layout mode merges the paginator's own edge
cases (`oversize-block`, `footnote-overflow` — computed since §21 but
previously dropped before reaching the embedder, a ROADMAP gap closed here)
with font entries; DOCX flow mode and PPTX carry the font entries alone.
`detectFontFidelity`/`fontFidelityWarnings` (`text-style.ts`, DOM-required —
deliberately not in `core/session.ts`, which Rule 3 keeps DOM-free) decide
per family whether it painted with something real (the requested name, or a
delivered alias) or fell to the browser's own unrelated default, by
comparing a `canvas.measureText` width against a same-generic-category
control — not by asking `document.fonts.check()`, which turned out unusable
for this: it returns `true` even for a family name that exists nowhere on
the calibration machine (verified, `tools/probe/font-metrics/
fonts-check-probe.ts` — on this Chromium build it reports whether *anything*
could paint the fallback stack, not whether the named family itself
resolved). The width comparison has one real gap of its own, also found by
testing against the real fixture rather than trusting the design on paper: Arial,
Times New Roman and Courier New measured "unavailable" on first pass even
though they are genuinely installed here, because Windows Chromium's own
`sans-serif`/`serif`/`monospace` generic keywords resolve to exactly those
three fonts — there is no way to distinguish "the real font, painting" from
"the generic default, which happens to be the same font" by width alone, and
it does not matter: the pixels are correct either way. Fixed by special-
casing those three names (`resolvesLikeGenericDefault`, `core/typography.ts`
— this is *why* they were picked as the generics' historical archetypes) to
never read as unavailable from that comparison alone. A working delivered
substitute is not itself a warning (`detectFontFidelity`'s `"substituted"`
status) — only `"unavailable"` (neither the requested name nor an alias
painted with anything real) produces a `RenderWarning`. Verified end-to-end
per format: a real font, a substituted font and a genuinely missing font each
render, and only the missing one's `.warnings` names it.

**A table row Word's own saved layout broke before is now real evidence
pagination uses, not dropped.** `w:lastRenderedPageBreak` leading a table
cell's first paragraph had nowhere to go under the mid-paragraph fix (§44):
tables break between rows, never inside a cell, so it was silently dropped
even at a row's very own start — real signal thrown away. It is real signal:
checked directly against `docProps/app.xml`'s own `<Pages>` count on two real
Word-saved files (the coverage fixture and the showcase term sheet), the
number of raw `w:lastRenderedPageBreak` occurrences in `document.xml`
(counting every one, including ones already dropped for other reasons)
matched `Pages − 1` exactly on both — Word's own account of how many times it
broke a page agrees with how many markers it left behind. `DocxTableRowNode.
pageBreakBefore` now carries a marker leading the row's first cell's first
paragraph (`parseTableCell` returns it alongside the cell; a marker on any
other cell, or after real text within the first cell's paragraph, still has
nowhere to go and stays dropped, unchanged). `splitTable` prefers a marked,
valid boundary over the latest-fitting one it would otherwise guess whenever
both agree it is usable (clean of a merge, no `w:keepNext` chain, fits the
room) — real evidence beats a heuristic. Fitting this in exposed a genuine
design gap, not a fixture accident: honoring the marked boundary as *which*
split to make is not the same as guaranteeing the resulting tail actually
*starts* a fresh page — a tail that fits some more content in the room left
on the current page would otherwise keep packing onto it, which is not what
`pageBreakBefore` means. A table encountered whole with a marked first row
still only breaks for it once the current page is already more than half
full, the same "prefer" gate a top-level hint uses — but a tail `splitTable`
itself just carved off at a marked boundary always starts fresh, ungated: that
split already used the strongest evidence available, so a further "is the
page half full enough" heuristic on top of it would second-guess a decision
real evidence had already settled. Caught by a test that mirrors what
actually happened building this: `splitTable` alone reported the right split
point, yet the rendered pages still put the marked row on the same page as
its predecessor — the second, narrower fix (tracked by object reference
through the one `pending` queue a table's fragments flow through, not by a
value that could coincidentally match) is what a passing test against the
real behaviour, not just the chosen split index, actually required.

**Not done in this pass, deliberately scoped out** (see ROADMAP M4): a
mid-paragraph `w:lastRenderedPageBreak` is still dropped outright rather than
promoted to a paragraph-level hint — splitting exactly where Word did, inside
a paragraph's own text, is a new pagination capability (paragraphs are atomic
today), not a small extension of what exists; the loss is real but bounded to
that one paragraph, and the next real marker (a following paragraph or table
row) resyncs pagination from there. Embedded fonts (DOCX `.odttf`, PPTX
`.fntdata`) are not read: a DOCX implementation was built and then removed
again, so the substitution table and `warnings` are the only font handling.

## 46. Arabic calibration measured the way it paints: through the alias, unrounded (Sep 2026)

Supersedes the calibration *numbers* in §45 (the method of choosing each
substitute stands). Found while benchmarking a proposal to replace Arial's
two-face fallback (Arimo + Lateef) with one font, Vazirmatn, which has both
Latin and Arabic glyphs.

**Vazirmatn, rejected on measurement.** `tools/probe/font-metrics/
vazirmatn-bench.ts` lays out every real paragraph in the three fixtures (142:
65 English, 45 Arabic, 32 mixed) in the DOM at 11pt across seven widths
(240–720px), in the real `arial.ttf`/`arialbd.ttf` and in each setup, and
counts the lines that start at the same word as Arial's. Vazirmatn was tested
the way aared would ship it: Latin and Arabic faces split by `unicode-range`,
each with its own fitted `size-adjust` and vertical overrides (so "one text box
cannot size both languages" does not apply here). Share of lines breaking like
Arial, regular / bold:

| Setup | English | Arabic | Mixed |
|---|---|---|---|
| Arimo + Lateef, Lateef refitted (below) | 100 / 100 | 96.2 / 98.4 | 98.2 / 99.7 |
| Vazirmatn, fitted per script | 97.1 / 96.4 | 95.6 / 98.6 | 96.8 / 97.6 |
| Vazirmatn as-is, same size | 97.1 / 79.4 | 86.4 / 86.4 | 93.3 / 92.1 |

It loses English (Arimo is a metric clone; Vazirmatn only averages out on a
sample line) and mixed paragraphs, and ties Arabic. After width-matching, its
Arabic alef is about 25% shorter than Arial's (refitted Lateef's, about 7%), so
its Arabic reads visibly smaller than Word's. It saves no download either:
Google Fonts serves it as per-script files, one Latin and one Arabic per
weight, the same count as Arimo + Lateef. As-is, its `line-height: normal` box
is 23px against Arial's 16px at 11pt (the overrides remove that).

**Bug 1: the Arabic fallback faces were calibrated alone, but paint beside a
Latin face.** An `ARABIC_SCRIPT_FALLBACKS` face covers `ARABIC_UNICODE_RANGE`
only, so the spaces, Western digits and punctuation inside Arabic text paint
in the alias's Latin face (Arimo, Tinos, Caladea, Carlito) at its own size.
§45 measured Lateef and Markazi Text on their own (`arabic-survey.ts --calib`),
counting *their* spaces, scaled by their `size-adjust`. As painted, every
fallback was too wide: Arial's Arabic 3.3%, Calibri's 2.5% (bold 3.3%), Times
New Roman's 2.3%, Calibri Light's 2.7%, Cambria's 0.9%. Fix:
`arabic-fallback-fit.ts` fits each through its real alias. The alias's width is
linear in the Arabic face's `size-adjust`, so two probes (100% and 80%) solve
for the value that matches the original's total width over every Arabic-only
run in the fixtures (59). Cambria is fitted against Times New Roman's Arabic
(what Word draws) through Caladea (what paints its spaces here), so it no
longer shares Times New Roman's numbers.

**Bug 2: vertical metrics were read rounded.** Canvas `fontBoundingBox*`
comes back in whole pixels, and the survey read it at 40px, so every
`ascentOverridePercent`/`descentOverridePercent` carried up to half a pixel of
error, 2.5% of the em at worst (Arabic Typesetting's descent 67.1% against a
correct 65.4%; Sakkal Majalla's 96.1%/57.7% against 97.1%/56.3%). Both probes
now read at 10000px; the values match each font file's `usWin` ascent/descent
(read with fontTools), which is what a browser on Windows paints the real font
with. All nine calibrations (the four main faces in `FONT_CALIBRATION`, the
five fallbacks) were corrected.

**Proof.** Arabic lines breaking like the original, before → after, regular /
bold (`arabic-fallback-fit.ts`, 77 paragraphs with Arabic, seven widths):

| Original | Arabic before | Arabic after | Mixed before | Mixed after |
|---|---|---|---|---|
| Arial | 90.5 / 94.0 | 96.2 / 97.6 | 99.3 / 99.3 | 98.2 / 99.7 |
| Times New Roman | 95.9 / 94.8 | 98.4 / 97.8 | 98.5 / 98.2 | 98.2 / 98.2 |
| Cambria | 98.4 / 97.5 | 98.4 / 96.5 | 97.8 / 91.4 | 96.7 / 92.5 |
| Calibri | 94.6 / 92.7 | 96.7 / 97.0 | 99.3 / 97.8 | 99.3 / 99.3 |
| Calibri Light | 95.9 | 96.4 | 97.1 | 97.8 |

Arabic paragraphs gain about 2 points on average at both weights. Mixed
paragraphs are flat overall: regular dips (Arial −1.1, Cambria −1.1) and bold
gains; single-line differences at this sample size are about one point, so
these are within noise. Every regular width now lands at 100.0% of the
original. The line-box check (`vertical-check.ts`, `line-height: normal` at
2000px so rounding costs 0.05%) puts every entry within 0.1% of the real font
except the two non-changes below.

**Blast radius.** Numbers only: `FONT_CALIBRATION` and
`ARABIC_SCRIPT_FALLBACKS` in `aared-fonts`; the resolver is unchanged. A
calibrated face paints only when the reader lacks the real font, so on the
calibration machine (which has every Office font) nothing renders differently
and the visual suite cannot show it; the probes above are the evidence. The
unit test that pinned Cambria to Times New Roman's numbers now pins the shared
substitute instead.

**Deliberate non-changes.**
- Bold keeps §45's rule: its own numbers only where it differs from regular
  by more than about 1.5%. Arial, Times New Roman and Cambria bold differ by
  0.7–0.9%, so their bold faces use the regular numbers and land at 99.3–99.4%
  width (Cambria bold's Arabic line match drops a point with it, within noise).
  Simplified Arabic bold's descent is 1% deeper than regular's; same rule.
- No `line-gap-override`. aared never paints text at `line-height: normal`
  (every renderer sets its own line height), so a face's line gap does not
  reach the page. It is why `vertical-check.ts` still shows Arabic
  Typesetting's `normal` box 0.7% short: the old, wrong descent happened to
  hide that gap.
- Whose spaces Word draws inside Cambria's Arabic (Times New Roman's or
  Cambria's) is unknown; `word-arabic-fallback`'s saved output cannot answer it
  (its sample order does not line up). A Word probe would settle it.
- The substitutes still have not been checked on a machine without the Office
  fonts (ROADMAP M4).

## 47. Paragraph spacing in lines: what a "line" is in Word and PowerPoint (Sep 2026)

Proposed: paint Word's `w:beforeLines`/`w:afterLines` with the CSS `lh` unit
(one line = the paragraph's own computed line height). Rejected on
measurement: neither app means that. Probe `tools/probe/line-spacing-units/`
builds paragraph pairs through each app's object model ("Before: 1 line" set
the way the UI sets it, beside control pairs with no spacing), exports
through the app itself and reads the gap off the PDF.

**Word: a line is always 12pt.** At 11pt and 24pt, single and double
spacing, Arial Arabic and Calibri, "1 line" added 12.0pt and "0.5 line"
6.0pt. `lh` would have given 12.7, 27.6, 25.3 and 16.1pt. Word also saves
the points beside the lines (`w:beforeLines="100" w:before="240"`), which
aared already read, so Word's own files were right before this entry. The
one exception is a line grid (`w:docGrid w:type="lines"`): Word then counts
the grid pitch (18pt at pitch 360, 24pt at 480) and snaps every line to the
grid. A stored pitch with no grid type changes nothing. aared has no line
grid, so that stays a gap.

Word reading hand-edited XML settled precedence: the line value and the
point value inherit separately, and a line value above zero wins, even over
the paragraph's own points (a style's `w:beforeLines="100"` beat a
paragraph's `w:before="120"`: 12pt). `beforeLines="0"` means "in points"
(`before="120"` then paints 6pt), and lines alone need no points
(`beforeLines="50"`: 6pt). Word writes a style's line value without points
(`<w:spacing w:beforeLines="100"/>`) and materialises `w:before` on each
paragraph, so only other producers' files differ. `resolveParaProps` now
carries both through the chain and `paragraphSpacePt` applies the rule.

**PowerPoint: a line is 1.2 × the paragraph's largest run size.** It is the
same in Arial, Calibri, Times New Roman, Segoe UI, Traditional Arabic and
Cambria Math, and at line spacing 0.8, 1.0 and 2.0. A 28pt word in an 18pt
paragraph makes the line count 28pt, before and after alike. `lh` matches
only at exactly 1.0 line spacing. aared dropped `a:spcPct` spacing
altogether (`spacingValue` returned it, and only points were used). It is now
`TextParagraphModel.spaceBeforeLines`/`spaceAfterLines` and paints as an `em`
margin on the paragraph's `<p>`, which already carries its largest run size.
So it also follows the shrink-to-fit pass that scales that size.

The same probe found three more rules, all fixed here because they belong to
the same spacing:
- PowerPoint drops space before a text body's first paragraph and after its
  last, in points or in lines, in shapes and table cells. aared painted both.
  `pptx-coverage.pptx` slide 5's bullets carry 10pt before every paragraph,
  and the list now starts where PowerPoint's does.
- One paragraph's space after adds to the next one's space before (1 + 1
  lines = 43.2pt at 18pt). A table cell with no vertical alignment was a
  block, where CSS collapses the two margins into the larger;
  `appendTextBody` now makes every text body a flex column.
- The Word renderer already stacks blocks apart, so Word is unaffected.

**Coverage.** Slide 12 ("تباعد الفقرات") was appended to
`pptx-coverage.pptx` through PowerPoint's object model. It has an RTL Arabic
text box and a table cell spaced in lines: the first/last edge, half a line,
the 28pt word, and after + before. A parser test reads the lines back, and a
DOM test (`pptx-paragraph-spacing.visual.test.ts`) re-reads the painted gaps
between paragraph boxes; it fails on the previous renderer. A synthetic
parser test pins Word's precedence cases, since Word's own saves always carry
the points too.

**Seen in the paired captures, not fixed here:**
- PowerPoint also lowers a paragraph's first line when its line spacing is
  above 100%; aared centres the extra leading.
- A text box that authors no insets gets none in aared, where PowerPoint
  uses 0.1in/0.05in.
- On slide 5, bullets larger than their text make aared's lines taller than
  PowerPoint's.

All three are in ROADMAP's fidelity-gap table.

## 48. PPTX autofit is what PowerPoint stored; bullets hang; a subtitle takes the master body style (Sep 2026)

The slide-5 screenshot (PowerPoint beside aared: a bullet list spread out
down the slide) and the paired captures of slide 1 turned up four causes with
one theme: aared worked out something PowerPoint reads from the file.

**Autofit follows the stored values, not a measurement.** Overturns the
shrink half of §45. Probe `line-spacing-units` built nine `normAutofit` boxes
and read PowerPoint's own PDF: PowerPoint does not measure on open. A bare
`<a:normAutofit/>` paints at 100%, overflow painted (not shrunk, not
clipped). Only the stored `fontScale` and `lnSpcReduction` change anything.
A painted size is `round(size × fontScale)` in whole points, half up; the line
step is 1.2 × that size × (line spacing − `lnSpcReduction`); spacing in points
is unscaled; spacing in lines counts 1.2 × the rounded size and ignores the
reduction. aared's old pass measured `scrollHeight`, took ×0.95 steps and
clamped at 0.6. That had no basis in Office, and it shrank slide 1's subtitle
to 28.95px where PowerPoint paints 32px (a bare `normAutofit`, tipped over
its threshold once the empty last paragraph was counted as a line).

`parseTextBody` now applies both stored values and the model carries them
(`TextBodyModel.fontScale`, `.lineSpacingReduction`; sizes and `lineSpacing`
are already scaled, so a renderer only paints). Only the slide's own `bodyPr`
counts: a layout or master `normAutofit` never carries a scale. Removed:
`shrinkTextToFit`, `scaleFonts`, `hasAutofitShrinkText`, the PPTX-only
`awaitDocumentFonts` wait (DOCX keeps it; nothing on a PPTX slide measures
text any more) and `data-aared-autofit-applied`. Tests: a parser test pins
the rounding, the reduction and the layout/master rule; the fidelity test
re-reads the painted size of the shrink box (24px).

Deliberate non-changes: `lineSpacingPt` (exact line spacing) and
`bullet.sizePts` stay unscaled, since PowerPoint's effect on both is
unprobed. **Limit:** `none` and `shrinkText` shapes still paint with
`overflow: hidden`, where PowerPoint paints the overflow.

**Bullets keep PowerPoint's hanging indent.** The marker was a grid column
with a gap, which put the text at the wrong place whenever the marker was
wider or narrower than the hanging gap, and a marker larger than its text
(150% bullets, a 42pt Wingdings square) made the line taller. Read off
PowerPoint's PDF: the marker starts the first line at the margin plus the
(negative) first-line indent, and the text starts at the margin or right where
a wider marker ends. The marker is now an inline box at least as wide as the
hanging gap, with no line height of its own, so both hold in either
direction. Along the way: `marL` counts from the right in a right-to-left
paragraph (`margin-inline-start`, not `margin-left`); an empty paragraph is
still one line at its own size; a `+mj-lt` bullet font is the theme's heading
font (`cssFontFamily`). The first-line indent is painted for every paragraph,
so `applyParagraphBox` lost its `firstLineIndent` option.
`pptx-bullet-geometry.visual.test.ts` re-reads the marker and text boxes of
slides 1 and 5 in both directions and fails on the grid.

**A subtitle placeholder takes the master's body style.** `masterTextStyle`
switched on the raw placeholder type, so `subTitle` (and content, table,
picture and chart placeholders) fell through to `otherStyle`, while the
master-placeholder lookup beside it already sent them to the master body. The
master's `bodyStyle` carries the 90% line spacing and 10pt before, so slide
1's subtitle painted lines 39px apart where PowerPoint's are 47, 51 and 48
(aared now 47, 51 and 47). Both lookups now share `masterPlaceholderType`;
only date, footer and slide number take `otherStyle`. Blast radius: every
placeholder that is not a title or one of those three, in every PPTX; the
layout's `lstStyle` and the slide's own properties still sit above the master
style, so nothing that authored its own values moved (no existing test
changed). `pptx-parser.test.ts` pins the subtitle's 0.9 and 10pt, and the
slide number's lack of both.

**Seen, not fixed.** On slides 1 and 5 aared's first line still sits about 3px
lower than PowerPoint's in paragraphs at 90% line spacing (the title and the
bullets alike); where PowerPoint puts the leading is unprobed (same family as
the 100%-and-over gap in §47). Both are in ROADMAP's fidelity-gap table.

## 49. Aptos: calibrated look-alikes measured against the real cloud fonts (Sep 2026)

Aptos has been Office's default theme font since 2023: Aptos for body text,
Aptos Display for headings, and Aptos Narrow as Excel's default. It had no
substitute, because §45 thought there was nothing to measure against. No
browser has Aptos, so its text painted in the generic sans-serif; on Windows
that is Arial, 5-17% wider than Aptos. Slide 1's "Next review: Thursday
12:00" measures 384.8px in Aptos and 407.3px in Arial; the paired captures
show 383px in PowerPoint and 406px in aared.

**The real fonts were on the machine all along.** Office downloads cloud
fonts into `%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts` and never
installs them as system fonts. The files are named by hash, so
`tools/probe/font-metrics/cloud-fonts.ts` finds them through their name
table. Versions 2.00 and 2.01 are both there, with identical advances. They
are used only for measuring and are never shipped or copied.

**Choice.** The method is §45's, applied to Latin text.
`latin-survey.ts` measured every Latin run in the five fixtures (194
samples, plus 77 digit strings reported on their own) in the real Aptos,
Aptos Display and Aptos Narrow, against 61 open candidates. Each candidate
was loaded with the resolver's own request (`fontCssUrl`), and every load was
verified.

Width spread alone did not decide it. The narrowest spreads came from
condensed faces whose letters end up 12-33% taller once size-adjust matches
the widths. So the rule became: keep Aptos's x-height and cap height. The bar
was set at 3% first. For Aptos Narrow no face met it; IBM Plex Sans Condensed
is closest, with its cap height 3-4% off. So the bar is 4%, at both weights,
and exactly one candidate per family passes it:

- Aptos → Radio Canada
- Aptos Display → Roboto
- Aptos Narrow → IBM Plex Sans Condensed

**Line breaks.** `latin-fit.ts` laid out real paragraphs at eight widths, in
both weights, for the 16 closest candidates. It counted paragraphs broken
exactly where Aptos breaks them:

| Family | Substitute (regular / bold) | Generic sans-serif (regular / bold) |
|---|---|---|
| Aptos | 96% / 93% | 75% / 63% |
| Aptos Display | 95% / 94% | 42% / 38% |
| Aptos Narrow | 96% / 95% | 56% / 51% |

The picks have the best average of the 16 for Aptos and Aptos Narrow, and
are 0.8 points under the best for Aptos Display. Visual weight ("ink") is
within 2-9%.

**Aptos Display** is Aptos spaced tighter for large sizes. A uniform
size-adjust cannot reproduce tighter spacing, so the Aptos pick lost ground
on it, and Display takes its own face.

**Numbers.** Each family gets its own `FONT_CALIBRATION` entry.

Bold carries its own numbers in all three families, although the weights
differ by only 1-3%. That departs from the Arabic entries' 1.5% bar, on
measured cost: at the regular numbers, Aptos Narrow's bold paints 1.3% wide,
and its identical paragraphs fall from 95% to 91%.

The vertical overrides are Aptos's typo ascent and descent (0.939/0.282 of
the em). The browser measured those, because Aptos sets USE_TYPO_METRICS.
The win metrics (1.010/0.275), which the other calibrations' "usWin" rule
would have picked, are unused. The line box now matches Aptos's within 0.1%;
Arial's was 5.8% short.

**Arabic under Aptos.** `aptos-arabic-fallback.ps1` sets Arabic in each
Aptos family, exports PDF, and reads the span fonts:

- Excel (one font per cell, no complex-script slot) draws it in Arial, at
  Arial's own width (1.001).
- Word draws it in Arial too, whether Aptos sits in the Latin slot of a
  left-to-right run or in the complex-script slot of an RTL one.

This is the common spreadsheet case, since Aptos Narrow is Excel's default.
So the three families take Arial's Lateef face in `ARABIC_SCRIPT_FALLBACKS`,
fitted through their own alias. `arabic-fallback-fit.ts` now registers the
Latin face with its own calibration, as the resolver does. The spaces come
from the size-adjusted Latin face, which is narrower than Arimo, so Lateef is
3-5% larger than under Arial. 96-100% of Arabic line starts land where
Arial's do.

PPTX and DOCX already set any run containing Arabic in its complex-script
font, so this face mostly matters for XLSX. Neither XLSX fixture has an
Aptos Narrow cell holding text, so the probe is the only evidence for that
route. No fixture was added.

**Blast radius.** `aared-fonts`'s table and tests change. For anyone passing
a resolver, Aptos documents now download fonts.

- The resolver eagerly loads every non-Arabic subset face of a substitute,
  in all four styles (§45 left that alone). The new substitutes cost about
  what shipped ones already do: Radio Canada 234 KB, Roboto 471 KB and IBM
  Plex Sans Condensed 175 KB, against Arimo's 445 KB and Carlito's 784 KB.
  A document in the default theme (Aptos plus Aptos Display) now fetches
  about 700 KB where it fetched nothing. Loading only the painted subsets is
  a separate, open decision.
- The site's Light Table demo passes `defaultResolver` for PPTX only. Its
  sample deck (`post-production-review.pptx`) has no Aptos, so the default
  demo is unchanged; a dropped Aptos deck now paints the substitute.
- The XLSX route reaches the resolver. A cell with no style takes style 0,
  whose font is the workbook default; `collectStrings` visits every distinct
  cell style, and the renderer paints `cssFontFamily(style.fontFamily)`.
  This was confirmed by reading the code and in the end-to-end check below;
  no test exercises it.
- End to end, built packages with the real `defaultResolver`:
  - Slide 1's "Next review: Thursday 12:00" paints 385.7px under the Aptos
    alias. Real Aptos measures 384.8px, and PowerPoint's capture spans the
    same x range.
  - An Aptos Narrow cell holding Arabic loads Lateef at 132.6%.
- No renderer code changed.
- XLSX column widths are a fixed 7px per character (`core/xlsx/sheet.ts`),
  never measured, so the substitute's digits move no column.
- No visual test uses the real resolver (they register faces by hand and stay
  offline), and neither do the bench or the paired captures. On this machine,
  those still paint Aptos in the generic sans-serif.
- The resolver unit test pins registration: Radio Canada under the Aptos
  alias, per-weight numbers, and Lateef limited to Arabic and loaded lazily.
- The probes are the acceptance, as for every calibration.

**Limits.**

- Italics use the upright numbers of their weight (`FontCalibration` is keyed
  by weight). They paint 2-3% narrower than Aptos's italic in Aptos and Aptos
  Display, and 1% in Aptos Narrow.
- Aptos Light, SemiBold, ExtraBold, Black, Mono and Serif are separate Office
  families and are not covered.
- Letterforms differ; widths and letter heights are what match.
- The substitutes are Google Fonts' current files, which can change under any
  calibration. Refitting Arial's and Cambria's Arabic faces while wiring this
  up came out 0.3-0.9% off the numbers §46 shipped (bold furthest). That was
  noticed, not changed.

## 50. Delivered fonts load when painted, not up front (Sep 2026)

The resolver used to call `.load()` on every non-Arabic subset face of a
substitute as soon as it registered it, in all four styles. Every Google
subset (Latin, Latin Extended, Greek, Cyrillic, Vietnamese…) came down,
whether or not the document used it. That was left alone in §45, and §49
made it cost more, since the default theme now has substitutes too.

**Change.** The resolver now registers faces and never loads them. The
browser fetches a face when text first needs it, going by its style, its
weight and its `unicode-range`.

Two things relied on the eager load, and both now load exactly what the
document paints:

- **The wait.** `collectFontUse` records, per family, the variants its runs
  paint in (their own `bold`/`italic`; an XLSX cell style has the same keys).
  `awaitDocumentFonts` loads those variants with a text sample, and that is
  what triggers the download of those faces and subsets only.
  - Before, it waited for regular and bold and never for italic. So DOCX
    layout could measure an italic run before its face arrived; the eager
    load only made that unlikely.
- **The fidelity check.** `detectFontFidelity` now measures in a variant the
  document paints. A bold-only family's regular face may never be fetched.
- **PPTX `render()`** now waits too, like both DOCX modes: it paints, waits
  (bounded, 2 s) for the slide's fonts, then reads the warnings and resolves.
  - Before, the eager loads usually finished in time. Lazily, nothing would
    have loaded, and every substituted family would have read as
    "font-unavailable".
  - The synchronous `paintSlide` still reads at once, and says so (API
    page).
  - XLSX neither measures nor warns, so it needs no wait.

**Measured.** `tools/probe/font-metrics/font-bytes.ts` records every face the
resolver registers with its file, and counts the ones the browser left
"unloaded" as not downloaded:

| Render | Before (loaded up front) | Now (loaded when painted) |
|---|---|---|
| `pptx-coverage.pptx` slide 1 (Aptos, Arial) | 58 files, 1620 KB | 5 files, 185 KB |
| `post-production-review.pptx` slide 1 | 26 files, 717 KB | 4 files, 167 KB |
| `docx-coverage.docx`, whole document | 87 files, 2307 KB | 11 files, 337 KB |
| `xlsx-coverage.xlsx` sheet 1, no Office fonts | 18 files, 572 KB | 4 files, 168 KB |

Most of what remains is the Lateef Arabic face, which was already loaded on
demand.

**Cost left.** A reader who has the real font still downloads what the wait
and the check load under the alias. On this machine the PPTX and DOCX
numbers equal the no-Office-fonts ones. Skipping those loads would need to
know the real font is installed, and that can't be measured reliably: Arial
measures the same as the generic it stands behind (§45). Kept, and bounded
by the table above.

**Tests.**

- The resolver tests now assert that no face is loaded up front, Latin or
  Arabic.
- `collectFontUse` has a test for variants: run and cell-style keys,
  including `bold: false`.
- The visual suite passes unchanged, and it now includes the no-Office-fonts
  pass (§51), which paints through the lazy faces in all three formats.

## 51. The no-Office-fonts review: substitutes painted where they matter (Sep 2026)

Every calibration was measured on a Windows machine with Office's fonts. The
real fonts are installed there and always win, so no substitute had ever
painted a real document. A phone, or a Mac or Linux machine without Office,
is where they paint.

**Simulated on this machine.**

- An installed font cannot be hidden from a page. A registered face named
  "Arial" that fails to load falls back to the installed Arial; the text
  still measured as Arial.
- So the review renames every family in the model to one no machine has
  (`"Arial absent"`). `packages/aared-fonts/test/absent-fonts.browser.ts`
  then registers the real resolver's faces under the renamed family's alias,
  through its `FontFaceCtor` hook: the same CDN files, calibrations and
  aliases a phone would get. A family with no substitute (Wingdings, Segoe
  UI) falls to the generic, as it would there.
- `bun run test:visual:artifacts --absent-fonts` adds that render as a third
  column of the PPTX and XLSX reviews. It uses the network.
  `tools/word-render/capture.ts --absent-fonts` writes the DOCX pages.

**What holds.** All three fixture families lay out as with the real fonts,
against the native Office renders:

- DOCX: the same 4 pages, the same English line breaks and the same table
  geometry.
- PPTX: all 12 slides in place.
- XLSX: every sheet, with no row or column moved.

**Fixed here.** Chart text in both PPTX and XLSX was painted with the bare
family name. It skipped `cssFontFamily`'s alias and generic, so a missing
font (Aptos, every default-theme chart) fell to the SVG default serif, even
on this machine. The chart titles and axes in the paired XLSX captures had
always been Times. `chart-svg.ts` now paints the same stack as every run;
`chart-svg.test.ts` re-reads the attribute.

**Found.** Symbol-font bullets became letters: Wingdings' square and check
mark painted as "q" and "ü" (`pptx-coverage.pptx` slide 5), because those
glyphs live at letter codes in Wingdings, and Android, iOS and Linux have no
Wingdings. Fixed in §52. The rest are accepted as good enough (ROADMAP
fidelity gaps; AGENTS.md "Good enough beats perfect"):

- **Lateef's Arabic draws taller than Arial's.** It is set 127-133% to match
  widths, so its strokes run past Arial's line box. In dense rows
  (`docx-coverage.docx`'s delivery tracker) they cross the borders.
  - Under a link (`pptx-coverage.pptx` slide 6), the underline sits at the
    Latin face's position and cuts through letters: the dot of "ج" merges
    into it.
- **Arabic breaks can shift by a word.** A justified Arabic paragraph did so
  (`docx-coverage.docx` page 2). This is within the calibration's measured
  spread (96-98% of line starts).
- **Text on the wrapping edge can flip.** A mixed-script table cell sitting
  right at its wrap point fits on one line where PowerPoint wraps it (slide
  2).

**Limits.**

- Characters no substitute covers fall to this machine's generic (Arial,
  Times New Roman). A phone's generic is Roboto or Noto, so unsubstituted
  families look different there.
- One look on a real device is still worth taking: ROADMAP M4.

## 52. Wingdings and Symbol bullets paint their Unicode shape (Sep 2026)

Found by §51. Wingdings and Symbol keep their glyphs at letter codes:
Wingdings' square is "q" and its check mark "ü", and Word's default
first-level bullet is Symbol's private-use U+F0B7. Only machines with Office
have those fonts. Anywhere else the bullet paints as the letter, or as an
empty box.

**Change.** `core/symbol-fonts.ts` maps the characters Office offers as
bullets to their Unicode shapes: 20 Wingdings, 5 Symbol.

- Both the plain and the private-use (U+F000 + code) forms are mapped.
- Each entry was matched by eye against the real glyph. The first guesses
  for four were wrong: Wingdings' check mark is the light ✓, not the heavy
  one.
- Shapes a phone may draw as colour emoji (☑, ▪, ◻, ♣, ♦) carry the
  text-style selector U+FE0E.

The PPTX parser applies it to `buChar` and the DOCX parser to list markers,
dropping the symbol font, so the model itself carries the shape and search
and copied text see it too. On a machine with Office, the marker now paints
from a standard font instead of Wingdings; the shapes are near-identical.

**Not covered (accepted):** runs set in a symbol font, `a:sym`/`w:sym`
symbols, and Wingdings 2/3 and Webdings, which still paint letters without
those fonts.

**Tests.**

- `symbol-fonts.test.ts` covers the table rules.
- The PPTX text-style test re-reads slide 5's markers from the real fixture.
- The bullet visual test re-reads the painted marker.
- A synthetic DOCX test covers Word's default Symbol bullet; no fixture has
  one.

## 53. `aared-fonts` frozen and unpublished: the browser's own fonts are the default (Sep 2026)

Overturns ROADMAP M4 ("`aared-fonts` publish + CDN default"). The real-device
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
rule is unknown, accept the rest in writing. Probes: `first-line-leading`,
`en-us-arabic-run`, `master-graphics` (all in the gitignored `tools/probe/`).

**Line spacing off 100% moves the text inside its box (overturns §47/§48's
"seen, not fixed").** PowerPoint's own PDF, Arial, Calibri and Times New Roman
at 18 and 40pt, 70% to 300%: the line step is always `spacing x 1.2 x size`
(aared already matched), but above 100% the first baseline sits
`0.9 x spacing x size` from the top of the text area, and below it the natural
baseline plus the removed leading, never above `0.9 x spacing x size`. CSS
centres the leading, so aared sat too low at 90% (the 3px on slides 1 and 5)
and too high above 100% (9pt at 200% and 40pt). `leadingShiftEm` paints the
difference as a relative `top` on the `<p>`: box heights, steps and flow are
untouched. The shift is measured against the height CSS really paints, which
`MINIMUM_LINE_HEIGHT` (1.1) floors, so 70%-90% lines still step 1.1em apart
where PowerPoint steps tighter: kept, because Arabic ink clips below it.
Worst gap to PowerPoint across the probe: 1.15pt (was 22.6pt).
`pptx-line-leading.visual.test.ts` pins PowerPoint's Arial 40pt numbers.
Limits: the natural baseline is taken as 0.95em (0.93-0.97 across the three
fonts); exact spacing (`spcPts`) is unprobed and unchanged. Blast radius: every
PPTX paragraph with a line spacing other than 100% - notably every title and
body on a default master (90%), table cells included; nothing with no authored
spacing moved. XLSX and DOCX do not run the PPTX paragraph painter, so they are
unchanged (DOCX's lines are Word's own model, §44).

**A text body's anchor and wrap cascade one attribute at a time (found in the
paired review, not in the table).** `parseTextBody` read `anchor` and `wrap`
only from the first non-empty `bodyPr` in the slide/layout/master chain. A
slide's own `<a:bodyPr><a:normAutofit/></a:bodyPr>` is non-empty, so it hid
the master's `anchor="ctr"`: every such title painted at the top of its box
where PowerPoint centres it (43px high on `pptx-coverage.pptx` slide 2, and
on slides 4, 7-12). Insets already cascaded per attribute (§47); anchor and
wrap now use the same list. Blast radius: every placeholder whose own `bodyPr`
carries a child but no `anchor`/`wrap`, in every deck, which is most
PowerPoint-authored titles; text boxes with no placeholder chain are
unchanged. Autofit and `prstTxWarp` still read the first non-empty `bodyPr`
(child elements, not attributes). Across the 12 coverage slides the pixels
differing from PowerPoint's export fell on every slide (about 30% in total)
with this and the changes below.

**Text frames with no authored insets get DrawingML's defaults.** The parser
only recorded insets some `bodyPr` in the chain stated, so an `AddTextbox` box
(which writes none) had no padding; PowerPoint uses 0.1in and 0.05in. The
defaults now resolve in `parseTextBody`, one constant pair shared with table
cell margins. Blast radius: every shape whose text body states no inset on any
side, in every deck; placeholders were unchanged (masters author theirs).
Existing tests did not move.

**Overflow paints in every autofit mode.** `none` and `shrinkText` no longer
clip (`overflow: visible`); the slide viewport still clips at the slide edge.
Text past its box now paints where PowerPoint paints it. Unprobed and
unchanged: exact spacing and bullet sizes under a stored `fontScale`.

**Colour-map overrides.** `p:clrMapOvr/a:overrideClrMapping` on a slide or
layout replaces the map above it; `a:masterClrMapping` keeps it. Slide beats
layout beats master. A layout's shared graphics use the layout's map; a
slide-level override does not reach them (the model is shared per layout).
Implemented from the spec; no fixture carries one.

**Table styles.** `a:lnRef` edge widths now come from the theme's
`a:lnStyleLst` (`DrawingTheme.lineWidths`). The old constant assumed 0.5, 1 and
1.5pt, which is the 2013-era Office theme (`post-production-review.pptx`); the
Aptos-era one (`pptx-coverage.pptx`) states 1, 1.5 and 2pt, so `lnRef` edges
were too thin there. Both fixtures now pin their own widths; the constant
stays only as the fallback for a theme that states none. A built-in Medium
Style 2 (seven ids, the default one included) missing from `tableStyles.xml`,
as in decks from generators that write only the id (python-pptx's default
table style is the Accent 1 id), is built from PowerPoint's own definition on
first use. Each id was applied to a table in real PowerPoint
(`tools/probe/table-style-guids`) and its stored definition parses identically
to the built one; any other missing style still paints unstyled.

**Arabic around ASCII punctuation in an `en-US` run.** PowerPoint's PDF: in a
left-to-right paragraph, a run tagged with a left-to-right language orders
Arabic words in written order around every ASCII punctuation character
(31 tried) - `W1 + W2`, `W1 (W2) W3` - while spaces alone (`W3 W2 W1`), an
Arabic `lang` and every right-to-left paragraph follow plain bidi. The parser
cuts such a run at the punctuation between two Arabic words and flags the
punctuation pieces `ltrRun` (existing mechanism: invisible marks in generated
content, never copied or searched). Untagged runs, Arabic `lang` and RTL
paragraphs are untouched; Arabic-specific punctuation is not treated as ASCII.
Neighbours are checked inside the run only.

**Punctuation stretches in right-to-left paragraphs.** The same probe method
on a right-to-left paragraph (`rtl-trailing-punct`, PowerPoint's own PDF): an
`ar-SA` run is plain bidi; an `en-US` run lays each stretch of two or more
punctuation marks out in written order, at the start, in the middle or at the
end (`.»` closing an Arabic quotation shows `.` then `»`; `(-)` shows in
written order), while the glyphs stay mirrored as right-to-left text (the left
guillemet on `pptx` slide 7 is the right one flipped, 82% against 43%
unflipped). That rules out a plain left-to-right override, which would draw
unmirrored glyphs. It was built (a parser split, a `TextRunModel` field and a
left-to-right isolate with each bracket in a right-to-left one) and matched
PowerPoint on `post-production-review.pptx` slide 7, then **removed**: the
only visible effect is which of `.` and `»` sits outermost, which a reader
does not notice, against a new public model field, nested bidi wrappers in
the DOM and a painter branch. Under "good enough" that is not worth carrying;
the measurement stays here for anyone who needs it (ASCII punctuation and
guillemets only; spaces, curly quotes, dashes and Arabic punctuation were
never probed). The left-to-right-paragraph rule above stays: it swaps whole
words, which a reader does see.

**Master and layout graphics verified, no change.** PowerPoint: a layout's
`DisplayMasterShapes = false` hides the master's shapes; a slide's hides layout
and master shapes. aared resolved all four probe slides identically, and a
master picture background resolves as a picture on every slide.

**Route completeness.** `DrawingTheme` is shared: XLSX and DOCX read the same
theme, so both now carry `lineWidths` too, which neither uses (a few numbers
per document); their colours and fonts are unchanged. `parseTextBody`, the
PPTX text painter, colour-map overrides and table styles are PPTX-only: XLSX
charts go through `core/chart.ts`, untouched, and DOCX resolves its own tables,
insets and bidi (`w:tblCellMar`, `w:rtl`, §43-§44), so neither gets the
defaults or the punctuation rule. Whether Word orders Arabic around ASCII
punctuation the way PowerPoint does is unprobed.

**Accepted, not built (recorded in the fidelity table):** fill approximations
(group pattern anchor, 96 DPI tiles, gradient stops); off-centre `circle` and
`shape` gradients stay marked fallbacks (COM cannot author them); live fields
other than slide number and date; `"Slide N"` titles (by design, §31); embedded
fonts (§45); notes (parsed onto `SlideModel.notes`, reachable through
`slide.read()`, not painted, not searched); `defaultTextStyle` versus
`otherStyle` (both fixtures agree, PowerPoint writes them consistently).
**Deferred by the maintainer:** charts, to be done in isolation; the quick-fix
route is written out in ROADMAP. **Found, not scheduled:** WordArt preset
warps (`prstTxWarp`) paint unwarped and marked. No earlier entry ruled them
out (the "explicitly out of scope" in `TextBodyModel.textWarp`'s comment had
no entry behind it and is reworded); the arch and circle presets could bend
real text along an SVG `textPath`, which keeps rule 1, and the other presets
are free-form envelopes with no DOM form.

## 55. DOCX pictures that text wraps around, placed as Word places them (Sep 2026)

An anchored picture (`wp:anchor`) used to paint as a plain left or right CSS
float whatever its wrap, so a centred picture sat at a column edge, a
top-and-bottom picture had text beside it, and a behind-text watermark pushed
the text away. Probe `docx-float-wrap` (gitignored): `author.py` writes 13
wrap cases through Word's COM object model (`wraps.docx`), `author-page.py` a
page watermark and a header logo (`page.docx`); each was exported by
`tools/word-render` and compared line by line with aared's capture.

**The model.** `DocxImageNode.anchor` (`DocxImageAnchor`) carries what Word
stores: wrap kind, `behindDoc`, `wrapText`, horizontal and vertical position
(`relativeFrom` with `align` or offset, in inches), the four wrap distances and
the wrap polygon as shares of the picture (Word writes it in a 21600 square).
The picture leaves the runs: it hangs off its paragraph or heading as
`floats`. A float-only paragraph stays a paragraph, so the picture has a line
to anchor to. The old `float` field is gone (pre-release; no alias).

**The placement** is pure numbers (`docx/float-layout.ts`, `planFloat`): the
same inputs give the same plan in a test without a browser. Four kinds:

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
and AGENTS.md now says so ("Good enough" covers only what a reader would not
notice). `docx/split-wrap.ts` lays the rows beside the picture out itself: each
row is a grid of (left slot | hole | right slot), filled in reading order
(right slot first in a right-to-left paragraph), with the rest of the
paragraph flowing normally below. Words are measured with canvas
`measureText` (measuring only; the text is DOM spans, selectable and
searchable) through a bounded word cache, then every painted slot is checked
against its width and the paragraph re-laid out with 1, 2 or 4px of slack if
the browser's shaping disagreed. Three approaches were timed on an Arabic
mixed-font paragraph: DOM `Range` measurement 20.5ms, canvas 2.4ms, identical
breaks. Reads, writes and checks run as three passes over every split
paragraph so layout is forced once per pass; a `ResizeObserver` re-runs flow
mode on width changes and fonts arriving re-run it once. Tight and through
take each row's hole from the polygon's extent over that row
(`polygonSpan`). The paragraph's minimum height reaches the picture's bottom
less its spacing after: Word starts the next paragraph at its ordinary place
(81px on the fixture's page 6, `docx-render.visual.test.ts`) and wraps it
around what remains, which a split cannot reach into, so the picture may hang
into the gap as in Word but never under the next line.

**Word applies the default paragraph style.** Found while comparing the
captures: aared's text was a pixel or two larger than Word's. A paragraph with
no `w:pStyle` takes the paragraph style marked `w:default="1"` (usually
`Normal`), not the bare document defaults; aared used the defaults, so the
probe painted 16px Aptos where Word painted 11pt Arial. `StyleData` records
`defaultParagraphStyleId` and `resolveParaProps` falls back to it. Blast
radius: every DOCX paragraph without a style; on the coverage fixture the
model gained `styleId: "Normal"` on 126 nodes and the pixels did not move,
since its Normal restates the defaults.

**Fixture.** Seven wrap cases were added to `docx-coverage.docx` through COM
(`tools/probe/docx-float-wrap/extend.py`, label stripped), first one per page,
then packed onto two landscape pages at the maintainer's request: the fixture
is six pages in Word and in aared.

**Review pass.** A review of this session's code for memory, speed and dead
code found ten points, all fixed: the word cache is bounded; the flow
`ResizeObserver` and font listener are disposed with the render
(`ImageWiring.cleanups`); constant picture styles are classes; an anchor
inside a table cell is placed against the cell's text width, not the page's;
a paragraph-relative vertical offset under an `align` no longer adds both; and
unused exports and a stale header comment went.

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
only `ZipError` to `invalid-document`. Malformed XML (`XmlError`, with no
`aared/` prefix either), a required part or element missing (plain
`throw new Error("aared/core: ...")` across `opc.ts`, `pptx.ts`, `theme.ts`,
`xlsx/`, `docx/`) and a corrupt deflate stream (the platform's `TypeError`)
crossed as a bare `{ name, message }` and rejected `open()`/`read()` with no
`code`, so `isAaredError()` and `error.code` saw nothing. Every one of them
means the same thing to a caller, this file cannot be previewed, so the
boundary every parser failure crosses now types them all `invalid-document`,
the `ZipError` precedent extended, instead of retyping ~35 throw sites. The
original text stays (prefixed `aared/core:` when it had none), followed by
the one action a reader can take: re-save the file in Office. An aared bug
on an odd file also reads as `invalid-document`; from the caller's side that
is still a file it cannot preview, and the original message is kept.

**A lost session reported `aborted`.** `sessionClosed()` answered a request
for a session the Worker no longer holds (closed, or its open cancelled) with
`code: "aborted"`, which `isAbort()` swallows: the element and the React
hooks then dropped a real usage error silently. It is `closed` now, as the
main thread already reports after `close()`.

Tests: a damaged main part inside a valid archive (malformed XML, and a
missing `p:sldSz`) rejects typed (`pptx-document.test.ts`); the Worker's
closed-session reply carries `closed` (`worker-cancellation.test.ts`).

**Worker errors read in the caller's entrypoint.** The Worker only knows it
is `aared/core`, so every failure it sent said `aared/core:` even when the
caller opened through `aared/pptx`, while main-thread errors already used the
caller's scope. The bridge now rewrites the prefix to its own scope as it
receives a failure. Cancellation keeps the one `aared/core: request was
cancelled` wording on both sides of the Worker: it is created on both, and a
reader never sees it. Headless `core` calls (no Worker) still say
`aared/core:`, which is what they are.

**A failing font resolver no longer fails the document; `font-error` is
gone.** A `fontResolver` that threw (or rejected) rejected the whole `read()`
and `render()` with its raw error, and the rejected promise was cached for the
session, so the family could never resolve again. A font is cosmetic: the
failure is now caught per family and not cached, the family paints in the
browser's fallback, and the render handle's `warnings` already names it
(`font-unavailable`, measured from what painted, §45). `font-error`, defined
and documented but never produced, is removed from `AaredErrorCode`
(pre-release; no alias). A failing `mediaResolver` was already handled in all
three renderers (an "(unavailable)" XLSX drawing, DOCX alt text, the marked
PPTX fallback; `MediaUrls` catches a synchronous throw too) and is unchanged.

Tests: the damaged-part test also pins the `aared/pptx:` prefix; a resolver
that throws and then rejects leaves both reads succeeding and is asked again
(`session-lifecycle.test.ts`).

**A damaged sheet says what it skipped.** The XLSX sheet parser skipped
anything with an unreadable address (a cell, hyperlink, merge,
conditional-format range) or an unreadable drawing anchor with a bare
`catch { continue; }`: the item vanished and nothing said the view was
incomplete. Excel never writes these; damaged files and buggy generators do.
They now go to `SheetModel.unsupported`, the list the parser already uses for
what a preview does not show, under a new kind, `damaged`, counted per kind
so a broken file adds one line per kind ("2 cells could not be read and were
skipped; the file may be damaged") rather than one per cell. The new kind
keeps "the file is broken" apart from "aared does not model this feature",
which a host words differently; the two existing "points at a missing part"
drawing entries move to it. Not render `warnings`: those are measured from
what painted, and this is known at parse time. An unreadable `dimension`
stays silent: the extent is then taken from the cells, so nothing is lost.
Test: a hand-built workbook with each kind of damage keeps its readable cell
and merge and reports four counted entries (`xlsx-document.test.ts`).

The audit's matching PPTX claim, that a chart whose part cannot be read
paints blank and unmarked, was wrong: a chart with no drawing already paints
the marked placeholder (`data-aared-chart-fallback="unsupported"`, a dashed
outline with its name, `pptx/index.ts` `appendChart`). Unchanged.

**Every embed layer reports a failed navigation.** The element fires `error`
and rejects `goTo()`; `preview()` rejects; the hook put only a failed mount
in state. Those are each layer's own idiom and stay. The gaps were these:

- **The hook lost page-change failures.** A host writing
  `onClick={() => goTo(index + 1)}` and reading `status`/`error` saw nothing
  when the next unit failed: the promise rejected unawaited, `status` stayed
  `"ready"`. The result gains `navigationError`, cleared by a unit change and
  every mount. Chosen over moving `status` to `"error"`: that would replace
  the whole view for one page that would not show, while this keeps the page
  on screen.
- **Reported where the failure happens, not where the call came from.** A
  first cut set `navigationError` inside the hook's own `goTo()` wrapper; but
  the hook also hands out `viewer`, the real handle, and hosts pass that
  around (the bench card, the site's paging bar), so `viewer.goTo()` bypassed
  it and the bench and site needed workarounds to route through the wrapper.
  Instead `preview()`'s handle fires `error` on a failed `goTo()` (still
  rejecting), the same event and `{ message, cause }` detail
  `<aared-viewer>` fires, and the hook listens to it as it does to
  `indexchange`. Every caller lands; the workarounds were reverted, and the
  two embed layers that are event targets now report navigation the same
  way. Cancellations and overtaken calls fire nothing.
- **Silent no-ops.** `preview().goTo()` after `destroy()` and the hook's
  `goTo()` before mount resolved as if they had moved. Both reject with
  `closed` now (no new code: nothing is open to move in).
- **A race found while there:** an overtaken `preview().goTo()` returned from
  its stale render and still set `index` and announced its own unit, so two
  quick calls could leave `index` naming a unit not on screen. `renderUnit`
  now reports whether it painted, and only the call that did moves `index`.

Tests: `react.visual.test.ts` (a failed `goTo()` keeps the slide and sets
`navigationError`, a unit change clears it, a failed `viewer.goTo()` sets it
too, `goTo()` while loading rejects `closed`), `preview.visual.test.ts` (a
failed `goTo()` fires one `error` with the element's shape; two `goTo()` in
one tick announce only the second, which failed against the old `goTo`;
`goTo()` after `destroy()` rejects). The bench's reference card keeps calling
`viewer.goTo()`; a failed slide or sheet change now shows as a notice bar
from `navigationError` instead of replacing the view with its error screen.

## 58. Error codes say whose fault it is (Sep 2026)

A design review of the error surface after §57 (a reviewer given only the
files and two questions: is this a good design, is anything wrong) judged
the model sound (one class, a closed code union, `aborted`/`closed`/
`worker-error` kept apart, one event shape) and found the codes coarser than
they look. Overturns two calls in §57.

**`invalid-document` meant "something failed", not "the file is broken".**
§57 typed every untyped Worker throw `invalid-document` with "re-save it in
Office". That blamed the file for aared's own defects (a `TypeError` in a
parser) and for valid content: `pptx.ts` threw on a paragraph alignment or
cell anchor it did not know, so a good deck was told to repair itself. Now
the fault decides the code:

```text
 the file is broken (a check that names what) -> damagedDocument(...) where the check is
                                                  invalid-document + repair hint
 not a ZIP / not an Office package              -> invalid-document, no repair hint
 malformed XML (XmlError)                       -> invalid-document + repair hint
 corrupt deflate (the platform's TypeError)     -> ZipError("corrupt") -> invalid-document
 an unknown value in a readable file            -> the default, read on (as Office shows it)
 anything else                                  -> internal, original error as `cause`
```

`damagedDocument` (`core/errors.ts`) replaces about 20 plain throws across
`opc`, `pptx`, `theme`, `xml`, `xlsx/` and `docx/`. The unknown-value
helpers (`align`, `verticalAlign`, `onOff`, a non-numeric colour modifier)
fall back instead of throwing: one odd attribute no longer costs a slide.
`internal` is new: "aared failed on this file", never "re-save it", with the
Worker's original error rebuilt as `cause` (name, message, stack) for the bug
report. The wire type's `code` is now required, and the main thread's
untyped-error branch, dead once every failure is typed, is gone.

**Usage errors had borrowed other codes.** "Index is not a whole number" was
`not-found`, the element's `open()` with no source was `not-found`, and the
hook's `goTo()` before mount was `closed` (§57). New codes split them the way
the platform does (`TypeError` vs `InvalidStateError`): `invalid-argument`
(a value the call does not take; `itemAt`, `preview`'s `assertIndex`, the
element's index check) and `invalid-state` (a call the viewer is not ready
for). `not-found` is left for whole numbers and names that do not exist;
`closed` for after teardown.

**A crashed Worker read as a closed document afterwards.** `terminate()`
stopped the bridge and every later request rejected `closed`. The bridge
keeps why it stopped and rejects later requests with it, so a call after a
crash reports `worker-error`. The crash itself now carries the browser's
event as `cause`, and its message when it has one.

**Smaller fixes from the same review and the open gaps:** a download that
drops mid-body is `network-error` (was the platform's raw `TypeError`);
`preview()`'s cancellation uses the shared `abortError()` (one wording); the
element and `preview()` share one empty-document check (`assertHasUnits`);
an element `open()` that a newer one replaced rejects `aborted` instead of
resolving as if it had finished (no `error` event: nothing failed);
navigation inside a unit (`viewer.xlsx`, `viewer.docx`) fires the handle's
`error` like `goTo()`, so the hook's `navigationError` hears it too; XLSX's
render handle reports font fidelity (`warnings`, measured after fonts
settle, as PPTX and DOCX do; it was always empty).

**`messageFor()` stays, documented for what it is:** one generic English
sentence per code for a UI, without the detail `message` carries, not
localized; a host in another language maps `code` itself. The API
reference's Errors section now has the code-to-meaning-to-caller table, the
message policy, `cause`, and the listen-or-catch rule for the event layers
(the event and the rejection carry the same instance by design).

**Not changed:** `AaredError.name` stays `"AbortError"` for cancellation so
existing `AbortSignal` handling keeps working; the docs now say to branch on
`code`. `scrollTo()`'s wrong-shape target stays `invalid-range` (it is a
range question). The reviewer's note on `DOMException` in `serializeError`
does not apply to the supported browsers.

Tests: the in-process Worker turns an injected `RangeError` into `internal`
with its stack and a corrupt deflate stream into `invalid-document`
(`worker-cancellation.test.ts`); a crash keeps reporting `worker-error`, an
`internal` failure arrives with `cause` in the caller's scope, a mid-body
drop is `network-error` (`session-lifecycle.test.ts`); `itemAt`'s split
(`error-helpers.test.ts`); `goTo(1.5)`, and `goToHeading`/`goToPage`
failures firing `error` (`preview.visual.test.ts`); a superseded element
`open()` rejecting `aborted` with no event (`element-viewer.visual.test.ts`);
the hook's before-mount `invalid-state` (`react.visual.test.ts`); XLSX
`warnings` naming a missing family (`xlsx-render.visual.test.ts`). The
unknown-value fallbacks have no test of their own.
