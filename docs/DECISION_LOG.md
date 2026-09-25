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
so a number is read, not guessed. The full rule list is in `HANDOFF.md`
(*DOCX Word-fidelity pass*). What it overturned:

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
