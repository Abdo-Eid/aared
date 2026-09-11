# Decision Log

> Purpose: a chronological record of what was decided and why, with the
> reasoning behind each call. Read this when you ask "why is it shaped this
> way" months from now. For the current architecture, see
> [ARCHITECTURE.md](ARCHITECTURE.md). For the public API, see
> [API.md](API.md). For what's queued next, see [ROADMAP.md](ROADMAP.md).

This is a distilled record from a single design session (2026-09-03), not a
rediscovered history — the raw decision table it's drawn from, with every
alternative considered, lives at
[`archive/ms-previewrs/DESIGN.md`](archive/ms-previewrs/DESIGN.md).

## 1. Why this project exists at all

The trigger was concrete, not speculative: evaluating client-side renderers
for an existing chat product's document-delivery panel surfaced two
disqualifying findings, preserved as primary evidence in
[`archive/ms-previewrs/PPTXVIEWJS-NOTES.md`](archive/ms-previewrs/PPTXVIEWJS-NOTES.md)
and
[`archive/ms-previewrs/SHEETJS-LIMITS.md`](archive/ms-previewrs/SHEETJS-LIMITS.md):

- `pptxviewjs` paints PPTX to `<canvas>`, which drops paragraph-level RTL
  (`algn="r"`, `rtl="1"`) with no way to patch it back in — canvas output is
  pixels, and pixels have no CSS escape hatch.
- SheetJS Community Edition returns an *empty* `cell.s` even when asked for
  styles (`cellStyles: true`), silently — it doesn't fail, it just returns
  nothing usable, which is the trap.

Neither is a bug in those libraries so much as a structural consequence of
how they're built. Searching for an alternative confirmed a real gap: no
lightweight, view-only, browser-native library spans PPTX/XLSX/DOCX with
correct Arabic/RTL. The choices are "own an entire office suite in the
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
gate fixture (`arabic-docx-bidi-coverage.docx`) carries one explicit `w:br w:type="page"`,
zero `w:lastRenderedPageBreak` hints, zero `w:pageBreakBefore` flags, and one
`nextPage` section boundary — yet Word prints four pages. The missing page is
layout overflow inside the first section, so page markers are evidence, not
pagination. Unpacked-OOXML counts and the Word print reference confirming
this are recorded in `DOCX_PAGINATION_AUDIT_AND_PROPOSAL.md`.

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
