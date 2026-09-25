# Architecture

> What's designed, how it fits together, and why it's shaped this way. For
> the reasoning behind each choice, see [DECISION_LOG.md](DECISION_LOG.md).
> For the public API this produces, see [API.md](API.md). For what's built
> so far vs. queued, see [ROADMAP.md](ROADMAP.md). For the
> landscape this project sits in, see [COMPETITORS.md](COMPETITORS.md).

## Why this exists

PPTX, XLSX and DOCX are, underneath, the same packaging format — Office Open
XML (OOXML): a ZIP of related XML parts. But they solve three different
rendering problems wearing one file-extension family:

| Format | Layout problem | Bound? |
|---|---|---|
| XLSX | A grid. The hard part is virtualizing 100k rows, not layout. | Bounded |
| PPTX | Absolutely-positioned shapes; text *inside* a shape still needs flow + autofit; shapes are DrawingML geometry. | Bounded |
| DOCX | Full flow layout — line breaking, pagination, tables, floats. This is a word processor's layout engine. | Unbounded |

No existing lightweight, view-only, browser-native library spans all three
with correct Arabic/RTL *through the browser's own DOM/SVG text stack*.
Single-format DOM renderers exist per format, and Canvas/WASM engines span
all three at suite weight — the missing middle is all-three, view-only,
Worker-parsed, DOM/SVG-painted, with demonstrated bidi. See
[COMPETITORS.md](COMPETITORS.md) and the measured snapshot in
[BENCHMARKS.md](BENCHMARKS.md).

## The one decision everything else follows from

**Text is rendered by the browser — DOM or SVG — never rasterized by us.**

The browser's text stack (HarfBuzz under the hood) gives correct shaping,
bidi, kashida, ligatures, and font fallback for free. The moment a renderer
draws glyphs into a `<canvas>` itself, it has taken on the job of
reimplementing that stack — and Arabic is exactly where that reimplementation
shows: paragraph-level RTL (`algn="r"`, `rtl="1"`) gets dropped, because
canvas output is pixels with no CSS escape hatch. This is a diagnosed,
specific failure — see `docs/archive/ms-previewrs/PPTXVIEWJS-NOTES.md`
(git history before `5a32bf0`) for
the library that hit it and couldn't be patched around it.

Consequence: WASM is a **leaf tool**, not the rendering core. It's used for
exactly two things where no native browser API or reasonable JS decoder
exists — EMF/WMF vector-image decoding (common in real Office files) — and
nowhere near text or layout. ZIP decompression uses the native
`DecompressionStream`; XML parsing uses core's headless XML parser so it
remains testable outside a browser.

## Package layout

```
  core            OPC (central-directory zip · rels · content types) · XML ·
                  DrawingML (shapes, 11 fast-path presets; the rest and
                  custGeom deferred to the main thread's ECMA table, see
                  DECISION_LOG §26 — pictures, text bodies) ·
                  charts → SVG · theme resolution · font resolver ·
                image pipeline · EMF/WMF (wasm)
                    runs inside a Worker; emits resolved, ready-to-paint
                    UnitModels — nothing format-specific touches main thread
  ├── pptx      slide model → one positioned <div> per shape: inline <svg>
  │             for geometry, HTML <div> for the text body
  ├── xlsx      own styles.xml parser (Community-edition SheetJS silently
  │             returns empty cell.s — see DECISION_LOG.md §1)
  │             → virtualized DOM grid: merges, frozen panes, cached values,
  │             number formats, sheet-level RTL
  └── docx      flow-rendered HTML with correct paragraph-level RTL and a
                 heading outline; `pagination: { mode: "layout" }` opts into
                 physical pages (section geometry + heuristic overflow).
                 See Fidelity contract below.
element         <aared-viewer>: viewport virtualization, element pooling,
                lazy media decode, shadow DOM, optional built-in controls
aared-fonts     OFL font files (Carlito, Liberation, Noto Naskh/Sans Arabic,
                Amiri) + a CDN default for the font resolver
```

`core` is a real shared layer, not an aspirational one, because the three
OOXML formats genuinely share machinery:

- **OPC** — the ZIP + `_rels` + `[Content_Types].xml` package model is
  identical across all three.
- **DrawingML** — shape geometry, pictures, and text bodies (`a:p` / `a:r` /
  `a:rPr`) use the same XML vocabulary in a slide, a spreadsheet drawing
  layer, and an inline Word image.
- **Charts** — the chart XML parts are the same regardless of host format,
  so one chart-to-SVG renderer serves all three.
- **Theme** — color and font-scheme resolution (`theme1.xml`) works
  identically everywhere it's referenced from.

### Shared modules

A rule that two formats both need lives once, and what differs per format is a
parameter, not a second copy. Look here before writing a helper.

**`core/`** (DOM-free; also bundled into the Worker):

| Module | Owns |
|---|---|
| `units.ts` | the px / pt / inch / EMU / twip constants and conversions. A leaf: anything may import it |
| `color.ts` | hex parse and format, sRGB mixing, linear light, luminance |
| `typography.ts` | the Arabic-script test and the line-height floor (`MINIMUM_LINE_HEIGHT`) |
| `chars.ts` | every character above U+00FF aared emits at runtime (default bullet, non-breaking hyphen, LRM/RLM, footnote return arrow), built at runtime from ASCII escapes (`JSON.parse`; rolldown folds `String.fromCharCode`). No other shipped string, template or regex literal may hold one, or a `\u` escape of one: V8 keeps a script with any wider character two bytes per character, the Worker included (DECISION_LOG §34) |
| `theme.ts` | `themeFontFamily`, how a theme font slot resolves: an empty complex-script slot falls to the theme's `Arab` script font (PPTX and DOCX) |
| `xml.ts`, `opc.ts` | `numberAttr`; `ownString` / `ownTree` (the one place strings are copied out of part text, for anything kept past a parse); `OpcPackage.readXml`, `readXmlIfPresent`, `officeDocument(format)`, `mediaContentType`, and the format registry `officeFormatOf`, the one place that decides which content types are which format |
| `xlsx/limits.ts`, `xlsx/range.ts` | the worksheet grid and the one A1 scanner |
| `errors.ts`, `search.ts`, `session.ts` | `isAbort`, `itemAt`; `foldText`, `runsText`, `findByTitle` (every title lookup: exact, then case-insensitive, never a substring); `WorkerSession.openAs`, `SessionMedia`, `sourceBytes` (the one source-to-bytes path, with progress, for `preview` and `<aared-viewer>`), and the Worker pool (`acquireBridge`: a closed document's clean Worker serves the next open, stopped after 15 s idle; DECISION_LOG §38) |
| `unit-cache.ts` | the bounded, most-recently-used cache of parsed slides and sheets a Worker session keeps |
| `index.ts` | the shared reference shapes: `UnitRef` (`{ index, title }`) and `CellRef` (`{ address, row, col }`) |

**`src/*.ts`** (main thread; reached from several entrypoints, so emitted as
shared chunks under `dist/core/`, where the worker URL also resolves):

| Module | Owns |
|---|---|
| `appearance.ts` | the light/dark resolve and the six published `--aared-*` design tokens (`--aared-ink`; `--aared-text` is `<aared-viewer>`'s consumer input and no renderer sets it) |
| `media-url.ts` | `MediaUrls`: one object URL per media part per render, all revoked together |
| `class-sheet.ts` | renderer defaults as classes: `cssClass()` interns constant declarations at module scope; one constructed stylesheet per render root (the document, or a shadow root), counted by the renders using it, adopted before a render attaches its DOM and removed with the last one. Inline style carries only what a node authors; no `<style>` text or style attribute is ever written (strict CSP); hot paths use `setAttribute`, not `dataset` (DECISION_LOG §36) |
| `text-style.ts` | a run's CSS and a paragraph's box, shared by PPTX and DOCX |
| `table-dom.ts` | the table grid and cell shells (PPTX, DOCX) |
| `chart-svg.ts` | the chart primitive painter (PPTX, XLSX) |
| `render-handle.ts` | `RenderHandleBase` (zoom, painted `scale`, change-only `zoomchange`, idempotent destroy — every format's handle, XLSX included), `safeScale`, `sameZoom`, and the fit math every format uses: `measureFitRoom` (room with the content set aside, so a fit can't read back its own result; scrolled ancestors are put back afterwards so setting content aside never moves the reader) and `fitScale` |
| `docx/outline.ts` | `headingOutline`: the one DOCX heading index space (body headings only), shared by `headings.list()`, `goToHeading`, the facade and search |
| `diagnostics.ts` | `MARK_FALLBACKS`: whether development diagnostics paint (PPTX dashed fallback outlines, XLSX sheet notice bar) — on from source, `false` in shipped bundles |
| `format-route.ts` | the lazy per-format open, `DEFAULT_ZOOM` (each format's opening zoom), unit announcements (`indexchange` + `slidechange`/`sheetchange` as `{ index, title }`), XLSX tab following, and event forwarding shared by `preview` and `<aared-viewer>` |

Tests and tools follow the same rule: `test/harness.ts`, `test/zip.ts` and
`test/docx-harness.ts` (non-test helpers), and `tools/shared.ts` for the three
`*-render` tools.

**Deliberately per-format** — do not "fix" these into one:

- Each renderer's extras on top of the shared base: PPTX's outline, gradient and
  shadows; DOCX's shading, super/subscript, dominant family and borders.
- `lnSpc`. DrawingML's multiple is the font size itself and Word's single line
  is about 1.15em, each validated against its own Office app, so
  `applyParagraphBox` takes it as a parameter (`singleLine`).
- Table padding and borders (PPTX pads in physical left/right with a faint
  fallback grid; DOCX pads inline-start/end and collapses shared edges) and each
  chart's `<svg>` shell and fallback box.
- `tint`. SpreadsheetML's is an sRGB lerp, DrawingML's is linear light and
  probe-backed; only the arithmetic underneath is shared.
- `TextRunModel.verticalAlign` (Word) against `baselinePct` (DrawingML): one
  concept in two dialects, each painted as measured against its own app.
- The facades' `render` calls: the element adds sheet tabs and the resolved
  theme; `preview()` forwards the caller's options. Their errors are the same
  `AaredError` codes (DECISION_LOG §31).
- The baseline adapters under `bench/`, which parse OOXML themselves so a
  baseline column never depends on the code it measures.

The decisions behind these are DECISION_LOG §23.

## Performance model

**Index first; parse the requested unit second.** Opening a document reads
only the package central directory plus its lightweight format index: for
PPTX that is `ppt/presentation.xml` and its slide relationships; for XLSX,
the workbook's sheet list. This produces `DocumentInfo` without inflating
every slide or sheet. Only `getUnit(i)` inflates and resolves that one unit.
This is a shared performance rule, not a PPTX-only trick: one visible unit
must not make a 200-slide deck or 100-sheet workbook behave like every unit
is visible. DOCX has a different flow-layout model, so its block
virtualization begins after its one document part is read.

**Parse in a Worker, render on main thread.** The worker walks the ZIP's
central directory and inflates entries on demand — never holding the whole
decompressed package in memory — turning each requested part into a resolved,
structured-cloneable `UnitModel` (geometry as SVG path data, fills as CSS
colors/gradients, text as styled runs with resolved direction, cached values
with formats already applied, font families already resolved). Large parts
never exist as a whole XML tree: `sheetData`, `sharedStrings` and DOCX's
`w:body` stream through `parseXml`'s `streamChildrenOf` mode. A Worker
session keeps nothing the page already holds: a DOCX session keeps the
package, its media set and a small index (styles, numbering, theme,
relationships, title), and `document.xml`, the notes and the parsed sections
are dropped once unit 0 is posted. Anything kept past a parse is copied with
`ownString` / `ownTree` (`core/xml.ts`), because a string sliced from a part
keeps the whole part text alive (DECISION_LOG §35). Cells live in a sorted
per-row grid rather than a per-cell map, so whole-column hyperlink and
conditional-format ranges cost the cells that exist, not the range's area.
`.rels` parts, layout/master chains, and DOCX header/footer parts are parsed
once per package and reused. Media bytes transfer to the main thread without
a copy. Format parsers load lazily via dynamic `import()`, so a splitting
bundler ships shared code plus one format. The main
thread never parses XML; it only paints models it's handed. Communication is
a typed request/response RPC over `postMessage`, with `AbortSignal` wired to
a cancel message so switching files mid-load doesn't leave the worker
computing something nobody will see.

**Virtualize everything.** Slides and spreadsheet rows are windowed to the
viewport; DOCX flow blocks are designed independently but caller-owned
slicing is the expected virtualization boundary (in-library flow
virtualization is deferred). Images decode lazily via
`createImageBitmap({ resizeWidth })` only once in view, and are revoked once
scrolled far enough away — memory stays proportional to the viewport, not to
the file.

**OPC laziness.** The package format itself makes this cheap: reading slide 1
of a 200-slide deck touches its index and that slide's ZIP entry, not the
whole archive. Renderers retain only a small near-viewport model cache; moving
away from a unit makes it eligible for eviction just as decoded media is.

## Fonts

Office documents routinely name fonts nobody has installed — Calibri, Sakkal
Majalla, Traditional Arabic, Simplified Arabic. `core` holds a resolver
interface and a default mapping to **metric-compatible, OFL-licensed**
substitutes:

| Requested | Substituted |
|---|---|
| Calibri | Carlito |
| Cambria | Caladea |
| Arial / Times New Roman / Courier New | Liberation Sans / Serif / Mono |
| Traditional Arabic / Sakkal Majalla / Simplified Arabic / Arabic Typesetting | Noto Naskh Arabic / Noto Sans Arabic / Amiri |

Metric compatibility matters because font metrics drive line breaking, which
drives everything downstream of it — a wrong-metrics substitute produces
wrong layout, not just a wrong typeface. Fonts load lazily via `FontFace`,
registered on `document.fonts` (so they cross shadow-DOM boundaries), and
only for families the open document actually references. Embedded fonts
(PPTX `ppt/fonts/*.fntdata`, DOCX `word/fonts/*.odttf`; both obfuscated in the
package) are meant to be de-obfuscated and preferred over the substitution
table when present — **not implemented yet**, see *Current state* below. The font files
themselves live in the separate `aared-fonts` package — `core` never forces
their download — with a CDN default so a plain `<script>`-tag consumer still
gets working fallback fonts without bundling anything.

Before optional font delivery exists, the Worker still resolves every run to
the family Office requested: direct and inherited run typefaces take priority,
then the applicable theme major/minor Latin or complex-script default. This is
not cosmetic. A run that inherits the embedding application's font can change
its width and clip inside an otherwise correct authored text box. The renderer
therefore paints the resolved family rather than depending on page-level CSS.

### Current state (Sep 2026)

How a font gets from a file to the screen, and what is and is not wired.

1. **Resolve.** The Worker resolves every run to the family Office asked for
   (direct, then inherited, then theme default) and the renderer paints that
   name. Shared by all three formats.
2. **Request.** `fontResolver` is opt-in and main-thread-only
   (`OpenOptions.fontResolver`, forwarded by `preview()`, the React hooks and
   `<aared-viewer>`). When a unit is read, the session
   (`core/session.ts`, `resolveFonts`) collects the family names in the
   resolved model (`requestedFontFamilies`) and calls the resolver once per
   family, memoized for the session; `unit(…, { fonts: false })` skips it. It
   awaits settlement of the resolver's promise, not the font file's load. Every
   format's `open()` reads through the same `unit()`, so DOCX and XLSX use the
   same hook as PPTX (read from the session code; not exercised by a DOCX or
   XLSX test).
3. **Deliver.** `aared-fonts` `defaultResolver` (`createResolver`).
   `substituteFamily()` maps the requested family (case-insensitive, quotes and
   whitespace stripped) to its substitute; an unknown family resolves
   `undefined` with no network traffic. A known one fetches the substitute's
   `@font-face` CSS from Google Fonts once per family, registers every subset
   face on `document.fonts` (with its `unicode-range`, `font-display: swap`),
   starts loading in the background, and resolves the regular face. It never
   rejects: an offline or failed fetch resolves `undefined` and the browser
   falls back to its own fonts. With no `FontFace` or `document` (headless) it
   is inert. The mapping is `FONT_SUBSTITUTIONS`: Calibri and Calibri Light →
   Carlito, Cambria → Caladea, Arial → Liberation Sans, Times New Roman →
   Liberation Serif, Courier New → Liberation Mono, Traditional Arabic and
   Sakkal Majalla → Noto Naskh Arabic, Simplified Arabic → Noto Sans Arabic,
   Arabic Typesetting → Amiri.
4. **Policy.** Opt-in, never a CDN default: nothing downloads unless the
   application passes a resolver (DECISION_LOG §12, which points to the
   zero-network-by-default rule in §6). The roadmap's "CDN default" names the
   package's default resolver, not a default-on behavior.

**Status.**

- Done: the resolver hook, the substitution table, and `defaultResolver`, with
  unit tests in `packages/aared-fonts` (table, CSS URL and parsing,
  once-per-family registration, failure resolves `undefined`, headless inert)
  against a fake `FontFace` and `fetch`.
- Not implemented: embedded fonts in PPTX (`ppt/fonts/*.fntdata`) and DOCX
  (`word/fonts/*.odttf`); XLSX is expected to be not applicable (Excel has no
  embed-fonts option), to be confirmed. All tracked in ROADMAP M4.
- The bench passes no `fontResolver`, so it paints with whatever fonts the
  machine has installed. On the development machine (Windows, Word installed)
  Arial and Cambria are real system fonts and DOCX text widths match Word's
  to about a pixel (a 509px title in Word against 508px in aared), so fonts
  were not the cause of the layout gaps found in the DOCX fidelity pass. On a
  machine without those fonts the bench would differ until a resolver is wired.

**DOCX layout depends on font metrics.** The layout-mode paginator does not
measure text: it guesses line wraps from a per-script average glyph advance
(`AVERAGE_CHAR_EM` in `core/docx/pagination.ts`, about 0.46em Latin and 0.40em
Arabic), calibrated against Arial and Arabic body text on Windows. A fallback
font with different widths changes the painted wrapping but not the estimate,
so pages can break differently from how they were filled. That is one more
reason substitutes must be metric-compatible.

**Open questions** (read from the code, not run — check before relying on them):

1. *Are delivered substitutes ever selected?* `createResolver` registers faces
   under the substitute's own name (`new FontFace("Carlito", …)`), while the
   renderers paint the requested name (`font-family: Calibri`).
   `resolveRunFamily` emits a single family rather than a stack, and nothing in
   `packages/aared/src` mentions the substitute names (a search for Carlito,
   Liberation, Caladea and `substituteFamily` finds nothing). Unless something
   aliases them, a substitute is registered but never chosen on a machine that
   lacks the requested font. Where the requested font is installed the browser
   uses it and the substitute is irrelevant. To verify: open a real deck or
   document in the bench with `defaultResolver` passed, on a machine without the
   font (or with it blocked), and check the computed font actually used.
2. *Browser-level coverage.* The `aared-fonts` tests are all unit-level with
   fakes. No test loads a real face in a browser or asserts that a renderer
   paints with it (`packages/aared/test` was not searched for one).
3. *Type stub.* `aared-fonts/src/index.ts` keeps a structurally identical local
   copy of `FontResolver` until it can import it from `aared/core`; the two
   must be kept in sync by hand.

## Fidelity contract

aared is a viewer, not a converter. Each format has an explicit, honest
promise:

| Format | In scope | Out of scope |
|---|---|---|
| **PPTX** | Shapes, fills, gradients, images, autofit text, paragraph-level RTL, tables, common chart types, embedded fonts, z-order, rotation, notes as text. Selectable, searchable text. | Animations, transitions, SmartArt (its fallback image renders if present), 3D effects, VBA, media playback |
| **XLSX** | Cached values with number formats, cell styles (fill/font/border/alignment), merges, column widths/row heights, frozen panes, sheet RTL, hidden rows/cols, cached conditional-formatting results, common chart types, images | Formula evaluation, pivot recalculation, data-validation UI, macros, sparklines |
| **DOCX** | Flow-rendered HTML with correct RTL, heading outline, tables, images, lists, footnotes, basic floats, tracked changes shown as final. Layout mode adds physical pages from section geometry plus heuristic overflow: correct page count/orientations, hard breaks, per-page headers/footers/footnotes, evaluated `PAGE`/`NUMPAGES` | **Pixel-identical Word pagination** — line wrapping and break positions match structurally, not exactly. Tables never split across pages yet; `nextColumn` degrades to continuous. |

This table is the boundary of what a bug report is versus a feature request.

## Quality strategy

**PowerPoint is the typographic reference.** A `bun:test` suite drives
`Bun.WebView` (built into Bun 1.4+) to load aared's render and capture real
DOM/SVG screenshots. `tools/powerpoint-render/` uses native PowerPoint
automation to export the same fixture to same-size PNGs. The retained pairs and
their `review.html` are inspected for text metrics, bidi placement, and
wrapping. The suite also re-reads the rendered DOM to assert semantic layout
properties such as RTL bullet side, table column order, resolved table bands,
and text-range containment under a nondefault host font. The latter matters
because the package renderer is embedded in arbitrary application styles while
the capture harness has a deliberately minimal page stylesheet.

**Excel is the print reference.** `tools/excel-render/` uses native Excel
automation to print every visible sheet to A4 (hardcoded policy: portrait,
fit sheet on one page, gridlines plus row/column headings, normal margins),
then rasterizes each sheet PDF with PyMuPDF at 150 DPI. The XLSX capture
test fits the aared grid — with the `headers` renderer option on — onto the
same A4 page box and pairs native print against aared fit in `review.html`.
Pairs are scaled to equal display width and judged by a human, never
pixel-diffed; the printed page header has no aared counterpart yet.

**Word is the layout reference for DOCX.** `tools/word-render/` uses native
Word automation on a private copy of the fixture (fields updated, never
saved), exports it to PDF, and rasterizes each page with PyMuPDF at 96 dpi —
CSS pixels, the size aared paints — while also dumping every text span's font,
size and position. Layout rules (header geometry, line boxes, spacing at page
tops, superscript geometry, bidi alignment) are read from those numbers rather
than judged from a screenshot; DECISION_LOG §22 records what that overturned.
`capture.ts` renders the same pages through aared for side-by-side review. It
is run by hand; it is not yet a `bun:test` capture suite like the PowerPoint
and Excel ones.

The fixture corpus must include **Arabic decks with tables from the first
commit**. That combination is precisely the failure mode that ruled out an
off-the-shelf PPTX renderer during evaluation (see the archive notes linked
above) — it's the regression this project exists to not reintroduce.

## Packaging

One published package, `aared`, with per-format subpath exports
(`aared/core`, `aared/pptx`, `aared/xlsx`, `aared/docx`, `aared/element`).
The web component `import()`s a format's renderer only the first time a file
of that type is opened, so bundle cost is decided by which formats a
consumer's users actually touch — not by package boundaries. The Worker is
one self-contained script (`dist/core/worker.mjs`, every format parser
bundled flat — no sibling chunks to deploy), which also backs Blob-URL
workers directly through the `worker` option. A second
package, `aared-fonts`, ships the OFL font files separately so consumers
without Arabic documents never download them. See `API.md` for the surface
this produces.

## Scope boundary: PDF

PDF is deliberately out of scope. `pdf.js` already solves client-side PDF
rendering well; duplicating it would spend effort on the one format that
isn't the gap. Consumers who need PDF too pair aared with `pdf.js` directly.

## Further reading

Why each decision was made, chronologically, is
[`DECISION_LOG.md`](DECISION_LOG.md). The first design session's raw
decision table — every alternative considered and why each was rejected —
was `docs/archive/ms-previewrs/DESIGN.md`, now in git history
(`git show 5a32bf0^:docs/archive/ms-previewrs/DESIGN.md`). This document is
the distilled, current-state version.
