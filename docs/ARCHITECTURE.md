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
aared-fonts     OFL substitutes (Carlito, Caladea, Arimo, Tinos, Cousine;
                Radio Canada, Roboto, IBM Plex Sans Condensed for Aptos;
                Amiri, Lateef, Markazi Text for Arabic) + a CDN default
                for the font resolver
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
| `errors.ts`, `search.ts`, `session.ts` | `isAbort`, `itemAt` (a whole number outside the list is `not-found`, anything else `invalid-argument`), `damagedDocument` (the one `invalid-document` for a broken file, thrown where the check is; the Worker types any other throw `internal`, §58); `foldText`, `runsText`, `findByTitle` (every title lookup: exact, then case-insensitive, never a substring); `WorkerSession.openAs`, `SessionMedia`, `sourceBytes` (the one source-to-bytes path, with progress, for `preview` and `<aared-viewer>`), and the Worker pool (`acquireBridge`: a closed document's clean Worker serves the next open, stopped after 15 s idle; DECISION_LOG §38) |
| `unit-cache.ts` | the bounded, most-recently-used cache of parsed slides and sheets a Worker session keeps |
| `index.ts` | the shared reference shapes: `UnitRef` (`{ index, title }`) and `CellRef` (`{ address, row, col }`) |

**`src/*.ts`** (main thread; reached from several entrypoints, so emitted as
shared chunks under `dist/core/`, where the worker URL also resolves):

| Module | Owns |
|---|---|
| `appearance.ts` | the light/dark resolve and the six published `--aared-*` design tokens (`--aared-ink`; `--aared-text` is `<aared-viewer>`'s consumer input and no renderer sets it) |
| `media-url.ts` | `MediaUrls`: one object URL per media part per render, all revoked together |
| `class-sheet.ts` | renderer defaults as classes: `cssClass()` interns constant declarations at module scope; one constructed stylesheet per render root (the document, or a shadow root), counted by the renders using it, adopted before a render attaches its DOM and removed with the last one. Inline style carries only what a node authors; no `<style>` text or style attribute is ever written (strict CSP); hot paths use `setAttribute`, not `dataset`; `pseudoElements` styles `::before`/`::after` (the `ltrRun` marks). DOCX, XLSX and PPTX renders all adopt it (DECISION_LOG §36, §44) |
| `text-style.ts` | a run's CSS and a paragraph's box, shared by PPTX and DOCX; `exactLinePitch`/`EXACT_PITCH_RUN` flattens every run on an exact-pitch line (`w:lineRule="exact"`, PowerPoint's `a:spcPts`) to `line-height:0` so a mixed-font line can't stretch past its authored pitch (§44) |
| `font-fidelity.ts` | which fonts a render painted with, for all three formats: `awaitDocumentFonts` (load the faces a model uses, bounded), `detectFontFidelity`/`fontFidelityWarnings` (the `font-unavailable` warnings). Apart from `text-style.ts` because the bundler splits shared code by module: XLSX needs only this |
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
  fallback grid; DOCX pads in physical left/right too, draws side borders on the
  inline start/end of a `w:bidiVisual` table, and collapses shared edges — a
  side border's half-width reaches into whichever cell does not paint it,
  `sideInsetPx`, §44) and each chart's `<svg>` shell and fallback box.
- `tint`. SpreadsheetML's is an sRGB lerp, DrawingML's is linear light and
  probe-backed; only the arithmetic underneath is shared.
- `TextRunModel.verticalAlign` (Word) against `baselinePct` (DrawingML): one
  concept in two dialects, each painted as measured against its own app.
- The facades' `render` calls: the element adds sheet tabs and the resolved
  theme; `preview()` forwards the caller's options. Their errors are the same
  `AaredError` codes (DECISION_LOG §31), and both fire `error` with the same
  `{ message, cause }` detail for a failure the caller may not await (§57);
  the empty-document check is one helper, `assertHasUnits` (`format-route.ts`).
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
virtualization is deferred). A content-sized XLSX sheet (no height of its
own, as `<aared-viewer>` and the bench give it) scrolls itself inside a
clipping ancestor's visible room rather than letting that ancestor scroll
the whole sheet — otherwise the sheet's sticky frozen rows, row numbers and
column letters would scroll away with it, and its own scrollbar would sit
below the fold (DECISION_LOG §44). Images decode lazily via
`createImageBitmap({ resizeWidth })` only once in view, and are revoked once
scrolled far enough away — memory stays proportional to the viewport, not to
the file.

**OPC laziness.** The package format itself makes this cheap: reading slide 1
of a 200-slide deck touches its index and that slide's ZIP entry, not the
whole archive. Renderers retain only a small near-viewport model cache; moving
away from a unit makes it eligible for eviction just as decoded media is.

## Fonts

Office documents routinely name fonts nobody has installed — Aptos, Calibri,
Sakkal Majalla, Traditional Arabic, Simplified Arabic. **The default is the
browser's own fonts:** with no resolver, a family the reader lacks falls to
its generic, nothing downloads, and on a real phone that was both faster
(0.3-0.5 s against 1.1-7.2 s) and easier to read than the substitutes below,
at the cost of some lines breaking a word differently from Office. So
`aared-fonts` is frozen and unpublished (DECISION_LOG §53); the rest of this
section describes it for a consumer that needs Office's line breaks exactly.

`core` holds a resolver
interface, `fontAlias()` (the naming rule a resolver's delivered substitute
registers under — see *Aliasing*, below), and `aared-fonts` holds a default
mapping to open substitutes:

| Requested | Substituted | Kind |
|---|---|---|
| Calibri / Calibri Light | Carlito | Metric clone (0 of 95 ASCII glyphs differ from real Calibri) |
| Cambria | Caladea | Metric clone (sub-pixel units-per-em quantization only, under 0.13%) |
| Arial | Arimo | Metric clone (0 differ) |
| Times New Roman | Tinos | Metric clone (0 differ) |
| Courier New | Cousine | Metric clone (0 differ) |
| Aptos / Aptos Display / Aptos Narrow | Radio Canada / Roboto / IBM Plex Sans Condensed, calibrated per family and weight | Not a clone — see *Aptos* |
| Traditional Arabic / Arabic Typesetting | Amiri, calibrated per font | Not a clone — see *Arabic calibration* |
| Simplified Arabic | Lateef, calibrated | Not a clone — see *Arabic calibration* |
| Sakkal Majalla | Markazi Text, calibrated | Not a clone — see *Arabic calibration* |
| Arabic text set in Arial / Times New Roman / Cambria / Calibri / Aptos | Lateef or Markazi Text as a second, Arabic-only face on the Latin substitute's alias | See *Arabic under Latin-only fonts* |

Metric compatibility matters because font metrics drive line breaking, which
drives everything downstream of it — a wrong-metrics substitute produces
wrong layout, not just a wrong typeface. Fonts load lazily via `FontFace`,
registered on `document.fonts` (so they cross shadow-DOM boundaries), and
only for families the open document actually references. Embedded fonts
(PPTX `ppt/fonts/*.fntdata`, DOCX `word/fonts/*.odttf`; both obfuscated in the
package) are not read yet — see *Current state* below. The font
files themselves live in the separate `aared-fonts` package — `core` never
forces their download — with a CDN default so a plain `<script>`-tag consumer
still gets working fallback fonts without bundling anything.

Before optional font delivery exists, the Worker still resolves every run to
the family Office requested: direct and inherited run typefaces take priority,
then the applicable theme major/minor Latin or complex-script default. This is
not cosmetic. A run that inherits the embedding application's font can change
its width and clip inside an otherwise correct authored text box. The renderer
therefore paints the resolved family rather than depending on page-level CSS.

### Aliasing (DECISION_LOG §45)

A resolver never registers a delivered substitute under the requested name
itself (that would shadow a reader who genuinely has the requested font
installed, and would apply a calibrated substitute's size/metric overrides to
that real, unrelated font) or under the substitute's own name (nothing then
connects it to the paragraph asking for "Calibri" — this was an unverified,
and as it turned out wrong, assumption before DECISION_LOG §45: a substitute
was registered but never actually selected). `fontAlias(family)` (`core/
typography.ts`, re-exported from `aared/core`) is the one shared naming rule:
`cssFontFamily` always emits a three-part stack — the requested name, its
alias, then the nearest generic — and any resolver that delivers something
for a family registers its `FontFace` under that family's alias. An unaliased
requested name that happens to be installed always wins; an alias nobody
registered costs nothing (CSS just skips to the generic).

### Arabic calibration (DECISION_LOG §45)

No Office Arabic font has an open metric clone — they're Microsoft-licensed,
and Arabic's contextual letterforms make cloning harder than Latin's simpler
outlines (a search for one found none). `aared-fonts` instead calibrates an
open face per font, chosen by measurement, not by name association: every
Arabic-only run in the fixtures (59 real paragraphs, real shaping via
`measureText`) in the real Windows font file, against eleven candidates loaded
from the same Google Fonts files the resolver delivers
(`tools/probe/font-metrics/arabic-survey.ts`). The criterion is the *spread*
of per-paragraph width ratios, since line breaks depend on it while the mean is
exactly what `size-adjust` removes; the height of alef and of the seen body is
then checked after adjusting. Amiri was the original, narrower choice for all
four; a candidate had to beat it by a quarter to displace it. Result: Amiri
for Traditional Arabic and Arabic Typesetting, Lateef for Simplified Arabic
(47% less spread), Markazi Text for Sakkal Majalla (41% less). Each gets
per-font `FontFace` descriptors: `sizeAdjust` scales the substitute's
advances to match real shaped-text widths, and `ascentOverride`/
`descentOverride` are computed to compensate for `size-adjust` scaling them
too (verified empirically — a browser multiplies an override percentage by
size-adjust's own factor, it does not apply the override to the font's
un-adjusted em). The ascent and descent are the original's own `usWin`
metrics, read unrounded: canvas returns them in whole pixels, and readings
at 40px were off by up to 2.5% of the em (DECISION_LOG §46). Widths were measured per weight and bold differs from
regular by up to 5.6%, so Traditional Arabic carries separate bold numbers.
Fonts sharing a substitute share one CDN fetch but register under different
aliases, each with its own calibration (`FONT_CALIBRATION` in
`aared-fonts/src/index.ts`). Arabic Typesetting is a modern reading of Naskh,
very compact: Amiri needs a 67% `size-adjust`, and a large adjustment is not
itself a defect — the letter heights land at the original's proportions.
Letterforms still differ; widths are what line breaks follow.

### Aptos (DECISION_LOG §49)

Aptos is Office's default theme font since 2023: Aptos for body text, Aptos
Display for headings, Aptos Narrow as Excel's default. No browser has it.
Office downloads it into a per-user cache
(`%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts`) and never installs it as a
system font, and it has no open clone. Each family is calibrated like the
Arabic fonts, against the real files from that cache
(`tools/probe/font-metrics/latin-survey.ts`, `latin-fit.ts`), on every Latin
run in the fixtures plus some written prose.

- **Choice.** Width must match, and so must the look: a face whose letters
  end up 10% taller once its widths are fitted is not a look-alike. Of 61 open
  candidates fetched with the resolver's own request, exactly one per family
  keeps x-height and cap height within 4% of the original at both weights:
  Radio Canada (Aptos), Roboto (Aptos Display) and IBM Plex Sans Condensed
  (Aptos Narrow).
- **Line breaks.** Laid out across eight widths, the three break 93-96% of
  paragraphs exactly where Aptos does. Painting the generic sans-serif
  (Arial's widths), what happened before, breaks 38-75% that way.
- **Aptos Display** is Aptos spaced tighter for large sizes. A uniform
  `size-adjust` cannot reproduce that, so it takes its own face.
- **Bold** gets its own numbers in all three families. Even a 1.3% weight
  difference cost Aptos Narrow's bold 4 points of identical line breaks.
- **Vertical metrics** are the ones Chromium paints Aptos with, its
  typographic ascent and descent (0.939/0.282 of the em). Aptos sets
  USE_TYPO_METRICS, and its win metrics (1.010/0.275) are unused. The line box
  matches the real font's within 0.1%; Arial's is 5.8% shorter.
- **Italic** is covered by the upright calibration of the same weight. Its
  widths are 2-3% narrower than Aptos's own italic in Aptos and Aptos Display.
- **Scope.** Not covered: Aptos Light, SemiBold, ExtraBold, Black, Mono and
  Serif, which are separate Office families.

### Arabic under Latin-only fonts (DECISION_LOG §45)

Most Arabic in real documents is not set in an Arabic-named font. In the
fixtures Arial carries 183 Arabic runs and Times New Roman 12, while the four
fonts above carry none; Arial's and Times New Roman's own Arabic glyphs match
Simplified Arabic's widths within 1.3% (they are the face Windows pairs with
them). Their substitutes Arimo and Tinos have no Arabic at all, so before this
that text fell through to whatever `sans-serif` the reader's system has — on
Windows that happens to be Arial's own Arabic, which is why it never showed on
the development machine. The resolver now adds a second face to the *same
alias*: the Arabic subset of Lateef (Arial, Times New Roman, Cambria) or
Markazi Text (Calibri, Calibri Light — Calibri 6 has its own Arabic),
restricted by `unicode-range` to the Arabic script blocks and calibrated to the
original's Arabic. The calibration is measured through the alias, never the
Arabic face alone: the spaces, Western digits and punctuation inside Arabic
text paint in the Latin face at its own size, so the Arabic face's
`size-adjust` is fitted with them in place (`tools/probe/font-metrics/
arabic-fallback-fit.ts`). Measured alone, it counted its own spaces and left
Arial's Arabic 3% too wide (DECISION_LOG §46). Composition was verified in a real browser before it was
built (`tools/probe/font-metrics/alias-composition.ts`): Latin measures
exactly as Arimo alone, Arabic exactly as calibrated Lateef alone, spaces and
punctuation stay with the Latin face (as they do in real Arial), and the add
order is irrelevant. Cambria has no Arabic glyphs; Word draws that text in
Times New Roman (read off Word's PDF export,
`tools/probe/font-metrics/word-arabic-fallback.ps1`), so it is calibrated
against Times New Roman's Arabic, through Caladea (which paints its spaces). Two costs are kept off documents that don't need them: the
face is left to load on demand rather than eagerly, and `awaitDocumentFonts`
and the fidelity probe put Arabic letters in their sample only when the model
has Arabic text (`collectFontUse`), so a Latin-only Arial document downloads
nothing extra.

Aptos has no Arabic glyphs, and Excel and Word both draw Arabic set in any
Aptos family in Arial, at Arial's widths. This was read off each app's PDF
export (`tools/probe/font-metrics/aptos-arabic-fallback.ps1`). In a
spreadsheet it is the ordinary case: a cell has one font and no
complex-script slot, and Aptos Narrow is Excel's default. So the Aptos
families take Arial's Lateef face, fitted through their own alias. The spaces
there come from the size-adjusted Latin face, narrower than Arimo's, so
Lateef is set 3-5% larger than under Arial. PPTX and DOCX already paint a run
containing Arabic in its complex-script font, so there this face only matters
for text that reaches the Aptos alias anyway.

Not covered: Courier New (its Arabic is monospace — every proportional
candidate is off by 45% or more on individual paragraphs) and Tahoma and
Segoe UI (not in the substitution table). A single font for both scripts
(Vazirmatn) was benchmarked and rejected: it loses English and mixed line
breaks and ties Arabic (§46).

### Current state (Sep 2026)

How a font gets from a file to the screen, and what is and is not wired.

1. **Resolve.** The Worker resolves every run to the family Office asked for
   (direct, then inherited, then theme default) and the renderer paints that
   name, plus its alias (see *Aliasing*), plus the nearest generic. Shared by
   all three formats.
2. **Request.** `fontResolver` is opt-in and main-thread-only
   (`OpenOptions.fontResolver`, forwarded by `preview()`, the React hooks and
   `<aared-viewer>`). When a unit is read, the session
   (`core/session.ts`, `resolveFonts`) collects the family names in the
   resolved model (`requestedFontFamilies`) and calls the resolver once per
   family, memoized for the session; `unit(…, { fonts: false })` skips it. It
   awaits settlement of the resolver's promise, not a font file: faces load
   when text needs them. A renderer that measures painted text or checks
   which fonts painted loads exactly the faces its model paints first
   (`awaitDocumentFonts`, per family and bold/italic variant, bounded at
   2 s): DOCX layout pagination before it lays out a page, both DOCX modes
   and PPTX `render()` before they read their fidelity warnings
   (DECISION_LOG §45, §50). Every format's
   `open()` reads through the same `unit()`, so DOCX and XLSX use the same
   hook as PPTX (read from the session code; not exercised by a DOCX or XLSX
   test beyond the shared session tests).
3. **Deliver.** `aared-fonts` `defaultResolver` (`createResolver`).
   `substituteFamily()` maps the requested family (case-insensitive, quotes and
   whitespace stripped) to its substitute; an unknown family resolves
   `undefined` with no network traffic. A known one fetches the substitute's
   `@font-face` CSS from Google Fonts once per *substitute* (shared across
   every requested family that maps to it), registers every subset face under
   `fontAlias(requestedFamily)` on `document.fonts` (with its `unicode-range`,
   any `FONT_CALIBRATION` descriptors, `font-display: swap`) and resolves the
   regular face. It loads nothing itself: the browser fetches a face when
   text first needs it, so a document downloads only the subsets and styles
   it paints (English in regular weight: one 10-30 KB file per family;
   DECISION_LOG §50). It never rejects: an
   offline or failed fetch resolves `undefined` and the browser falls back to
   its own fonts. With no `FontFace` or `document` (headless) it is inert.
4. **Policy.** Opt-in, never a CDN default: nothing downloads unless the
   application passes a resolver (DECISION_LOG §12, which points to the
   zero-network-by-default rule in §6). The roadmap's "CDN default" names the
   package's default resolver, not a default-on behavior. Independent of a
   resolver, a render reports which of its own families actually painted with
   something real (`DocxRenderHandle.warnings` / `PptxRenderHandle.warnings`)
   — see *Fidelity flag*, below — so an embedder without a resolver still
   knows when a render is approximate.

**Status.**

- Done: the resolver hook, the alias mechanism, the substitution table, the
  Arabic calibration, and `defaultResolver`, with unit tests in `packages/
  aared-fonts` (table, calibration sanity, alias registration — never the
  requested name or the substitute's own name, two requested families sharing
  one CDN fetch but registering separately, CSS URL and parsing, failure
  resolves `undefined`, headless inert) against a fake `FontFace` and `fetch`,
  plus an end-to-end visual test in `packages/aared` that renames a run to a
  family guaranteed not to exist anywhere and proves the registered alias is
  what actually paints, not the generic fallback (this used to be an
  unverified assumption — see *Aliasing*, above).
- The `warnings` fidelity flag (below): done for both DOCX and PPTX, unit and
  visual-tested against a real render.
- Embedded fonts (DOCX `.odttf`, PPTX `.fntdata`) are not read — see
  § Known limits.
- The bench passes no `fontResolver`, so it paints with whatever fonts the
  machine has installed. On the development machine (Windows, Word installed)
  every substitution above except Aptos's is calibrated against the real,
  installed Office font, so this machine cannot exercise the substitute path
  in the bench — the end-to-end visual test (above) is what actually proves
  it. Aptos is not installed there either (it lives in Office's cloud-font
  cache), so its text paints in the generic sans-serif in the bench and in
  the paired PowerPoint/Excel captures, which pass no resolver either.
- The no-Office-fonts review (`bun run test:visual:artifacts
  --absent-fonts`; `tools/word-render/capture.ts --absent-fonts`) is where
  the substitutes paint real fixtures: every family is renamed to one no
  machine has and the real resolver delivers under that name, since an
  installed font cannot be hidden from a page (DECISION_LOG §51). For a real
  phone, serve the bench on the network (`.claude/launch.json` `bench-lan`,
  `vite --host`) and open it with and without `?fonts=aared` (§53).
- Recalibrating, if `aared-fonts` is ever reopened: fit an Arabic fallback
  face through its alias, never alone, since the Latin face paints the
  spaces (`tools/probe/font-metrics/arabic-fallback-fit.ts`); read vertical
  metrics unrounded, at 10000px; measure Aptos from Office's cloud-font cache
  (`%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts`, filled only after a
  document using the font is opened; `cloud-fonts.ts` finds the hashed files
  by name table), and take its typo metrics as the browser reads them. The
  calibration machine has every real font, so a wrong number never shows in
  the visual suite without `--absent-fonts` (§46, §49).

**DOCX layout depends on font metrics.** The layout-mode paginator measures
real block and row heights off screen (`measureFlow`, DECISION_LOG §44), which
now waits for the fonts a section actually uses to finish loading first
(DECISION_LOG §45) rather than measuring whatever was painted first. A
headless caller (no DOM) still falls back to a per-script average glyph
advance estimate (`AVERAGE_CHAR_EM` in `core/docx/pagination.ts`). A fallback
font with different widths than the estimate's calibration changes the
painted wrapping in that headless path, so estimated pages can break
differently from measured ones — one more reason substitutes should be
metric-compatible (Latin) or calibrated (Arabic).

### Fidelity flag

Every render handle carries `.warnings: readonly RenderWarning[]`
(`{code, message}`, `core/model.ts` — shared by every format, not DOCX-only;
`RenderHandleBase.setWarnings()`, `render-handle.ts`). A DOCX layout-mode
render merges the paginator's own edge-case warnings (`oversize-block`,
`footnote-overflow` — computed since DECISION_LOG §21 but previously dropped
before reaching the embedder, a ROADMAP gap) with `code: "font-unavailable"`
entries; DOCX flow mode and every PPTX and XLSX render carry the font
entries only. Detected by `detectFontFidelity`/`fontFidelityWarnings` in
`font-fidelity.ts` (DOM-required, so a renderer module, never
`core/session.ts`, which Rule 3 keeps DOM-free; its own module, apart from
`text-style.ts`'s paragraph painting, so XLSX downloads only the check): each
family's rendered width is compared, via `canvas.measureText` (no DOM
insertion), against a same-generic-category control guaranteed to resolve to
nothing of its own. `document.fonts.check()` cannot do this job at all —
verified empirically (`tools/probe/font-metrics/fonts-check-probe.ts`), it
reports `true` even for a family that exists nowhere on the calibration
machine. One real exception this comparison must not flag: Arial, Times New
Roman and Courier New are themselves what a browser's own generic keyword
commonly already resolves to (`resolvesLikeGenericDefault`,
`core/typography.ts`) — measuring the same as the control there is
inconclusive, not evidence of absence, and the pixels are correct either
way. A working delivered substitute (`detectFontFidelity`'s `"substituted"`
status) is not itself a warning — only `"unavailable"` (neither the
requested name nor an alias painted with anything real) is.

## Format rules

What Office actually writes and paints, each found against the native app
(DECISION_LOG entries named). Read the one for your format before touching
its renderer; when a paired capture is off, check what the file stores
before changing paint code.

**All formats.**

- An exact line pitch is the paragraph's, never a run's: every run on a
  `w:lineRule="exact"` or PowerPoint `a:spcPts` line paints at
  `line-height:0` (`EXACT_PITCH_RUN`, `text-style.ts`), so a
  differently-metriced run can't stretch the line; `exactLinePitch()` leaves
  an `atLeast` floor free to grow (§44).
- A "line" of paragraph spacing is not the CSS `lh` unit. Word's is 12pt,
  and a line value above zero wins over points, inherited separately
  (`paragraphSpacePt`). PowerPoint's is 1.2 × the paragraph's largest run,
  painted as `em` on the `<p>`; PowerPoint drops space before a text body's
  first paragraph and after its last and adds after + before between
  paragraphs, which is why every PPTX text body is a flex column (§47).
- A left-to-right run in a right-to-left paragraph keeps "71%" or
  "15 Oct 2026" together (`TextRunModel.ltrRun`, marks added in CSS
  pseudo-content, §44).
- OOXML words are not CSS words: `w:vertAlign` `superscript` maps to
  `super`/`sub`, and a browser silently drops anything it doesn't know. Pin
  such mappings with computed-style assertions.
- Line-box clipping is a metrics defect, never fixed by turning clipping
  off: DrawingML's `0.9` stays in the model and the renderer clamps the
  browser line height to `1.1` (PPTX's leading shift accounts for the clamp).

**PPTX.**

- PowerPoint reads its layout from the file; it does not compute it (§48).
  `normAutofit` shrinks only by the stored `fontScale`/`lnSpcReduction`
  (applied in `parseTextBody`, rounded to whole points); a bare one paints at
  100% with the overflow showing, so there is no measuring pass.
- A bullet is a marker box at least as wide as the hanging gap, with no line
  height of its own; `marL` counts from the right in RTL (§48).
- `masterTextStyle` and `masterPlaceholderType` must agree, or a subtitle
  loses the master body spacing (§48). Inset precedence is nearest-present
  master → layout → slide (`tools/probe/title-insets/`).
- Every line and connector is `p:cxnSp` (§44).
- `bodyPr` attributes (`anchor`, `wrap`, the four insets) cascade one at a
  time, nearest first: slide, layout, master. A slide `bodyPr` that only holds
  `a:normAutofit` still takes the master's anchor (§54).
- Text past its box paints in every autofit mode; only the slide edge clips
  it. A text body no `a:bodyPr` in the chain gives insets to takes DrawingML's
  0.1in left/right and 0.05in top/bottom, resolved in `parseTextBody` like a
  table cell's margins (§54).
- Line spacing off 100% moves the text, not the box: PowerPoint sits it
  `0.9 x spacing x size` from the top of the line above 100%, and puts the
  removed leading above it below 100%; CSS centres it. `leadingShiftEm`
  (`pptx/index.ts`) paints the difference as a relative `top` on the `<p>`,
  measured against the height actually painted, floor included (§54).
- Colours resolve through the slide's own `p:clrMapOvr`, else the layout's,
  else the master's `p:clrMap` (`overrideColorMap`). `a:lnRef` widths come from
  the theme's `a:lnStyleLst` (`DrawingTheme.lineWidths`; the 2013-era Office
  theme states 0.5, 1 and 1.5pt, the Aptos-era one 1, 1.5 and 2pt, so it is
  read, never assumed). A table style the deck does not define but PowerPoint
  ships (Medium Style 2, seven ids) is built from PowerPoint's own definition
  on first use (`tableStyleFor`); any other missing style paints unstyled (§54).
- A run tagged with a left-to-right `lang` in a left-to-right paragraph keeps
  Arabic words in written order around ASCII punctuation (`W1 + W2`,
  `W1 (W2) W3`); spaces alone, an Arabic `lang` and right-to-left paragraphs
  follow plain bidi. `parseTextBody` cuts the run at the punctuation and flags
  those pieces `ltrRun` (`arabicPunctuationPieces`, §54). In a right-to-left
  paragraph the same kind of run stays plain bidi (a punctuation-order
  difference PowerPoint has there is accepted, §54).
- A slide's `showMasterSp="0"` hides layout and master graphics; a layout's
  hides the master's (`tools/probe/master-graphics/`, §54).

**XLSX.**

- Context-ordered cells render `dir="auto"`: the sheet's `dir="rtl"` mirrors
  columns only, and inheriting it reordered `180,000 EGP` and `13 Sep 2026`.
- An empty cell has no text span, no `dir` and no alignment; a row's fill
  columns are one cell (`data-aared-fill="tracks"`) and a fill row one cell
  (`data-aared-fill="row"`) with gradient gridlines. Fill lines come from
  `syncFillLines()` custom properties, which run before the first window, on
  scale and on `resize` (device pixel ratio).
- The grid paints its first window in the same task it is attached (a frame
  later showed a short sheet without its fill), and a `max-height`
  clipper's room is its limit, not its current height (§40).
- A content-sized sheet caps its own scroll box to a clipping ancestor's
  visible room (above `MIN_SCROLL_ROOM_PX`, 120px), so frozen rows, row
  numbers and column letters stay sticky (`syncViewport`, §44). Code that
  resizes the scroller after it picked up a cap (the print-capture harness
  grows it to fit every row) must lift the cap before each measurement and
  reapply it after, or the stale cap masks the resize.
- `headers` chrome is sticky, exclusion-aware and off by default;
  virtualized re-renders restore focus to the selected cell
  (`xlsx-keyboard.test.ts`).
- The Excel review compares A4 **print**, never screen; print-title rows and
  the page header have no aared counterpart.

**DOCX.**

- Word writes a right-to-left table as `w:bidiVisual`: borders logical, cell
  margins physical, `tblInd` from the start side (§43). It centres a table's
  side border on the grid line, so the text on either side sits a margin
  beyond the border's inner half, whichever cell paints it (`sideInsetPx`,
  §44).
- `alignFor` swaps `left`/`right` in RTL paragraphs; a header hangs from its
  `w:header` distance and is never clipped; line boxes follow run sizes
  (`lineFontSizePt`) at about 1.15em (`WORD_SINGLE_LINE`); first-block
  spacing drops after breaks (`dropsTopSpacing`). Shared line-box constants
  live in `core/docx/pagination.ts`.
- A paragraph with no `w:pStyle` takes the style marked `w:default="1"`
  (`StyleData.defaultParagraphStyleId`), not the bare document defaults (§55).
- Anchored pictures leave the runs for the paragraph's `floats`.
  `planFloat` (`docx/float-layout.ts`) turns the anchor and the text column
  (`AnchorFrame`: the section column in flow mode, the page in layout mode, a
  cell's text width inside a table) into a one-side CSS float, a full-width
  band, an absolute overlay, or a split. A split (`bothSides` with room on
  each side) is laid out by `split-wrap.ts`: rows of (slot | hole | slot)
  filled in reading order, words measured with canvas and checked against the
  painted DOM (§55).
- There is no `paragraph-text` wrapper; layout pages are
  `content-visibility: auto` and lay out their tabs when first shown
  (`settlePageTabStops`, §37). One `MediaUrls` owner per render, one Blob URL
  per part.

## Fidelity contract

aared is a viewer, not a converter. Each format has an explicit, honest
promise:

| Format | In scope | Out of scope |
|---|---|---|
| **PPTX** | Shapes, fills, gradients, images, autofit text, paragraph-level RTL, tables, common chart types, z-order, rotation. Selectable, searchable text. Speaker notes are read onto `SlideModel.notes`, not painted. | Embedded fonts (§45), WordArt warps (painted flat and marked), animations, transitions, SmartArt (its fallback image renders if present), 3D effects, VBA, media playback |
| **XLSX** | Cached values with number formats, cell styles (fill/font/border/alignment), merges, column widths/row heights, frozen panes, sheet RTL, hidden rows/cols, cached conditional-formatting results, common chart types, images | Formula evaluation, pivot recalculation, data-validation UI, macros, sparklines |
| **DOCX** | Flow-rendered HTML with correct RTL, heading outline, tables, images, lists, footnotes, pictures with Word's text wrap (square, tight, through, top and bottom, behind and in front), tracked changes shown as final. Layout mode adds physical pages from section geometry plus heuristic overflow: correct page count/orientations, hard breaks, per-page headers/footers/footnotes, evaluated `PAGE`/`NUMPAGES` | **Pixel-identical Word pagination** — line wrapping and break positions match structurally, not exactly. Tables never split across pages yet; `nextColumn` degrades to continuous. |

This table is the boundary of what a bug report is versus a feature request.

## Known limits

Differences from Office that are settled: accepted as not worth the code
(rare, or not something a reader notices) or kept by design. They are not
open work; differences still meant to be closed are in
[ROADMAP.md § Open gaps](ROADMAP.md#open-gaps-not-blocking-pre-release).
A limit that turns out to matter goes back there.

| Format | Limit | Decision |
|---|---|---|
| PPTX | Off-center `path="circle"` and `path="shape"` gradients are not painted (shapes, cells, WordArt and backgrounds alike): a dashed outline in dev builds, `data-aared-fill-fallback` / `data-aared-background="unsupported"`. COM cannot author either, so a fix would need a hand-made slide appended to `pptx-coverage.pptx`, plus a probe. Every other fill kind paints (§28) | Accepted — rare in real decks |
| PPTX | Fill approximations: patterns in group children anchor to their child-space position rather than the slide origin; picture tiles assume 96 DPI; gradient stops follow a fitted smoothstep, within a few colour levels of PowerPoint | Accepted |
| PPTX | Tables: a built-in style PowerPoint ships but the deck does not define paints unstyled unless it is one of the seven Medium Style 2 ids (§54) | Accepted |
| PPTX | Arabic around ASCII punctuation in a left-to-right paragraph's `en-US` run follows PowerPoint (§54); the neighbour check stays inside one run, and Word is unprobed (DOCX unchanged) | Accepted |
| PPTX | In a right-to-left paragraph's `en-US` run, PowerPoint lays a stretch of punctuation out in written order (`.»` ending a quotation shows `.` outermost; aared shows `»`). Built once, then removed as not worth a model field and nested bidi wrappers (§54) | Accepted — not worth the code |
| PPTX | Text with no run size takes `p:defaultTextStyle`'s size per level; which of it and the master's `p:otherStyle` PowerPoint uses when they disagree is unprobed (both fixtures agree) | Accepted — rare |
| PPTX | A slide-level `p:clrMapOvr` does not reach the layout's shared graphics (the layout model is shared per layout). Implemented from the spec; no fixture carries an override (§54) | By design |
| PPTX | Paragraphs at 70%-90% line spacing step 1.1em apart, not tighter, so Arabic ink is not cropped; exact spacing (`a:spcPts`) and bullet sizes under a stored `fontScale` are unprobed (§54) | Accepted — good enough |
| PPTX | Only `slidenum` and `datetime` `a:fld` fields are live; other field types keep their cached text | Accepted |
| PPTX | Synthetic slide titles (`"Slide N"`) in `info.titles` and `slide.info`: the real title would parse every slide at `open()` (§31); `slide.read()` has the real text | By design |
| PPTX | Notes are parsed onto `SlideModel.notes` (reachable through `slide.read()`) but nothing paints or searches them | By design — the viewer is slides only |
| XLSX | The sheet notice bar paints in development builds only; a host that wants to tell users reads `SheetModel.unsupported` itself | By design |
| DOCX | Justified paragraphs approximate Word 2013's space-shrinking conservatively (§44): one row still wraps a word earlier than Word, and English Calibri measures 0.3-0.5% wider in Word, so a few English lines wrap differently | Accepted — deliberately conservative |
| DOCX | A mid-paragraph `w:lastRenderedPageBreak` hint is dropped: paragraphs are atomic in pagination | Accepted |
| Cross | Embedded fonts (DOCX `word/fonts/*.odttf`, PPTX `ppt/fonts/*.fntdata`) are not read, with no unsupported marker: almost no real files embed fonts, and reading them costs every DOCX open a font-table read. A working DOCX implementation is kept only in a local `git stash` named "DOCX embedded fonts (.odttf): working implementation, deferred"; PPTX was never investigated (§45) | Accepted — may never be built |
| Cross | `aared-fonts` is frozen and unpublished; the default is the browser's own fonts (§53). Reopen only for a consumer that needs Office's line breaks exactly, starting from the network cost (the resolver fetches Google Fonts stylesheets before the first paint) | Decided |
| Cross | Font substitutes, with `aared-fonts` only: Arabic set in Courier New is not substituted (its Arabic is monospace; no proportional candidate calibrates); Tahoma and Segoe UI are not in the table (Office's themes set Arabic in Arial and Times New Roman, and both ship with Windows); Aptos Light, SemiBold, ExtraBold, Black, Mono and Serif have no substitute; Aptos italics paint with the upright calibration of their weight (1-3% narrow), and the CDN's files can drift up to 0.9% (§49) | Accepted — good enough |
| Cross | Without Office fonts: Lateef, the Arabic face under Arial, Times New Roman, Cambria and Aptos, is set 127-133% to match widths, so its strokes cross dense table borders and an underline cuts through letters; justified Arabic can break a word earlier or later (§51) | Accepted — good enough |
| Cross | Symbol-font characters outside bullets and list markers, and Wingdings 2/3 and Webdings anywhere, keep their letter codes, so a device without those fonts paints letters (§52) | Accepted |

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
package, `aared-fonts` (a font resolver with calibrated substitutes), stays in
the repo frozen and unpublished (DECISION_LOG §53). See `API.md` for the
surface this produces.

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
