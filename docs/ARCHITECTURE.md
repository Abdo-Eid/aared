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
with correct Arabic/RTL. The ones that exist either rasterize to canvas (and
lose RTL, selection, and accessibility in the process) or pull in an entire
office suite (Univer, LibreOffice-WASM) to do it. aared exists to be the
missing middle: bounded scope, browser-native text, view-only.

## The one decision everything else follows from

**Text is rendered by the browser — DOM or SVG — never rasterized by us.**

The browser's text stack (HarfBuzz under the hood) gives correct shaping,
bidi, kashida, ligatures, and font fallback for free. The moment a renderer
draws glyphs into a `<canvas>` itself, it has taken on the job of
reimplementing that stack — and Arabic is exactly where that reimplementation
shows: paragraph-level RTL (`algn="r"`, `rtl="1"`) gets dropped, because
canvas output is pixels with no CSS escape hatch. This is a diagnosed,
specific failure — see `docs/archive/ms-previewrs/PPTXVIEWJS-NOTES.md` for
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
                DrawingML (shapes, ~180 preset geometries, pictures, text
                bodies) · charts → SVG · theme resolution · font resolver ·
                image pipeline · EMF/WMF (wasm)
                    runs inside a Worker; emits resolved, ready-to-paint
                    UnitModels — nothing format-specific touches main thread
  ├── pptx      slide model → one positioned <div> per shape: inline <svg>
  │             for geometry, HTML <div> for the text body
  ├── xlsx      own styles.xml parser (Community-edition SheetJS silently
  │             returns empty cell.s — see docs/archive/.../SHEETJS-LIMITS.md)
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
with formats already applied, font families already resolved). The main
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
| Arial / Times New Roman / Courier New | Liberation Sans / Serif / Mono |
| Traditional Arabic / Sakkal Majalla / Simplified Arabic / Arabic Typesetting | Noto Naskh Arabic / Noto Sans Arabic / Amiri |

Metric compatibility matters because font metrics drive line breaking, which
drives everything downstream of it — a wrong-metrics substitute produces
wrong layout, not just a wrong typeface. Fonts load lazily via `FontFace`,
registered on `document.fonts` (so they cross shadow-DOM boundaries), and
only for families the open document actually references. Embedded PPTX fonts
(`ppt/fonts/*.fntdata`, obfuscated in the package) are de-obfuscated and
preferred over the substitution table when present. The font files
themselves live in the separate `aared-fonts` package — `core` never forces
their download — with a CDN default so a plain `<script>`-tag consumer still
gets working fallback fonts without bundling anything.

Before optional font delivery exists, the Worker still resolves every run to
the family Office requested: direct and inherited run typefaces take priority,
then the applicable theme major/minor Latin or complex-script default. This is
not cosmetic. A run that inherits the embedding application's font can change
its width and clip inside an otherwise correct authored text box. The renderer
therefore paints the resolved family rather than depending on page-level CSS.

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

The fixture corpus must include **Arabic decks with tables from the first
commit**. That combination is precisely the failure mode that ruled out an
off-the-shelf PPTX renderer during evaluation (see the archive notes linked
above) — it's the regression this project exists to not reintroduce.

## Packaging

One published package, `aared`, with per-format subpath exports
(`aared/core`, `aared/pptx`, `aared/xlsx`, `aared/docx`, `aared/element`).
The web component `import()`s a format's renderer only the first time a file
of that type is opened, so bundle cost is decided by which formats a
consumer's users actually touch — not by package boundaries. A second
package, `aared-fonts`, ships the OFL font files separately so consumers
without Arabic documents never download them. See `API.md` for the surface
this produces.

## Scope boundary: PDF

PDF is deliberately out of scope. `pdf.js` already solves client-side PDF
rendering well; duplicating it would spend effort on the one format that
isn't the gap. Consumers who need PDF too pair aared with `pdf.js` directly.

## Further reading

The full decision log — including alternatives considered and why each was
rejected — lives in `docs/archive/ms-previewrs/DESIGN.md`. This document is
the distilled, current-state version of it; the archive is the working
history behind it.
