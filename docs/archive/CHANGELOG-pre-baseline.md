# Pre-baseline development changelog

> **Archived — provenance only.** This is the development history distilled
> from the former Decision Log before establishing the October 14, 2026
> baseline. These are not published releases or promises of current support.
> Future changes belong in the root [CHANGELOG](../../CHANGELOG.md).
> Current contracts and limits are in [ARCHITECTURE](../ARCHITECTURE.md);
> the full reasoning and evidence survive in [Notes](../Notes.md).

Design discussions began September 3, 2026; the first Git commit was authored
September 11. The five periods below follow consecutive groups of actual
commits, not a retrospective regrouping of files. No package version is
invented for these periods. Dates and results inside Notes retain their
original context; a historical claim is not a new fidelity acceptance.

## September 11–18, 2026 — Engine, formats and preview API

### Added

- Repository scaffold, canonical design documents and typed package entries.
- Worker-backed PPTX parsing and DOM/SVG rendering, pictures, charts, autofit
  and resolved tables.
- XLSX parsing and virtualized grids with Arabic/RTL coverage.
- DOCX parsing and rendering, viewer routing for all three formats and the
  optional font resolver.
- DOCX navigation rail, zoom and structural Word-style layout pagination.
- The one-call preview facade, React hooks and a package page.

### Changed

- Reviewed XLSX against native Excel rather than treating another renderer
  as the reference.
- Redesigned the public API around resolved data and shared rendering contracts.

**Experience:** Notes' [design](../Notes.md#design-scope-and-native-evidence),
[API](../Notes.md#api-viewer-lifecycle-and-errors) and
[DOCX](../Notes.md#docx-layout-tables-and-drawings) topics preserve the initial
choices and their later corrections, including the early pagination limits.

## September 19–25, 2026 — Streaming, documentation and measurement

### Added

- Documentation navigation, code copying, playground snippets and measured
  performance figures on the home page.
- Native Word export tools and browser/Worker memory calibration.
- Deferred main-thread geometry, shared fill painting, lazy DOCX pages,
  Worker reuse and bounded retained state.
- Project branding and an npm name placeholder, not a package release.

### Changed

- Streamed large parts, cached relationships and reduced retained parse state.
- Tried a split Worker, then restored one self-contained file after consumer
  bundlers exposed Worker URL and loading problems.
- Standardized the async object-form API across formats.
- Used class-backed defaults and Latin-1 shipped JavaScript.
- Retired adaptive document dark mode in favor of source-faithful paper.

### Fixed

- Native PPTX typography, table and fill differences.
- Measurement rulers: page and Worker memory are read through distinct CDP
  sessions; heap readings, download sizes and failed passes are not conflated.
- Layout jumps when replacing or fitting a unit.

**Experience:** Notes' [Worker](../Notes.md#worker-packaging-and-memory),
[paint](../Notes.md#shared-rendering-paint-and-host-isolation) and
[measurement](../Notes.md#performance-benchmarks-and-navigation) topics.

## September 26–October 1, 2026 — Office fidelity, fonts and packaging

### Added

- Native probes for line spacing, script-specific emphasis, RTL tables,
  numbering, picture wrapping and preset geometry.
- Alias-based substitute-font calibration, readiness checks and absent-font
  reviews; Aptos comparisons and Unicode mappings for symbol-font bullets.
- Production consumer checks for Worker URLs under webpack and Next.js.
- Typed document errors distinguishing session closure, cancellation and
  caller/document/environment faults.

### Fixed

- Word pagination and line layout, PPTX connector geometry, RTL/LTR runs,
  placeholder inheritance, table styling and stored autofit behavior.
- DOCX floating pictures, including text flowing on both sides where Word
  does so.
- Delivered-font use: fonts are selected through their alias and awaited
  when painted, before layout measurements depend on them.

### Changed

- Froze the optional, unpublished font package after real-device latency and
  Arabic legibility results; browser fonts became the default.
- Closed the PPTX fidelity pass while keeping accepted differences explicit.

**Experience:** Notes' [fonts](../Notes.md#fonts-script-and-line-measurement),
[PPTX](../Notes.md#pptx-fidelity) and
[DOCX](../Notes.md#docx-layout-tables-and-drawings) topics retain measurements,
rejected approaches, native findings and the limits of the reviews.

## October 3–6, 2026 — Unified viewer, search and charts

### Added

- One viewer behind the embedding layers, toolbar methods, navigation,
  cancellable rendering and a Start page.
- Search owned by the viewer, Arabic spelling normalization, occurrence
  positions and cross-format highlighting.
- A resolved chart model read in the Worker and laid out on the page;
  chart-type coverage with explicit surface/map boundaries.
- Native chart corpora and vector/word comparisons against Office.
- DOCX text boxes, line grids, header/footer search and extended real fixtures.

### Fixed

- Chart defaults, layout assumptions and Arabic text behavior using native
  evidence, not guesses from specifications alone.
- Missing drawing identities, list counters, first-page/header behavior,
  empty section-ending paragraphs, text-box clipping/rotation/fills and
  preset text rectangles.
- Stale renders, cancelled handles and lifecycle failures across embedding
  layers.

### Changed

- Consolidated Office export and rasterization, with a standard probe shape.
- Search reads words instead of constructing whole models where appropriate;
  XLSX search follows the active-sheet scope.

**Experience:** Notes' [search](../Notes.md#search-and-highlighting),
[charts](../Notes.md#charts) and
[quality](../Notes.md#quality-tooling-and-documentation) topics.

## October 11–14, 2026 — Rendering, interaction and developer workflows

### Added

- Shared drawing/group effects and SVG-first images with Office raster fallback.
- XLSX conditional formatting for top/bottom and averages, data bars, icons,
  date periods, colour scales and bounded expressions over saved values.
- Spreadsheet formula bar, range/reference selection and inline aggregates.
- Borrowed content capabilities and explicit Worker resource cleanup.
- A single developer-tools dispatcher with grouped implementations and one
  ignored output home; preserved benchmark/research evidence.
- Shared visual page/server/build helpers and deferred-coverage proofs.

### Fixed

- Spreadsheet fit, selection past authored data, centred overflow and Summary
  layout; validated IFS value-range shapes.
- Action-owned error reporting and duplicate failure channels.
- Inherited host typography and cross-render/cross-bundle SVG resource names;
  search-sheet restoration and repeated default CSS.
- Tool import, report, argument, output-naming and declaration-build boundaries.

### Changed

- Worker parser imports became static and symbol-level; main-thread renderer
  laziness remains intentional.
- Removed only verified completed probe scratch topics, preserving ambiguous
  investigations, shared helpers and tracked chart/font evidence.
- Retired ROADMAP into ARCHITECTURE's one gap/limit inventory, release readiness
  and unscheduled proposals.
- Established this baseline: all 128 existing Decision Log entries move to
  topic-organized Notes with stable decision IDs; this development changelog
  is archived and the active changelog starts empty at `Unreleased`.

**Experience:** Notes' [XLSX](../Notes.md#xlsx-layout-and-conditional-formats),
[paint](../Notes.md#shared-rendering-paint-and-host-isolation) and
[quality](../Notes.md#quality-tooling-and-documentation) topics preserve the
full root causes, native evidence, blast radius and deliberate non-changes.

The complete original history is preserved in verified external Git bundles.
Only `main` is squashed; other branch, tag, stash and checkpoint refs retain
their original objects. This archive is not a replacement for that backup.

## Primary research appendices

The three source snapshots formerly named `DESIGN.md`, `PPTXVIEWJS-NOTES.md`
and `SHEETJS-LIMITS.md` are retained **inside this one archive**, not as separate
files. They were recovered from the original root commit because the Decision
Log's old hash-based retrieval instruction was already invalid. The snapshot
body text is preserved; heading levels are lowered to nest inside this file.
Names, API claims, proposed tools and predecessor-app recipes are historical,
not present-day contracts or claims about newer third-party versions.

- [Original design-session table](#research-design)
- [PPTXViewJS observations](#research-pptxviewjs)
- [SheetJS observations](#research-sheetjs)

<a id="research-design"></a>

### Original design-session table

> **Archived — provenance only.** Superseded design snapshot, including D1–D21 and alternatives considered. Its original introductory links were obsolete; use Notes, ARCHITECTURE and the API reference for current contracts.

#### aared (عارض) — design

Client-side previewer for PPTX, XLSX and DOCX. Web component + headless core.
Fast, low-memory, Arabic/RTL-correct, honest about fidelity.

Status: **design complete, ready to scaffold.** Decisions are recorded as
they're made. This doc is the source of truth — update it, don't fork it.

Author / copyright holder: **Abdulhamed Eid** (GitHub `Abdo-Eid`).
Name: عارض — "the one who displays / presents." Romanized `aared` (long ā of
ʿāriḍ). npm `aared`, `aared-fonts`; GitHub `Abdo-Eid/aared`; element
`<aared-viewer>`. All verified free on 2026-09-03 (`ared` is taken by an
unrelated Redis library).

##### 1. What we're building

One public open-source package that previews **PPTX, XLSX and DOCX** entirely
in the browser. Not an editor. Not a converter. A viewer.

**PDF is out of scope** — pdf.js already does it well; users pair `aared`
with it if they need both.

##### 2. Decisions

###### Scope & policy

| # | Decision | Why |
|---|---|---|
| D1 | **View-only, forever.** | No formula engine (render cached `<v>` values — every saved workbook has them), no round-tripping, no lossless-edit obligations. Halves the problem. |
| D2 | **Build order: PPTX → XLSX → DOCX.** | PPTX/XLSX have bounded layout and the biggest ecosystem gap; Arabic matters most there. DOCX is the bottomless pit — last, at an honest tier. |
| D3 | **Public OSS, Apache-2.0.** Copyright Abdulhamed Eid. `NOTICE` file carries the attribution and must be preserved by redistributors. | Attribution + NOTICE survive forks and rebrands; explicit patent grant; adoptable by companies — which is what builds credit. Fallback if stronger protection is ever wanted: MPL-2.0. Never AGPL. Dep licenses all compatible (ExcelJS MIT if used for reference; fonts OFL). |
| D4 | **Browser floor: Baseline 2023** — Chrome/Edge 110+, Firefox 115+, Safari 16.4+. | Safari 16.4 is the gate: `DecompressionStream`, `OffscreenCanvas`, `FontFace` in workers. Below that we'd polyfill exactly the things that make it fast. |
| D11 | **Per-format fidelity contract, stated up front.** See §4. | So nobody expects Word-perfect pagination from a DOCX preview. |

###### Rendering

| # | Decision | Why |
|---|---|---|
| D5 | **Text is rendered by the browser (DOM / SVG), never rasterized by us.** | The single decision that makes Arabic work: shaping, bidi, kashida, ligatures, font fallback for free. Owning pixels = shipping HarfBuzz + bidi + fonts. This is exactly why `pptxviewjs` fails on Arabic and can't be patched. |
| D6 | **WASM is a leaf tool, not the core.** Only planned use: EMF/WMF → SVG. | Common in real Office files, no JS decoder exists. Not for ZIP (`DecompressionStream` is native) or XML (`DOMParser` is native and fast). |
| D12 | **PPTX: one positioned `<div>` per shape**, containing an inline `<svg>` for geometry and an HTML `<div>` for the text body. Rotation/flip via CSS `transform` on the wrapper. | DOM order = z-order for free; text sits above its own fill and below shapes above it. No `foreignObject` (Safari bugs, broken printing). Text stays selectable, searchable, RTL-correct. |
| D13 | **XLSX: DOM grid**, row + column virtualized; sticky headers; merges as spanning cells; frozen panes as sticky regions; sheet `rightToLeft` flips the whole grid. | Keeps copy/select/accessibility and Arabic for free. Canvas would shape Arabic fine (`fillText` uses the browser engine) but loses all of that for no gain at preview scale. |
| D14 | **DOCX: flow-rendered HTML**, correct paragraph-level RTL, headings outline. **Not paginated.** | Faithful pagination is a word-processor layout engine. Say it loudly in the README. |
| D15 | **Charts: one DrawingML chart → SVG renderer** shared by all three formats. | Identical chart parts across PPTX/XLSX/DOCX. Bar, column, line, pie, doughnut, area, scatter first. |

###### Performance & memory

| # | Decision | Why |
|---|---|---|
| D7 | **Parse in a Worker, render on main thread.** Worker turns the ZIP into a small, structured-cloneable render model, per part, lazily. | OPC lets us read slide 1 without touching slide 200. Main thread only paints — no XML on main. |
| D8 | **Virtualize everything.** Slides, rows, columns, DOCX blocks. Decode images only in-viewport, at display resolution (`createImageBitmap` + `resizeWidth`), revoke when far off. Element pooling. | Memory stays proportional to the viewport, not the file. |
| D16 | **ZIP: parse the central directory only; inflate entries on demand** via `DecompressionStream`. | Never hold the whole decompressed package in memory. |
| D17 | **Worker protocol** (O6): typed request/response RPC over `postMessage` with a promise map and `AbortSignal` → cancel message. `open(source)` → `DocumentInfo` (format, unit count, titles/outline, sizes, sheet names, theme). `getUnit(i)` → `UnitModel`, fully resolved: geometry as SVG path data, fills as CSS colors/gradients, text as paragraphs/runs with resolved style + direction, number formats already applied to cached values, fonts as resolved family names. `getMedia(id)` → transferred `ArrayBuffer`. Main thread keeps an LRU of decoded media keyed by id and revokes on eviction. | Everything expensive or format-specific happens once, off-main. The model is the public headless API (D10). |

###### Fonts

| # | Decision | Why |
|---|---|---|
| D18 | **Core ships only the mapping table + `fontResolver` hook.** Font files live in **`aared-fonts`** (separate optional package) and are also served from a CDN as the resolver's default. Loaded lazily via `FontFace`, only families the document uses, registered on `document.fonts` (crosses shadow roots). | Keeps core small; users without Arabic never download Amiri. Metrics drive line breaks, so this is a top-3 fidelity lever. |
| D19 | **Mapping to metric-compatible OFL fonts:** Carlito ↔ Calibri; Liberation Sans/Serif/Mono ↔ Arial/Times/Courier; Noto Naskh Arabic / Noto Sans Arabic / Amiri ↔ Traditional Arabic / Sakkal Majalla / Simplified Arabic / Arabic Typesetting. **Embedded PPTX fonts** (`ppt/fonts/*.fntdata`, obfuscated) are de-obfuscated and loaded in preference to the table. | Right metrics → right line breaks → right layout. |

###### Packaging & API

| # | Decision | Why |
|---|---|---|
| D9 | **Single npm package `aared`, subpath exports, lazy-loaded per format.** `exports`: `.`, `./core`, `./pptx`, `./xlsx`, `./docx`, `./element`. The element `import()`s a format's renderer the first time that type is opened. | Bundle size is decided by code-splitting at the format boundary, not npm boundaries. One published package avoids core/renderer version skew. |
| D10 | **Headless core is a public API; the element is one consumer.** `open(file) → Document`; `doc.info`; `doc.unit(i) → UnitModel`; per-format `render(model, container)`. | This is the "flexible to some extent." Build your own UI on the model. |
| D20 | **Element surface** (O7): `<aared-viewer>`. Attributes `src` (URL), `type` (override; auto-detected from `[Content_Types].xml` otherwise), `page` (0-based unit index — slide / sheet / DOCX is one unit), `zoom` (`fit-width` \| `fit-page` \| number), `controls` (boolean, like `<video controls>` — built-in nav/zoom bar). Property `file` (File/Blob/ArrayBuffer). Events `load` (detail: `DocumentInfo`), `pagechange`, `progress`, `error`. Methods `open()`, `goTo(n)`. Shadow DOM; styling via `::part(stage)`, `::part(page)`, `::part(controls)` and CSS custom properties. | Minimal chrome; flexibility lives in core, not the element. |
| D21 | **Repo layout** (O8): single repo `Abdo-Eid/aared`, bun workspaces. `packages/aared` (published), `packages/aared-fonts` (published), `fixtures/` (corpus), `bench/` (drop-a-file harness), `tools/reference-render/` (headless LibreOffice → PNG). TypeScript strict; `tsdown` build; `vitest` unit; Playwright visual regression; Biome lint/format; Changesets for releases. | Small team, two publishables, one version story. |

##### 3. Architecture

```
core          OPC (central-directory zip, rels, content types) · XML · DrawingML
              (shapes, ~180 preset geometries, pictures, text bodies) · charts →
              SVG · theme resolution · font resolver · image pipeline · EMF/WMF
              (wasm)            → runs in a Worker, emits UnitModels
  ├── pptx    slide model → positioned div per shape: <svg> geometry + HTML text
  ├── xlsx    OWN styles.xml parser (the thing SheetJS CE refuses to read) →
  │           virtualized DOM grid, merges, frozen panes, cached values, formats
  └── docx    flow-rendered HTML with correct RTL, outline. Not paginated.
element       <aared-viewer>: viewport virtualization, element pooling, lazy
              media decode, shadow DOM, optional controls
aared-fonts   OFL font files + CDN default for the resolver
```

What the three formats share, and why `core` is real: the OPC package model,
**DrawingML** (identical shape/picture/text-body XML in all three), **charts**
(identical chart parts — one renderer serves everything), **theme** resolution.

##### 4. Fidelity contract

| Format | Promise | Explicitly out of scope |
|---|---|---|
| PPTX | Shapes, fills, gradients, images, text with autofit, paragraph-level RTL, tables, charts (common types), embedded fonts, z-order, rotation, notes (as text). Selectable, searchable text. | Animations, transitions, SmartArt (render its fallback image if present), 3D effects, VBA, media playback |
| XLSX | Cached values with number formats, cell styles (fill/font/border/alignment), merges, column widths/row heights, frozen panes, sheet RTL, hidden rows/cols, conditional-formatting *results* where cached, charts (common types), images. | Formula evaluation, pivot recalculation, data validation UI, macros, sparklines |
| DOCX | Flow-rendered HTML (docx-preview tier) with correct RTL, headings outline, tables, images, lists, footnotes, basic floats, tracked changes shown as final. | **Faithful pagination.** Page count and page breaks will not match Word. Headers/footers rendered once, not per page. |

##### 5. Quality strategy

**Reference truth = headless LibreOffice.** `soffice --headless --convert-to
pdf` → `pdftoppm` gives a reference PNG for any fixture (`tools/reference-render/`).
Playwright screenshots of our render vs. the reference, per fixture. Not
pixel-equal — structurally comparable — but it catches "table columns
reversed" and "Arabic bullets on the wrong side," the failures that matter.

Fixture corpus seeds from `tests/` here and must include **Arabic decks with
tables from day one** — the failure mode that disqualified `pptxviewjs`.

##### 6. Open decisions

None. Next step is scaffolding `Abdo-Eid/aared` per D21 and starting PPTX
per D2.

##### 7. Inputs / prior art

Evidence in this folder: `PPTXVIEWJS-NOTES.md` (canvas + dropped paragraph
attrs fails Arabic; clobbers `window.FileReader`), `SHEETJS-LIMITS.md` (CE
returns empty `cell.s` even with `cellStyles: true`; merges and `!cols[].wpx`
are readable — styling is paywalled).

Landscape: docx-preview, mammoth · pptx-preview, pptxviewjs · ExcelJS (parses
styles; useful as a reference implementation) · **Univer** (closest
all-in-one, but an editor suite, heavy) · LibreOffice-WASM / allotropia (real
fidelity, ~300 MB — the nuclear option, not a competitor).

<a id="research-pptxviewjs"></a>

### PPTXViewJS observations

> **Archived — provenance only.** Observations about pptxviewjs@1.1.9 and the preceding chat application, not today's aared code or newer upstream versions.

#### PptxViewJS — what to know before keeping it

Findings from evaluating `pptxviewjs@1.1.9` as the browser-side renderer for
delivered `.pptx` decks. Written down because two of these are invisible until
they bite, and one of them bites somewhere else entirely.

The library does the headline job well: it unzips and paints a deck client-side,
fast, with no server and no conversion step. Everything below is the cost.

---

##### 1. It overwrites `window.FileReader` — the worst one

On module initialisation, unconditionally and with no namespace:

```js
if (typeof window !== 'undefined') {
    window.Logger = Logger;
    window.EnhancedLogger = EnhancedLogger;
    window.ProgressTracker = ProgressTracker;
    window.FileReader = FileReader;      // ← its own utility class
    window.ErrorHandler = ErrorHandler;
    window.CoordinateTransform = CoordinateTransform;
    window.StringUtils = StringUtils;
}
```

`FileReader` here is one of the library's own helpers. It is not the DOM API and
has no `readAsArrayBuffer`. So **importing the library breaks `new FileReader()`
for the rest of the page session** — file uploads, image previews, anything
reading bytes in the browser.

How it surfaced: open a deck, then open a Word document. The document fails with

```
e3.readAsArrayBuffer is not a function
```

because JSZip inside `docx-preview` reads a `Blob` through `new FileReader()`.
Nothing in that error points at a slide deck, and the two formats have no
obvious relationship — this is expensive to debug from cold.

Note the `globalThis.JSZip` assignment a few lines away *is* guarded with
`if (!globalThis.JSZip)`. The `FileReader` one is not, so it is a slip rather
than a considered decision — which is its own signal about the code around it.

**What we do about it.** Two independent defences, because either alone leaves a
gap:

- `use-pptx-deck.ts` imports through a wrapper that snapshots those seven names
  and restores them in a `finally`. Restoring after the import is sufficient
  because the assignments happen once, at module scope. One place knows the
  library misbehaves, and one place to delete if it is ever fixed upstream.
- `use-docx-document.ts` hands `renderAsync` a `Uint8Array` rather than a
  `Blob`. JSZip only reaches for `FileReader` when given a Blob, so this removes
  the dependency on the global instead of merely repairing it. It also skips a
  copy we never needed.

The wrapper is the one that matters for the rest of the app: uploads elsewhere
would still have broken without it.

---

##### 2. Canvas output — no text, and no accessibility

Slides are painted to `<canvas>`. There is no text layer, and there cannot be
one: nothing selects, nothing is searchable with Ctrl+F, nothing copies, and a
screen reader gets an `aria-label` and nothing else.

This is structural, not a gap to fill later.

Worth being precise about the comparison, though: what a deck had *before* was
`docQa`'s page images, which are equally textless. So this loses nothing that
existed — but it does mean the deck surface stays below the accessibility floor
the rest of the viewer meets, and no amount of work on this library changes
that. Fixing it needs a renderer with a DOM.

What it *does* gain over the images is completeness: the QA renders are a
sample, capped at 12 slides (24 for long decks) and seeded at quartiles, so a
forty-slide deck showed about half of itself.

---

##### 3. Paragraph-level RTL is dropped, and cannot be patched

Character shaping and bidi *within* a run are correct. What is ignored is
direction at the paragraph level — `algn="r"` and `rtl="1"` — so Arabic lists
render left-aligned with bullets on the wrong side, and the specified font can
fall back to a serif.

Because the output is painted pixels, **there is no CSS escape hatch.** With
`docx-preview` a direction problem is recoverable in our own stylesheet scoped
to the affected sections, since that renderer produces a real DOM. Here the
options are a patch to the library, an upstream issue, or accepting it.

**Check an RTL table before committing.** Reversed column order is a worse
failure than misalignment, because it misrepresents which column is which — a
reader cannot tell they are looking at the wrong data.

This matters more for us than for most adopters: Arabic decks are a live use
case, not an edge case.

---

##### 4. Chart.js is a hard dependency, not an optional one

The docs describe `chart.js` as needed only for decks containing charts. The ESM
build imports `chart.js/auto` unconditionally at module scope, so the bundler
fails without it regardless of deck content. Install it. It is inside the
dynamic import, so it only loads with the deck panel's chunk.

---

##### 5. Rendering is not concurrency-safe

One `PPTXViewer` instance holds its own current-slide state, and `renderSlide`
moves that state before it paints. Two overlapping calls race: the second
reassigns the slide while the first is still drawing.

This is easy to hit, because the stage and every filmstrip thumbnail share one
viewer. The symptom is a strip full of the same slide repeated, or slides in the
wrong frames — which reads as a bug in our own code.

Handled by serialising through a promise chain per viewer (`renderPptxSlide`).
Cheap, and it makes the order deterministic.

---

##### Verdict as it stands

Three marks against, and they are different in kind:

| | |
|---|---|
| No text layer | Structural. No worse than the page images it replaces. |
| Paragraph RTL dropped | **The deciding one.** Unfixable on our side, and Arabic decks are a real use case. |
| Trashes `window.FileReader` | Contained, but tells you how carefully the rest is written. |

The first is a known trade. The third we have defended against. The second is
the one to test before committing — render a real generated Arabic deck,
including a table, and look hard at alignment and column order.

If it fails there, dropping the library removes all three problems at once, and
the fallback is the delivery-time slide images we already produce.

<a id="research-sheetjs"></a>

### SheetJS observations

> **Archived — provenance only.** Observations about SheetJS Community Edition 0.20.3 and the preceding chat application, not today's aared renderer or current vendor offerings.

#### XLSX rendering — what we can and cannot show

Why a delivered workbook looks plainer in the viewer than it does in Excel, what
is reachable with the library we already have, and what SheetJS Pro would add.

Measured against **SheetJS Community Edition 0.20.3**, the version the app
resolves. (A stale `0.18.5` also sits in the Bun cache — check
`XLSX.version` at runtime rather than trusting a `node_modules` path.)

---

##### What the viewer does today

`use-workbook.ts` reads the file and calls:

```ts
XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false })
```

That returns **values and nothing else** — a rectangular array of strings. Every
merge, fill, border, font, colour and chart in the file is discarded before
`buildSheetGrid` ever sees it. The grid then adds back what it can infer from
the values alone: column widths from content length, and right-aligned monospace
for columns whose values are all numeric.

`raw: false` is doing real work and is worth keeping: it takes each cell's
*formatted* text, so a date reads as the date the sheet displays and a currency
cell as money, rather than as an underlying serial number.

So the viewer is faithful about **data** and silent about **presentation**.

---

##### Measured: what Community Edition actually returns

Written and read back through CE 0.20.3, with `cellStyles: true`:

| Feature | Key | Result |
|---|---|---|
| Merged ranges | `sheet["!merges"]` | ✅ `[{s:{r,c}, e:{r,c}}]` — intact |
| Column widths | `sheet["!cols"]` | ✅ intact, including `wpx` (pixels) |
| Row heights | `sheet["!rows"]` | ✅ key exists; populated when the file sets them |
| Number formats | `cell.z`, `raw:false` | ✅ already used |
| Cell styling | `cell.s` | ❌ `{ patternType: "none" }` — no fill, font, or border |

The last row is the decisive one. `cellStyles: true` does not fail, it just
returns nothing usable — which is the easy trap here, because it looks like it
worked.

---

##### The gap, on a real file

A generated tracker sheet next to the viewer's rendering of it:

| In Excel | In the viewer |
|---|---|
| Title merged across A–H, bold, large | Plain text in A1 |
| KPI band — dark fill, white type, four merged cells | Four separate cells, no band |
| Big figures (12, 74.0, 2, 1) in 20pt | Same figures at body size |
| Table header row with a dark teal fill | Ordinary row |
| Green / red conditional values | All one colour |
| Embedded bar chart | Absent entirely |

Nothing here is wrong — every value is correct and correctly formatted. What is
missing is the visual structure the author used to say *which* numbers matter.

---

##### Reachable now, without changing library

Both are ordinary reads with the current CE build:

- **Merged cells.** `!merges` gives exact ranges, which map to `colSpan` /
  `rowSpan` on the rendered table. This recovers most of the document's shape:
  the spanning title, the KPI band, section headers. Highest value per unit of
  work by a distance.
- **Author's column widths and row heights.** `!cols[].wpx` is already in
  pixels. The grid currently guesses widths from content length, so a column the
  author deliberately made narrow or wide is ignored.

Neither needs a licence, a new dependency, or a change to how the file is read —
only that `useWorkbook` stops throwing the sheet object away after
`sheet_to_json`.

---

##### Not reachable with Community Edition

- Fills, fonts, weights, borders, text colour — `cell.s` is empty.
- Conditional formatting.
- Data validation.
- Embedded charts. These are separate OOXML chart parts; CE does not read them,
  and drawing one means interpreting the chart XML and rendering it.
- Images and shapes.

---

##### What SheetJS Pro adds

From the vendor's own module list ([sheetjs.com/pro](https://sheetjs.com/pro/)):

| Module | Adds |
|---|---|
| **Pro Basic** | Cell, column, table and worksheet styling; text styling and rich text; styled HTML table read/write; data validations and conditional formatting; locale support |
| **Pro Image** | Reading and writing images and shapes |
| **Pro Chart** | Reading and writing charts and chart sheets |
| **Pro Edit** | Updating data in a file with perfect fidelity — everything it cannot model is preserved rather than dropped |

Separate components cover PivotTable / PivotChart generation and processing, and
password protection with file encryption.

For our gap specifically: **Pro Basic** covers the fills, fonts, borders and
conditional formatting, and **Pro Chart** covers the missing bar chart.

Pricing is not published — it is a commercial licence arranged with SheetJS
directly, so treat cost as unknown until someone asks them.

---

##### Recommendation

Two tiers, and they are very different in cost.

**Do the structural half now.** Merges plus the author's column widths are a
contained change to `use-workbook.ts` and `document-sheet-source.ts`, need no
licence, and recover most of what makes a sheet readable at a glance. A merged,
correctly-proportioned grid reads as the author's document even without colour.

**Treat full visual fidelity as a separate decision.** It needs a paid licence
(Pro Basic, plus Pro Chart if charts matter), and the honest question to answer
first is what a reader actually needs from a spreadsheet *preview*. If the
answer is "recognise it and check the numbers", the structural half is enough
and the licence is not worth it. If delivered workbooks are presentation
artefacts that people forward, that changes the answer.

Worth keeping in proportion: XLSX already reads **correctly**. This is about how
much of the author's formatting we reproduce, not about whether the data is
right.

## Baseline migration record

> **Frozen provenance.** The following administrative record and complete
> ID mapping describe the original October 14 migration. Statements about
> verbatim bodies and an empty Unreleased section refer to that snapshot;
> later editorial changes do not rewrite this evidence. Notes remains the
> experience reference, not the audit log of this migration.

### Provenance and preservation

The original log ran from the first design session (September 3, 2026) through
October 2026. The raw first-session D1–D21 table, including alternatives, now
lives in [the archived design appendix](#research-design),
beside [PPTXViewJS observations](#research-pptxviewjs)
and [SheetJS observations](#research-sheetjs)
in the same archived markdown file.
They were recovered from the original root commit; no absent historical hash
is needed to read them in a fresh clone.

Before this migration the log had already been compacted in October: decisions,
reasons and code-cited rules were kept, measurement tables, test lists and
step-by-step stories were cut, and superseded entries became stubs. This
migration does **not** repeat that compaction or claim to reconstruct material
previously cut. Verified external bundles preserve the full original history
and the exact pre-migration document. The frozen
[pre-baseline changelog](#pre-baseline-development-changelog) provides a concise
chronological view, not a replacement for the detailed experience below.

Probe directories named in entries are gitignored, local scratch. The entry
carries the finding; the topic name identifies the interaction, not a dependency
on a directory that exists in a fresh clone. Previously completed probes may
already have been removed after codification; source comments, real fixtures
and tests are the durable proof.

### Development baseline and record ownership

The maintainer chose the October 14, 2026 state as the package's development
starting point, without pretending the work began that day. The original root
author date is September 11, 2026, 18:00:10 +03:00; its distinct committer date
is September 12, 2026, 00:06:51 +03:00. Both are retained on the squashed root.
The design session predates Git and remains September 3, not September 11.

Five consecutive history groups end at the real September 18, September 25,
October 1 and October 6 snapshots, then at the completed October 14 baseline.
Each commit carries the original group-ending tree; the last includes this
documentation migration. This is a chronological squash, not file regrouping
or a claim that intermediate commits never existed. Full original history and
the ordinary migration commit are backed up separately before updating main.
Only main changes; the integration branch, name-reservation tag, deferred
embedded-font stash and checkpoint refs are preserved. Consequently an all-ref
log may still show old commits. Ignored captures/probes are not in a Git bundle
and are neither cleaned nor included in the baseline by this operation.

The former Decision Log mixed reusable rationale with progress reporting.
Notes keeps that experience, including superseded choices; the archived
changelog distills development chronology and root CHANGELOG starts tracking
subsequent changes under Unreleased. No release/version bump, public API,
renderer, parser, fixture, font calibration or performance measurement changes
as part of this migration; source/test references change only in comments.
ARCHITECTURE retains its one inventory and release gates. Documentation audit
compares every inherited title/body, checks the complete ID index and validates
local destinations; tree equality and bundle restoration are the history proof.
Per AGENTS' docs-only rule, this work does not rerun runtime/build verification.

### Decision ID and migration index

Every row maps one original Decision Log entry to its destination in Notes.
The index is frozen at the migration; later experience notes need not be added
to this historical audit record.

| Original ID | Entry and destination | Topic |
|---|---|---|
| §1 | [Why this project exists at all](../Notes.md#note-1) | [Design, scope and native evidence](../Notes.md#design-scope-and-native-evidence) |
| §2 | [PDF was scoped out, deliberately](../Notes.md#note-2) | [Design, scope and native evidence](../Notes.md#design-scope-and-native-evidence) |
| §3 | [Text is rendered by the browser, never rasterized — the load-bearing decision](../Notes.md#note-3) | [Design, scope and native evidence](../Notes.md#design-scope-and-native-evidence) |
| §4 | [View-only, permanently — not a phased-in decision](../Notes.md#note-4) | [Design, scope and native evidence](../Notes.md#design-scope-and-native-evidence) |
| §5 | [Build order: PPTX → XLSX → DOCX, not alphabetical or "easiest first"](../Notes.md#note-5) | [Design, scope and native evidence](../Notes.md#design-scope-and-native-evidence) |
| §6 | [Packaging: one package with lazy subpath exports, not `@scope/*` packages](../Notes.md#note-6) | [Design, scope and native evidence](../Notes.md#design-scope-and-native-evidence) |
| §7 | [DOCX pagination is explicitly not promised](../Notes.md#note-7) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §8 | [PowerPoint is the text reference; LibreOffice is a companion artifact](../Notes.md#note-8) | [Design, scope and native evidence](../Notes.md#design-scope-and-native-evidence) |
| §9 | [PowerPoint probes settle renderer disagreements](../Notes.md#note-9) | [Design, scope and native evidence](../Notes.md#design-scope-and-native-evidence) |
| §10 | [DOCX layout mode: structural pagination without Word parity](../Notes.md#note-10) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §11 | [DOCX layout fidelity corrections from bench comparison](../Notes.md#note-11) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §12 | [`aared/preview`: superseding the deferred facade decisions](../Notes.md#note-12) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §13 | [DOCX `render()` joins the shared async object envelope](../Notes.md#note-13) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §14 | [`aared/react`: the mount lifecycle as hooks, not docs prose](../Notes.md#note-14) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §15 | [`capabilities` cut before first publish; `search()` and the headless boundary kept](../Notes.md#note-15) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §16 | [Search cost, DOCX search coverage, `info.sizes`, and the Worker bundle](../Notes.md#note-16) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §17 | [PPTX master/layout surface: flatten the small, share the large](../Notes.md#note-17) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §18 | [Core performance pass: streaming parse, cached relationships, split worker](../Notes.md#note-18) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §19 | [Split worker vs. consumer bundlers: Vite verbatim copy and the abort taxonomy](../Notes.md#note-19) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §20 | [Worker back to one self-contained file: the split is reverted](../Notes.md#note-20) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §21 | [Measuring text for DOCX pagination: measure in the DOM; Pretext considered, not adopted](../Notes.md#note-21) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §22 | [DOCX layout follows Word's own numbers; supersedes parts of §10 and §11](../Notes.md#note-22) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §23 | [Shared-code pass: one home per rule, and the decisions it forced](../Notes.md#note-23) | [Shared rendering, paint and host isolation](../Notes.md#shared-rendering-paint-and-host-isolation) |
| §24 | [Dark mode, looked at: automatic text follows its fill; authored text on the dark surface is an open gap](../Notes.md#note-24) | [Design, scope and native evidence](../Notes.md#design-scope-and-native-evidence) |
| §25 | [Adaptive dark retired: every format is source-faithful](../Notes.md#note-25) | [Design, scope and native evidence](../Notes.md#design-scope-and-native-evidence) |
| §26 | [Deferred geometry: the worker describes, the main thread evaluates](../Notes.md#note-26) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §27 | [Shape style references with `idx="0"` mean "no reference"](../Notes.md#note-27) | [Shared rendering, paint and host isolation](../Notes.md#shared-rendering-paint-and-host-isolation) |
| §28 | [DrawingML fills: one model, one painter, and what PowerPoint actually draws](../Notes.md#note-28) | [Shared rendering, paint and host isolation](../Notes.md#shared-rendering-paint-and-host-isolation) |
| §29 | [Real zoom fit for PPTX and DOCX; the DOCX heading/page facade gap closed](../Notes.md#note-29) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §30 | [PPTX tables and text lines, matched to PowerPoint](../Notes.md#note-30) | [PPTX fidelity](../Notes.md#pptx-fidelity) |
| §31 | [One API grammar across formats and layers](../Notes.md#note-31) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §32 | [Measured numbers on the home page, and how they're measured](../Notes.md#note-32) | [Performance, benchmarks and navigation](../Notes.md#performance-benchmarks-and-navigation) |
| §33 | [The benchmark rulers were wrong; corrected before any product change](../Notes.md#note-33) | [Performance, benchmarks and navigation](../Notes.md#performance-benchmarks-and-navigation) |
| §34 | [Shipped JS is Latin-1](../Notes.md#note-34) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §35 | [The DOCX Worker keeps nothing the page holds](../Notes.md#note-35) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §36 | [Renderer defaults are classes in one adopted sheet; a lighter grid](../Notes.md#note-36) | [Shared rendering, paint and host isolation](../Notes.md#shared-rendering-paint-and-host-isolation) |
| §37 | [DOCX layout pages paint lazily](../Notes.md#note-37) | [Performance, benchmarks and navigation](../Notes.md#performance-benchmarks-and-navigation) |
| §38 | [The Worker is reused between opens](../Notes.md#note-38) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §39 | [Bounded Worker caches; minified dist](../Notes.md#note-39) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §40 | [No layout jumps when a unit is replaced](../Notes.md#note-40) | [Performance, benchmarks and navigation](../Notes.md#performance-benchmarks-and-navigation) |
| §42 | [npm name: `@aared/viewer`, one package, reserved with a placeholder](../Notes.md#note-42) | [Design, scope and native evidence](../Notes.md#design-scope-and-native-evidence) |
| §43 | [DOCX right-to-left tables: `w:bidiVisual`, logical borders, physical margins](../Notes.md#note-43) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §44 | [Fixes found by the showcase documents (Sep 2026)](../Notes.md#note-44) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §45 | [Font fidelity: a delivered substitute is actually used, waited for, and reported (Sep 2026)](../Notes.md#note-45) | [Fonts, script and line measurement](../Notes.md#fonts-script-and-line-measurement) |
| §46 | [Arabic calibration measured the way it paints: through the alias, unrounded (Sep 2026)](../Notes.md#note-46) | [Fonts, script and line measurement](../Notes.md#fonts-script-and-line-measurement) |
| §47 | [Paragraph spacing in lines: what a "line" is in Word and PowerPoint (Sep 2026)](../Notes.md#note-47) | [Fonts, script and line measurement](../Notes.md#fonts-script-and-line-measurement) |
| §48 | [PPTX autofit is what PowerPoint stored; bullets hang; a subtitle takes the master body style (Sep 2026)](../Notes.md#note-48) | [PPTX fidelity](../Notes.md#pptx-fidelity) |
| §49 | [Aptos: calibrated look-alikes measured against the real cloud fonts (Sep 2026)](../Notes.md#note-49) | [Fonts, script and line measurement](../Notes.md#fonts-script-and-line-measurement) |
| §50 | [Delivered fonts load when painted, not up front (Sep 2026)](../Notes.md#note-50) | [Fonts, script and line measurement](../Notes.md#fonts-script-and-line-measurement) |
| §51 | [The no-Office-fonts review: substitutes painted where they matter (Sep 2026)](../Notes.md#note-51) | [Fonts, script and line measurement](../Notes.md#fonts-script-and-line-measurement) |
| §52 | [Wingdings and Symbol bullets paint their Unicode shape (Sep 2026)](../Notes.md#note-52) | [Fonts, script and line measurement](../Notes.md#fonts-script-and-line-measurement) |
| §53 | [`aared-fonts` frozen and unpublished: the browser's own fonts are the default (Sep 2026)](../Notes.md#note-53) | [Fonts, script and line measurement](../Notes.md#fonts-script-and-line-measurement) |
| §54 | [PPTX fidelity pass: where PowerPoint puts text, colour maps, table styles, Arabic around punctuation (Sep 2026)](../Notes.md#note-54) | [PPTX fidelity](../Notes.md#pptx-fidelity) |
| §55 | [DOCX pictures that text wraps around, placed as Word places them (Sep 2026)](../Notes.md#note-55) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §56 | [The Worker URL resolves under webpack and Next.js (Sep 2026)](../Notes.md#note-56) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §57 | [Every document failure is typed; a lost session is `closed`, not `aborted` (Sep 2026)](../Notes.md#note-57) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §58 | [Error codes say whose fault it is (Sep 2026)](../Notes.md#note-58) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §59 | [Small API faults from a developer-ergonomics review (Oct 2026)](../Notes.md#note-59) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §60 | [Benchmarks measure aared; competitors are measured once (Oct 2026)](../Notes.md#note-60) | [Performance, benchmarks and navigation](../Notes.md#performance-benchmarks-and-navigation) |
| §61 | [The hook's result narrows by `status`; its zoom and index are live (Oct 2026)](../Notes.md#note-61) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §62 | [Bytes in as a typed array; every event typed (Oct 2026)](../Notes.md#note-62) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §63 | [The API page says less, and its examples are compiled (Oct 2026)](../Notes.md#note-63) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §64 | [Options start it, the viewer changes it; the element runs on `preview()`'s viewer (Oct 2026)](../Notes.md#note-64) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §65 | [A stale render never paints; the hook owns teardown; the element lays out pages (Oct 2026)](../Notes.md#note-65) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §66 | [Emptying `<aared-viewer>` is not an error (Oct 2026)](../Notes.md#note-66) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §67 | [The viewer has toolbar methods and says its default zoom (Oct 2026)](../Notes.md#note-67) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §68 | [An aborted render never hands back a handle (Oct 2026)](../Notes.md#note-68) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §69 | [The viewer searches its own document; building from the contents stays low level (Oct 2026)](../Notes.md#note-69) | [Search and highlighting](../Notes.md#search-and-highlighting) |
| §70 | [A Start page for newcomers; the API page is the reference (Oct 2026)](../Notes.md#note-70) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §71 | [Charts: a real chart model, read in the Worker, laid out on the page (Oct 2026)](../Notes.md#note-71) | [Charts](../Notes.md#charts) |
| §72 | [Hostile counts are bounded where the file is read; three XLSX and PPTX rules corrected (Oct 2026)](../Notes.md#note-72) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §73 | [Charts: the reader-visible gaps closed, every type but maps and surface drawn (Oct 2026)](../Notes.md#note-73) | [Charts](../Notes.md#charts) |
| §74 | [Charts: the assumed rules checked against Office, and corrected (Oct 2026)](../Notes.md#note-74) | [Charts](../Notes.md#charts) |
| §75 | [A whole-repo static review: shared failure classes and deliberate behavior (Oct 2026)](../Notes.md#note-75) | [Quality, tooling and documentation](../Notes.md#quality-tooling-and-documentation) |
| §76 | [Four Office probes: bold by script, number-format ids, empty paragraphs, first-page headers (Oct 2026)](../Notes.md#note-76) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §77 | [List counters per list definition, and the coverage fixtures extended (Oct 2026)](../Notes.md#note-77) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §78 | [Search matches the Arabic spellings of one word (Oct 2026)](../Notes.md#note-78) | [Search and highlighting](../Notes.md#search-and-highlighting) |
| §79 | [XLSX search covers one sheet, as Excel's Find does (Oct 2026)](../Notes.md#note-79) | [Search and highlighting](../Notes.md#search-and-highlighting) |
| §80 | [Search reads words, not models (Oct 2026)](../Notes.md#note-80) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §81 | [One theme rule on every viewer layer: light by default, `system` follows the OS (Oct 2026)](../Notes.md#note-81) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §82 | [DOCX text boxes and section line grids (Oct 2026)](../Notes.md#note-82) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §83 | [DOCX header and footer search (Oct 2026)](../Notes.md#note-83) | [Search and highlighting](../Notes.md#search-and-highlighting) |
| §84 | [Viewer search owns highlighting across formats (Oct 2026)](../Notes.md#note-84) | [Search and highlighting](../Notes.md#search-and-highlighting) |
| §85 | [PPTX search results count text occurrences (Oct 2026)](../Notes.md#note-85) | [Search and highlighting](../Notes.md#search-and-highlighting) |
| §86 | [Object-form preview and reactive React appearance (Oct 2026)](../Notes.md#note-86) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §87 | [Active find position and compact navigation (Oct 2026)](../Notes.md#note-87) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §88 | [Complete viewers from single-format entries (Oct 2026)](../Notes.md#note-88) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §89 | [Reference apps use the convenient viewer APIs (Oct 2026)](../Notes.md#note-89) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §90 | [Empty section-ending page-break paragraphs do not add pages (Oct 2026)](../Notes.md#note-90) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §91 | [DOCX pagination stays DOM-measured; fidelity work is bounded by visible value (Oct 2026)](../Notes.md#note-91) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §92 | [Numbered DOCX headings and refreshed table-of-contents results (Oct 2026)](../Notes.md#note-92) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §93 | [Tools consolidation: one Office export driver, one rasterizer, a standard probe shape (Oct 2026)](../Notes.md#note-93) | [Quality, tooling and documentation](../Notes.md#quality-tooling-and-documentation) |
| §94 | [A DOCX correctness pass: two limits ruled, autofit and fills probed, numbered headings with a live TOC, and a zoomed container (Oct 2026)](../Notes.md#note-94) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §95 | [DOCX text boxes: resize, clipping, geometry, rotation, fills (Oct 2026)](../Notes.md#note-95) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §96 | [Probe cleanup and the label-strip record (Oct 2026)](../Notes.md#note-96) | [Quality, tooling and documentation](../Notes.md#quality-tooling-and-documentation) |
| §97 | [A row's height is read from where Word writes it (Oct 2026)](../Notes.md#note-97) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §98 | [A preset's text rectangle, and a fallback that keeps its fill (Oct 2026)](../Notes.md#note-98) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §99 | [A table's style is drawn from the style, not from the cells (Oct 2026)](../Notes.md#note-99) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §100 | [Word theme ink includes its shade or tint (Oct 2026)](../Notes.md#note-100) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §101 | [Group fill is inherited paint; group effects are composite paint (Oct 2026)](../Notes.md#note-101) | [Shared rendering, paint and host isolation](../Notes.md#shared-rendering-paint-and-host-isolation) |
| §102 | [Word drawings use the shared effect painter, and a picture has no text to justify (Oct 2026)](../Notes.md#note-102) | [Shared rendering, paint and host isolation](../Notes.md#shared-rendering-paint-and-host-isolation) |
| §103 | [PPTX visual review accepted (October 10, 2026)](../Notes.md#note-103) | [PPTX fidelity](../Notes.md#pptx-fidelity) |
| §104 | [Table-level borders and padding inherit through the full style chain (October 10, 2026)](../Notes.md#note-104) | [DOCX layout, tables and drawings](../Notes.md#docx-layout-tables-and-drawings) |
| §105 | [XLSX conditional formats: top/bottom and above-average evaluated; Excel's behaviour measured first (October 10, 2026)](../Notes.md#note-105) | [XLSX layout and conditional formats](../Notes.md#xlsx-layout-and-conditional-formats) |
| §106 | [XLSX data bars paint Excel's x14 gradient and default bar lengths (October 10, 2026)](../Notes.md#note-106) | [XLSX layout and conditional formats](../Notes.md#xlsx-layout-and-conditional-formats) |
| §107 | [XLSX conditional formats: icons, date periods and standard deviations painted; native behaviour measured first (October 10, 2026)](../Notes.md#note-107) | [XLSX layout and conditional formats](../Notes.md#xlsx-layout-and-conditional-formats) |
| §108 | [SVG-first pictures, shared render settlement, lazy-painter boundary, XLSX picture effects (October 10, 2026)](../Notes.md#note-108) | [Shared rendering, paint and host isolation](../Notes.md#shared-rendering-paint-and-host-isolation) |
| §109 | [Three-colour scales have Excel-authored, paired A4 proof (October 10, 2026)](../Notes.md#note-109) | [XLSX layout and conditional formats](../Notes.md#xlsx-layout-and-conditional-formats) |
| §110 | [Bounded conditional-formula evaluation over cached values (October 11, 2026)](../Notes.md#note-110) | [XLSX layout and conditional formats](../Notes.md#xlsx-layout-and-conditional-formats) |
| §111 | [Closing the conditional-formula gap: dates, text search, conditional aggregates, x14 and cross-sheet (October 11, 2026)](../Notes.md#note-111) | [XLSX layout and conditional formats](../Notes.md#xlsx-layout-and-conditional-formats) |
| §112 | [One-pass preview opening and real lazy viewer/effect boundaries (October 11, 2026)](../Notes.md#note-112) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §113 | [Corrections to earlier status statements (Oct 11, 2026)](../Notes.md#note-113) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §114 | [Worker parsers move from dynamic to static imports (October 11, 2026)](../Notes.md#note-114) | [Worker, packaging and memory](../Notes.md#worker-packaging-and-memory) |
| §115 | [Spreadsheet viewer chrome: formula bar, range select and inline aggregates (October 11, 2026)](../Notes.md#note-115) | [XLSX layout and conditional formats](../Notes.md#xlsx-layout-and-conditional-formats) |
| §116 | [XLSX: two Excel differences kept by choice, missing fonts measured, centred overflow fixed, the Summary redesigned (October 11, 2026)](../Notes.md#note-116) | [XLSX layout and conditional formats](../Notes.md#xlsx-layout-and-conditional-formats) |
| §117 | [XLSX fit: the file's sizes, the browser's layout, fonts as residue (October 12, 2026)](../Notes.md#note-117) | [XLSX layout and conditional formats](../Notes.md#xlsx-layout-and-conditional-formats) |
| §118 | [XLSX selection past the data: a click, a drag, the arrow keys and Excel's plus pointer (October 12, 2026)](../Notes.md#note-118) | [XLSX layout and conditional formats](../Notes.md#xlsx-layout-and-conditional-formats) |
| §119 | [Viewers lend their document, take any search result, and merge appearance updates (October 13, 2026)](../Notes.md#note-119) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §120 | [Failures report through the action's owner, not two channels (October 13, 2026)](../Notes.md#note-120) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §121 | [Explicit Worker inputs own special-deployment resource cleanup (October 13, 2026)](../Notes.md#note-121) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §122 | [Viewers lend content capabilities, not their document lifecycle (October 13, 2026)](../Notes.md#note-122) | [API, viewer lifecycle and errors](../Notes.md#api-viewer-lifecycle-and-errors) |
| §123 | [One developer-tools entry point, shared review backends and checked tooling (October 13, 2026)](../Notes.md#note-123) | [Quality, tooling and documentation](../Notes.md#quality-tooling-and-documentation) |
| §124 | [Test cleanup preserves distinct routes and named fixture behavior (October 13, 2026)](../Notes.md#note-124) | [Quality, tooling and documentation](../Notes.md#quality-tooling-and-documentation) |
| §125 | [Tools have task-owned source folders and one generated-output home (October 13, 2026)](../Notes.md#note-125) | [Quality, tooling and documentation](../Notes.md#quality-tooling-and-documentation) |
| §126 | [IFS range shapes and the deferred coverage proofs (October 14, 2026)](../Notes.md#note-126) | [Quality, tooling and documentation](../Notes.md#quality-tooling-and-documentation) |
| §127 | [One visual harness: a shared page runner and per-run browser builds (October 14, 2026)](../Notes.md#note-127) | [Quality, tooling and documentation](../Notes.md#quality-tooling-and-documentation) |
| §128 | [Rendered content owns inherited text traits and paint-resource names (October 14, 2026)](../Notes.md#note-128) | [Shared rendering, paint and host isolation](../Notes.md#shared-rendering-paint-and-host-isolation) |
| §129 | [One gap inventory: the roadmap retires into ARCHITECTURE (October 14, 2026)](../Notes.md#note-129) | [Quality, tooling and documentation](../Notes.md#quality-tooling-and-documentation) |
