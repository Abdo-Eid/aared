# Architecture

> The current engine, its rendering rules, and the gaps that remain. This is
> the canonical present-state document, including release readiness.
> [Notes.md](Notes.md) owns reasoning and engineering experience;
> [CHANGELOG](../CHANGELOG.md) owns changes after the development baseline;
> [the API reference](../site/src/content/docs/api.mdx) owns public signatures;
> [BENCHMARKS.md](BENCHMARKS.md) owns measured performance;
> [COMPETITORS.md](COMPETITORS.md) owns the external landscape.

Read this for the current data flow, ownership, rendering rules and dispositions.
Linked Notes entries explain the experiments and choices behind them; they are
not alternate specifications. A historical test run or measurement does not
become a current guarantee merely because its note is cited here.

## Purpose and boundaries

`aared` previews real PPTX, XLSX and DOCX files client-side, with particular
attention to Arabic and RTL. All three formats have a Worker parser, a
DOM/SVG renderer and an end-to-end viewer. Fidelity gaps remain open. Public
semantic data is supported; every format's rendering model is internal.

The formats share Office Open XML packaging but not their layout problem:

| Format | Layout model |
|---|---|
| PPTX | A fixed scene of positioned shapes, with flow and stored autofit inside text boxes |
| XLSX | A grid with authored tracks, merges, frozen panes and viewport-windowed rows |
| DOCX | Flowing paragraphs, tables and drawings; optional pagination from section geometry and browser measurements |

Four boundaries govern every layer:

1. **Text is painted by the browser as DOM or SVG, never rasterized by aared.**
   Browser shaping, bidi and fallback remain available. Canvas measurement is
   permitted; canvas text painting and a separate WASM text/layout engine are not.
2. **ZIP/XML parsing runs in a Worker.** Main-thread renderers consume resolved
   models, not XML. `core` has no DOM dependency and is testable in plain Bun.
3. **View-only is permanent.** There is no editing, write path, round-tripping,
   workbook recalculation or pivot recalculation. Conditional formatting reads
   cached values; it does not calculate the workbook.
4. **PDF is outside scope.** Consumers pair aared with `pdf.js`. Browser print
   and browser-free thumbnail export are undecided proposals, not delivered APIs.

No WASM ships in the engine. A future EMF/WMF image decoder could be a leaf
dependency, never a text renderer. The project's place between single-format
DOM libraries and heavier suite engines is documented in COMPETITORS.

## Package layout

One zero-dependency package exposes eight subpaths. `aared-fonts` is a separate,
frozen, unpublished workspace package; no font files ship with either package.

| Entry | Responsibility |
|---|---|
| `@aared/viewer/core` | Shared public types, errors, warnings and appearance; internal OPC/XML/theme/RPC machinery lives beside it |
| `@aared/viewer/worker` | One self-contained module Worker containing all three parsers |
| `@aared/viewer/pptx` | Open presentations, read slides, render DOM/SVG scenes, and provide fixed-format preview |
| `@aared/viewer/xlsx` | Open workbooks, read sheet ranges, render a virtualized grid, and provide fixed-format preview |
| `@aared/viewer/docx` | Open documents, expose section metadata/headings, render flow/layout pages, and provide fixed-format preview |
| `@aared/viewer/preview` | Detect the format during open and lazily load its viewer route |
| `@aared/viewer/react` | `usePreview`, `usePptx`, `useXlsx`, `useDocx`; lifecycle owned by the hook |
| `@aared/viewer/element` | Register `<aared-viewer>`; shadow DOM, attributes, optional controls and element-owned lifecycle |

```text
file bytes --> Worker: OPC + XML + format/theme/style resolution
                        |
                        v
                 resolved UnitModel
                        |
                        v
                main thread: DOM/SVG
                        |
              preview / hooks / element
```

The **low-level** route is `open()` plus `render()`. It separates document and
render lifetimes. It is not browser-free rendering: only parsing and pure
layout helpers are DOM-free. The **viewer** route owns opening, navigation and
teardown. `preview` names the operation, `element` its HTML delivery mechanism;
they are not separate rendering engines.

Use **low-level**, not **headless**, for this consumer-facing layer. Browser-free
means no browser at all; it does not describe the format renderers.

### Shared modules

Paths below are relative to `packages/aared/src/`. A rule needed by two
formats lives once; format differences are parameters, not duplicate helpers.
This map identifies responsibility, not every exported symbol.

| DOM-free module | Owns |
|---|---|
| `core/zip.ts`, `core/opc.ts` | Central-directory ZIP reads, on-demand inflation, content types, relationships, format detection and media types |
| `core/xml.ts` | Headless XML parsing, streaming children, numeric attributes, `ownString`/`ownTree` for retained data |
| `core/units.ts`, `core/color.ts` | px/pt/inch/EMU/twip conversions; colour arithmetic and formatting |
| `core/theme.ts`, `core/typography.ts` | Theme slots, colour maps, Arabic detection, generic stacks, font aliases and line-height floor |
| `core/chars.ts`, `core/symbol-fonts.ts` | Runtime non-Latin-1 constants and supported symbol-font mappings |
| `core/drawingml-shape.ts`, `core/drawingml-engine.ts`, `core/drawingml-resolve.ts`, `core/presets.ts` | Fast-path geometry, deferred geometry evaluation and the generated PPTX preset catalogue |
| `core/fill.ts`, `core/drawingml-effects.ts`, `core/image-source.ts` | Resolved fills, drawing effect lists and SVG-first embedded picture sources |
| `core/table-style.ts`, `core/table-span.ts` | Conditional table-style regions, legacy look bitmasks and bounded spans |
| `core/chart.ts`, `core/chart/read*.ts`, `core/chart/spec.ts`, `core/chart/text.ts` | Classic/Office 2016 chart parts to a bounded resolved spec; chart words and accessible descriptions |
| `core/chart/draw.ts`, `engine.ts`, `scale.ts`, `metrics.ts`, `scene.ts`, `trend.ts`, `bounds.ts` | Chart layout from spec/frame to path/text primitives; DOM-free but loaded on the page, not by the Worker |
| `core/xlsx/limits.ts`, `range.ts`, `numberfmt.ts`, `formula.ts` | Worksheet bounds/A1 scanning, number formats and bounded conditional-formula evaluation |
| `core/errors.ts`, `core/search.ts`, `core/collect.ts` | Typed errors/abort helpers, search folding/title lookup and resolved-model font-use collection |
| `core/rpc.ts`, `core/worker.ts`, `core/session.ts`, `core/worker-init.ts`, `core/unit-cache.ts` | Protocol/dispatch, main-thread session bridge, source normalization, Worker configuration/pooling and bounded unit retention |
| `core/index.ts`, `core/model.ts` | Public shared contracts and internal structured-cloneable rendering shapes, respectively |

`core/` means DOM-free, **not** that every module belongs in the Worker bundle.
Chart layout stays behind `chart-svg.ts`'s dynamic import. Its dependency closure
must not import page-entry dependencies that create a top-level-await loading
cycle; values are resolved upstream or passed into layout (§73).

| Main-thread module | Owns |
|---|---|
| `appearance.ts` | Appearance defaults/merge, theme resolution/system listener, automatic ink and chrome paint tokens |
| `class-sheet.ts`, `render-defaults.ts` | Interned CSS defaults, root-owned adopted sheets, content text baseline and repeated paint-layer classes |
| `resource-id.ts`, `svg-element.ts` | Unique paint-resource names and SVG construction |
| `text-style.ts`, `font-fidelity.ts` | Run/paragraph CSS, exact line pitch, bounded font loading and fidelity checks |
| `media-url.ts` | One object URL per media part per render, owned cleanup and source fallback |
| `table-dom.ts` | PPTX/DOCX table shells and cell grid |
| `fill-paint.ts`, `pattern-presets.ts`, `fill-loader.ts` | Shared SVG fill painter and cached lazy loading for charts/DOCX boxes |
| `effect-paint.ts`, `effect-loader.ts`, `transform-paint.ts` | DOM/SVG effects, lazy XLSX/DOCX loading, positioned rotation/flips |
| `chart-svg.ts` | Lazy chart layout, browser font calibration, SVG paint, accessible naming and settlement |
| `render-handle.ts`, `background-error.ts` | Common zoom/fit/events/teardown/settlement; built-in-action error reporting |
| `format-route.ts`, `embed-viewer.ts`, `auto-viewer.ts` | Shared controller, format adoption/content/render adapters and generic format routing |
| `content-view.ts`, `search-highlight.ts` | Restricted borrowed-content associations and lifetime guards; live DOM/SVG highlights/navigation |
| `docx/outline.ts`, `docx/stories.ts` | Body heading index and story traversal |
| `<format>/index.ts`, `<format>/renderer.ts`, `<format>/data.ts` | Public facade/interfaces, private handle/model painting, and isolated semantic projections |
| `diagnostics.ts` | Development diagnostic switch; production removes diagnostic chrome |

Keep genuine dialect differences: DrawingML line multiples use font size,
Word single-line pitch is about 1.15em; Word superscript uses `verticalAlign`,
DrawingML `baselinePct`; table edge/padding rules differ; SpreadsheetML tint
moves HLS lightness while DrawingML uses its own colour-modifier pipeline.
The shared code takes those differences explicitly (§23). Keep renderer extras
local too: PPTX outline/gradient/shadow text treatment is not DOCX shading,
super/subscript, dominant-family and paragraph-border treatment.

## Performance model

### Opening and resolved data

`sourceBytes` copies caller-owned buffers/views once and owns bytes read from
Blob/URL inputs. A preview transfers those bytes through `WorkerSession.openBytes`;
the `open` response includes the detected format. The chosen route adopts that
same session, with no preliminary `detect` request or second parse. Standalone
`detect(source)` is a separate operation; following it with preview is another
open. An asserted format is checked before importing its route (§112).

PPTX/XLSX opening reads a lightweight index before individual units. PPTX
installs the page-side preset table before any slide/layout read; XLSX constructs
its sheet collection from the index. DOCX's `docx/session-document.ts` reads
unit 0 and resolves fonts before returning its document: sections are already
resolved, with no `section.read()` operation.

Internal models are structured-cloneable data, read-only by contract. Shared
XLSX styles and hyperlinks stay shared after structured clone; clone before
modifying data. Public reads project isolated semantic data: XLSX cached cells,
PPTX paragraphs/optional notes and DOCX section metadata in inches. XLSX data
reads skip page-side font/geometry resolution; PPTX text reads never load slide
models. Private model reads feed painters, sheet switching, outlines and search.
Public rendering accepts registered handles/documents or borrowed content only;
no detached-model or public snapshot route exists. See the
[data contract](#resolved-model-contract) and Notes [§135](Notes.md#note-135).

The Worker inflates parts on demand rather than retaining a decompressed archive.
`sheetData`, shared strings and `w:body` stream rather than materializing entire
XML trees. Retained strings/trees use `ownString`/`ownTree`: a sliced string can
otherwise retain a whole XML part. Worksheet cells occupy a sorted per-row grid,
not a per-cell map expanded across empty ranges (§35).

Relationships, layout/master chains and header/footer parts reuse package-level
reads. A DOCX session retains the package/media set and a small index: styles,
numbering, theme, relationships/hyperlinks, settings, notes-part names and title.
Its body/notes trees and sections are not retained after parsing. PPTX/XLSX do
retain bounded parsed-unit caches: **four slides, two sheets**, most recently used.
XLSX's first range read still parses a full sheet; its limits are recorded below.

### Worker execution and deployment

`core/worker.ts` uses **static symbol-level parser imports**. The Worker build is
one non-splitting, self-contained module; dynamic imports would not defer its
bytes and would weaken tree-shaking (§114). Main-thread format routes, chart
layout and optional fill/effect painters are the useful lazy boundaries.

RPC uses typed `postMessage` requests for open/detect, unit/text/layout/media
reads, cancel and close. Media buffers transfer without a second main-thread
copy. Abort signals send cancellation; failed imports/adoption close any opened
session. Browser imports themselves cannot be cancelled, so callers observe
late rejections without allowing a stale open to paint.

`WorkerInit` accepts exactly one construction mode (§121):

| Input | Behaviour |
|---|---|
| `{ url }` | Direct module Worker; URL may be a string or `URL` instance |
| `{ url, crossOrigin: "blob" }` | Explicit Blob bootstrap importing a CDN module |
| `{ source }` | Blob Worker containing a self-contained module source |
| `{ create }` | Caller factory supplying a fresh native Worker; unpooled |

Page-relative URLs resolve at open time. Generated Blob URLs belong to aared
and are revoked on ready, constructor/boot failure or termination; caller URLs
remain caller-owned. URL/source inputs pool by value and boot mode, not option
object identity. An unready Worker never enters the pool. A clean closed session's
Worker can serve the next open; it terminates after **15 seconds idle** (§38).
Tests replacing `globalThis.Worker` drain the pool in `finally`.

CDN bootstrap requires CORS and appropriate `worker-src`/`script-src` permissions,
including `blob:` and the CDN origin where applicable. Chrome also checks the
imported CDN module against `worker-src`; a page's `securitypolicyviolation`
listener can miss Worker-scope failures. Refused/crashed Workers reject requests
as `worker-error`, not `aborted`. There is no implicit fallback, main-thread
parsing, security-policy relaxation or automatic crash retry.

### Viewports, fitting and media

PPTX preview paints one slide; callers choose how many additional slides to render.
XLSX windows rows to the viewport. DOCX flow is not virtualized by the library;
layout pages use `content-visibility: auto`. Caller-owned slicing is the current
flow boundary, not a promise of an in-library block virtualizer.

Fit measures room with existing content temporarily set aside, then restores
scrolled ancestors. A slide occupies its full-width box before fitting.
Replacing a unit must not move the reader. Content-sized spreadsheets cap their
own scroller to a clipping ancestor's visible room, so frozen rows/headings stay
sticky and their scrollbar is not below the fold (§40, §44).

`MediaUrls` owns one object URL per media part per render, reused for each use
within that render and revoked on replacement/destroy. Media is not
viewport-bounded: scrolling does not release or decode parts. Reusing a document
does not imply sharing object URLs between its render handles.

### Input budgets

Limits apply where claims are read, before count-driven page loops (§72):

| Claim | Current bound |
|---|---|
| ZIP entry inflation | 256 MiB per entry; central-directory offsets validated |
| XML | DOCTYPE refused; depth bounded at 512; only predefined/numeric references resolved |
| XLSX grid | 1,048,576 rows, 16,384 columns; `<col>` ranges cover the grid once |
| XLSX merges | Cut to extent; 250,000 cells per merge, 1,000,000 merged cells per sheet |
| Table spans | Cut to the remaining rows/columns through `core/table-span.ts` |
| Chart points | 100,000 per series, 250,000 per chart; bounded read budget |

Hostile-input tests and a seeded corruption sweep require either a result or a
typed error. These controls are **not** a total input-memory guarantee: aggregate
ZIP/download budgets and remaining audits are [open](#resources-and-input-safety).

## Viewer behavior and ownership

### One controller, several entry points

`embed-viewer.ts` implements `openViewer({ scope, source, target, options,
onProgress, route })`. `format-route.ts` supplies the adopt/content/render
contract and event forwarding without importing a renderer at runtime.
Fixed-format `viewer.ts` adapters use that controller; `auto-viewer.ts` chooses
the generic preview/element route from the already-open session.

All embed entries defer the controller; format entries also defer their adapter.
A fixed-format consumer emits its own renderer/shared helpers, not siblings or
auto-detection code. The Worker still carries all parsers. Only DOCX local preview
accepts pagination; local preview does not accept a `type` option (§88, §112).

Low-level callers destroy each render and separately `close()` its document.
A preview owns both primary render and document; `destroy()` releases both and
clears its target. Hooks own unmount/source replacement; elements own removal,
source replacement and their appearance. Their exposed handle types omit the
members owned by the embedding layer.

### Borrowed content

`viewer.content` lends reads, metadata, media and ordinary format `render()`
inputs from the same parse (§122). It is a runtime facade, not a type-only view
of the owned document:

| Format | Borrowed content |
|---|---|
| PPTX | Slide collection with `info`/`read`, presentation metadata and media |
| XLSX | Sheet collection with `info`/range `read`, workbook metadata and media |
| DOCX | Projected section metadata, headings, document metadata and media; the whole facade is a render input |

Frozen closure-backed collections/media expose no owner, session, `close()` or
search method. A renderer-private WeakMap associates slide/sheet inputs and DOCX
content with the backing source and lifetime signal. The viewer's document/route
are ECMAScript-private. Rendering keeps automatic media, PPTX shared layout media,
XLSX tab navigation and DOCX outline navigation without another open.

Collection access, reads, media and render unwrapping check lifetime first.
After owner teardown they fail `closed`, even through a retained unit input.
Pending borrowed renders combine caller/owner signals and cancel as `aborted`.
Completed secondary handles are caller-owned: destroy them before the viewer;
the viewer does not track/destroy them. Already-returned metadata/data/bytes
remain readable. Semantic projections clone nested containers and are not
deep-frozen; changing a local copy does not edit or repaint the file. DOCX's
fresh section projections bypass the stable-handle identity cache to avoid
retaining every read. This is an ownership
boundary, not a security sandbox. There is no `preview({ document })` API.

### Navigation, appearance and cancellation

Navigation is newest-wins. Every render has its own signal; a newer move, aborted
open or destroy prevents late paint and removes only its own painted nodes.
`goTo()` returns early for the visible unit only when no move is in flight;
otherwise it renders that unit again to cancel a pending replacement. Next/previous
step from the pending move. Zoom steps are bounded to 25%–300%; reset restores
the format's default fit. XLSX cell and DOCX heading/page controls operate on the
current render through stable viewer control objects (§64–§68).

Appearance changes chrome, never recolours authored document content.
Imperative `setAppearance()` merges specified knobs, one level into `spreadsheet`.
React's appearance option is a complete state: it replaces previous options
through `REPLACE_APPEARANCE`, supplies defaults beneath them, and resets omitted
knobs. Value comparisons avoid reopening for inline object identity/order changes.
Updates during load apply when ready; imperative updates survive unrelated
rerenders until option values change. Index/zoom options are starting values.

Viewer theme defaults to light. Explicit light/dark ignores OS changes; `system`
is watched live by the shared viewer only while selected, and the watcher stops
on replacement/destroy. The element separately watches its own chrome. Low-level
render resolves `system` when painted; its caller owns subsequent updates (§81).

PPTX/DOCX restyle existing roots; XLSX repaints the cached visible grid while
preserving scroll/zoom/selection. No appearance update reparses a file (§86, §119).

### Search and highlights

Search stays on `viewer.search()`. Low-level PPTX/DOCX documents and XLSX sheet
handles offer data-only search; XLSX has no workbook-level search method.
Viewer XLSX search follows the sheet on screen, not a move in flight. A workbook
search is an explicit loop over sheets. DOCX searches source blocks and each
distinct header/footer part once; inherited painted repetitions are not extra
source results. Unpainted story variants reject navigation with `not-found`.

PPTX/XLSX search asks the Worker for **text, not models** (§80). PPTX's
`readPptxSlideText` walks own shapes/groups, table `contentCells` and chart words,
without resolving inheritance, geometry, theme, background or notes. Its text
must equal full-model paragraph text for every fixture slide. XLSX uses
`sheetText`; a search reuses a retained sheet but does not populate/refresh the
unit cache. DOCX already holds resolved text. Search keys live on the page and
are dropped with the document.

`searchKey` folds case, tashkeel/tatweel, alef variants, alef maqsura, Arabic-Indic/
Persian digits and presentation forms (NFKC); results retain original text.
Title lookup instead uses exact then case-insensitive matching, never substring
or Arabic search folding (§78).

The lazy `search-highlight.ts` controller owns CSS Custom Highlights on the
stable target. New queries win; failed/cancelled queries preserve previous
highlights. Empty queries clear them. `{ highlight: false }` returns data and
clears prior paint. Nonempty highlighted search without browser support rejects
`unsupported-feature`; data-only search remains available.

Highlights survive replaced slides/sheets, virtualized rows and DOCX split-wrap
rebuilds. Each viewer has independent ranges/styles/observers in its document or
shadow root. Displayed page-field values are highlighted. `searchNavigation`
steps through painted occurrences (including repeated DOCX story copies); active
rectangles live outside the text tree, pointer-transparent and clipped, and
follow scroll/resize. `goToSearchMatch()` uses the same activation path (§119).

The application owns search UI and results. A sheet change requires rerunning
an active query to update its result list; persistent highlight paint alone is
not enough. Clear/opt-out/destroy removes active markers and observation. The
element supplies the API, not a search panel.

### Errors

Public calls throw/reject only, including explicit `element.open()`. Applications
handle their promises. Built-in automatic loads, attribute changes, toolbar
actions and XLSX tabs catch their own failures and emit one `error` event.
`background-error.ts` suppresses cancellation, normalizes unexpected failures
to `AaredError("internal")`, and preserves typed renderer errors (§120).

React exposes one `error` for mount/built-in failures. A failed built-in action
keeps ready content intact; a successful unit change/mount clears it. Hook state
and events do not duplicate failures from application-initiated calls. Worker
failure remains request-level `worker-error`, with no second global crash event.

`isAbort` does not swallow `closed`: cancellation and use-after-close are distinct.
Broken-file checks use `invalid-document`; unexpected parser failures become
`internal`. `itemAt` distinguishes integer out-of-range (`not-found`) from
malformed input (`invalid-argument`). `messageFor()` is an optional presentation
helper, not a required handler or an error-policy switch. The API reference owns
the full error vocabulary.

## Embedding, styling and packaging

### CSS and resource ownership

`class-sheet.ts` interns constant declaration blocks and adopts one constructed
sheet per document/shadow root, reference-counted by renders. Defaults use
classes; authored geometry/formatting and measured values stay inline through
CSSOM. No `<style>` text or serialized style attribute is required for strict
CSP. Tests read computed style for defaults (§36, §44, §59).

Sheet adoption follows a detached target when it acquires a document/shadow root
and is restored if the host replaces `adoptedStyleSheets`; shared rules are
released only when their last render on that root ends.

`render-defaults.ts` resets inherited emphasis, font variant/stretch, tracking,
casing and indentation on PPTX's viewport, XLSX's whole host, DOCX's flow article
and layout page class. Pagination measurement uses the same class. It does not
replace resolved family/size/ink/line pitch or exact-pitch rules. Class specificity
is (2,1,0), using `.aared-<hash>:not(#aared-0#aared-1)` internally; authored
inline values still win. Interning distinguishes pseudo-element/declaration keys and rejects
detected local hash collisions (§128).

`resource-id.ts` gives each SVG fill/filter, chart description and highlight name
a lazy 128-bit bundle namespace plus counter. Independent bundles/concurrent
views cannot reuse source shape IDs as paint-resource IDs. Definitions belong to
their render; table subdefinitions derive from its fill name. Search's sheet is
lifecycle-owned and restored on refresh if the host replaces `adoptedStyleSheets`.
Generated classes/resource IDs are internal, not styling/fragment API.

The published chrome tokens are `--aared-surface`, `--aared-ink`, `--aared-border`,
`--aared-muted-text`, `--aared-accent`, `--aared-focus-ring`. The element's
`--aared-text` is a consumer input. Host resets still require explicit inline
element display; root baselines are not protection from direct descendant
selectors. The embedding limit is recorded in the gap inventory.

### Build and lazy boundaries

`tsdown` emits minified Latin-1 `.mjs`, declarations and maps. Shared chunks
must stay under `dist/core/`. The default constructor stays literally
`new Worker(new URL("./worker.mjs", import.meta.url), { type: "module" })` so
consumer bundlers discover/copy it. Worker splitting stays disabled; no sibling
Worker chunks need deployment. Element registration is marked as a side effect.
Entry imports are SSR-safe and do not touch DOM/Worker globals until use.

Shipped literals stay Latin-1 to avoid widening whole scripts in V8. Wider
characters come from `core/chars.ts`'s runtime constants (bullets, bidi marks,
no-break hyphen, return arrow, text-presentation selector, Arabic fold targets
and probe text). Escapes alone are insufficient if the bundler folds them to
raw characters. Build-output and consumer-build checks guard this (§34).

Preset geometry data stays in the PPTX renderer entry, never shared session code.
`fill-loader.ts` and `effect-loader.ts` cache retryable lazy imports. Charts/DOCX
text boxes share optional fills; XLSX/DOCX load effects only when needed. PPTX
imports its effects directly. `chart-svg.ts` fetches layout on the first chart;
DOCX starts `chartLayoutReady` during its font wait. These are mechanisms, not
fixed bundle-size promises; measured sizes live in BENCHMARKS.

Production checks cover Vite/webpack, generic/fixed-format and multi-page builds,
CDN/source Worker modes, lazy charts/highlighted search and borrowed secondary
views. Blocked CSP, missing CORS and direct cross-origin construction must reject
without hanging. Next.js 16 Turbopack/webpack/static export/basePath/CDN evidence
is the earlier manual review (§56), not a fresh run of every later change.

## Format rules

These rules describe current parse/paint behaviour. Numbered citations refer to
Notes findings. Read shared rules before changing a format-specific painter;
limitations and unprobed cases live in the [gap inventory](#gaps-limits-and-fidelity-gaps).

### Text, bidi and theme resolution

- Exact line pitch belongs to the paragraph. On Word `lineRule="exact"` and
  DrawingML `spcPts`, `EXACT_PITCH_RUN` sets run line height to zero so mixed-font
  runs cannot enlarge it; `atLeast` remains a floor (§44).
- A spacing "line" is not CSS `lh`: Word uses 12pt (positive line values win
  over points, inherited separately); PowerPoint uses 1.2 times the largest run.
  PowerPoint drops first-before/last-after spacing and adds adjacent before+after,
  so text bodies are flex columns (§47).
- `ltrRun` CSS pseudo-content keeps Latin dates/percentages together in RTL.
  Word `superscript`/`subscript` resolve to CSS `super`/`sub`. DrawingML's authored
  line multiple remains in the model, but browser paint floors at 1.1 to avoid
  clipped ink; PPTX leading placement accounts for that floor.
- Theme complex-script lookup falls from an empty slot to its `Arab` script font.
  PPTX colours use slide override, layout override, then master colour map;
  each master owns its theme/table colours. XLSX/DOCX charts use Office's default
  map. `lnRef` widths come from the actual theme's line list, not fixed guesses.
- DrawingML named colours resolve before shade/tint/alpha. Word theme ink applies
  `themeShade`/`themeTint` to HSL lightness with channel truncation; cached `w:val`
  is fallback only when the theme slot is unavailable. This shared run-property
  path covers styles, lists, stories, notes and text boxes (§95, §100).

### Geometry, fills, pictures and effects

The Worker carries **12 hand-coded presets**. Other PPTX presets/custom geometry
travel as deferred programs and are resolved on the main thread against the
PPTX-only generated catalogue before a read returns. DOCX text boxes use the
fast-path geometry, without importing that catalogue (§26).

Geometry may carry fill-only/stroke-only subpaths. A preset's `textRect` is local
viewBox geometry: PPTX and DOCX convert it against the painted frame and add
body insets. Vertical insets use pixels, not CSS percentage padding (which is
width-relative). Ellipse insets follow `(1 - 1/sqrt(2)) * axis / 2`; rounded-box
text insets use `radius * 29289/100000` (§98).

Unresolvable geometry stands in for the **outline**, not the whole shape: the
authored fill/text still paint. Development draws a dashed fallback; shipped
paint draws no fallback outline. Text-only/missing-geometry placeholders are not
diagnosed as authored unsupported shapes.

Every PPTX fill, DOCX text-box shape fill, chart fill and non-solid XLSX cell
fill uses `fill-paint.ts`: solid, linear/path gradients, **54** DrawingML preset
patterns plus SpreadsheetML's **17** cell patterns, pictures stretched/cropped/
tiled. All gradient stops use `officeGradientStops`, including WordArt. A fill
the painter cannot paint returns `undefined` and is diagnosed, not silently
treated as authored `none`.

XLSX cell fills follow Excel's dialect, read from its screen bitmaps (§130):
- **Gradients** (`gradientFill`) blend straight in sRGB (`blend: "srgb"`), not
  PowerPoint's linear-light curve. A `degree` is a scaled angle, so 45° runs
  corner to corner. `type="path"` is a rectangular gradient from the box its
  `left`/`right`/`top`/`bottom` positions give.
- **Patterns** anchor at each cell's own top-left, where PowerPoint anchors at
  the slide origin.
- **Painting:** a solid fill stays the cell's CSS background. Anything else is
  an SVG layer behind the text (`cell-fill`), and the painter loads only for a
  sheet that has such a cell. A conditional-format fill stays a plain colour
  over it.

`core/image-source.ts` resolves embedded SVG (`asvg:svgBlip`) first, with the
main `blip` raster as fallback. A read/decode failure switches to Office's raster;
a lone SVG source paints alone. Pictures, picture bullets and picture fills
share that route. External media relationships are not fetched (§108).

`DrawingEffectsModel` resolves conventional outer/inner shadows, glow, soft edges
and bottom reflections. `effect-paint.ts` uses CSS drop-shadow or SVG filters over
live DOM/SVG. Reflection uses native CSS where available, otherwise inert,
aria-hidden mirrored DOM with a mask: no duplicate selection, links or search.
Media refresh/observers and filters belong to the render and work in shadow roots.
Unsupported effect families remain explicit (§101, §102).

### Tables

PPTX/DOCX share cell shells and conditional-style-region ordering, not all border
rules. `TableModel.direction` controls visual order without reordering XML cells.
Header exclusion from band numbering is shared through `core/table-style.ts`.

PPTX uses physical left/right padding and paints shared edges once. Missing
deck definitions for the seven Medium Style 2 IDs resolve from PowerPoint's own
styles on first use; other absent built-in definitions remain unstyled (§54).

DOCX style regions come from `styles.xml` (`tblStylePr`), under direct formatting.
Region run/paragraph properties sit below paragraph/character styles: a paragraph
colour can win while a region's otherwise-unset weight remains. `tblLook` explicit
attributes win over its legacy bitmask. Table borders/margins merge base-first
through the full `basedOn` chain, edge/side individually; direct zero padding and
explicit no-border win last (§99, §104).
Word `tcBorders` names four cell sides, not DrawingML's inside/outside regions;
each style region overlays those sides.

Word `bidiVisual` makes borders logical, margins physical and indentation
start-relative. Cell edges collapse; `sideInsetPx` reserves the border's inner
half in either neighbour so text stays a margin beyond it (§43, §44).

### Charts

All formats share `ChartSpecModel`, one Worker reader and one page-side layout/
SVG painter. Classic types include column/bar, line/area, pie/doughnut, scatter/
bubble/radar, stock, pie/bar-of-pie and combinations. Office 2016 types include
waterfall, funnel, histogram, Pareto, box-and-whisker, treemap and sunburst.
XLSX chartex formulas resolve hidden `_xlchart.*` names to cells; DOCX reads
chart parts ahead because drawing-node parsing is synchronous.

Chart layout uses the frame's real size in points, not a scaled large-chart
template. A DOCX chart is placed like a picture; a floating XLSX chart uses its
drawing anchor. Layout rules (§71–§74) are:

| Topic | Rule |
|---|---|
| Value axes | Keep zero unless all data are in the top sixth of their magnitude. Floating ends gain half the data span; zero-held ends gain 5% headroom (none on radar). Steps are the smallest 1/2/5 times a power of ten giving at most ten intervals and labels at least 1.5 font sizes apart. Error bars extend the axis. |
| Category labels | Wrap at spaces up to three lines, then rotate 45°, then 90°, then thin only if upright labels still overlap. |
| Layout constants | `K` in `core/chart/engine.ts`: value-label gap 0.92 font sizes, category centres 1.34 below the axis (outward ticks fit in that gap), plot right room 4pt unless labels use it, title height 1.5 font sizes plus 3.3pt gap. End labels reserve room beyond the plot only where they protrude. Stacked legend rows add 5.9pt. Top-right legends reserve their band. |
| Legends | Side legends take at most a third of the chart. Top/bottom wrapping uses equal cells (widest entry plus 0.14 font sizes), minimum-row columns, and line plus 0.48-font-size row spacing. Nothing paints outside the frame. |
| Pie labels | Category/percent stack unless a separator is authored. Whole percentages sum to 100, residue to the largest slice. Best-fit uses slice room; inside wrapping uses 0.85 radius, outside wrapping a fifth of plot width. Outside labels spread with leaders; manual positions are chart-fraction offsets. Dragged labels get a thin black leader with a level end unless touching the pie. |
| Grouped categories | First `c:lvl` is innermost; outer labels centre on spans with separators through label rows. |
| Date axes | Days below a 28-day gap, months below a year, otherwise years; absent major units label each base unit while room permits, then rotate 45° and step/thin. Date axes never rotate upright; non-date codes paint serial numbers. |
| Unstyled defaults | Axis/tick/trend/error/leader/hi-low lines resolve Office defaults; stock automatic markers/down bars resolve too. General trend equations use five significant digits and R² four places; the label is centred at the line end. |
| Stock and of-pie | Hi-low/up-down bars use slot/(1+gap); default of-pie split takes the last third, Other at 3 o'clock, secondary size from `secondPieSize`. |
| Office 2016 | Histograms start at minimum with right-closed bins, width `3.5 * sampleSD / n^(1/3)` rounded up to two digits. Pareto sorts/sums categories with a secondary cumulative-percent line. Box plots default to exclusive quartiles, inclusive when the series specifies it. Treemaps use squarified layout with grey parent banners (`tx1` 35%, two label sizes). Sunbursts begin at 12 o'clock. |
| Dense lines | M4 keeps first/min/max/last per half-point plot column. |

Chart text is an SVG paragraph: its first strong letter controls all label lines.
Mixed scripts use separate tspans so Arabic takes `a:cs`, theme cs/Arab or Arial,
not a Latin face merely because it has Arabic glyphs. Joining-form advance tables
include lam-alef and zero-width marks. The outer SVG is LTR; RTL lines swap their
anchors to preserve layout positions on RTL slides/sheets/hosts. Browser-painted
face measurements scale Office layout metrics when the face differs (2.5%
tolerance). Stored chart `normAutofit` scale/reduction applies after inheritance.

Data labels preserve typed text and fields (`VALUE`, `CATEGORYNAME`, `SERIESNAME`,
`PERCENTAGE`, `CELLRANGE`) with cached fallback. Value-from-cells lives in c15
extensions. Chart picture fills carry their owning part relationships/media types.
SVG charts have role/name/description and real searchable text; settlement waits
for charts to finish painting.

Oversized caches and unsupported plot types produce a **whole-chart placeholder
with a reason**, never a partial plot. Surface/map and layout/master-chart limits,
metric residue and unmeasured oracle cases are grouped below.

### PPTX scene and text bodies

The slide is a fixed scene at **96 CSS px per inch** under one uniform scale.
Back-to-front paint is effective background, shared layout/master graphics,
slide elements. Background references recolour through the owning master's map;
`showMasterSp="0"` suppresses inherited graphics. Layout models share by reference.
Groups map `chOff`/`chExt` into placed bounds; lines/connectors use `p:cxnSp`.

Body attributes cascade independently, nearest-present slide → layout → master:
anchor, wrap and four insets. `masterTextStyle` and placeholder type agree so
subtitles inherit body spacing. Absent insets resolve to DrawingML's 0.1in
left/right, 0.05in top/bottom. Text past its box paints in all autofit modes;
only the slide edge clips (§48, §54).

PowerPoint autofit uses only stored `fontScale`/`lnSpcReduction`, applied during
parse with whole-point rounding. Bare `normAutofit` stays 100%; no measurement
pass invents a shrink. Runs without an authored size take `p:defaultTextStyle`'s
per-level size. `leadingShiftEm` accounts for PowerPoint's non-centred
leading: above 100%, text sits `0.9 * spacing * size` from the line top; below
100%, removed leading belongs above it, with the browser floor included.

Bullets reserve at least the hanging gap, without their own line height; RTL
`marL` counts from the right. Colour/size follow `buClr`/`buSzPct`/`buSzPts` or
text defaults. Picture bullets use the session and a character fallback;
autonumber sequences restart by type. Runs resolve underline fill/line,
strike/caps/spacing/baseline/highlight and ea/sym typefaces. Ink falls through
shape/layout/master `fontRef` and table text to Automatic (white on dark,
otherwise black), never the host's inherited ink.

In an LTR paragraph, an LTR-language run containing Arabic around ASCII
punctuation is split into `ltrRun` pieces (`arabicPunctuationPieces`); Arabic
language, whitespace alone and RTL paragraphs use normal bidi (§54).
WordArt paints outline as text-stroke, linear gradient as background-clip and
glow/shadow as text-shadow. Unsupported path gradients are marked;
noFill/text-grpFill runs retain source markers. Warp/path-gradient limitations
are in the inventory.
Run/shape links emit `linkactivate`; a run link wins inside a shape.

`grpFill` explicitly requests the parent group's paint; direct child fill wins.
`fillBounds` carries the ancestor paint box into child space so gradients span
the group. Groups have no rectangular backing. Text-run `grpFill` instead
resolves noFill, retaining its source flag. Group effects apply to the combined
child silhouette; shape effects remain local (§101).

Slide-number fields show the displayed slide number, including shared graphics.
Other fields, including datetime text, keep cached text. Notes are resolved onto
`SlideModel.notes`; title/notes presentation policy is in the inventory.

### XLSX grid, selection and conditional formatting

File column widths, row heights, merges, hidden tracks and frozen panes stay
authored. The browser decides fit/wrap/spill/clipping from text it actually
paints; the renderer does not resize tracks to compensate for font residue.
Context-ordered cells use `dir="auto"`; sheet RTL mirrors columns, not every
cell's text direction (§116, §117).

Unwrapped text spills into empty neighbours, stopping at content. Left/right
alignment spills that way; centre alignment stays centred on its own cell even
when one side is blocked (`TextSpill.after`). After paint, `settleSpillLines`
reads text/cells in a batch and hides only grid edges reached by ink; authored
right borders stay. Font settlement repeats that pass. Cell padding is 1.5px
inside its 1px border (`CELL_PAD_PX`). Short text beside blank neighbours leaves
their uncrossed gridlines intact.

Empty cells have no text span/direction/alignment. Painted spare tracks/rows use
single fill cells (`data-aared-fill="tracks"` / `"row"`) with gradient gridlines.
`syncFillLines` runs before the first window and on scale/resize. The first
window paints in the attachment task, not a later animation frame. A max-height
clipper supplies its room limit, not current content height. `syncViewport`
caps a content-sized scroller above 120px room; measurement tooling must lift
any old cap before resizing and reapply it afterward (§40, §44).

Headings are sticky/exclusion-aware and default **on** in every layer. Low-level
render leaves sheet tabs off; preview enables them unless explicitly disabled.
Virtualized repaint
restores selected-cell focus. Selection is a cursor plus optional extended
range (Shift/arrows or drag); ordinary movement reanchors. The formula bar
shows saved formula/formatted text and range aggregates (Sum/Average/Min/Max/
nonempty Count), never calculation or a bottom status bar (§115, §118).

Visible same-sheet formula references have cell tints and one dashed perimeter
per range; extended selection has its own wash and cursor. Selection-restore
also restores these on scroll. Empty area past data is selectable by geometry
(`fillCellAt`, `paintSelection`), with a cell pointer (not on headings); arrows enter painted fill,
while Home/End/page keys stay on data. Floating drawings extend the grid beyond
used range using default track dimensions. Their transforms/effects remain
session-owned and are requeued while unattached or the lazy painter loads;
only a failed import marks painter-unavailable (§112).

**Conditional formatting is a snapshot over cached values**, in priority and
`stopIfTrue` order (§105–§111, §126). Supported rule families are cellIs, text
rules, duplicate/unique, colour scales, data bars, top10, aboveAverage (population
SD thresholds), timePeriod, icon sets and expression, including x14 expression
parts with xm formulas/ranges and inline differential styles. `dxfId` names a
style, not a stored match; conditions must be evaluated at parse time.

| Conditional-formula subset | Semantics |
|---|---|
| Core | Bounded arithmetic/comparison/concatenation and supported logical/error tests, over cached values |
| Dates | TODAY/NOW at parse time; DATE rollover/1900 leap day; YEAR/MONTH/DAY; WEEKDAY kinds 1–3 |
| Text | SEARCH with case-insensitive wildcards; literal case-sensitive FIND; 1-based MID; range-joining CONCAT |
| Aggregates | COUNTIF/COUNTIFS with wildcard/comparison criteria and Excel blank rules; SUMIF/AVERAGEIF top-left-aligned value ranges; SUMIFS/AVERAGEIFS identical value/criteria row and column dimensions |

Relative references shift from the first range's origin per target, including
cellIs operands. Other sheets load values-only, not their rules, preventing
evaluation recursion. Missing formula cache is unknown, not blank. Unsupported
syntax/missing cache/budget exhaustion leaves the entire rule plain and reported.
Scalar Excel errors never directly match; zero/FALSE/blank are false and a bare
empty-string condition is VALUE error. IFS shape mismatches are VALUE errors,
even with equal cell counts, and ISERROR can match them (§126).

Blank recipients materialize only for small ranges (at most 50,000 cells per
rule); larger blank-sensitive ranges remain authored-only. COUNTIF-family blank
counting stays arithmetic. Two/three-colour stops resolve percent over min/max
span and percentile by rank interpolation, clamping endpoints. Native
`CF Rules!E2:E12` pins the min/median/max scale (§109).

Icons use a 16px SVG catalogue at inline start: RTL mirrors position, not arrow
direction. The first threshold is ignored, remaining bounds are inclusive;
showValue false hides text, NoIcons paints none for that tier. timePeriod tests
date-formatted numbers against local calendar time captured when parsing.
Unsupported x14 linked data-bar extensions remain unpainted.

### DOCX flow, pages and drawings

Flow mode concatenates section bodies and paints the last section's default
header/footer once. Layout mode waits for fonts, measures eligible top-level
blocks and table row tracks off-screen at section body width (`measureFlow`),
then paginates. Rect heights are divided by `paintedToLayoutZoom`; all comparisons
use layout pixels, not zoom-scaled rect pixels (§94). Standalone drawings and
note/header/footer reservations remain estimated; DOM-free pagination is fully
estimated from per-script average advances (`AVERAGE_CHAR_EM`). A different
painted font can wrap differently from those estimates. Page placement is
structural, not Word-pixel-identical.

Tables split between rows. A leading `lastRenderedPageBreak` in the first cell's
first paragraph marks a preferred row boundary; `splitTable` prefers it over
its latest-fitting guess when soft breaks are preferred. Fragmentation/repeated
headers and mid-paragraph break residue are recorded separately below.

`alignFor` swaps physical left/right in RTL. Word single-line pitch follows
run sizes at about 1.15em; first-block spacing drops after breaks. A paragraph
without `pStyle` uses the default paragraph style, not bare doc defaults. Arabic
or `w:rtl` runs take bCs/iCs (including Latin text inside RTL runs); other runs
take b/i. Empty paragraphs paint a line at paragraph-mark size plus spacing,
not a heading; a table cell's lone empty paragraph gives its row that line.
A break-only paragraph is not an extra empty line. Empty
pageBreakBefore+section-ending paragraphs are section markers, not blank pages.
docGrid snaps eligible line spacing upward to its pitch (18pt if omitted);
snapToGrid can disable it, exact spacing stays exact, before/afterLines use
grid pitch, and table rhythm follows adjustLineHeightInTable (§55, §76, §90).

Headers/footers inherit **each kind independently** from preceding sections.
titlePg chooses first on a section's first page; otherwise evenAndOddHeaders
chooses even by printed number, else default. Missing kinds show nothing.
pgNumType start restarts printed PAGE numbers and parity; titlePg/restarts do
not inherit. Headers hang from their authored distance without clipping.
PAGE/NUMPAGES update for layout; other ordinary fields retain saved results.

Anchored drawings live in paragraph `floats`. `planFloat` uses the section/page/
cell `AnchorFrame` for four placement kinds:

| Kind | Current placement |
|---|---|
| Wrap | One-side CSS float; wrap distances pad the border box and `shape-outside` uses that box or the tight/through polygon |
| Band | A column-width float excludes text above and below the drawing |
| Overlay | Absolute no-wrap paint; z-index -1 behind text or 1 in front, using paragraph origin or page/margin origin in layout mode; a header translates its origin |
| Split | Text slots on both sides of a centred exclusion, then normal flow below |

`split-wrap.ts` builds slot/hole/slot rows in reading
order (right first for RTL), preserving styled words; canvas measurement is
checked against painted DOM. Resizes and font settlement rebuild while alive,
not after teardown (§55, §126).

Text-box exclusions replan from **effective frame size** after settlement;
initial plans use authored extent. Ordinary pictures keep authored size.
Block pictures/charts with unauthored/justified alignment use paragraph start,
not a nonexistent text justification; anchors determine anchored placement.
effectExtent reserves drawing ink room: block padding or inline insets without
moving an inline picture off its baseline (§102).

#### Text boxes

DrawingML and selected VML AlternateContent branches reuse flow nodes for
paragraphs, lists, tables and search. DrawingML shape fill comes only from
direct spPr fill, not an outline's colour/hiddenLine. spPr/bodyPr belong to the
first owning wsp, not a nested box (§95).

`docx/text-box.ts` uses three layers:

1. **Frame:** untransformed layout size, consumed by wrapping and pagination.
2. **Surface:** geometry/paint, insets, anchor, rotation and clipping.
3. **Content:** browser text as one flex item, preserving collapsing paragraph margins.

Rotation cannot contaminate pagination's frame rect. `settleTextBoxes` batches
size reads then SVG writes, with one ResizeObserver per rendered root; lazily
skipped pages retain last valid paint. spAutoFit sizes from content in both
directions; fixed frames clip regardless of vertOverflow. Word ignores stored
normAutofit scale, so DOCX does too. wrap none uses white-space pre and widens
to the line; anchor positions top/middle/bottom. Rotation/flips reuse the shared
transform reader; preset dash resolves to the existing border style. Clipped
words remain in the resolved model and are still searchable.

An SVG inside the surface routes fill/outline through the shared painter and
geometry text rectangle. Effects settle after boxes, never in measurement, so
filters do not move pagination. Their observers/media clean up with the render.

#### Lists, headings and fields

List counters run by abstractNum, not numId; startOverride restarts a level on
first use and shallower items restart deeper levels. Empty numbered paragraphs
and numbered headings consume numbers. Headings use the normal inline marker
path without inventing indent; heading text includes the marker while runs
retain authored words (§77, §94).

TOC tracking is context-level/body-only, not per paragraph or inside tables,
boxes, stories or notes. Refresh occurs only when every result-changing switch
is understood: o/u levels, h links, n numbers, p separator, b bookmark range;
w/x/z are accepted no-ops. Caption/sequence/TC switches (a/c/s/d/f/l) or missing
bookmark range keep the **whole saved result**. Entries use TOC styles, body and
table-cell headings, and stable heading/toc anchors; the public outline remains
body-only. Bookmark ranges use body block-index space, merging same-named ranges
and dropping cross-section ranges. Layout PAGEREF entries follow aared pages;
flow keeps saved numbers and gives new entries none (§94).

Tracked changes show final content: insertions/moves-to in, deletions/moves-from
out. Their native fixture evidence gap is distinct from the implemented reader.
Layout pages settle tab stops when first shown; no paragraph-text wrapper is
required (§37). The tab pass reads every paragraph, then writes every paragraph,
a round per tab ordinal, with a read-only check afterwards, so opening costs a
few forced layouts however many paragraphs hold tabs; a read after each write
per paragraph made a 700-page file take ten seconds in layout and minutes in flow (§37).

## Fonts

**Default: the browser's fonts, with no download.** Requested installed fonts
win; missing families fall through to generic. That can change wrapping and
pagination. `aared-fonts` stays frozen/unpublished (§53); its optional substitutes
are calibrated look-alikes, not a promise of identical Office lines.

### Font delivery and aliases

1. The Worker resolves direct/inherited/theme family slots, including complex
   script. Renderers do not inherit a host's font as document content.
2. On a unit read, the main-thread session calls optional fontResolver once per
   requested family. Search-only text reads carry no model and request no fonts.
   Preview/hooks forward the resolver; the element takes none.
3. Paint uses requested family → `fontAlias(family)` → generic. A resolver
   registers delivered FontFaces under the **alias**, never the requested or
   substitute name, so a real installed requested font still wins (§45).
4. Faces fetch when painted. `awaitDocumentFonts` loads used family/weight/style
   subsets with a bounded 2s wait before measurements/fidelity checks in every
   format. DOCX layout measures after this wait; XLSX keeps file track sizes.

`collectFontUse` samples PPTX/DOCX run text and XLSX value/formatted text, including
Arabic digits. FontFace registration is on document.fonts, across shadow roots.
Offline delivery falls back; a subsequent render can recognize a face that
loaded after the bounded wait (§50, §126).

### Frozen substitute resolver

`createResolver`/`defaultResolver` normalize requested names, fetch Google Fonts
CSS once per substitute and register subset faces separately for each requested
alias. Unknown families cost no network; failed delivery resolves undefined;
without DOM/FontFace the resolver is inert. It does not eagerly load font files.

The [frozen family choices](Notes.md#frozen-calibration-outcome), the reasons for
them and the alias/metric experiments are experience records in Notes §45–§53.
Pinned descriptors and recalibration procedures live in
[tools/font-calibration](../tools/font-calibration/README.md). Substitute coverage
limits stay in the inventory below; neither those experiments nor the Windows
fallback measurements (§116) authorize resizing authored tracks or glyphs.

## Fidelity contract

aared is a reader, not an Office converter. These capabilities are delivered
with the specific [limits](#gaps-limits-and-fidelity-gaps) below, not a blanket
pixel-equality guarantee. Missing in-scope behaviour is a bug; permanent scope
boundaries and undecided proposals are not promised capabilities.

| Format | Delivered reading scope | Permanent/excluded scope |
|---|---|---|
| PPTX | Positioned scenes, geometry/fills/pictures/effects, browser text/bidi/stored autofit, tables, shared charts, links, notes as data | Animations, transitions, 3D rendering, VBA, media playback; SmartArt uses a fallback image if present |
| XLSX | Cached values/formats, cell styling, authored tracks/merges/hidden rows and columns, frozen panes, RTL, selection/chrome, supported conditional snapshots, pictures and shared charts | Workbook/pivot recalculation, macros, data-validation UI, sparklines |
| DOCX | RTL flow, tables/lists/notes/headings, text boxes and wrapped drawings/charts, final tracked-change view; optional physical pages with headers/footers/footnotes and page fields | Pixel-identical Word pagination; multi-column nextColumn degrades to continuous |

WordArt warps and pagination fragmentation are **limitations with their own
dispositions**, not permanently excluded merely because absent.

### Warnings and diagnostics

Every render handle carries readonly `RenderWarning[]`. DOCX layout combines
pagination warnings (oversize-block, footnote-overflow) with font warnings;
flow/PPTX/XLSX also report font-unavailable. `detectFontFidelity` compares painted
widths with a nonexistent generic control; document.fonts.check alone cannot
identify availability. Requested generic-default faces can be indistinguishable
and are not flagged merely for matching; a working alias is substituted, not
unavailable. This is a heuristic, not browser font-identity introspection.

Model unsupported notices and DOM fallback markers are distinct from warnings.
Relevant markers include fill/geometry/effect/chart fallback and text-warp;
styling hooks are documented by the API page. MARK_FALLBACKS removes development
dashed outlines and the XLSX notice bar from production. **WordArt text-warp is
written unconditionally**, outside that flag. Missing/text-only geometry is not
marked as an authored unsupported preset. Consumers needing XLSX notices read
`SheetModel.unsupported`, not the development bar.

## Gaps, limits and fidelity gaps

### Known limits

This is the **one inventory**, grouped by related topic rather than discovery
date. The status describes disposition; proof coverage is separate:

- **Open:** still intended to be closed; not a first-release blocker unless named
  in Release readiness. An unstarted capability remains unstarted.
- **Deferred:** retained gap, with its dependency or trigger where known.
- **Accepted:** a deliberate stopping point for a specific difference, not an
  assertion that every similar failure is cosmetic.
- **By design:** intentional behaviour; **Settled/Decided** preserves existing
  scope decisions. Those are not commitments to implement an alternative.
- **Evidence gap:** implemented/assumed behaviour lacks the named native fixture
  or oracle proof; it is not automatically a known rendering failure.
- **Documented boundary:** a current architectural/review constraint, not a
  separate backlog promise or a decision to exclude every future alternative.
- **Unscheduled limitation:** absent support without an existing implementation
  commitment; unlike an accepted limit, no stopping-point decision is implied.

Close open entries by removing/shrinking residual work and moving durable rules
to the relevant section, why to Notes, proof to tests/fixtures. Record subsequent
changes in the active CHANGELOG. Do not
leave a completed backlog narrative. Accepted/by-design differences stay visible
and may be reconsidered with new evidence or a maintainer decision. Unscheduled
proposals are separate and do not become commitments here.

#### Resources and input safety

| Item | Status and impact / next proof |
|---|---|
| Total input budgets | **Open.** No cross-entry inflated-byte accounting/high-ratio aggregate ZIP budget. URL `readResponse` accumulates without a ceiling. Character-reference-heavy XML still costs raw bytes/parse time although entity expansion is refused. Add budgets/typed errors and audit remaining count-driven loops; audited table spans/merges are not reopened by this residue (§72). |
| XLSX range-first parsing | **Deferred.** A first read parses the full sheet, streamed, then slices. A small requested range is not a bounded parse; scroll-driven range rereads/range-first parsing are absent. |
| XLSX window row metadata | **Accepted.** Rows with all cells outside the column window are dropped, including their authored heights; full reads/defaultRowHeight are unaffected. Height-only rows remain within sheet extent, up to 1,048,576 row models. |
| Media memory | **Documented boundary.** URLs/decoded media are per render, not viewport-released; resource reductions remain conditional proposals, not an implemented memory budget. |

#### Shared text, fonts, search and embedding

| Item | Status and impact / evidence |
|---|---|
| Embedded fonts | **Decided.** DOCX odttf/PPTX fntdata are unread with no unsupported marker. Revisit only for a consumer needing Office line breaks, starting with network/ownership cost (§45, §53). |
| Substitute coverage | **Accepted.** No Courier New Arabic, Tahoma/Segoe UI or Aptos Light/SemiBold/ExtraBold/Black/Mono/Serif substitutions. Aptos italic uses upright weight calibration (about 1–3% width residue); CDN files can drift. Lateef's large size adjustment can cross dense borders/underlines and alter justified breaks. Calibration evidence belongs to §49/§51, not a current-device guarantee. |
| Symbol fonts | **Accepted.** Unsupported symbol characters outside mapped bullets/list markers, and Wingdings 2/3/Webdings, keep source codes; devices lacking those faces can show letters (§52). |
| Browser minimum font size | **By design.** Honour the reader setting. Larger painted text can change typography, wrapping and page count relative to Office (§94). |
| Run direction beyond left-to-right runs | **Open: evidence gap.** A run that is not right-to-left keeps its numbers together in an RTL paragraph (`ltrRun`, §44). Two cases remain unprobed: a right-to-left run holding a number ("بنسبة 71%" with the number in an `ar-SA`/`w:rtl` run), where Word and PowerPoint were seen to disagree and aared keeps plain bidi; and, in Word, a left-to-right run holding Arabic words in a left-to-right paragraph ("Mixed: عربي + إنجليزي"). PowerPoint's side of the second case is answered: written order around ASCII punctuation (§54, `pptx-fidelity.test.ts`). Next proof: author those paragraphs in Word and PowerPoint, export, and pin the order Office paints. |
| Arabic search variants | **Accepted.** Do not fold ta marbuta/ha, hamza-on-waw/ya to bare letters, lone hamza, Quranic annotations U+06D6–06ED, Persian ya/kaf. Reopen for a real miss rather than merging distinct words speculatively (§78). |
| DOCX result activation | **Accepted.** Text-based mapping of a search match can fall back to its source place without active-occurrence highlighting when painted fragments no longer hold whole text or identical paragraphs reorder (§119). |
| Direct host CSS | **Documented boundary.** Root baselines prevent inherited traits, not arbitrary descendant selectors. Keep application prose/resets out of painted content or use shadow DOM; internal classes/resource IDs are not public hooks (§128). |

#### Shared graphics and accessibility

| Item | Status and impact / evidence |
|---|---|
| Fill residue | **Accepted.** Off-centre circle/shape path gradients are unpainted/diagnosed; reopening needs a native probe plus a case COM cannot directly author. Pattern anchoring can be local to child/box rather than page; tiled pictures assume 96 DPI; fitted gradient stops retain small colour residue (§28, §95). |
| Effect families/softness | **Accepted.** DAGs, 3D/perspective/skewed shadows and nonstandard scaled/skewed reflections remain diagnosed. Blur/alpha soft edges and glow can differ cosmetically. Rotated/flipped nested-group inherited paint uses mapped rectangular bounds, not a full inverse affine transform (§101, §102). |
| Dense effect stacks | **Accepted.** Overlapping stacks can slow zoom despite usable normal-case effects. The measured normal/stress/fallback comparisons and optimization history are in [Notes §101](Notes.md#note-101) and [§102](Notes.md#note-102), not a device budget. |
| Vector image formats | **Unscheduled limitation.** EMF/WMF have no decoder in the engine; this does not authorize rasterizing browser text. |
| Table ARIA | **Deferred.** PPTX/DOCX table shells lack role=row children; role=table/cell alone is non-conforming. XLSX's grid route is separate. |
| Chart data table | **Open.** Charts have accessible name/description/text but no screen-reader data-table representation. |

#### PPTX scene, geometry and text

| Item | Status and impact / next proof |
|---|---|
| Sub-path fill modes | **Open, not started.** darken/darkenLess/lighten/lightenLess are lost; cube/can/bevel faces paint flat. Carry modes through compile-presets, presets, drawingml-engine, resolved geometry and PPTX/shared fill paint. Prove Office amounts with one shape per mode, engine tests and paired capture. This is a separate capability, not a local colour patch. |
| WordArt warps | **Deferred until a real deck needs them.** Preset warps paint flat with text retained and text-warp marker shipped. Only textArchUp has fixture proof. ArchUp/ArchDown/Circle/Button might use SVG textPath after a probe; other envelopes reshape letters and have no ordinary DOM form. No promise of all-warps support. |
| Built-in table styles | **Accepted.** Absent deck definitions beyond the seven Medium Style 2 IDs remain unstyled (§54). |
| RTL punctuation | **Accepted.** In RTL en-US runs, PowerPoint's written punctuation ordering can differ from browser bidi. LTR-neighbour handling stays within one run; Word parity is unprobed (§54). |
| Tight line spacing | **Accepted.** 70–90% authored line multiples paint at the 1.1em safety floor rather than clipping Arabic ink (§54). |
| Text metrics evidence | **Evidence gap within accepted text scope.** defaultTextStyle versus master otherStyle when they disagree, exact spacing and bullet sizes under stored fontScale lack distinct native cases (§54). |
| Layout colour overrides | **By design; evidence gap.** Slide-level clrMapOvr does not recolour a shared layout model's graphics. No fixture carries that override (§54). |
| Titles and notes | **By design.** Index/handle titles are synthetic Slide N to avoid parsing every slide on open. Resolved slide text is readable; notes are data only, neither painted nor searched (§31). |

#### XLSX styles, text fit and interaction

| Item | Status and impact / dependency |
|---|---|
| Border style representation | **OPEN: renderer representation.** Raw Excel names map to CSS in borderCss; DOCX already resolves solid/double/dashed/dotted before its model. Normalizing XLSX is internal work, independent of the supported data contract. |
| Cell textRotation | **Deferred.** SpreadsheetML alignment textRotation is absent from the style model; PPTX's textRotation is a different, implemented route. |
| General alignment | **By design.** General strings use sheet start, not Excel's first-strong-letter placement. Explicit alignment is unaffected. RTL English differs from native print; Arabic-first General in LTR and RTL General numbers are unprobed (§116). |
| Shared borders | **By design.** Adjacent authored CSS borders both paint: 2px internally versus native Excel's 1px, while outer edges stay 1px. Single-sided rules avoid that; PPTX's shared-edge rule is different (§116). |
| Font fit residue | **Accepted.** Missing faces/generics and small-size browser/native metric differences can wrap/cut sooner despite correct file tracks. CF Rules/Notes and tracker Runtime header cases pin it. Oversized numbers are cut, not replaced by hashes, to avoid reporting font substitution as numeric overflow (§117). |
| Selection/reference overlays | **Accepted.** Natural-coordinate perimeters/fill cursor overlays do not follow sticky frozen tracks; cell content/tint still sticks. Reopen with a meaningful crossing-frozen-pane case (§118). |
| Number-format locale | **Accepted/By design.** Weekday/month names stay English; locale tokens do not select a name table. Missing regional codes use ECMA en-US currency/accounting or short date/time for era IDs. Accounting star-fill spacing is unpainted (§76). The Arabic-locale weekday example remains unprobed. |
| Print-only elements | **Documented boundary.** A4 native print headers/print-title rows have no aared counterpart; review is not a screen-layout oracle. |

#### XLSX conditional formatting

| Item | Status and impact / next proof |
|---|---|
| Negative/middle-axis data bars | **Deferred.** Metadata resolves, but ranges requiring a zero-axis split remain unpainted/reported. Reopen with negative-value native fixture and paired Excel review (§106). |
| Formula subset | **By design.** Names, table refs, arrays/spills, lookups, TEXT, unlisted/volatile functions and exotic WEEKDAY kinds remain unsupported; affected rules fail closed. This is not workbook recalculation (§111). |
| Colour-scale interpolation | **Accepted.** Intermediate colours can differ by one RGB channel level; authored stops match native fixture (§109). |
| Icon artwork | **Accepted.** Twenty SVG families are aared artwork, not embedded Excel glyphs. Six arrow tiers use shaftless filled triangles; tier/direction/approximate colours remain readable (§107). |
| Clock-dependent rules | **By design.** timePeriod follows local calendar time at parse; identical bytes may paint differently on another date, as Excel reevaluates them. |

#### DOCX pagination, wrapping and text boxes

| Item | Status and impact / next proof |
|---|---|
| Cross-page content | **Open.** Table rows and paragraph/list blocks are indivisible; long content can exceed a page. Contiguous tblHeader rows do not repeat. Word permits row continuation subject to rules. Deliver fragmentation/repeated headers as a usable capability with real native fixture cases, not by calling current between-row splitting complete. |
| Anchored pictures | **Open.** Page/margin-relative vertical positioning starts at paragraph top; following paragraphs clear remaining both-side pictures; a top-and-bottom picture after a hanging picture is pushed down; through wrap misses Word's wider lowest-row exclusion; VML pict has no placement. Band/through cases need native probes. Prior paired placement/drift measurements remain in [Notes §55](Notes.md#note-55) and [§91](Notes.md#note-91). |
| Large-document first paint | **Accepted.** Layout mode measures every block and builds every page before `render()` resolves, so it is linear in document size: about 3.4 ms a page measured (a 409-page file opens in 1.4 s; Chrome 155, one session). Heap is linear in painted elements (58 MB there). Flow holds more (97 MB; no lazy pages) and the default mode is flow. First-page-first pagination is an unscheduled proposal below (§37). |
| Header/footer reservation | **Accepted.** Pagination uses each section's default story bands; a taller first/even header moves body inset but does not recompute page breaks. |
| Page rhythm | **Accepted.** Estimated drawings/notes/story reservations and conservative justification leave vertical/wrapping residue. Reopen for visible clipping/wrong-page content, not sub-pixel parity. The paired anchors and dated offsets behind that stopping point are in [Notes §91](Notes.md#note-91) and [§102](Notes.md#note-102). |
| Justification | **Accepted.** Conservative Word 2013 space shrinking can wrap a word earlier; small native/browser Calibri width differences remain (§44). |
| Breaks inside runs | **Accepted rare residue.** Only the first explicit page break in one run splits it; later breaks there are ignored and inline drawings precede that break. Mid-paragraph lastRenderedPageBreak is dropped. This does not settle the open fragmentation capability. |
| Table band size | **Accepted (§99's deliberately unread limit).** tblStyleRowBandSize/ColBandSize are unread; bands assume one row/column. Coarser styles therefore paint too frequently. The native style probe uses band sizes of one; no new implementation commitment is introduced. |
| Refreshed TOC formatting | **Accepted; reconsider after 0.1.0 with a probe.** TOC styles correctly replace heading style appearance, but direct formatting within heading runs is not retained in entries (§94). |
| VML boxes | **Accepted legacy limit.** Fixed wrap type/distance/origin stand in for authored VML placement; VML gradient/pattern fill is unread. This differs from unplaced VML pictures (§95). |

#### Chart fidelity and proof coverage

| Item | Status and impact / evidence |
|---|---|
| Unsupported chart hosts/types | **Accepted.** Surface/map show whole-chart placeholders; PPTX layout/master charts are unpainted. 3D chart data/colours draw flat. |
| Axis/title metrics | **Accepted.** Axis titles retain small offsets; smoothed-scatter axes ignore curve overshoot, bubble x axes ignore bubble radius. An empty-defRPr minimal title differs from PowerPoint's size; reopen with another meaningful case (§74). |
| Pie/date policy | **By design.** Outside pie labels avoid overlap rather than copying Office collisions. Date formats paint as written, not host system short date; different Office label widths can change rotation/plot room. |
| Unstyled border | **Settled.** Excel's thin-grey chart-area default is used in all formats; PowerPoint omits it, Word was not probed. No host-specific hairline default. |
| Text/paint residue | **Accepted.** Arabic ascent/line height follow Latin face, bold Arabic uses regular advances, unlisted Arabic Supplement uses 0.45em. Gradient/pattern chart line/text uses one colour; shape-sector gradients use bounds, patterns use chart origin. |
| Advanced chart residue | **Accepted.** dropLines are read not drawn; stacked hi-low lines use raw values; moving averages skip blanks; radar has no trends/error bars; chartex per-point labels/style parts and chart theme overrides are unread. |
| Assumed rules | **Evidence gaps, not established misses.** Histogram overflow/underflow/count rounding, box-plot mean line, grouped separator reach, minimum pie radius and moving-average blanks still need native cases. They remain within the prior accepted algorithm scope (§74). |
| Oracle reach | **Open proof gap.** The 94-case vector corpus does not score every host/font route. Stock/of-pie/2016 types have eye-reviewed Excel exports; XLSX and devices without Office fonts are not scored. Add cases rather than presenting those routes as quantitatively verified. |

#### Fixtures, diagnostics and tooling

| Item | Status and next proof |
|---|---|
| DOCX trHeight | **Open evidence gap.** No Word-authored coverage fixture carries it; row-height parsing proof is synthetic-only. |
| DOCX Office 2016 chart | **Open evidence gap.** cx:lvl cache reading has synthetic coverage, no native chartex in docx-coverage. Paste through Word and pair it as for PPTX. |
| DOCX SVG picture | **Open evidence gap.** Shared resolver has proven PPTX/XLSX routes, but no Word-authored SVG coverage picture. Add through Word in its own probe/review pass; appended content changes pagination/drawing expectations. |
| DOCX tracked changes | **Evidence gap.** Final-view parsing is covered synthetically; no Word-authored tracked-change fixture. No claim of native fidelity proof. |
| Word token diagnostics | **Open.** compare-pagination pairs repeated tokens by occurrence without context; tokenization differences can create false missing/extra/page mismatches. Improve matching/expose ambiguity before using it as an automatic fidelity gate; paired images and unambiguous anchors remain evidence. |
| Chart metric drift | **Open.** Generated metrics require Python/Windows fonts; generated-drift tests check presets only. Add a manual gen_metrics --check path; do not imply bun test regenerates chart metrics. |
| Competitor bench rows | **Open.** pptx-browser/Extend/Apollo have download sizes without open/navigation/memory rows. Measure once or leave those comparisons out; numbers/dates live in BENCHMARKS. |
| Review environment/output | **By design.** Native review needs Windows/Office; same sanitized filename stems share a replaced output folder. XLSX parts are curated merge-clean A-start row bands, not arbitrary viewports. DOCX uses built output; absent-font columns require network. No parallel harness/bulk cleanup (§123, §125). |

## Quality strategy

Renderer capabilities must re-read actual DOM/model output, not trust an
intermediate. Format-shape cases use real Office-authored fixtures, with Arabic
and tables in every format. Pure-function edges may be synthetic. Tests prefer
computed styles, text ranges, track positions and numeric SVG geometry over
serialized CSS/attribute mechanisms.

Fixture tests pin authored labels, order, geometry, break causes and story
placement rather than frozen whole-corpus totals. A fixture-only XML break oracle
independently grows with explicitly appended pages; public page metadata matches
rendered DOM and PAGE/NUMPAGES use that final count. Focused cases retain exact
counts. This avoids both brittle totals and unexplained trailing pages (§124).

`test/harness.ts` owns visual gating, bundled entries and page/server lifecycle.
`browserModule` memoizes resolved entry paths per run and evicts failed builds.
`withVisualPage` creates a private server/WebView per call and also stops a
supplied server unless keepServer explicitly transfers that responsibility.
Tests keep route/MIME/font gates and within timeouts where needed. Memoized module
text shares no page state (§127).

Visual commands select `*.visual.test.ts` **files**, not titles. Plain bun test
skips browser proof; run test:visual after fonts, wraps or content lifetime work.
The same file selection backs artifacts/absent-fonts. All three formats have
registered-alias success coverage; bounded offline delivery and later successful
rendering are also tested. PPTX captures independently compare painted aspect
ratio to public sizes, not just forced screenshot dimensions (§126).

### Native references and developer tools

Use `bun run tools --help` and [tools/README.md](../tools/README.md) for current
commands/prerequisites. The dispatcher composes existing backends, not a second
Office/browser harness. Shared export/rasterization live in tools/review;
shared process/paths/CDP in tools/shared; probe COM helpers in probe/_common.

| Reference | Evidence pipeline |
|---|---|
| PowerPoint | Same-size native PNG and Bun.WebView DOM/SVG capture, checked for typography, bidi and wrapping. Tests additionally read bullet side, tables and text containment. |
| Excel | Native visible-sheet A4 print (portrait, fit-one-page, headings/gridlines, normal margins), PyMuPDF 150 DPI; aared A4 fit at equal display width. Human comparison, never pixel equality. |
| Word | Private-copy fields updated without saving; PDF rasterized at 96 DPI with exact span font/size/positions. aared layout pages and token reports accompany images. Native review is explicit, not ordinary bun test. |
| Chart oracle | Specialized vector-PDF corpus/scoring and metrics generation, distinct from A4 document review. Proof limits are listed in the inventory. |

`bun run tools review <file>` handles arbitrary files through those same routes.
Absent-fonts adds a third column by renaming requested fonts and delivering
the frozen aliases, since installed faces cannot be hidden; DOCX normal and
absent-font tokens stay separate. LibreOffice is not a PPTX/XLSX oracle because
it substitutes fonts. Native tool prerequisites and Windows automation caveats
belong in the tool catalog/format READMEs and AGENTS.

Generated output lives under ignored tools/out/{pptx,xlsx,docx,charts,benchmarks,
search-fixtures}. Each native export replaces its input-stem folder; keep derived
work elsewhere. Chart corpora and font research are pinned evidence, not disposable
output. Promotion of generated chart inputs/source is explicit. No automatic or
blanket deletion of unfinished probes; codified topic findings follow the probe
lifecycle in AGENTS and tools/probe/README.

Tools TypeScript participates in typecheck; declaration generation uses source-only
configuration so tools/tests do not emit library declarations. CLI parsing is
side-effect-free, validates before launching backends and retains failure artifacts
while returning failure status for failed/empty measurements. Review parts preserve
visible-sheet order and grid intersection, including deep-LTR/RTL crops (§123–§125).

The bench's React app and measurement HTML build as separate module graphs;
the second Vite pass retains app assets. Shared package helpers must never be
borrowed from the app bootstrap: lazy format imports on the measurement page
would then execute its unrelated React mount. A production-output browser test
opens all three Arabic fixtures without a React root (§133).

The site has four consumer pages: Home (evaluation and evidence), Getting
started at `/start/` (first success and task recipes), Playground (live
integrations) and API reference (supported contracts). Getting started leads
with registration, bounded host and cleanup before presenting integration
choices. Data schemas live beside their format operations; shared lifetime,
appearance, search and deployment rules have canonical API sections. Repository
docs own parser/Worker implementation history, not consumer instructions (§132).
Home uses fixed-format previews; the playground demonstrates element, hooks,
JavaScript viewer and manual handle rendering, never detached models. Minimal
snippets explicitly exclude the app-owned demo toolbar. Pointer and keyboard
tab changes carry the selected file/format; manual paging reports failures,
allows retry and removes controls on teardown. Snippet compilation, cross-page
prose guards and fixture-backed demo DOM tests cover these boundaries.

The site build also emits plain Markdown for agents: `/aared/llms.txt` (an index),
`/aared/start.md`, `/aared/api.md` and `/aared/llms-full.txt`. They are generated
by `site/scripts/generate-llms-txt.ts` from `start.mdx`, `choosing-a-layer.mdx` and
`api.mdx`, so they cannot drift from the pages the tested snippets and
`docs-parity` already guard; regenerate, never hand-edit. Only MDX module import
nodes bind partials, which expand at their component use sites; fenced examples
retain their code, including relative imports. Components are unwrapped to their
children and in-site links become absolute, because a link in a file read
off the deployment resolves against nothing. Notes and ARCHITECTURE are never
inputs. The package README is the agent entry point that points here.

Bench measurement uses fixed previews (XLSX tabs off, DOCX layout on),
with forwarded cancellation/one destroy. Interactive cards use usePreview;
mounting cards own lifetime and frames navigate without destroy ownership.
100% in that UI means default fit. Baseline adapters parse independently so a
competitor column does not depend on aared (§89). The interactive bench supplies
no fontResolver unless opened with `?fonts=aared`; installed fonts therefore win
in ordinary bench views, unlike the explicitly renamed absent-font review.

Benchmarks are manual, rare, on an idle mains-powered machine. Run/report commands
and dates live in BENCHMARKS/tools catalog; compatible same-machine runs preserve
competitor observations. CDP Worker reads use tools/shared/cdp.ts, not WebView.cdp
sessionId (which still reads its page); recalibrate after touching routing.
There is no push/PR test CI. Checks run locally before pushing; the only GitHub
workflow validates a published stable release's tagged checkout and deploys the
docs site. It rejects tags whose version differs from the viewer manifest,
unannotated tags and commits outside `main`. Pages write/OIDC permissions belong
only to the deployment job. Bun is pinned to the locally validated version (§134).

The site is hosted at `https://abdo-eid.github.io/aared/`, with Astro/Vite base
`/aared/`. Astro UI uses `sitePath`; MDX links use sibling-relative routes.
Consumer code examples keep application-owned paths unchanged. A production
WebView test serves only the base-prefixed site, checks direct page navigation
and re-reads all three formats through the four playground integrations.

### Recorded verification

The evidence records live once, with their dates and boundaries:

- [Notes §103](Notes.md#note-103): the maintainer's PPTX review acceptance.
- [Notes §126](Notes.md#note-126), [§127](Notes.md#note-127) and
  [§128](Notes.md#note-128): deferred-coverage, harness and renderer-isolation
  runs, including test totals, native pairs and deployment checks.
- [BENCHMARKS](BENCHMARKS.md): published measurement tables, dates and method.

These are records of prior runs, not verification performed for this editorial
pass or a new Word fidelity acceptance. Release gates are stated below, not
inferred from a test total or benchmark date.

## Release readiness

The data/model boundary is decided (§135). Fidelity inventory entries
are not release blockers unless explicitly promoted to a gate. PPTX visual
review acceptance is already recorded; this rewrite neither reopens it nor
declares every format pixel-identical.

### Resolved-model contract

**Decided:** supported consumer operations and semantic data, private rendering
models in all formats. Format entries export the complete small inspectable
schemas; `core` exports shared grammar/errors/warnings, not rendering models.
Public document/handle/collection types are interfaces, with construction and
model state internal. Render inputs are registered handles/documents or borrowed
content, never projected data or detached payloads. No experimental subpath,
stability tier or opaque public snapshot is introduced.

XLSX `read()` returns isolated cached cells/ranges, not styles, conditional paint,
drawings or track dimensions. PPTX `read()` returns nonempty trimmed source-order
paragraphs and optional speaker notes; chart words occur at the chart's element
position, shared layout/master graphics are excluded, and search neither parses
nor searches notes. DOCX sections project page setup in inches; headings, search
and navigation retain their semantic contracts. Rich-content/custom-model
rendering and detached persistence are not supported public capabilities.

Compatible additions do not require a breaking release; incompatible changes
follow the versioning policy below. Declaration dependency checks and compiled
consumer tests cover constructors, methods, events and borrowed content, not
merely barrel names. XLSX fill/border fidelity work remains OPEN independently.

For published releases, 0.x breaking changes are minor releases with migration
notes; versions are immutable (fix via a new version/deprecation). Each release
matches its annotated tag. `1.0.0` needs frozen API/model decisions and closed
release gates. Update API/README/current readiness, CHANGELOG and any changed
rationale in Notes for each release, under AGENTS' release rules. The first
release also removes the root README's "Not published yet" install note; the
package README, Getting started and the home install card already read as
published and go live only with the release.

## Unscheduled proposals

These are **undecided, not delivered APIs or commitments**. Keep related ideas
visible without presenting a wishlist as engine behaviour.

| Topic | Proposal / unresolved boundary |
|---|---|
| Bundle/memory budgets | Only for demonstrated consumer need: per-format React entries (today one lazy react entry; its own weight not separately measured), element registry variant, one Blob per shared PPTX picture, deferred evaluator loading, detached layout/master skeleton (§39). Measured sizes remain in BENCHMARKS. |
| Progressive DOCX pagination | Measure and paginate in slices and publish pages as they are decided, so first paint costs the visible pages, not the document. It would make `pageCount`, `pages`, `PAGEREF` and `NUMPAGES` provisional, and DOM windowing was set aside (§37). Only worth deciding for a consumer file where the linear cost matters. |
| Browser print | viewer.print through the browser dialog/Save as PDF, one slide/range/page per printed page from DOM/SVG; no exportPdf engine. Not designed beyond that job. |
| XLSX content size | A supported contentSize operation may help custom host sizing; units remain undecided (inches or CSS px at scale 1). Track dimensions are private; fit-width/fit-page zoom is the current supported way to fit a sheet, not summing tracks from public data. |
| Viewer decorations | Display-only range annotations/overlays; revisit with a real consumer need, not editing. |
| Preview sheet selection by title | A custom sheet selector currently uses viewer.goTo(index); only the manual sheet render handle accepts goToSheet(title). Consider title-based preview selection against overlap with shared unit navigation, not as an already-delivered control. |
| Whole-workbook search | Data-only search currently loops over sheet.search(); viewer Find stays sheet-local. A workbook operation would need explicit hidden-sheet scope, cancellation and first-search cost. Cached values/formatted text remain the boundary, not formula recalculation. |
| React start-value naming | index/zoom are initial values, unlike reactive appearance. Consider clearer naming only for demonstrated confusion; controlled props would require synchronization/loop rules and are not approved by this naming question. |
| Sparse cell access | A convenience helper may be useful if repeated consumer tasks show sparse RowData iteration is cumbersome. First compare a helper over existing SheetData; do not expose rendering styles or reshape the data contract for hypothetical convenience. |
| Picture enumeration / retained assets | Public data does not enumerate picture identifiers or provide detached rendering payloads. A gallery or asset snapshot needs a demonstrated job and a complete media/copying/lifetime/version design; current painting models are not a shortcut to that API (§135). |
| Browser-free thumbnails | Image output for agent pipelines: output format, DOM-free scope and fidelity contract undecided. DOM-free DOCX pagination is a possible foundation, not proof of a browser-free renderer. Do not promise it. |
| CSV/TSV | May or may not land as one sheet through XLSX. Needs non-ZIP routing/name/type, positive detection confirmed by content (never HTML error/legacy XLS), BOM/strict UTF-8/windows-1256 encoding probes, RFC4180/delimiter/sep= parsing and sheet-direction choice. info.format=xlsx and verbatim values are preferences, not decisions; alternatives are csv and Excel inference. Arabic fixture requires sign-off; native comparison must record Excel regional settings. |

To adopt an idea, first define the complete normal-use capability and its proof,
then add only the actual open commitment to the inventory and record why.
Consumer recipes belong in Getting started, contracts in the API reference,
implementation rules here, changes in CHANGELOG and engineering reasoning in
Notes. The former consumer-friction audit was distilled into these owning
documents (§132), not replaced by a second inventory.
