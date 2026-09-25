# Competitive Landscape

> Who else renders Office documents client-side, what they actually cover,
> and where aared differs. For why aared is shaped the way it is, see
> [DECISION_LOG.md](DECISION_LOG.md) and [ARCHITECTURE.md](ARCHITECTURE.md).
> Public summary table: [../README.md](../README.md#compared-to-other-libraries).

**The recurring pattern:** single-format libraries exist and are individually
reasonable (`docx-preview`, ExcelJS), but every one that tries to *paint*
rather than build a real DOM gives up paragraph-level RTL correctness to do
it — and nothing spans all three formats at aared's target weight with DOM/SVG
output. The all-in-one options that do exist are full office suites, WASM
Canvas engines, or commercial SDKs — not lightweight previewers.
"RTL verified" below means a specific public implementation claim was found;
"not verified" does not mean unsupported, only that no affirmative evidence
was found. Sizes are given only where published; otherwise "not published".

## At a glance (closest comparators)

| Library | Formats | Client-side | Renders via | Paragraph-level RTL | License | Weight |
|---|---|---|---|---|---|---|
| **aared** (this) | PPTX, XLSX, DOCX | ✅ | DOM/SVG | ✅ (design goal — see [DECISION_LOG.md §3](DECISION_LOG.md#3-text-is-rendered-by-the-browser-never-rasterized--the-load-bearing-decision)) | Apache-2.0 | Targeting small — per-format lazy load |
| [`@silurus/ooxml`](https://github.com/yukiyokotani/office-open-xml-viewer) | PPTX, XLSX, DOCX | ✅ | Canvas 2D (Rust→WASM parsers) | Verified fixes (DOCX RTL-number, UAX#9 mixed-text; Arabic-layout test work) | MIT | Published per-format: DOCX ~4.0 MB / 1.2 MB gzip; XLSX ~2.6 / 0.82 MB; PPTX ~2.5 / 0.78 MB |
| [`@opentf/office-docs-viewer`](https://github.com/Open-Tech-Foundation/office-docs-viewer) | DOCX today; PPTX/XLSX slot in as engine modules | ✅ | Canvas (`ImageBitmap` LRU, DPR-correct) + Worker host | Not verified | Check repo (engine Rust→WASM + TS core/react) | Not published |
| [`omni-doc-viewer`](https://github.com/akbhuker/omni-doc-viewer) | DOCX/XLSX/PPTX (+PDF/image/md/csv) | ✅ | Per-format engines (docx-preview, SheetJS, pptx-preview+emf-converter) behind one API, lazy-loaded | Inherits per-engine limits | Permissive (wraps permissively-licensed engines) | Small core, engines lazy-loaded |
| [`docx-preview`](https://github.com/VolodymyrBaydalka/docxpreview) | DOCX | ✅ | DOM | ✅ (`w:bidi`/`w:rtl` → `direction:rtl`; real DOM, CSS-fixable) | Apache-2.0 | Small |
| [`mammoth`](https://github.com/mwilliamson/mammoth.js) | DOCX | ✅ | DOM (semantic HTML, not visual-fidelity) | Known gap — [no default `dir=rtl`](https://github.com/mwilliamson/mammoth.js/issues/336) | BSD-2 | Small |
| [`@aiden0z/pptx-renderer`](https://github.com/aiden0z/pptx-renderer) | PPTX | ✅ | DOM (HTML/SVG, selectable text; charts via ECharts canvas) | Minimal, verified: parses `a:pPr rtl` → CSS `direction` + `textAlign`; no bidi algorithm or table-cell `dir` — browser-bidi-reliant | Apache-2.0 | Measured entry ~774 KB / 169 KB gzip (standalone ESM ~1783 KB bundles JSZip+ECharts) — see [BENCHMARKS.md](BENCHMARKS.md) |
| [`pptx-preview`](https://www.npmjs.com/package/pptx-preview) | PPTX | ✅ | DOM (front-end viewer from ArrayBuffer) | Not verified | npm says ISC, but README has restrictive source-use language — legal review needed | Not published (ECharts, JSZip, Lodash) |
| [`pptxviewjs`](https://www.npmjs.com/package/pptxviewjs) | PPTX | ✅ | Canvas | ❌ — drops paragraph `algn`/`rtl` entirely; see [DECISION_LOG.md §1](DECISION_LOG.md#1-why-this-project-exists-at-all) | MIT | Medium (bundles Chart.js unconditionally) |
| [`@datagridxl/spreadsheet-preview`](https://github.com/DataGridXL/spreadsheet-preview) | XLSX | ✅ | Grid (self-described "PDF.js for spreadsheets") | Not verified | Community badge / commercial to remove | Not published |
| [`@xlscalc/xlsx-preview`](https://github.com/Victor-EU/xls-renderer-calc) | XLSX | ✅ (+Worker path) | DOM grid (React `ExcelView` / `renderToHtml`), recomputes formulas | Not verified | MIT | Not published |
| [SheetJS Community Edition](https://sheetjs.com/) | XLSX (parse only, no renderer) | ✅ | — `sheet_to_html` gives a plain HTML table, not a grid renderer | N/A — parses `sheet.rightToLeft`, but returns empty `cell.s` even with `cellStyles: true` | Apache-2.0 | Small |
| [ExcelJS](https://github.com/exceljs/exceljs) | XLSX (parse/write, no renderer) | ✅ | — you build the grid | N/A — does parse `styles.xml` correctly, useful as a reference implementation | MIT | Medium |
| [Univer](https://github.com/dream-num/univer) | PPTX*/XLSX/DOCX-ish (via Docs) | ✅ (XLSX import path can require a conversion backend) | End-to-end bidi work documented ([PR #7011](https://github.com/dream-num/univer/pull/7011)) — re-test against release | Partial, improving | Apache-2.0 | Large — full editable-office-suite runtime |
| [SpreadJS](https://developer.mescius.com/spreadjs/docs/tutorials/how-to-import-and-export-excel-xlsx-using-javascript) | XLSX/XLSM/XLTX/XLTM/CSV | ✅ (client import/export via IO module) | Verified: worksheet RTL orientation + Arabic/Hebrew cell direction (charts/shapes/comments/slicers limited) | View/edit | Commercial (developer + deployment-hostname) | Not published |
| [Apryse WebViewer](https://docs.apryse.com/web/guides/office) | DOCX, XLSX, PPTX (+legacy) | ✅ (true client-side WASM; server optional) | Verified: Arabic/Farsi/Hebrew/Urdu localization + RTL layout/view/edit | View/annotate (+paid DOCX/sheet editing) | Commercial | Not public |
| [Nutrient Web SDK](https://www.nutrient.io/guides/web/viewer/office-documents/) (ex-PSPDFKit) | DOC/DOCX, XLS/XLSX, PPT/PPTX | ✅ (browser Office→PDF via WASM, no Office/server needed) | Not verified | View/annotate; editing applies to converted PDF | Commercial | Not public |
| [LibreOffice-WASM (allotropia)](https://github.com/allotropia/lowasm) / [ZetaJS](https://github.com/allotropia/zetajs) | Everything LibreOffice supports | ✅ | Own renderer, compiled to WASM | ✅ (real LO engine) | MPL-2.0 | Very large (~300 MB) — the fidelity ceiling, not a practical dependency |
| [`pdf.js`](https://github.com/mozilla/pdf.js) | PDF only | ✅ | DOM + Canvas hybrid, real text layer | N/A (PDF glyphs are pre-positioned) | Apache-2.0 | Medium — the reason PDF is out of aared's scope, see [DECISION_LOG.md §2](DECISION_LOG.md#2-pdf-was-scoped-out-deliberately) |

## Measured snapshot

Qualitative table above; numbers live in
[BENCHMARKS.md](BENCHMARKS.md) (generated from `tools/bench-measure.ts`
JSON — never hand-filled). Until the first full pass runs, treat every
weight claim ("small", "~300 MB", per-format MB figures) as vendor-stated or
unpublished, not as aared-measured.

## Direct three-format local renderers (closest to aared)

- **`@silurus/ooxml` / `office-open-xml-viewer`** — the single closest
  comparator: identical DOCX+XLSX+PPTX scope, fully local, read-only like
  aared. Opposite output choice: Rust/WASM parsers + Canvas 2D (with
  worker/`OffscreenCanvas` option and a transparent-DOM selection overlay as
  a stop-gap) instead of DOM/SVG. Has the strongest public payload benchmark
  (see table). Treat as the Canvas reference implementation to beat on
  accessibility/selectability/weight.
- **Verified in WebView (Sep 2026)** against
  `fixtures/pptx-coverage.pptx` (then 6 slides): all 6 slides paint — RTL
  shaping, table column order, shapes, gradients, and the embedded picture
  are all correct at a glance — and none of the text is selectable or
  reachable by assistive tech. That is the DOM-vs-canvas gap stated as a
  measured fact, not a design slogan: same fixture, their pixels vs our
  DOM. One suspect, unconfirmed: slide-5 RTL bullets paint on the left;
  needs the paired PowerPoint capture before calling it a failure.
- **`@opentf/office-docs-viewer`** — newer Rust→WASM engine (OPC container,
  display model, virtualization) + framework-agnostic TS core + React
  binding + Worker host. DOCX implemented; PPTX/XLSX are engine modules that
  slot in. Canvas paint with page virtualization and `ImageBitmap` cache.
- **`shrimp-doc-viewer`** — fully local Rust→WASM display-list/CPU-raster in
  a Worker, `ImageBitmap`/Canvas. Covers DOCX/XLSX/PPTX (+PDF/ODF). CJK
  advertised; RTL not verified. ~2 MB WASM core; bundled default CJK font
  makes install ~12 MB (lazy-loaded).
- **`Oxi`** (MPL-2.0 core) and **`A3S Office`** — Rust/WASM-rooted suites
  covering all three formats with view *and* edit/round-trip. A3S claims RTL
  overrides/complex-script handling (independently unverified). Both are
  editor-weight, not previewer-weight.
- **`omni-doc-viewer`** — one API + format detection (extension → MIME →
  zip-magic sniffing `word/`/`ppt/`/`xl/`) + per-format lazy load over proven
  engines: PPTX via `pptx-preview` (DOM output, `pages` re-sorted to
  `presentation.xml` order, DOM text search provider, EMF/WMF rasterized via
  cached `emf-converter`), DOCX via docx-preview, XLSX via the `@e965/xlsx`
  SheetJS fork for parsing plus its own virtualized DOM table
  (`createVirtualTable`) for display. 100% client-side,
  offline, no keys. PPTX fidelity is explicitly "readable preview", and the
  PPTX path has zero RTL handling (no `rtl`/`bidi`/`direction` in its chunks).
  Legacy `.doc`/`.ppt` throw `UnsupportedFormatError`.

## DOCX-only

- **`docx-preview` / `docxjs`** proves the DOM-first approach works for one
  format and handles `w:bidi`/`w:rtl`. Section fidelity gap, verified
  Sep 2026 on `docx-coverage.docx` (A4 portrait + A4 landscape
  sections): its `breakPages` pagination flows the landscape section onto
  portrait-width pages (2 pages, all 794px — the 1122px landscape page
  never surfaces), while aared paginates it as a true landscape page.
  aared's DOCX tier targets this level
  of fidelity (see [ARCHITECTURE.md](ARCHITECTURE.md#fidelity-contract)) but
  as one renderer inside a shared `core`, not a standalone library.
- **`docx-renderer`** (fork of `docx-preview-sync` → `docx-preview`, for
  Recon) — paginated HTML DOM, headers/footers/footnotes, nested tables,
  anchored drawings, tracked changes, overlay system. No longer synced
  upstream. Apache-2.0, depends on JSZip/Lodash/Konva.
- **`docxodus`** (+ `react-docxodus-viewer`) — DOCX→HTML in WASM, no server;
  keeps justification, style inheritance, numbering, tables, comments,
  headers/footers, footnotes; paginated mode; block editor + `save()` and
  doc-compare/redlines in the same engine.
- **`SuperDoc`**, **EigenPal `docx-editor`**, **`docx-html-renderer`** —
  browser OOXML→DOM/paginated editors or converters. Editing-capable, so
  heavier than aared's view-only scope; useful as fidelity references.
- **`mammoth`** converts to semantic HTML, not Word-page fidelity, with the
  RTL gap linked above.

## PPTX-only

- **`@aiden0z/pptx-renderer`** (Apache-2.0) — browser-only OOXML→HTML/SVG
  DOM with selectable text (charts via ECharts canvas), the closest PPTX-side
  proof that a DOM renderer is viable. Zip limits, lazy media/slides with
  windowed list rendering, model-based text search with overlay highlights,
  and EMF fallbacks (embedded-PDF→pdf.js or bitmap→PNG; WMF has no vector
  render). RTL is minimal but verified: `a:pPr rtl` → CSS `direction` +
  `textAlign`, no bidi algorithm or table-cell `dir` — aared's paragraph +
  table + cell `dir` pinning and bidi fixtures still differentiate.
- **`mdurbar/pptx-viewer`** — lightweight TS parser→SVG slides, one
  dependency (`fflate` ~8 KB gzip). Weight floor only — **not a fidelity
  reference.** Verified broken in WebView (Sep 2026) on
  `pptx-coverage.pptx`: its relationship walk puts *slides* into
  the `slideMasters` map (keyed `ppt/slides/slideN.xml`) and leaves
  `slideLayouts` empty, so every slide falls back to the first "master" and
  slide 2's table paints on all 6 slides — silently, no error, correct slide
  count. A renderer that fails loud is a bug; one that renders the wrong
  slide content without telling you is disqualified as an oracle.
  Why it downloads so little (read from its 0.2.2 `dist/pptx-viewer.js`,
  Sep 2026): its shape table maps 94 of Office's ~187 preset names onto
  56 drawings of its own and turns every other preset into a rectangle
  (lines ~1893–2000); it has no RTL or bidi handling; it parses every part
  up front on the main thread with `DOMParser`. aared ships the full preset
  catalog (`src/core/presets.ts`, about 20 KB gzip on its own).
- **`PPTXjs`** (mesheshа) — local JS/JSZip→HTML `div`s + D3/NVD3 charts +
  optional Reveal.js mode. DOM output but legacy jQuery-era stack. RTL not
  verified.
- **`pptx-browser`** — zero-dependency Canvas renderer (native
  `DecompressionStream` + `DOMParser` + Canvas 2D). Interesting weight data
  point; Canvas, so same RTL objection as `pptxviewjs`.
- **`ChristopherVR/pptx-viewer`** — full TS monorepo (core + React/Vue/
  Angular/Svelte/vanilla) parsing, rendering as HTML/CSS/SVG, editing,
  presenting, collaboration (Yjs), export. Editor-suite weight, not
  previewer weight.
- **`pptx-preview` / `pptxviewjs`** prove canvas rendering is *fast* and
  *simple to ship* — and prove exactly why that trade fails for RTL content.
  This is the evidence base the whole project sits on. Note the
  `pptx-preview` license caveat in the table.
- Smaller/older converters: **`PPTX2HTML`**, **`pptx-react-renderer`**,
  **`pptx-parser`** (parser→HTML strings, not a viewer shell),
  **`pptx-wasm`**, **`SlideGlance`** (Rust/WASM→SVG/PNG). Mostly
  single-format, incomplete-fidelity, RTL unverified.

## XLSX-only

- **SheetJS / ExcelJS** are parsers, not renderers — neither one paints a
  grid. aared's `xlsx` module needs its own `styles.xml` reader regardless
  of which parser (if either) it builds on, because SheetJS CE's styling gap
  is exactly the thing that needs fixing. (`sheet_to_html` is a plain table
  dump, not a grid.)
- **Grids that need a parser:** `FortuneSheet` (MIT, React spreadsheet UI),
  `x-data-spreadsheet` / `x-spreadsheet` (MIT, ~1.2 MB unpacked),
  `Handsontable` (commercial/source-available since 7.x; MIT ended at
  6.2.2) — all view/edit grids, XLSX normally via SheetJS, RTL unverified.
- **View-only previews (most comparable):** `@datagridxl/
  spreadsheet-preview` ("PDF.js for spreadsheets": virtual scrolling,
  multi-sheet, frozen panes, images, cached formula values) — verified
  Sep 2026 (0.1.0) on `xlsx-coverage.xlsx`: sheet names and RTL
  order correct, but the grid is virtualized and needs an explicit `height`
  option (without it only the chrome paints — blank grid), and cells paint
  to canvas inside shadow DOM, so cell text is not in the DOM (unselectable,
  unsearchable) and scrolling smears under repaint lag — aared's DOM grid
  has neither failure mode;
  `@xlscalc/xlsx-preview` (MIT, Worker path, DOM grid, *recomputes* formulas
  in-browser, `renderToHtml` shares style composition with the React
  renderer); `@hewliyang/xlsx-preview` (Rust WASM extraction → Canvas);
  `@extend-ai/react-xlsx` (viewer + optional editing, worker parsing);
  `@nuptechs/nup-xlsx-preview` (commercial, CSS-grid renderer, virtual
  scroll, 7 themes).
- **Commercial grid:** **SpreadJS** (fully client-side, verified worksheet
  RTL + Arabic/Hebrew cell direction) and **Mescius Document Solutions JS
  Data Viewer** (read-only grid for XLSX/CSV/Arrow/Parquet — not
  DOCX/PPTX). **Aspose.Cells for JavaScript via C++** is browser WASM for
  processing/render-to-image/PDF/HTML, but no out-of-box interactive viewer.

## Commercial embeddable SDKs

- **Apryse WebViewer** and **Nutrient Web SDK** are the two true client-side
  (WASM, no Office/server) Office viewers. Both commercial, weight not
  public. Apryse has verified RTL coverage; Nutrient converts Office→PDF in
  the browser (edits/annotations apply to the converted PDF).
- **Syncfusion** (DOCX editor + spreadsheet UI; complex import paths commonly
  need service endpoints; verified RTL layout/localization for its controls),
  **GroupDocs.Viewer for Node.js** (Node+Java, server-side render to
  HTML/PDF/images — not browser-local), **TX Text Control .NET Server**
  (browser editor over WebSocket to a backend — not browser-local).

## Full suites and hosted alternatives (not local previewers)

- **Univer / LibreOffice-WASM / ZetaJS** prove full fidelity is *achievable*
  — at the cost of being an entire application runtime rather than an
  embeddable previewer. aared is deliberately not competing on total
  fidelity; it's competing on being small enough to drop into someone else's
  chat UI or document panel, which none of these is built for.
- **ONLYOFFICE Docs/DocSpace** (browser editor + Document Server doing
  conversion; AGPL/community + commercial) and **Collabora Online** (server
  renders 256×256 tile PNGs/deltas to a browser canvas over WOPI/WebSocket;
  documents never leave the server) are server-dependent by architecture.
- **Hosted iframes:** Microsoft Office Web Viewer / Office for the web and
  Google Docs Viewer fetch a *public URL* server-side and render a hosted
  experience — no local-file/privacy equivalence, no offline, not a
  substitute for a no-server renderer.
- **Wrapper components** (`@cyntler/react-doc-viewer` — unmaintained,
  Office via MS iframe + public URLs only; `@iamjariwala/react-doc-viewer`
  — maintained fork, local DOCX via docx-preview + local XLSX table, other
  Office via Office Online; `react-file-viewer-v2`) wire per-format
  renderers together but add no engine of their own.

## Reading the table

- **Three-format DOM/SVG whitespace is real:** docx-preview/docx-renderer/
  Mammoth cover DOCX; PPTXjs/PPTX2HTML/pptx-preview/pptx-react-renderer/
  mdurbar/aiden0z cover PPTX; SheetJS's `sheet_to_html` covers basic
  XLSX→table. Most are single-format, incomplete-fidelity, or lack explicit
  Arabic/bidi assurances — spanning all three with demonstrated bidi is
  still the gap.
- **Correct Arabic/bidi is a meaningful differentiator but no longer
  unique.** Direct evidence exists for docx-preview paragraph bidi,
  silurus's UAX#9 fixes, Apryse, and SpreadJS; Univer has explicit bidi work
  to acceptance-test. For every other entry, treat Arabic as a measured
  compatibility test, not an inferred feature.
- **`pdf.js`** isn't a competitor so much as a neighbor — aared is designed
  to be paired with it, not to replace it.

## Explicit non-competitors

- **Authoring/template packages** (`PptxGenJS`, `docx` npm, `docxtemplater`)
  generate/mutate files rather than render an opened document — not
  renderers.
- **ExcelJS / Handsontable / x-spreadsheet** are models or grids, not OOXML
  renderers by themselves (XLSX ingestion normally via SheetJS).
- **GroupDocs, TX Text Control, ONLYOFFICE, Collabora, Microsoft/Google
  viewers** are buying alternatives, not substitutes for a no-server,
  local-file renderer.
