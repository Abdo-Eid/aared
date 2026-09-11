# Competitive Landscape

> Who else renders Office documents client-side, what they actually cover,
> and where aared differs. For why aared is shaped the way it is, see
> [DECISION_LOG.md](DECISION_LOG.md) and [ARCHITECTURE.md](ARCHITECTURE.md).
> Public summary table: [../README.md](../README.md#compared-to-other-libraries).

**The recurring pattern:** single-format libraries exist and are individually
reasonable (`docx-preview`, ExcelJS), but every one that tries to *paint*
rather than build a real DOM gives up paragraph-level RTL correctness to do
it — and nothing spans all three formats at aared's target weight. The
all-in-one options that do exist are full office suites, not previewers.

## At a glance

| Library | Formats | Client-side | Renders via | Paragraph-level RTL | License | Weight |
|---|---|---|---|---|---|---|
| **aared** (this) | PPTX, XLSX, DOCX | ✅ | DOM/SVG | ✅ (design goal — see [DECISION_LOG.md §3](DECISION_LOG.md#3-text-is-rendered-by-the-browser-never-rasterized--the-load-bearing-decision)) | Apache-2.0 | Targeting small — per-format lazy load |
| [`docx-preview`](https://github.com/VolodymyrBaydalka/docxpreview) | DOCX | ✅ | DOM | ✅ (real DOM, CSS-fixable) | MIT | Small |
| [`mammoth`](https://github.com/mwilliamson/mammoth.js) | DOCX | ✅ | DOM (semantic HTML, not visual-fidelity) | ✅ | BSD-2 | Small |
| [`pptx-preview`](https://www.npmjs.com/package/pptx-preview) | PPTX | ✅ | Canvas | ❌ (canvas, no CSS fix) | MIT | Small |
| [`pptxviewjs`](https://www.npmjs.com/package/pptxviewjs) | PPTX | ✅ | Canvas | ❌ — drops paragraph `algn`/`rtl` entirely; see [DECISION_LOG.md §1](DECISION_LOG.md#1-why-this-project-exists-at-all) | MIT | Medium (bundles Chart.js unconditionally) |
| [SheetJS Community Edition](https://sheetjs.com/) | XLSX (parse only, no renderer) | ✅ | — you build the grid | N/A — parses `sheet.rightToLeft`, but returns empty `cell.s` even with `cellStyles: true` | Apache-2.0 | Small |
| [ExcelJS](https://github.com/exceljs/exceljs) | XLSX (parse/write, no renderer) | ✅ | — you build the grid | N/A — does parse `styles.xml` correctly, useful as a reference implementation | MIT | Medium |
| [Univer](https://github.com/dream-num/univer) | PPTX*/XLSX/DOCX-ish (via Docs) | ✅ | Canvas (own engine) | Partial, improving | Apache-2.0 | Large — full editable-office-suite runtime |
| [LibreOffice-WASM (allotropia)](https://github.com/allotropia/lowasm) | Everything LibreOffice supports | ✅ | Own renderer, compiled to WASM | ✅ (real LO engine) | MPL-2.0 | Very large (~300 MB) — the fidelity ceiling, not a practical dependency |
| [`pdf.js`](https://github.com/mozilla/pdf.js) | PDF only | ✅ | DOM + Canvas hybrid, real text layer | N/A (PDF glyphs are pre-positioned) | Apache-2.0 | Medium — the reason PDF is out of aared's scope, see [DECISION_LOG.md §2](DECISION_LOG.md#2-pdf-was-scoped-out-deliberately) |

## Reading the table

- **`docx-preview` / `mammoth`** prove the DOM-first approach works for one
  format. Neither spans PPTX or XLSX. aared's DOCX tier targets
  `docx-preview`'s level of fidelity (see
  [ARCHITECTURE.md](ARCHITECTURE.md#fidelity-contract)) but as one renderer
  inside a shared `core`, not a standalone library.
- **`pptx-preview` / `pptxviewjs`** prove canvas rendering is *fast* and
  *simple to ship* — and prove exactly why that trade fails for RTL content.
  This is the evidence base the whole project sits on.
- **SheetJS / ExcelJS** are parsers, not renderers — neither one paints a
  grid. aared's `xlsx` module needs its own `styles.xml` reader regardless
  of which parser (if either) it builds on, because SheetJS CE's styling gap
  is exactly the thing that needs fixing.
- **Univer / LibreOffice-WASM** prove full fidelity is *achievable* — at the
  cost of being an entire application runtime rather than an embeddable
  previewer. aared is deliberately not competing on total fidelity; it's
  competing on being small enough to drop into someone else's chat UI or
  document panel, which neither of these is built for.
- **`pdf.js`** isn't a competitor so much as a neighbor — aared is designed
  to be paired with it, not to replace it.
