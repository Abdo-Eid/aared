# Roadmap

> Queued sequencing. For the shape each milestone builds toward, see
> [ARCHITECTURE.md](ARCHITECTURE.md) and [API.md](API.md). For why
> the order is PPTX → XLSX → DOCX, see
> [DECISION_LOG.md §5](DECISION_LOG.md#5-build-order-pptx--xlsx--docx-not-alphabetical-or-easiest-first).

This is the open-work tracker — it changes often. Treat it as current status,
not a spec; the canonical docs (`API.md`, `ARCHITECTURE.md`) are the spec.

## M0 — Repo scaffold

- [x] Design complete, distilled into `docs/{ARCHITECTURE,API,DECISION_LOG,COMPETITORS}.md`
- [x] `LICENSE` (Apache-2.0), `NOTICE`
- [x] Root `AGENTS.md`, `HANDOFF.md`, `README.md`
- [x] Bun workspace scaffold: `packages/aared`, `packages/aared-fonts`, typed
      placeholder modules that compile but don't render anything yet
- [x] `fixtures/` seeded with real generated PPTX/XLSX/DOCX/PDF files
- [x] `bench/` — visual open-speed harness (built early, ahead of M1:
  baseline renderers, sample hot-buttons, per-open ms, aared compare column)
- [ ] CI (`.github/workflows/ci.yml`): typecheck + unit tests
- [x] Verified workspace bootstrap: `bun install` → `bun run typecheck` →
      `bun test` → `bun run build` all pass against the initial scaffold

## M1 — PPTX core + renderer

- [x] `core`: lazy OPC reader, Worker RPC (`open`/`unit`/`media`/cancel),
      DrawingML shape/text-body parsing, theme resolution, and basic chart →
      SVG. The shared public route dispatches PPTX, XLSX, and DOCX while
      preserving the unit-lazy boundary.
- [x] `pptx`: shape-to-`<div>` renderer per
      [ARCHITECTURE.md — the DrawingML rendering decision](ARCHITECTURE.md)
      (inline `<svg>` geometry + HTML text body, CSS `transform` for rotation)
      - Current route covers real-DOM text/tables, preset geometry, solid and
        linear-gradient paint, transforms, pictures, charts, notes, autofit,
        and resolved table styling. Native PowerPoint exports remain the text
        fidelity reference; LibreOffice output is a companion artifact only.
      - Deliberately unsupported until their parser/renderer paths land:
        custom/unknown geometry uses a visible marked fallback; path gradients,
        pattern/image fills, embedded fonts, and animations
        do not render. Do not turn these into silent approximations.
- [x] Main-thread font resolver integration and metric-compatible CDN delivery.
- [x] `bench/` runs the public Worker-to-renderer route for PPTX, XLSX, and
  DOCX in its aared column beside baseline renderers
- [x] **Fixture requirement, non-negotiable:** at least one Arabic deck with a
      table, passing visual regression, before this milestone is considered
      done — this is the exact failure mode that motivated the project (see
      [DECISION_LOG.md §1](DECISION_LOG.md#1-why-this-project-exists-at-all))
- [x] `docs-parity` test: every public package export named in `API.md`, with
      documented callable examples verified against implementations
- [x] **API redesign implemented** (see
      [`docs/archive/aared-api-redesign.md`](docs/archive/aared-api-redesign.md)):
      `open()` moved to format entrypoints (`aared/pptx`, `aared/xlsx`,
      `aared/docx`), `aared/core` is shared types/errors only.
      Format-specific documents (`PptxDocument`, `XlsxDocument`,
      `DocxDocument`) with collection grammar (`slides.get`/`sheets.get`/
      `sections.list`/`headings.list`), typed errors (`AaredError` with 10
      codes), `MediaCollection`, `Appearance`/dark mode, `ZoomValue`,
      `<aared-viewer>` element with `theme` attribute, and comprehensive
      API reference (now `site/src/content/docs/api.mdx`).

## M2 — XLSX

- [x] `xlsx`: `styles.xml` reader (fills/fonts/borders/number formats),
      virtualized DOM grid, merges as spanning cells, frozen panes as sticky
      regions, and sheet-level RTL.
- [x] Chart rendering reused from the M1 chart-to-SVG module (shared by
      design, see [ARCHITECTURE.md](ARCHITECTURE.md#package-layout))
- [x] Fixture requirement: RTL sheet with merges and a frozen pane
- [x] `tools/excel-render/`: native Excel A4 print artifacts (hardcoded
      portrait fit-to-page with headings, rasterized with PyMuPDF), paired
      with aared A4 fits in the XLSX capture visual test
- [x] `headers` renderer option: sticky Excel-style row/column headings
      that skip hidden rows/columns; context-ordered cells render
      `dir="auto"` so LTR values keep visual order in RTL sheets
- [x] Fixture requirement, second half: `arabic-xlsx-rtl-coverage.xlsx` —
      RTL grid plus LTR summary plus deep Shot Log and Notes assembled in
      Excel from Excel-owned sheets, with the reserved
      `gray125` fills slot intact and a text-selector manifest
      (`arabic-xlsx-rtl-coverage.json`) so tests locate sheets by authored
      text, never by index

## M3 — DOCX

- [x] `docx`: flow-layout HTML renderer, paragraph-level RTL, heading outline,
      tables, lists, images, footnotes, and headers/footers.
- [x] `docx` layout mode (`pagination: { mode: "layout" }`): section
      containers with real `pgSz`/`pgMar` geometry, heuristic overflow
      pagination, per-page headers/footers/footnotes, evaluated
      `PAGE`/`NUMPAGES`. The gate fixture paginates to the four structural
      pages of the Word reference (see
      [DECISION_LOG.md §10](DECISION_LOG.md#10-docx-layout-mode-structural-pagination-without-word-parity)).
- [x] Explicit non-goal reaffirmed in the README/API:
      pixel-identical Word pagination is not attempted (see
      [DECISION_LOG.md §7](DECISION_LOG.md#7-docx-pagination-is-explicitly-not-promised))
- [x] Fixture requirement: a bidirectional document (Arabic + Latin runs in
      the same paragraph)

## M4 — `aared-fonts` publish + CDN default

- [x] Metric-compatible substitutions (Carlito, Liberation family, Noto
      Naskh/Sans Arabic, Amiri) mapped through the CDN-backed `defaultResolver`.
- [ ] Embedded-PPTX-font extraction (`ppt/fonts/*.fntdata` de-obfuscation)
      reviewed against real decks that carry embedded fonts

## Completion sequence

- [x] **API redesign** (format entrypoints, typed errors, collections,
      dark mode, element rewrite, docs parity test) — fully implemented
- [x] **Audit conformance** (see `api-redesign-audit.md`): doc falsities
      fixed (README, column-virtualization claim, CSS variables), link
      safety closed (element event forwarding, DOCX auxclick, PPTX
      hyperlink scope documented), element vocabulary migrated (`index`/
      `indexchange`, `count`), new headless tests for search semantics,
      controller APIs, dark-mode provenance, and capabilities.
- [ ] **Paired visual-inspection pass (next):** review the real fixtures in
      `bench/` and the available native Office/browser artifacts together.
      For each discrepancy, trace the source OOXML, resolved Worker model,
      renderer DOM/SVG, and Office output before changing code. Record the root
      cause, affected formats, deliberate non-changes, and remaining limits in
      the handoff; do not treat screenshot differences as isolated CSS patches.
- [ ] **Final polish and optimization pass:** resolve the confirmed visual and
      behavioral defects from inspection, remove inconsistency across the three
      routes, and make only measured performance improvements. Finish with the
      full build/typecheck/unit/visual gate, current documentation, and a clean
      package surface ready for pre-release completion.

## Known fidelity gaps (not blocking pre-release)

These are deliberate or known limitations. Do not treat them as bugs unless
the audit explicitly marks them as missing from the contract.

| Format | Gap | Status |
|---|---|---|
| PPTX | Slide backgrounds (`p:cSld/p:bg`) hardcoded white — authored dark/gradient backgrounds render white | Deferred |
| PPTX | Grouped shapes (`grpSp` with nested `sp`) silently dropped — no visible fallback marker | Deferred |
| PPTX | Synthetic slide titles (`"Slide N"`) — no title-placeholder extraction from slide layout | Deferred |
| PPTX | `notes: true` capability overclaims — notes are parsed but nothing renders them | Deferred |
| PPTX | Embedded fonts silently deferred — no unsupported marker (unlike geometry/fills) | Deferred |
| XLSX | `fit-width`/`fit-page` zoom accepted but only clears the CSS zoom — no fit logic | Deferred |
| XLSX | `rotation` missing from cell style model | Deferred |
| XLSX | First `read()` is unbounded — full sheet parsed then sliced; scroll-driven re-reads deferred | Documented |
| DOCX | Adaptive dark mode always-dark — no `matchMedia`/`"system"` resolution; diverges from XLSX | Deferred |
| DOCX | Paginator `warnings` computed then silently dropped — never reach the embedder | Deferred |
| DOCX | Search excludes table and footnote text — paragraph + heading + list only | Documented |
| DOCX | Table ARIA missing `role="row"` children — non-conforming (PPTX too) | Deferred |
| DOCX | Element renders flow-only — no layout-mode attribute | Documented |
| Cross | "system" theme has no `change` listener — OS flips apply only incidentally | Deferred |

## Design decisions (settled)

| Decision | Rationale |
|---|---|
| `aared/preview` facade (shipped; supersedes two deferrals below) | One-call auto-detect embed (`preview()` + `detect()` + `PreviewHandle`) extracted from the element's lifecycle, not forked. One-format consumers unaffected: separate subpath, runtime `import()` loads only the detected format chunk. See DECISION_LOG §12. |
| ~~No `preview()` convenience API~~ (superseded) | Was: `open()` + `render()` + `<aared-viewer>` cover the simple path. Now: bench's 205-line adapter proved the simple path wasn't simple; superseded by `aared/preview`, scoped to open → render → navigate → destroy. |
| ~~No all-format auto-detect entrypoint~~ (superseded) | Was: deferred to avoid forcing one-format consumers to pay for all formats. Now: that rationale holds via the `aared/preview` subpath (root `.` stays `aared/core`); detection runs in the Worker via `WorkerSession.detectFormat`. |
| No viewer decorations API in v1 | Decorations (`decorations: [{ range, kind }]`) are a nice-to-have for rich integrations but not part of the core previewer contract. Defer to a follow-up. |
| No tab RTL reversal | Sheet tabs intentionally stay in LTR visual order even for RTL sheets — flipping tab order on direction change would be disorienting. |
| CSS variables minimal expansion | Added `--aared-border`, `--aared-muted-text`, `--aared-accent`, `--aared-focus-ring` to renderers. Deeper theming uses headless API. |
| `::part(toolbar)` alias | Added as alias on controls div. Sheet tabs/row headers/column headers are inside format-specific DOM, not shadow parts. |

## Not yet scheduled

- Viewer decorations API (`decorations: [{ range, kind }]`) — viewer-only
  annotations overlaying rendered content; deferred from v1, revisit when a
  real consumer needs it
- Publishing to npm (name/scope already reserved — see
  [DECISION_LOG.md](DECISION_LOG.md) and `AGENTS.md`)
- VitePress or similar docs-site build (plain Markdown is enough for now)
- Any editing capability — explicitly out of scope, see
  [../DECISION_LOG.md §4](../DECISION_LOG.md#4-view-only-permanently--not-a-phased-in-decision)
