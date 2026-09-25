# Roadmap

> Queued sequencing. For the shape each milestone builds toward, see
> [ARCHITECTURE.md](ARCHITECTURE.md) and [API.md](API.md). For why
> the order is PPTX → XLSX → DOCX, see
> [DECISION_LOG.md §5](DECISION_LOG.md#5-build-order-pptx--xlsx--docx-not-alphabetical-or-easiest-first).

This is the open-work tracker — it changes often. Treat it as current status,
not a spec; the canonical docs (`API.md`, `ARCHITECTURE.md`) are the spec.

## Shipped (M0–M3 + API redesign)

- Scaffold, design docs, seeded fixtures, `bench/` harness; verified
  `install → typecheck → test → build`.
- Worker-parsed, DOM/SVG-painted renderers for PPTX, XLSX, and DOCX, each
  with its Arabic fixture (deck with table, RTL sheet with merges and
  frozen pane, bidi document) and visual regression; native
  PowerPoint/Excel/Word review pipelines under `tools/*-render/`.
- API surface: format entrypoints with typed errors and collection grammar,
  `aared/preview` facade, `aared/react` hooks, `<aared-viewer>` element,
  `docs-parity` test, and the shared-code pass (one home per rule).
- PPTX master/layout surface (DECISION_LOG §17) and source-faithful theme
  everywhere (DECISION_LOG §25): dark `theme` drives chrome only.
- Grouped shapes (`grpSp`) with child-space mapping, pinned by real-fixture
  parser and visual tests.
- DrawingML fills on shapes, table cells and slide backgrounds through one
  painter: pattern, picture (stretched, cropped, tiled), rect-path and
  centered radial gradients, with PowerPoint's linear-light stop blending;
  unpaintable fills are marked in dev builds, never dropped silently
  (DECISION_LOG §28, `pptx-coverage.pptx` slides 7–10).
- Development diagnostics behind one build flag (`src/diagnostics.ts`):
  PPTX dashed fallback outlines and the XLSX sheet notice bar paint from
  source only, never in shipped bundles.
- Deploy (`.github/workflows/deploy.yml`: push to `main` → build packages
  → build site → GitHub Pages). The package build must run first because
  the site imports `aared`/`aared-fonts` through their `dist/` exports map,
  with no source fallback.
- Real PPTX and DOCX zoom fit: both formats measure their host and apply a
  real scale (`fit-width`/`fit-page`), instead of leaving the page at its
  authored size or relying on a consumer's own CSS `zoom` trick that a
  percentage-width box silently ignores (DECISION_LOG §29).
- PPTX tables and text lines matched to PowerPoint exports: rows grow to
  their text, default cell insets, the full table-style region surface
  (every region, six edges each, fonts, diagonals, dashes, rotated text),
  boundary-centred edges, line boxes sized by their runs, and PowerPoint's
  1.2× single line (DECISION_LOG §30).
- `PreviewHandle.docx` (`aared/preview`, `aared/react`): DOCX's page count,
  laid-out page sizes, and heading outline, plus `goToHeading()` navigation
  — real API instead of a consumer scraping the painted DOM for what the
  parsed document and render handle already know (DECISION_LOG §29).
- One API grammar across formats and layers (DECISION_LOG §31): units vs
  places (`goTo` vs `goToCell`/`goToHeading`/`goToPage`), `{ index, title }`
  unit references and `{ address, row, col }` cell locations everywhere,
  async `goTo…` and `search()`, change-only events, a `"closed"` error code,
  zoom owned by the facade across navigation, `PreviewHandle.xlsx`
  controls, and DOCX `goToPage` in layout mode.
- Real XLSX zoom: opens at 100% like Excel; `fit-width`/`fit-page` fit the
  used area, +/− scale the whole grid (DECISION_LOG §31).
- True benchmark numbers (DECISION_LOG §33): memory read from the page and
  each Worker in its own DevTools session, with first-open, after-full-read
  and idle-Worker readings; download size measured as what the browser
  fetches from one production build, dependencies and versions included.
- Memory and size work against those numbers (DECISION_LOG §34–§39):
  Latin-1 shipped JS, a DOCX Worker that keeps nothing the page holds
  (streamed `w:body`), renderer defaults as classes in one adopted
  stylesheet, a lighter XLSX grid, lazily painted DOCX layout pages, Worker
  reuse between opens, bounded Worker unit caches, and minified dist.
- Viewers no longer jump when a unit is replaced (DECISION_LOG §40): no
  scroll to the top on Next, no fill that appears a frame late or grows row
  by row; the playground's React, JS and full-control demos show the same
  slide size and paging bar as the element.

## M4 — `aared-fonts` publish + CDN default

- [x] Metric-compatible substitutions (Carlito, Liberation family, Noto
      Naskh/Sans Arabic, Amiri) mapped through the CDN-backed `defaultResolver`.
- [ ] Cloud-only Office faces (Aptos, the Office default theme font since
      2023): PowerPoint fetches them, browsers lack them, so text paints in a
      generic substitute — Latin table text in `pptx-coverage.pptx` measures
      ~12% larger than PowerPoint's (DECISION_LOG §30).
- [ ] Embedded-PPTX-font extraction (`ppt/fonts/*.fntdata` de-obfuscation)
      reviewed against real decks that carry embedded fonts
- [ ] Embedded-DOCX-font extraction (`word/fonts/*.odttf`, obfuscated with the
      `w:fontKey` GUID; faces are named by `w:embedRegular`/`Bold`/`Italic`/
      `BoldItalic` in `word/fontTable.xml` and resolved through
      `word/_rels/fontTable.xml.rels`), preferred over the substitution table
      when present, reviewed against real documents that carry embedded fonts
- [ ] Embedded-XLSX-font check: Excel has no embed-fonts option (Word and
      PowerPoint do), so this is expected to be not applicable. Confirm against
      real workbooks and record the answer, rather than assume it.

Polish, bench measurement, and visual inspection are continuous practice
during development (see AGENTS.md Root-cause discipline), not sequenced
phases, so they are not tracked here. The gate ladder in HANDOFF.md records
what is proven; the N-run bench table stays TBD there until generated from
JSON — no hand-filled numbers. No CI is wanted — deliberately.

## Known fidelity gaps (not blocking pre-release)

These are deliberate or known limitations. Do not treat them as bugs unless
the audit explicitly marks them as missing from the contract.

| Format | Gap | Status |
|---|---|---|
| PPTX | Off-center `path="circle"` and `path="shape"` gradients are not painted (shapes, cells, WordArt and backgrounds alike): a dashed outline in dev builds, `data-aared-fill-fallback` / `data-aared-background="unsupported"`. COM cannot author either, so both wait on a hand-made slide appended to `pptx-coverage.pptx`, plus a probe. Everything else paints through one fill painter, checked against PowerPoint renders (DECISION_LOG §28): solid, linear, centered radial and `rect`-path gradients (linear-light, smoothstep-eased stops), the 54 preset patterns, and pictures stretched, cropped or tiled. That covers shapes, table cells and slide backgrounds (slide → layout → master, `bgRef` theme styles included) | Documented |
| PPTX | Fill approximations: patterns in group children anchor to their child-space position rather than the slide origin; picture tiles assume 96 DPI (images with other DPI metadata tile differently in PowerPoint); gradient stops follow a fitted smoothstep, within a few colour levels of PowerPoint | Documented |
| PPTX | Tables: built-in styles referenced by GUID but missing from `tableStyles.xml` paint unstyled (PowerPoint regenerates them from the theme); `lnRef` edge widths assume the default theme's 0.5/1/1.5pt because `a:lnStyleLst` is not read (DECISION_LOG §30) | Deferred |
| PPTX | Charts: fixed-template drawing (no chart model) — value axis always 0 to the data maximum in four steps; tick marks, authored text sizes, legend position, `gapWidth`, number formats and manual layout ignored; stacked/percent bars draw clustered, combo charts keep one type; area, scatter, radar and chartex show a placeholder | Deferred — quick fix vs. rebuild undecided |
| PPTX | A run tagged `en-US` with Arabic words around a neutral (`عربي + إنجليزي`, `post-production-review.pptx` slide 5) keeps its words left-to-right in PowerPoint; aared applies the Unicode bidi algorithm and reverses them | Deferred — needs a ground-truth probe |
| PPTX | Text boxes and table cells with no run size take `p:defaultTextStyle`'s size per level; which of it and the master's `p:otherStyle` PowerPoint uses when they disagree is unprobed (both fixtures agree) | Deferred — needs a fixture where they differ |
| PPTX | Layout/slide colour-map overrides (`p:clrMapOvr`) are not read, only the master's `p:clrMap`; a layout or slide using a dark background style may resolve `bg1`/`tx1` against the wrong map | Deferred — needs an edited fixture to reproduce |
| PPTX | Layout/master graphics: charts are not painted, and the `showMasterSp` rule (slide flag hides layout and master graphics; layout flag hides the master's) is unverified against PowerPoint output; a picture background *inherited* from a layout or master is unexercised by the fixtures (slide 9 covers a slide's own tiled picture background) | Deferred — needs an edited fixture to verify |
| PPTX | Only `slidenum` and `datetime` `a:fld` fields are live; other field types keep their cached text | Documented |
| PPTX | Synthetic slide titles (`"Slide N"`) in `info.titles` and `slide.info` — kept by design: extracting the title placeholder would parse every slide at `open()` (DECISION_LOG §31); `slide.read()` has the real text | Documented |
| PPTX | Notes are parsed but nothing renders them (the former `capabilities.presentation.notes: true` overclaimed this; `capabilities` was cut — DECISION_LOG §15) | Deferred |
| PPTX | Embedded fonts silently deferred — no unsupported marker (unlike geometry/fills) | Deferred |
| XLSX | The sheet notice bar (formulas never recalculated, unsupported features) paints in development builds only; shipped builds show the sheet alone, and a host that wants to tell users reads `SheetModel.unsupported` itself | Documented |
| XLSX | `rotation` missing from cell style model | Deferred |
| XLSX | Border styles reach the renderer as raw Excel names (mapped in `xlsx/index.ts`), where DOCX resolves its borders in core. Resolving them in core changes the public `XlsxCellStyleModel`, so it waits on the model-contract decision below | Deferred |
| XLSX | First `read()` is unbounded — full sheet parsed then sliced; scroll-driven re-reads deferred | Documented |
| XLSX | Embedded fonts: Excel has no embed-fonts feature, so none are expected; unconfirmed against real workbooks (see M4) | Deferred — confirm |
| DOCX | Paginator `warnings` computed then silently dropped — never reach the embedder | Deferred |
| DOCX | Search excludes headers, footers, and text boxes (text boxes are not parsed) — headings, paragraphs, lists, table cells, and footnotes/endnotes are covered | Documented |
| DOCX | Table ARIA missing `role="row"` children — non-conforming (PPTX too) | Deferred |
| DOCX | Element renders flow-only — no layout-mode attribute | Documented |
| DOCX | Embedded fonts (`word/fonts/*.odttf`) are not read — text falls back to the substitution table or system fonts; no unsupported marker | Deferred |
| Cross | `system` theme: `<aared-viewer>` listens for OS flips; `preview()`/hooks rely on the host re-forwarding appearance | Deferred |

## Before first publish (open decisions)

Decisions that cost nothing to change now and a major version to change
later. AGENTS.md "Pre-release breaking changes" expires at first publish, so
resolve each of these first. Add to this list as new ones surface; delete an
entry (and record the outcome in DECISION_LOG) when it is settled.

The name is taken: `@aared/viewer@0.0.0` is a placeholder (DECISION_LOG §42,
git tag `v0.0.0`). What remains is the first real release at `0.1.0`: rename
the packages to `@aared/viewer` / `@aared/fonts`, drop `private`, pack-check,
maintainer publishes, tag `v0.1.0`. HANDOFF "Publishing" has the checklist.

### Are the resolved model types a public contract?

**Status: open. Not decided.**

`read()` returns the parser's resolved models, and `aared/core` exports about
60 types describing them (`ShapeModel`, `TextRunModel`, `TableModel`,
`SheetRangeModel`, `CellModel`, `DocxParagraphNode`, …). The API reference
documents them under "Resolved models". Exporting and documenting them makes
the parser's internal output a public contract: after publish, renaming a
field, splitting a node type, changing a unit, or making a field required
breaks every consumer that reads the models. The parser is still moving (see
the fidelity gaps above: two unpainted gradient kinds, heuristic DOCX
layout), so a promise made now is one you keep paying for. Changed since
this entry was written: `SlideLayoutModel` was added (shared **by
reference** across slides, so treat it as read-only), and the fill work
(DECISION_LOG §28) replaced `ImageBackgroundModel` with `PictureFillModel`
and turned `TableModel.cellFills` from colour strings into fills — exactly
the kind of reshaping a stability promise would have forbidden.

Why the models are public at all: `render()` accepts a handle *or* a model,
and `read()` is the seam the renderers themselves consume. That justifies
`read()` existing. It does not by itself justify a stability promise on every
model shape.

Options:

1. **Stable public API.** Keep every model type exported and documented as
   stable; any change after publish is a semver-major with a migration note.
2. **Tiered (leaning this way).** Keep the XLSX range types
   (`SheetRangeModel`, `CellModel`, `CellRange`) stable — `sheet.read("A1:M200")`
   is a headline use and the only place data reading is a natural need. Mark
   the PPTX and DOCX shape/node models **experimental** (shape may change in
   a minor version). `render()` keeps accepting a model; the shape simply is
   not promised.
3. **Internal.** Stop exporting the model types; `read()` returns opaque or
   loosely typed values. Smallest surface, but drops typed custom rendering
   and the XLSX data use.

Rationale for the lean: nobody has a strong reason to write their own PPTX
renderer against `ShapeModel`, while reading cell values is a plausible need.
Confirm by checking what real consumers use before publishing (playground,
bench, and any early adopters).

When settled: update `site/src/content/docs/api.mdx` § Resolved models (state
the stability tier per format, and label experimental types), trim
`aared/core` exports if going with option 3, update `docs-parity.test.ts`
accordingly, and record the outcome in DECISION_LOG.

### Bundle size

Status: decided — single self-contained worker (41 KB gzip minified for all
three formats; per-format workers were re-checked with measured numbers and
rejected again — DECISION_LOG §20). Main-thread code is already lazy per
format at runtime. Measured against the other libraries (download size,
BENCHMARKS.md): smallest for XLSX and for all three formats; about 8 KB
behind docx-preview + jszip for DOCX; well behind mdurbar for PPTX, whose
lightness comes from dropping most preset shapes and RTL. Remaining risk:
Worker URL resolution under Vite/webpack (HANDOFF.md). Deferred unless a
real consumer size budget demands it: per-format `preview`/`react` subpaths
and an element registry variant.

Memory follow-ups (weighed, not scheduled — DECISION_LOG §39): one Blob per
shared PPTX picture; the PPTX-only geometry evaluator's `import()` out of
the shared session chunk; a detached layout/master skeleton instead of the
full trees.

### Search follow-ups (observed, not scheduled)

- No Arabic-aware matching anywhere (tashkeel, alef/hamza, ya/alef-maqsura,
  digit forms): search is a plain case-insensitive substring.
- PPTX search skips notes; grouped-shape text is covered (groups recurse).
- A Worker-side text-only `search` message would cut the cold first search
  (see DECISION_LOG §16).

## Not yet scheduled

- Viewer decorations API (`decorations: [{ range, kind }]`) — viewer-only
  annotations overlaying rendered content; deferred from v1, revisit when a
  real consumer needs it
- Search-text highlighting — headless `search()` already returns model ranges;
  what's missing is painting them as a non-mutating overlay (ranges in,
  highlight layer out, source text untouched). Small job; rides on the
  decorations overlay when that lands, or ships standalone before it
- Headless thumbnail export (nice to have) — render a slide/page/sheet to an
  image without a browser, for agent pipelines that need a fast high-fidelity
  glance instead of a full view. Open design, not scheduled: output format
  (SVG serialize vs raster), DOM-free requirement, per-format scope
  (slides/pages/sheets), and fidelity-vs-speed contract all undecided.
  Do not promise until the mechanism is proven. For DOCX the DOM-free
  paginator's estimator (`core/docx/pagination.ts`) is the natural fit: it is
  the only page layout that runs with no browser, so a headless render takes
  its page breaks from it. That is why the estimator stays even after the
  browser renderer moves to DOM-measured pagination (DECISION_LOG §21) — and
  why its accuracy then needs a real bar of its own instead of today's
  constants fitted to one fixture (for example measured font metrics injected
  into it)
- Shape-fidelity batch — diagnosed Sep 2026 from slide 3 of
  `pptx-coverage.pptx` vs PowerPoint and aiden0z: (a) shipped — the full ECMA-376 preset catalog (176 presets, 16 of them proven symmetry aliases, in one eager table shipped only in the `aared/pptx` entry, oracle-checked endpoint-by-endpoint against LibreOffice's independent conversion) resolves through the deferred-geometry engine (DECISION_LOG §26 and its addendum); the 11 hand-coded fast-path presets stay worker-side; (b) shipped — unsupported geometry renders an empty
  dashed outline (like picture/chart fallbacks) instead of a filled rect
  that reads as a wrong shape; (c) shipped — custGeom path commands (moveTo/lnTo/arcTo/quadBezTo/cubicBezTo/close, guide-computed coordinates included) evaluate through the same engine; (d) `grpSp` rendering — shipped:
  groups parse to `GroupModel` and render nested with child-space mapping
  (pinned by `pptx-group.test.ts` + the group visual test; group-level
  effects and `grpFill` inheritance stay unresolved); (e) shipped — the
  `line` preset was invisible because a sub-pixel authored axis rounded the
  SVG viewport to zero (fixed with a 1px-minimum viewport + viewBox growth
  that keeps 96px per inch), and `ellipse` painted zero-area (fixed with two
  half-arcs). Golden path-string unit tests pin the fast-path formulas; the
  oracle snapshot pins the engine; DOM re-read tests cover marking,
  visibility, and layer painting
- Chart fidelity (PPTX, and XLSX through the same `core/chart.ts`) —
  diagnosed Sep 2026 from `post-production-review.pptx` slide 4 vs
  PowerPoint. Today the chart XML turns straight into drawing primitives on a
  fixed 1000×600 box: the value axis always runs 0 to the data maximum in
  four steps (0/10.5/21/31.5/42 where PowerPoint draws 0–45 by 5); tick
  marks, `txPr` text sizes, legend position, `gapWidth`/`overlap`, `showVal`
  and number formats, `autoTitleDeleted` and manual `c:layout` are ignored;
  stacked and percent bars draw clustered; a combo chart keeps one type;
  doughnut draws as a pie; area, scatter, radar, stock and chartex show the
  dashed placeholder. Two paths, not yet chosen:
  (a) quick fix, about a day, keeping the current structure — nice-number
  axis scaling, tick marks, authored text sizes in points, centred legend,
  `gapWidth`, `showVal` + `numFmt`, `autoTitleDeleted`; brings slide 4 close;
  (b) rebuild, multi-week — a semantic chart model in core (plot groups,
  axes with scaling/format/ticks/gridlines, data labels, legend, text and
  theme/chart-style colours), Excel-compatible auto axis scaling and number
  formatting (reusing the XLSX number-format code), a layout engine sized to
  the real frame (legend placement, axis-label reservation, manual
  `c:layout`), then more types step by step (stacked/percent, area, scatter,
  true doughnut, radar, combo with secondary axes, trendlines, error bars)
  and a picture fallback for chartex from `mc:AlternateContent`
- Publishing to npm (name/scope already reserved — see
  [DECISION_LOG.md](DECISION_LOG.md) and `AGENTS.md`)
- VitePress or similar docs-site build (plain Markdown is enough for now)
- Any editing capability — explicitly out of scope, see
  [../DECISION_LOG.md §4](../DECISION_LOG.md#4-view-only-permanently--not-a-phased-in-decision)
