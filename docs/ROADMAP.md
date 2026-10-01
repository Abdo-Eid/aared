# Roadmap

> Open work only. For how the engine works, see
> [ARCHITECTURE.md](ARCHITECTURE.md); for what landed and why, the numbered
> entries in [DECISION_LOG.md](DECISION_LOG.md); for the build order, §5
> there.

Keep it current: finished work leaves this file (the DECISION_LOG entry is
its record), and a difference settled as accepted or by design moves to
[ARCHITECTURE.md § Known limits](ARCHITECTURE.md#known-limits).

**Status (Sep 2026): the API is complete; fidelity and pre-publish gates
remain.** All three formats ship through their entrypoints (`aared/pptx`,
`aared/xlsx`, `aared/docx`: `open()` → `slides`/`sheets`/`sections`/`headings`
→ `read()` → `render()` → `destroy()` → `close()`, typed `AaredError`,
`MediaCollection`, theme-only `Appearance`), and through the embed layers
(`aared/preview`, `aared/react`, `<aared-viewer>`). Content is
source-faithful in every theme; unsupported geometry and fills paint marked
fallbacks, never silent approximations. DOCX layout mode paginates from
DOM-measured blocks with Word's own numbers: close to Word, not a Word-exact
promise.

Gates (AGENTS.md § Testing rule): typecheck, build, unit tests,
`docs-parity`, the PPTX/XLSX/DOCX DOM and visual tests, and `bench-measure`
all pass. One is open: **PPTX visual review sign-off** (same-size PowerPoint
and aared captures, the Arabic deck with a table included), a `1.0.0` gate.
No CI is wanted, deliberately; `bench-measure` runs rarely (`--calibrate`
first on an idle machine, about 30 minutes a pass).

## Next up

In order:

1. **Pre-publish:** settle the resolved-model contract tier ("Before first
   publish" below), then the first release (checklist there).
2. **PPTX visual review sign-off.**
3. **DOCX picture placement:** a top-and-bottom picture after a paragraph
   whose picture still hangs, and the lowest rows beside a through-wrapped
   picture (limits table); both visible, both need a Word probe of the rule.
4. **DOCX residuals:** block images and footnotes still estimate their
   pagination height; page-3 drift (per-font single-line metrics, block-image
   spacing); promote the anchor-comparison script into `tools/word-render`.
5. **Charts, in isolation** (the maintainer's): size and performance first;
   the quick-fix route is under "Not yet scheduled".
6. **Final polish pass:** confirmed defects only, then the full gate and docs.

## Open gaps (not blocking pre-release)

Known differences still meant to be closed. Differences settled as accepted
or by design are not here: they are in
[ARCHITECTURE.md § Known limits](ARCHITECTURE.md#known-limits).

| Format | Gap | Status |
|---|---|---|
| PPTX | Charts: fixed-template drawing (no chart model) — value axis always 0 to the data maximum in four steps; tick marks, authored text sizes, legend position, `gapWidth`, number formats and manual layout ignored; stacked/percent bars draw clustered, combo charts keep one type; area, scatter, radar and chartex show a placeholder. Charts on layouts and masters are not painted | Deferred — the maintainer's, in isolation |
| PPTX | WordArt preset warps (`a:prstTxWarp`) paint the text flat, marked `data-aared-text-warp`; the text is all there. The four curve presets (`textArchUp`, `textArchDown`, `textCircle`, `textButton`) could follow an SVG `textPath` as real text (a PowerPoint probe for the arc, then the painter); the other presets reshape the letters and have no DOM form. Low priority (Sep 2026): visible when it occurs, but warped WordArt is rare in current decks and unseen in Arabic ones; PPTX only, since DOCX does not parse text boxes | Deferred — pick up when a real deck needs it |
| PPTX | Master and layout graphics (a master logo, `showMasterSp="0"` on a layout or slide, a picture background) match PowerPoint on a probe (§54) but no fixture carries them, so the suite does not pin them: three `test.todo` in `pptx-layout.test.ts` | Open — needs master content added to a coverage deck |
| XLSX | `rotation` missing from cell style model | Deferred |
| XLSX | Border styles reach the renderer as raw Excel names (mapped in `xlsx/index.ts`), where DOCX resolves its borders in core. Moving them changes the public `XlsxCellStyleModel`, so it waits on the model-contract decision | Deferred |
| XLSX | First `read()` is unbounded — full sheet parsed then sliced; scroll-driven re-reads deferred | Deferred |
| DOCX | Layout pagination measures top-level blocks and table rows in the DOM (§21, §44); block images and footnotes still estimate. Table rows never split, and header rows (`w:tblHeader`) do not repeat | Open — estimates in "Next up" |
| DOCX | Anchored pictures (§55): a wrapped picture positioned vertically from the page or a margin starts at its paragraph's top; paragraphs after a picture with text on both sides do not wrap around what of it remains (they start below it); a top-and-bottom picture right after a paragraph whose picture still hangs is pushed below that picture (about 14px on `docx-coverage.docx` page 5); with a through wrap Word keeps a wider hole on the lowest rows than the polygon at the line's height, so it wraps about a word more there (probe case 8); VML pictures (`w:pict`) have no placement | Open — the band and through-wrap rows are in "Next up" |
| DOCX | No line grid (`w:docGrid w:type="lines"`): aared counts a line as 12pt, right only with the grid off (§47) | Deferred |
| DOCX | Text boxes are not parsed (their text is lost); search also excludes headers and footers | Open — not scheduled |
| DOCX | Table ARIA missing `role="row"` children — non-conforming (PPTX too) | Deferred |
| DOCX | Element renders flow-only — no layout-mode attribute | Deferred |
| Cross | `system` theme: `<aared-viewer>` listens for OS flips; `preview()`/hooks rely on the host re-forwarding appearance | Deferred |

## Before first publish (open decisions)

Decisions that cost nothing to change now and a major version to change
later. AGENTS.md "Pre-release breaking changes" expires at first publish, so
resolve each first. Delete an entry (and record the outcome in DECISION_LOG)
when it is settled.

The name is taken: `@aared/viewer@0.0.0` on npm is a name-holding placeholder
(README only; org `aared`, owner `mego0eid`; §42, git tag `v0.0.0`). The
repo's packages are still named `aared` / `aared-fonts`, `private: true`,
`0.0.0`; nothing of the real code is public.

### Release checklist (`0.1.0`), in one piece of work

1. Settle the open decision below.
2. Rename `aared` to `@aared/viewer`. `aared-fonts` is frozen and stays
   unpublished (§53); it keeps its workspace name, and its `aared/core`
   import renames with the rest. Imports become `@aared/viewer/pptx`,
   `/xlsx`, `/docx`, `/preview`, `/react`, `/element`, `/worker`. Touch the
   `package.json` name and `exports`, every import in `bench/`, `site/`,
   `tools/` and the tests, `tsdown.config.ts`,
   `tools/check-consumer-build.ts`, the API reference page (docs-parity), the
   README install block, and the site's install card and playground
   snippets. The `<aared-viewer>` tag name does not change.
3. Drop `private` and add `publishConfig.access: "public"` (the `files`
   list, `["dist"]`, is already there); run `bun run build`, the unit and
   visual tests,
   `bun tools/check-consumer-build.ts` and `npm pack --dry-run` (the tarball
   has `dist/` and no fixtures or sources you did not intend).
4. The maintainer runs `npm publish` in their own terminal (2FA on; no token
   is ever given to an agent), then `git tag -a v0.1.0` on the release
   commit. What changes after it is in AGENTS.md "Pre-release breaking
   changes".

### Are the resolved model types a public contract?

**Status: open. Not decided.**

`read()` returns the parser's resolved models, and `aared/core` exports about
60 types describing them (`ShapeModel`, `TextRunModel`, `TableModel`,
`SheetRangeModel`, `CellModel`, `DocxParagraphNode`, …), documented under
"Resolved models" in the API reference. That makes the parser's output a
public contract: after publish, renaming a field, splitting a node type,
changing a unit, or making a field required breaks every consumer that reads
the models. The parser is still moving — the fill work replaced
`ImageBackgroundModel` with `PictureFillModel` (§28), and DOCX pictures traded
`float` for `anchor` and `floats` (§55) — exactly the reshaping a stability
promise would forbid.

`render()` accepts a handle *or* a model, and `read()` is the seam the
renderers themselves consume. That justifies `read()` existing, not a
stability promise on every model shape.

Options:

1. **Stable public API.** Every model type exported and documented as stable;
   any change after publish is a semver-major with a migration note.
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

## Follow-ups (observed, not scheduled)

### Bundle and memory

Deferred unless a real consumer size budget demands it: per-format
`preview`/`react` subpaths and an element registry variant. Memory follow-ups (weighed, not scheduled —
§39): one Blob per shared PPTX picture; the PPTX-only geometry evaluator's
`import()` out of the shared session chunk; a detached layout/master skeleton
instead of the full trees.

### Search

- No Arabic-aware matching anywhere (tashkeel, alef/hamza, ya/alef-maqsura,
  digit forms): search is a plain case-insensitive substring.
- PPTX search skips notes; grouped-shape text is covered (groups recurse).
- A Worker-side text-only `search` message would cut the cold first search
  (§16).

### XLSX content size

- No single content-size getter: a consumer that wants a sheet's natural
  width and height must add up `columnWidths`, `defaultRowHeight` and each
  row's own `height` from the read model. A `contentSize()` (inches, or CSS
  px at scale 1) on the read model or the render handle would remove that
  sum.
- The numbers exclude the viewer's own chrome (row-number column, column
  letters, sheet tabs), so a caller sizing a box to the sheet has to add the
  chrome by hand. `fit-width` and `fit-page` already account for it.

## Not yet scheduled

- Viewer decorations API (`decorations: [{ range, kind }]`) — viewer-only
  annotations overlaying rendered content; revisit when a real consumer
  needs it
- Search-text highlighting — headless `search()` already returns model ranges;
  what's missing is painting them as a non-mutating overlay (ranges in,
  highlight layer out, source text untouched). Small job; rides on the
  decorations overlay when that lands, or ships standalone before it
- Headless thumbnail export (nice to have) — render a slide/page/sheet to an
  image without a browser, for agent pipelines that need a fast high-fidelity
  glance instead of a full view. Open design: output format (SVG serialize vs
  raster), DOM-free requirement, per-format scope, and fidelity-vs-speed
  contract all undecided. Do not promise until the mechanism is proven. For
  DOCX the DOM-free paginator's estimator (`core/docx/pagination.ts`) is the
  natural fit: it is the only page layout that runs with no browser, which is
  why it stays now that the browser renderer measures in the DOM (§21) — and
  why its accuracy needs a real bar of its own instead of constants fitted to
  one fixture
- PPTX group residuals — group-level effects and `grpFill` inheritance are
  unresolved
- Chart fidelity (PPTX, and XLSX through the same `core/chart.ts`) —
  diagnosed Sep 2026 from `post-production-review.pptx` slide 4 vs
  PowerPoint. Today the chart XML turns straight into drawing primitives on a
  fixed 1000×600 box: the value axis always runs 0 to the data maximum in
  four steps (0/10.5/21/31.5/42 where PowerPoint draws 0–45 by 5); tick
  marks, `txPr` text sizes, legend position, `gapWidth`/`overlap`, `showVal`
  and number formats, `autoTitleDeleted` and manual `c:layout` are ignored;
  stacked and percent bars draw clustered; a combo chart keeps one type;
  doughnut draws as a pie; area, scatter, radar, stock and chartex show the
  dashed placeholder. The maintainer takes this in isolation (Sep 2026), with
  size and performance as first-class goals: the chart code is shared with
  XLSX (`core/chart.ts`, `src/chart-svg.ts`), so any change is checked against
  both formats' captures, and it should not grow the Worker: check what
  `core/chart.ts` adds to the shared session chunk (the geometry table set the
  precedent of keeping bulk out of it, §26).
  Two paths:
  (a) quick fix, about a day, keeping the current structure. The list, each
  item small and independent, in the order that moves slide 4 the most:
  1. nice-number value-axis scaling (round the data max up to a 1/2/5 step,
     five or so gridlines; PowerPoint draws 0 to 45 by 5 where aared draws
     0 to 42 in fours), and honour `c:scaling` `min`/`max`/`majorUnit` when
     the file states them;
  2. tick marks and axis line from `c:majorTickMark`, and `c:delete` for an
     axis the author hid;
  3. authored text sizes in points from each `c:txPr`, not the fixed ones;
  4. legend position from `c:legendPos` (right, top, bottom, left) with the
     plot area making room, centred on its side;
  5. `c:gapWidth` and `c:overlap` for bar widths;
  6. data labels (`showVal`, `showPercent`) with `c:numFmt`, reusing the XLSX
     number-format code so a `0.0%` or `#,##0` reads as Excel prints it;
  7. `c:autoTitleDeleted`, and a single-series chart's title from the series
     name when the title is automatic;
  8. stacked and percent-stacked bars (`c:grouping`) and a real doughnut hole
     (`c:holeSize`), the two shapes a reader sees wrong most often.
  Leave placeholders for area, scatter, radar, stock and chartex; leave manual
  `c:layout` ignored. Each item gets a DOM test that re-reads the painted
  SVG (`chart-svg.ts`), the coverage deck's charts as the fixture, and the
  paired PowerPoint capture of `post-production-review.pptx` slide 4 as the
  review; the `unsupported` placeholder keeps its marker so nothing paints
  silently wrong;
  (b) rebuild, multi-week — a semantic chart model in core (plot groups,
  axes with scaling/format/ticks/gridlines, data labels, legend, text and
  theme/chart-style colours), Excel-compatible auto axis scaling and number
  formatting (reusing the XLSX number-format code), a layout engine sized to
  the real frame (legend placement, axis-label reservation, manual
  `c:layout`), then more types step by step (stacked/percent, area, scatter,
  true doughnut, radar, combo with secondary axes, trendlines, error bars)
  and a picture fallback for chartex from `mc:AlternateContent`
- CSV/TSV preview — considered Sep 2026, may or may not land. Open a
  `.csv`/`.tsv` as a one-sheet workbook through `aared/xlsx`, the way Excel
  does. The renderer side is nearly free: a parsed CSV is a plain
  `SheetModel` (default 8.43-character columns, default row height, no
  styles) that the XLSX renderer already paints, overflow, frozen panes and
  dark mode included. The cost is all on the input side:
  (a) routing — the Worker's `detect` and `open` both start with
  `OpcPackage.open`, so non-ZIP bytes fail today as `invalid-document`, and
  `sourceBytes` drops `File.name`, `Blob.type` and the URL before the Worker
  sees anything; the element, `preview()` and `aared/xlsx`'s `open()` would
  carry that hint through, and a CSV Worker session has no package behind it
  (the media route assumes one);
  (b) detection needs a positive signal (`.csv`/`.tsv`, `text/csv`,
  `text/tab-separated-values`, or an explicit type) confirmed by a content
  check — bytes alone are unsafe: a 200-status HTML login page or a JSON body
  is valid text, and an encrypted `.xlsx` or a legacy `.xls` (OLE,
  `D0 CF 11 E0`) is not a ZIP either; none of them may paint as a one-column
  sheet;
  (c) encoding — byte-order mark, then strict UTF-8, then `windows-1256`
  (Bun's `TextDecoder` decodes it headless, so `core` stays testable without
  a browser). This is the Arabic win: classic Excel reads a BOM-less UTF-8
  CSV as the system code page and garbles it — confirm on current Excel 365
  before claiming that anywhere public;
  (d) parsing — RFC 4180 quoting (delimiters and newlines inside quoted
  fields); delimiter auto-detected among `,` `;` tab `|`, honouring a `sep=`
  first line (Excel's own convention; otherwise Excel uses the locale's list
  separator); cut at the `isInsideGrid` bounds with a visible `unsupported`
  entry; the tab name is the file name without its extension;
  (e) sheet direction — a CSV carries no right-to-left signal (Excel follows
  its own "default direction" option), so a content heuristic or a render
  option; any Arabic-range test goes through `core/chars.ts` (shipped JS
  stays Latin-1).
  Open decisions: `info.format` stays `"xlsx"` (preferred — a one-sheet
  workbook like Excel, no new entrypoint, events or API-page entries) or a
  new `"csv"` `DocumentFormat`; cell values shown verbatim, with number
  detection used only for General alignment (preferred — Excel's inference
  drops leading zeros, turns long IDs into `1.23457E+11` and some text into
  dates) or Excel-style inference; and the fixture — rule 4 wants an Arabic
  fixture with a table, and a CSV cannot live inside the existing coverage
  workbooks, so it would be a new file in `fixtures/` and needs sign-off.
  Ground truth: `tools/excel-render/` can open the CSV in Excel for the
  paired review, but Excel's result depends on that machine's list
  separator, ANSI code page and default direction, so a capture only counts
  as an oracle when it records those regional settings
