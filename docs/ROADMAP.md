# Roadmap

> Open work only. For how the engine works, see
> [ARCHITECTURE.md](ARCHITECTURE.md); for what landed and why, the numbered
> entries in [DECISION_LOG.md](DECISION_LOG.md); for the build order, §5
> there.

Keep it current: finished work leaves this file (the DECISION_LOG entry is
its record), and a difference settled as accepted or by design moves to
[ARCHITECTURE.md § Known limits](ARCHITECTURE.md#known-limits).

**Status (Oct 2026): the API is complete; fidelity and pre-publish gates
remain.** All three formats ship through their entrypoints (`aared/pptx`,
`aared/xlsx`, `aared/docx`: `open()` → `slides`/`sheets`/`sections`/`headings`
→ `read()` → `render()` → `destroy()` → `close()`, typed `AaredError`,
`MediaCollection`, theme-only `Appearance`), and through the embed layers
(`aared/preview`, `aared/react`, `<aared-viewer>`). Each format entry also offers
`preview({ source, target })` with shared navigation, search and one `destroy()`;
its consumer build omits sibling renderers (§88). Content is
source-faithful in every theme; unsupported geometry and fills paint marked
fallbacks, never silent approximations. DOCX layout mode paginates from
DOM-measured blocks with Word's own numbers: close to Word, not a Word-exact
promise.

Gates (AGENTS.md § Testing rule): typecheck, build, unit tests,
`docs-parity`, the PPTX/XLSX/DOCX DOM and visual tests, and `bench-measure`
all pass. One is open: **PPTX visual review sign-off** (same-size PowerPoint
and aared captures, the Arabic deck with a table included), a `1.0.0` gate.
No CI is wanted, deliberately; `bench-measure` runs rarely, on mains power
(§60).

## Next up

In order:

1. **Pre-publish:** settle the resolved-model contract tier, then the first
   release ("Before first publish" below).
2. **PPTX visual review sign-off** (the open gate above).
3. **DOCX table styles** (first capability below): the largest difference a
   Word reader sees.
4. **DOCX, visible differences:** the numbered-heading marker, stale saved
   table-of-contents results, and the two picture-placement cases (both need a
   Word probe of the rule). Pagination residuals follow (see the DOCX gaps).

## Open gaps (not blocking pre-release)

Known differences still meant to be closed. Differences settled as accepted
or by design are not here: they are in
[ARCHITECTURE.md § Known limits](ARCHITECTURE.md#known-limits).

### DOCX

- **Numbered heading:** a numbered paragraph with a heading style takes its
  number in the sequence (§77) but its marker is not painted. Reproduced on
  `tools/word-render/out/numbered-heading-probe/word-source.docx`: Word reports
  `ListString` `1.` for its Heading 1 paragraph and paints `1. Project Overview`;
  aared paints `Project Overview`. The Word-authored list places the marker at
  the paragraph start and the heading text after the list's marker tab. The
  paired page-2 captures are in `tools/word-render/out/word-source/`. The earlier
  coverage capture was not a valid test because its numeral was typed into the
  heading text. A heading is not a list item in the model, so painting its
  computed marker needs data on the heading node. *Open; visible on numbered-
  heading documents.*
- **DOCX table-of-contents fields:** Word's page-1 export of the updated
  `numbered-heading-probe/word-source.docx` includes `5. Delivery tracker /
  متابعة التسليم` and `الإجمالي`; aared's page 1 still shows the earlier saved
   TOC result without either entry. `tools/office-export.ts --format docx` updates fields
  in Word's private read-only export copy, while aared reads the field's cached
  result from the DOCX and only recalculates live `PAGE`/`NUMPAGES` fields.
  Refresh supported TOC entries from the parsed body outline (including the
  document's TOC switches and level range), or define a faithful stale-cache
  policy. The updated Word/aared page-1 pair is in `tools/word-render/out/word-source/`.
  *Open; visible when headings change after the TOC cache was last saved.*
- **Anchored pictures (§55):** a wrapped picture positioned vertically from the
  page or a margin starts at its paragraph's top; paragraphs after a picture
  with text on both sides start below what remains of it; a top-and-bottom
  picture right after a paragraph whose picture still hangs is pushed below it
  (about 14px on `docx-coverage.docx` page 5); with a through wrap Word keeps a
  wider hole on the lowest rows than the polygon at the line's height, so it
  wraps about a word more there (probe case 8); VML pictures (`w:pict`) have no
  placement. The October 6 paired capture also shows accumulated vertical
  spacing and wrapping differences on page 5, reaching about 54px at a unique
  text anchor near the bottom; keep that review with these picture-placement
  cases. *Open; the band and through-wrap cases need a Word probe.*
- **Word comparison diagnostics:** `tools/word-render/compare-pagination.ts`
  runs on Windows, but pairs repeated tokens by sorted occurrence without
  surrounding context. Different PDF/browser tokenization can shift those
  pairs and produce false page mismatches or missing/extra text reports. Improve
  matching and expose ambiguous pairs before using the report as an automatic
  fidelity check.
  *Open tooling gap; paired images and unambiguous anchors remain the evidence.*
- **DOCX content across pages:** tables already break between rows, but a row
  and a paragraph/list block remain indivisible; long content can exceed a page.
  `w:tblHeader` rows are not repeated. Word permits row continuation by default
  (subject to row/style rules) and repeats contiguous leading header rows.
  Treat row/paragraph continuation and repeated headers as a separate,
  reader-visible capability, with a real fixture case before expanding scope.
  *Open.*
- **Text-box autofit:** DrawingML and VML text boxes use their authored frame
  and text size; modes that shrink text or grow the frame remain unimplemented.
  Parsing, browser text rendering and search are delivered (§82–84).
- **Office 2016 chart:** a chart in a DOCX reads its `cx:lvl` cache, tested on
  synthetic parts only; `docx-coverage.docx` holds none. *Open: paste one into
  the document through Word, as was done for the PPTX deck.*

### PPTX

- **WordArt preset warps** (`a:prstTxWarp`) paint the text flat, marked
  `data-aared-text-warp`; the text is all there. The four curve presets
  (`textArchUp`, `textArchDown`, `textCircle`, `textButton`) could follow an
  SVG `textPath` as real text (a PowerPoint probe for the arc, then the
  painter); the other presets reshape the letters and have no DOM form. Rare in
  current decks, unseen in Arabic ones. This gap covers PPTX WordArt.
  *Deferred: pick up when a real deck needs it.*
- **Group residuals:** group-level effects and `grpFill` inheritance are
  unresolved. *Open, not scheduled.*
- Path fill modes: see "Capabilities not built".

### XLSX

- **Conditional formats:** `top10`, `aboveAverage`, `iconSet` and `timePeriod`
  rules, and `formula` bounds are reported in the sheet's notices but not
  evaluated. Investigated, not started — read
  [XLSX conditional formats](#xlsx-conditional-formats) before starting. *Open.*
- **`rotation`** is missing from the cell style model. *Deferred.*
- **Border styles** reach the renderer as raw Excel names (mapped in
  `xlsx/index.ts`), where DOCX resolves its borders in core. Moving them
  changes the public `XlsxCellStyleModel`, so it waits on the model-contract
  decision. *Deferred.*
- **First `read()` is unbounded:** the full sheet is parsed then sliced;
  scroll-driven re-reads are deferred. *Deferred.*

### Cross-format

- **Untrusted input (§72), still unbounded:** the zip's total inflated bytes and
  compression ratio (a 1000:1 entry opens, about 630 MB), 100 MB of repeated
  entities (4 s, 2 GB), the "download from URL" path's size, and
  `gridSpan`-style loops in code not yet audited. Each is a budget or a typed
  error away; none is a small-file freeze. *Open, S each.*
- **Chart border:** an unstyled chart's border is per app (PowerPoint none,
  Excel a hairline; aared draws Excel's everywhere); a per-format default needs
  a Word probe first. *Open, small.*
- **Table ARIA** is missing `role="row"` children, non-conforming (PPTX and
  DOCX). *Deferred.*

### Tools and bench

- **Chart oracle (94 cases, §74):** the few rules still assumed (ARCHITECTURE
  Known limits, Charts) need cases of their own; the 2016 types, stock and
  of-pie are compared with Excel's exports by eye, not scored; the XLSX route
  and devices without Office's fonts are not scored; a data table for screen
  readers is not offered. *Open.*
- **Generated tables** not checked in `bun test`: `core/chart/metrics-data.ts`
  needs Python and Windows fonts (`tools/chart-oracle/gen_metrics.py`); only
  `core/presets.ts` has a drift test (§72). *Open: a `--check` mode in
  `gen_metrics.py`, run by hand.*
- **Bench:** pptx-browser, Extend and Apollo joined after the last `--all` run:
  they have download sizes but no open, navigation or memory rows. *Open:
  measure once, or leave out.*

### Capabilities not built (each its own piece of work)

Two differences are missing features, not defects. Each would be one
capability under AGENTS.md "Capability delivery": none is started, and none
lives in Known limits (they are meant to be closed, not accepted).

**DOCX table styles.** A table's `w:tblStyle` is not read, so a table that
Word draws from a style (Table Grid, Grid Table and List Table: banded rows, a
shaded or bold header row, style borders, the `w:tblLook` switches) paints
plain. The tables in `docx-coverage.docx` carry their formatting directly, so
the suite does not show it.
- *Layers:* `core/docx/styles.ts` (the table style, its `w:basedOn` chain and
  `w:tblStylePr` regions), `core/docx/document.ts` (resolve regions per cell
  under direct formatting), the model's existing cell fields (no new
  renderer). The PPTX route already does this for `a:tblStyle` (§54,
  ARCHITECTURE "Table styles"); share what two formats need.
- *Proven by:* a Word probe first (precedence of style regions against direct
  formatting; how a band counts around a merged cell or header row), then a
  table styled through Word's COM added to `docx-coverage.docx` (needs
  sign-off), a parser test, and a layout-mode comparison with
  `tools/word-render/`.

**PPTX path fill modes (`darken`, `darkenLess`, `lighten`, `lightenLess`).**
A preset's sub-path can carry these, but the geometry engine keeps only
whether a sub-path is filled and stroked (`GeometryPaths`), so the shaded faces
of a cube, can or bevel paint in the shape's flat fill. Visible on those
shapes, rare in decks.
- *Layers:* `scripts/compile-presets.ts` and `core/presets.ts` (carry the mode
  per sub-path), `core/drawingml-engine.ts`, the model's geometry result, and
  `pptx/index.ts` plus `fill-paint.ts` (a shaded copy of the fill, through the
  shared painter). Only PPTX has presets.
- *Proven by:* a PowerPoint probe of one shape per mode (Office's exact
  amounts), an engine test against it, and a paired capture.

## XLSX conditional formats

*Investigated Oct 2026, not started. The findings below are verified and worth
keeping; the choices are not made, and this item is not settled.*

The rule kinds Excel evaluates from a cell's cached value that aared reports in
`SheetModel.unsupported` instead of painting: `top10`, `aboveAverage`,
`iconSet` and `timePeriod`, and a `formula` colour-scale or data-bar bound. What
already paints is unchanged: `cellIs`, the four text rules, `duplicateValues`,
`uniqueValues`, colour scales and data bars, with `stopIfTrue` and priority
order (`core/xlsx/sheet.ts`, `EVALUATED_RULES` at `sheet.ts:432`).

**A correction to the earlier record.** The old wording of this item said
formula bounds "never will be" evaluated, on view-only grounds. That is wrong as
stated, and the distinction matters:

- **A `formula` bound naming a cell is readable, so it is not a "never".** That
  is only true of a bound containing a function or arithmetic. `<cfvo
  type="formula" val="$A$1"/>` is a bare cell reference: read that cell's cached
  `<v>`, exactly as aared already does for a cell formula, and no formula engine
  is involved. `scaleBound` (`sheet.ts:393`) currently rejects the whole
  `type="formula"` case through one `default:` arm, so today it is all-or-nothing
  by omission, not by policy.
- **`expression` has no shortcut in the file.** That is a fact about the format,
  not a decision — see the evidence below, and the cost study in "Open
  decisions" before treating it as closed.

**Evidence, already gathered.** Re-deriving any of this is wasted work:

- *The file stores no outcome.* Unzipped `fixtures/xlsx-coverage.xlsx`: a
  `conditionalFormatting` block carries the condition only, and `dxfId` names a
  **style** in the styles part, not a result. The `x14` extension, where Excel
  duplicates data-bar rules, duplicates only bounds and colours
  (`x14:cfvo`, `x14:negativeFillColor`). There is no per-cell "matched" flag
  anywhere — Excel recomputes which cells match on every load. The contrast that
  settles it: a cell formula stores both sides, `<f>SUM(F5:F16)</f><v>31.4</v>`,
  which is why aared can show formula results with no engine. So "use Excel's
  cached value" works wherever the *inputs* are the cell's own values, and there
  is nothing to read when the condition is an arbitrary formula.
- *`expression` has no shortcut in the file.* It is the one rule type where the
  `<formula>` is the entire condition, with no companion attribute carrying the
  meaning. Compare `containsText`, which works only because Excel duplicates its
  meaning into `cfRule/@text`, so `matchesText` (`sheet.ts:375`) evaluates from
  `@text` and ignores the `<formula>NOT(ISERROR(SEARCH(…)))` sitting beside it.
  There is no equivalent duplicate for `expression`, so the only way in is
  evaluating the formula itself — the open cost study in "Open decisions", not a
  settled refusal. Two constraints any such evaluator must respect: it must
  never be *wrong* (aared does not guess — a wrong colour is worse than none),
  and a bounded subset is what makes it tractable, since arbitrary workbook
  formulas reach named ranges, tables and volatile functions.
- *`CT_CfRule` (ECMA-376 18.3.1.10), attribute defaults:*
  `aboveAverage` **true**, `percent` **false**, `bottom` **false**,
  `equalAverage` **false**, `stopIfTrue` **false**; `rank` is `unsignedInt`,
  `stdDev` is `int`; up to three `<formula>` children. `top10`'s N is the
  `rank` attribute, and `percent` says whether `rank` is a percentage or a count
  (MS-OI29500) — read `rank`, fall back to `<formula>[0]` if a file omits it.
- *`ST_IconSetType`: 17 base families* (`3Arrows`, `3ArrowsGray`, `3Flags`,
  `3Signs`, `3Symbols`, `3Symbols2`, `3TrafficLights1`, `3TrafficLights2`,
  `4Arrows`, `4ArrowsGray`, `4Rating`, `4RedToBlack`, `4TrafficLights`,
  `5Arrows`, `5ArrowsGray`, `5Quarters`, `5Rating`), plus four that only the
  `x14` extension adds (`3Stars`, `3Triangles`, `5Boxes`, `NoIcons`).
- *`CT_IconSet` (18.3.1.49):* `iconSet`, `showValue`, `reverse`, `percent`
  attributes; thresholds in `<cfvo>` children. Note there are **two** unrelated
  `percent` attributes — `cfRule@percent` is `top10` only, `iconSet@percent`
  says the icon thresholds are `percent` rather than `num`.
- *`ST_TimePeriod`:* eleven values — `today`, `yesterday`, `tomorrow`,
  `last7Days`, `next7Days`, `lastWeek`, `thisWeek`, `nextWeek`, `lastMonth`,
  `thisMonth`, `nextMonth`. No `<formula>` child.
- *Date detection already exists:* `isDateFormatCode`
  (`core/xlsx/numberfmt.ts:199`), which returns per-section date flags for a
  format code. Only the serial→date conversion and the clock are missing.

**Open decisions.** Not settled; pick these when the item is picked up.

- *Icon sets: draw one set, map the rest onto it.* The thresholds are free
  either way (`scaleBound` reused unchanged), so the whole question is how many
  shapes to draw. Leading idea, from the maintainer: draw a small canonical set
  and **map every undrawn family onto the closest drawn one** rather than
  reporting it, so an icon set always paints something. That is defensible
  because the families differ in shape, not in meaning — all of them encode the
  same one-to-five tiers, so painting `3Flags` as the three symbols carries the
  same reading Excel gives it. Still open: which families form the canonical
  set, what the mapping is for each of the rest (including the four `x14`-only
  ones), and whether a mapped icon should be reported at all or silently
  substituted. If mapping turns out to mislead — a shape carrying an unintended
  association — the fallback is to report the rule instead. Whatever is drawn
  should be recognisable at a glance and checked against a paired Excel capture,
  per AGENTS.md "Good enough beats perfect". The grey arrow variants are the
  coloured geometry in grey, so they are close to free once their coloured
  counterparts are drawn. All 17 remains a legitimate option.
- *Whether `timePeriod` follows the clock.* It depends on the reader's date, so
  the same file would paint differently depending on when it is opened. Excel
  re-evaluates these rules on every load, so matching it is arguably the
  faithful choice; the counter-argument is that aared is otherwise deterministic
  for the same bytes. If it is evaluated, the clock should be injected into
  `parseXlsxSheet` so `core` stays headless and tests pin a fixed date, and the
  dependence stated in the API page and Known limits. Leaving it unsupported and
  reporting it is also defensible; so is deferring until a real workbook with a
  date-based rule shows up.
- *The formula engine: study its cost before ruling anything out.* `expression`
  rules and a `formula` bound containing a function cannot be read from the file
  (evidence above), so today they are reported and the reasoning stops at rule
  5. That is a stance, not a measurement, and rule 5 is about not
  recalculating a document — evaluating a conditional rule is arguably a
  different act. Before closing this item for good, the open work is a **cost
  study**: what a minimal evaluator would cost (bundle weight in the Worker, the
  parse surface, the maintenance of a second expression language next to the
  number-format parser), and **where it would earn its keep** — which real rule
  shapes it would actually serve, against `expression` rules in the wild and
  against formula bounds like `PERCENTILE($B$2:$B$50,0.9)`. Only then decide
  whether to build it, and if not, make the notice say *why* rather than just
  that the rule is unevaluated. `timePeriod` is a related question in miniature:
  `TODAY()` is a clock-dependent function, so it sits on the same boundary.
  Do not treat "view-only" as the last word on this without doing the study.

**What building it means.** `top10`, `aboveAverage` and `timePeriod` are one
coherent capability, plus the `formula`-bound fix: a matcher each beside
`matchesCellIs` (`sheet.ts:350`), over `numericValuesIn` (`sheet.ts:434`),
reusing the existing `differential` fill/font path, added to `EVALUATED_RULES` —
no model change, no renderer change, no new notice. The bound fix is smaller
still: resolve `val` through the existing `parseRef`
(`core/xlsx/range.ts:60`), read that cell from the grid, and give `scaleBound` a
resolver parameter. `timePeriod` adds a serial→date conversion over the
1899-12-30 epoch and a `today` parameter on `parseXlsxSheet`, skipping non-date
cells.

Icon sets are a second capability, the only one needing a model and renderer
change. The thresholds are the identical computation as a colour scale, so
`scaleBound` is reused unchanged across every family and that part is free; the
cost is drawing — a `cell.conditional.icon` field, then an SVG in
`appendCellContent` (`xlsx/index.ts:591`) following the `DATA_BAR` pattern:
absolutely positioned at the inline-start so a RTL sheet mirrors it, text lifted
with `CELL_TEXT_OVER`, honouring `showValue` and `reverse`. `isEmptyCell`
(`xlsx/index.ts:587`) must learn that an icon counts as content, or icon-only
cells would paint nothing.

**Before any of it, measure.** A ground-truth probe through Excel's COM object
model on a scratch workbook, then `tools/probe/_common/strip_label.py`. This is measurement, not
investigation: it settles the cases where the spec and Office diverge, so no
matcher is written against a guess. Open questions: **top10** — how a tie at the
cut-off is broken (which of the equal cells Excel paints), whether `percent`
rounds the count down or up, and whether Excel writes `rank` or a `<formula>`.
**aboveAverage** — which cells count toward the mean (blanks and text
excluded?), the sign handling of `stdDev`, `equalAverage`. **timePeriod** — does
`lastWeek` start on Sunday, does `last7Days` include today, where do month
boundaries fall. **iconSet** — the XML Excel actually writes, and a rendering
reference per drawn family.

**Proven by:** that probe; `xlsx-conditional.test.ts` reading real parsed cells,
plus a renderer test that re-reads the DOM for the icon overlay; each new test
confirmed to **fail** with its feature switched off, as the theme work did (§81);
`bun run typecheck`, `bun test`, `bun run build` with the Latin-1 `dist` test,
and the xlsx visual suites; and a paired Excel capture for the new fixture and
the drawn icon families.

**The fixture gap is part of this.** The new rules go into the existing
`fixtures/xlsx-coverage.xlsx` through COM — the sanctioned path per
`fixtures/README.md`, so no new fixture file and no sign-off. That is what
replaces the injected XML: it is the reason the existing cases are pinned on
hand-written XML, and closing it is what makes the rest trustworthy.
`tools/excel-render` then produces the native A4 pages for side-by-side review,
including the three-colour-scale capture that does not exist yet. Docs to update
alongside: ARCHITECTURE *Format rules* and *Known limits* (whatever the icon
decision leaves undrawn, and the `timePeriod` clock if it follows the clock), a
DECISION_LOG entry carrying the probe findings and blast radius, api.mdx's
`conditional` overlay prose, and `fixtures/README.md`.

**Stop condition.** The probe is expected to contradict at least one reading
above — top-10 tie-breaking is the likeliest. Treat what it finds as the answer,
not as a problem with the probe, and record it in the DECISION_LOG entry rather
than encoding the spec's version of it.

**Not part of this item.** The `x14` data-bar extras visible in the coverage
fixture (negative-value fill, axis colour, gradient, border, direction) are
separate residue.

## Before first publish (open decisions)

Decisions that cost nothing to change now and a major version to change
later. AGENTS.md "Pre-release breaking changes" expires at first publish, so
resolve each first. Delete an entry (and record the outcome in DECISION_LOG)
when it is settled.

`@aared/viewer@0.0.0` on npm only holds the name (§42); nothing of the real
code is public.

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

`aared/core` exports the resolved-model types (`ShapeModel`,
`SheetRangeModel`, `DocxParagraphNode`, …; the list is the API page's
"Resolved models" section) that `read()` returns. Published,
they become a contract, while the parser still reshapes them (§28, §55).

Options:

1. **Stable:** every model type promised; any change is a major with a
   migration note.
2. **Tiered (the lean):** the XLSX range types (`SheetRangeModel`,
   `CellModel`, `CellRange`) stable, since reading cells is a real need; the
   PPTX and DOCX models **experimental** (may change in a minor).
   `render()` still accepts a model.
3. **Internal:** stop exporting them; smallest surface, but loses typed
   custom rendering and the XLSX data use.

Confirm the lean against what real consumers use before publishing.

When settled: update `site/src/content/docs/api.mdx` § Resolved models (state
the stability tier per format, and label experimental types), trim
`aared/core` exports if going with option 3, update `docs-parity.test.ts`
accordingly, and record the outcome in DECISION_LOG.

## Not yet scheduled

- Bundle size and memory, only if a real consumer budget demands it:
  per-format `preview`/`react` subpaths, an element registry variant, one Blob
  per shared PPTX picture, the geometry evaluator's `import()` out of the
  shared session chunk, a detached layout/master skeleton (§39)
- `viewer.print()` through the browser's own print dialog (Save as PDF
  included), one slide, sheet range or page per printed page painted from the
  same DOM/SVG, never an `exportPdf` engine; from the superseded API proposal
  in `docs/archive/`, not designed beyond that
- XLSX `contentSize()` (inches, or CSS px at scale 1) on the read model or the
  render handle: today a caller adds up `columnWidths`, `defaultRowHeight` and
  each row's `height`, and adds the viewer's own chrome (row numbers, column
  letters, sheet tabs) by hand
- Viewer decorations API (`decorations: [{ range, kind }]`) — viewer-only
  annotations overlaying rendered content; revisit when a real consumer
  needs it

- Headless thumbnail export — a slide/page/sheet as an image without a
  browser, for agent pipelines. Output format, DOM-free scope and fidelity
  contract all undecided; do not promise it. For DOCX, the DOM-free
  paginator (`core/docx/pagination.ts`, §21) is the natural base
- CSV/TSV preview (may or may not land) — a `.csv`/`.tsv` as a one-sheet
  workbook through `aared/xlsx`, as Excel does. Painting is nearly free (a
  plain `SheetModel`); the work is input: routing non-ZIP bytes past
  `OpcPackage.open` with the file name or type carried through; detection
  from a positive signal confirmed by content (an HTML error page or a legacy
  `.xls` must never paint as a sheet); encoding (BOM, strict UTF-8, then
  `windows-1256`: the Arabic win over Excel, to confirm on Excel 365 first);
  RFC 4180 parsing with delimiter detection and `sep=`; and sheet direction,
  which a CSV doesn't carry. Open decisions: `info.format` stays `"xlsx"`
  (preferred) or a new `"csv"`; values verbatim (preferred) or Excel-style
  inference; and a new Arabic fixture, which needs sign-off. An Excel capture
  is an oracle only with its regional settings recorded
