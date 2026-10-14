# tools/probe

Ground-truth probes: reverse-engineering real Office behavior by
interacting with it ([AGENTS.md](../../AGENTS.md#ground-truth-probe)). Scratch evidence,
not a deliverable — each finished probe is deleted once its finding is
codified as a comment plus a fixture and a test.

## Why probes are disposable

A probe is tied to one code change or bug question. Once answered and
codified, it is irrelevant: the comment, fixture, and test reproduce the
result, so the snapshot trees can be re-derived or dropped anytime.
Losing an ignored probe directory (fresh clone, folder delete) loses
nothing except unfinished investigations — and an unfinished probe is the
only kind worth keeping.

## Workflow

For unclear OOXML semantics, test a targeted interaction in real
PowerPoint/Excel/Word rather than relying on the spec alone:

1. Save the confusing file in a gitignored, nested Git repo at
   `tools/probe/<topic>/`, not the main repo or `%TEMP%`.
2. Unzip every part and commit the snapshot as "before" in that nested repo.
3. Open the file in Office, perform the one interaction in question and save.
4. Re-unzip over the snapshot and inspect `git diff`. Ignore incidental save
   noise (properties, thumbnails, whitespace); identify the relevant XML change.
5. Codify the finding as a standalone parse-site comment describing the Office
   interaction and XML result, plus real-fixture coverage and a regression test.
   Record reasoning/evidence in Notes and current behavior in ARCHITECTURE.
   A bare pointer to the scratch topic is not a durable explanation.
6. Once comment, fixture and test reproduce the finding, delete that completed
   topic. Until then, preserve it; never move/reset/delete another contributor's
   investigation without checking. The retained inventory below is not blanket
   cleanup authorization.

Use the standard shape and shared helpers below for new probes. Font calibration
is separate frozen research under `tools/font-calibration/`, not a probe topic.

## Standard shape

One directory is one question, and holds exactly one role script:

| Role | Script | Writes | Snapshot? |
|---|---|---|---|
| `probe` (throwaway measurement) | `probe.py` (or `probe.ps1`) | stdout only | never |
| `author` (builds a probe-local file) | `author.py` / `author.ps1` | `<name>.docx/.pptx/.xlsx` in-dir | optional `snapshot/` |
| `extend` (mutates `fixtures/`) | `extend.py` / `extend.ps1` | temp copy of `fixtures/*` | never (the diff lives in `fixtures/`) |

Every directory carries a `README.md`: the question, the exact Office
steps (or the script command), the finding (which XML changed), and where
it is codified (parse-site comment, fixture, test) — or "throwaway, no
fixture" when nothing landed.

New scripts import the shared helpers instead of copying them:

```python
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "_common"))
from office import word_app, fixture_copy, snapshot
from strip_label import main as strip_label
```

## `_common/`

`office.py` (Word/PowerPoint/Excel COM guards, fixture temp-copy,
unzip snapshot) and `strip_label.py` (org sensitivity-label removal,
run on a fixture after every Office save).

## Retained-topic inventory (October 2026)

The tools-layout cleanup removed only seven topics whose Office finding,
source rule, real-fixture coverage and regression tests were verified. Their
durable records are:

| Completed topic | Finding and surviving evidence |
|---|---|
| `group-id` | PowerPoint draws a group without an id; shapes/pictures draw after repair restores an id. `core/pptx.ts` records the interaction, `pptx-group.test.ts` checks the real coverage deck with ids removed |
| `title-insets` | Setting title margins at master/layout/slide scope established independent inset inheritance, with explicit zero preserved. `core/pptx.ts`, `pptx-parser.test.ts`, the coverage deck |
| `line-spacing-units` | Native exports established line-unit spacing, Word style precedence and stored PPTX autofit scaling. `core/pptx.ts`, `core/docx/document.ts`, `pptx/index.ts`, parser/paragraph-spacing/bullet-geometry tests and both coverage fixtures |
| `table-style-colors` | PowerPoint's table-style tint is linear-light. `core/theme.ts`, the coverage deck's table, exact resolved table-fill assertions in `pptx-parser.test.ts` |
| `svg-pictures` | Office SVG extension references and optional PNG fallback; Excel picture effects from the second COM pass. `core/image-source.ts`, fixture notes, image-source/PPTX/XLSX parser tests and `svg-pictures.visual.test.ts` |
| `xlsx-cf-rules` | Excel top/bottom ties, percent rank and mean/population-deviation behavior. `core/xlsx/sheet.ts`, `xlsx-conditional.test.ts` and Notes §105 |
| `xlsx-summary-redesign` | Completed Summary authoring, differential-style reorder and explicit alignment. `fixtures/xlsx-coverage.xlsx`, fixture notes, parser/chrome/render/conditional tests and Notes §116 |

A second pass (October 14, 2026) audited each remaining topic one at a time
and closed these:

| Completed topic | Finding and surviving evidence |
|---|---|
| `docx-rtl-table`, `docx-delivery-tracker`, `exact-pitch`, `ltr-runs`, `docx-lists-headers` | Authoring records for cases now in `docx-coverage.docx`/`pptx-coverage.pptx` (Delivery tracker section, slide 11, section 8). `core/docx/document.ts`, `core/pptx.ts`, `text-style.ts`, docx-parser, pptx-parser, docx-render and pptx-fills tests; Notes §43, §44, §76. `docx-rtl-table`'s promised standalone fixture was never made; the tracker section covers it |
| `docx-cell-border` | Word centres a side border on the grid line. `docx/renderer.ts`, `docx-render.visual.test.ts` (the fixture's 3pt border); the 0.5pt and 6pt measurements are in Notes |
| `xlsx-builtin-numfmt` | Excel's displayed text for 60 code/value cases and its resolution of built-in ids 0-58 are test data (`xlsx-numfmt-oracle.json`, `xlsx-numfmt-builtin-oracle.json`) checked by `xlsx-numberfmt.test.ts`; the regional ids are listed in `core/xlsx/numberfmt.ts`. Id 47 is `mm:ss.0`, as Excel resolves it (ECMA-376 misprints `mmss.0`; Notes §130) |
| `first-line-leading` | PowerPoint's 96 first-line baselines are `pptx-first-line-oracle.json`; `pptx-line-leading.visual.test.ts` checks all six font/size cases at eight spacings |
| `table-style-guids` | PowerPoint's stored definitions of the seven Medium Style 2 ids are `pptx-table-styles-oracle.xml`; `pptx-fidelity.test.ts` checks the built style equals each |
| `en-us-arabic-run` | PowerPoint keeps written order for all 31 ASCII marks in an en-US run of an LTR paragraph; `pptx-fidelity.test.ts` checks every mark |
| `run-direction` | Its README had the rule backwards; Notes §44 and the code agree. The two unprobed cases are an OPEN row in ARCHITECTURE (Run direction beyond left-to-right runs) |
| `docx-bidi-visual` | `w:bidiVisual` makes borders logical, keeps cell margins physical and measures `tblInd` from the start side. `core/docx/document.ts`, docx-parser and docx-render tests, the fixture's RTL table, Notes §43 |
| `docx-cs-bold` | Complex-script runs take `w:bCs`/`w:iCs` only, others `w:b`/`w:i`. `core/docx/document.ts`, `docx-parser.test.ts`, the fixture's bold appendix header |
| `docx-first-header` | First, even and default headers inherit independently by kind; a first page with no first header is blank. `core/docx/document.ts`, `docx/pagination.ts`, `docx-headers.test.ts`, `w:titlePg` in the fixture |
| `docx-justify` | Word shrinks spaces in justified lines only in `compatibilityMode` 15+. `core/docx/document.ts`, `docx/renderer.ts`, docx-parser and docx-render tests; the raw width sweep was not kept |
| `docx-empty-paragraph` | An empty paragraph is one line at its mark's size plus its spacing. `core/docx/document.ts`, `docx-empty-paragraph.visual.test.ts` (rebuilds the probe document), Notes |
| `master-graphics` | `showMasterSp` at layout and slide scope, and the inherited master picture background, matched PowerPoint with no change. `core/pptx.ts` (`parsePptxLayout`), pptx-layout tests, deck slides 15-18 |
| `pptx-lines` | PowerPoint writes lines and connectors as `p:cxnSp`. `core/pptx.ts`, deck slide 11, pptx-parser and pptx-fills tests |
| `pptx-merged-chartex` | Merged table cells and a pasted Excel waterfall (`cx:lvl` cache). `core/pptx.ts`, deck slides 13-14, pptx-parser and chart-chartex tests |
| `pptx-second-master` | Per-master themes; aared had used the first theme for every slide and was fixed. `core/pptx.ts`, deck slides 15-18, pptx-layout tests, Notes (the moved-table capture is described there) |
| `xlsx-data-bar-gradient` | Empty folder |

Source/test names above are relative to `packages/aared/src/` and
`packages/aared/test/`. Historical topic names in comments identify an Office
interaction; they are not links to scripts a fresh clone must have.

**Six local topics remain. This does not declare them finished or
unfinished: it records why they were not deleted.**

| Reason to preserve | Topics |
|---|---|
| Unique Office evidence or an authoring recipe not recorded elsewhere | `docx-float-wrap` (six wrap cases not in the fixture), `docx-charts` (the chart-authoring route and why others failed), `docx-list-counters` (Word's printed markers and the empty-cell height) |
| README describes an investigation without a completed finding | `drawing-id` |
| Live work | `model-contract`, `xlsx-cell-fills` |

This is a local cleanup inventory, not a promise that ignored folders exist
in a fresh clone. For the next probe, record closure and surviving evidence in
its own README; never batch-delete the retained list. Shared `_common/`
helpers are live code and are not cleanup candidates.
