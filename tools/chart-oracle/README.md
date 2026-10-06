# Chart oracle

The proof loop for the chart engine (`packages/aared/src/core/chart/`).
PowerPoint draws a fixed corpus of 94 charts. aared draws the same charts
through its own route: read as the Worker reads them, laid out as the page
lays them out. Then the two are compared word by word.

| File | Role |
| --- | --- |
| `gen_corpus.py` | Writes `chart-oracle.pptx`, `chart-oracle.xlsx` and `cases.json` next to itself (or into the folder given), one chart per slide or sheet. The chart XML is written by hand, so each case states exactly what Office is asked to draw. The `office` flavour is what PowerPoint and Excel 365 insert; the `bare` flavour is minimal XML, so Office falls back to its own defaults. Each sheet also holds its chart's data in the cells the chart names, because Excel recalculates chart links on open. Needs `python-pptx`, `openpyxl` and `lxml` |
| `export-office.ps1` | Windows with Office: PowerPoint exports the deck to `out/chart-oracle.pdf` (vector), and Excel exports each sheet's chart to `out/xlsx-<sheet>.png` (`-Sheets c1,c2` limits it to those sheets, `-Dir` is the folder holding the workbook) |
| `render.ts` | aared draws every chart in the deck: `out/engine.json`, plus `out/svg/NN.svg` for eyeballing |
| `measure.ts` | Lays the engine's text out in a real browser (`Bun.WebView`) and reads back each word's box: `out/measured.json` |
| `score.py` | Compares with PowerPoint's PDF using PyMuPDF: words matched by string (centre error), the value axis's numbers, and the inner plot box from the gridlines. Prints a table and writes `out/score.json` |
| `gen_metrics.py` | Regenerates the advance-width tables in `core/chart/metrics-data.ts`: Latin from metric-compatible font files, Arabic per letter and joining form from Arial and Tahoma (the Windows files; `--arabic-only` prints just those). Needs `fontTools` |
| `arabic-width.py` | Scores the engine's joined-letter Arabic widths (`joined_width` in `gen_metrics.py`, the same rule as `core/chart/metrics.ts`) against PowerPoint's PDF, for the cases marked `"script": "arabic"` in `cases.json`. `score.py` also reports those cases apart from the Latin ones |
| `bench.ts` | Speed of the two halves (read, layout) on small and very large charts, with primitive counts and model size |
| `gen_types.py`, `types-oracle.xlsx` | Needs `pywin32` and Excel. Excel (a private `DispatchEx` instance) makes stock and pie-of-pie / bar-of-pie charts, and exports Excel's own PNG of each; `test/chart-types.test.ts` reads the workbook |
| `gen_chartex.py`, `chartex-oracle.xlsx` | Needs `pywin32` and Excel; its output goes to the current folder, so run it from `tools/chart-oracle/`. The same for the Office 2016 types (waterfall, funnel, histogram, Pareto, box and whisker, treemap, sunburst); `test/chart-chartex.test.ts` reads it. Excel writes these parts with no data cache (their formulas name hidden `_xlchart.*` names) |

```sh
powershell -ExecutionPolicy Bypass -File tools/chart-oracle/export-office.ps1
bun tools/chart-oracle/render.ts
bun tools/chart-oracle/measure.ts
python tools/chart-oracle/score.py
```

Every Office distance the layout uses is in `K` in `core/chart/engine.ts`.
Calibrating means fitting those numbers to the score, not changing code
paths.

The corpus is a tool corpus, not a fixture. Its XML is written by hand to
pin one feature per case. Real-file tests use the decks and workbooks in
`fixtures/`.

The first 64 cases are in-sample: `K` was fitted on them. The ten added in
DECISION_LOG §73 (six Arabic, `ar-*`, scored apart; four legend and pie
stress cases) were not, except the bottom-legend grid and the inside-label
wrap, each fitted on one of them. The twenty `g-*` cases (§74) each check a
rule the engine had assumed; the case's `guess` field names it, and
`test/chart-oracle-guesses.test.ts` pins what PowerPoint drew. Cases the
generator's options do not cover use its raw-XML hooks: `serExtra` and
`serExt` (series XML: dLbls, trendline, errBars; an extLst), `catXml`,
`catNums` with `catFmt`, an axis `tag` of `c:dateAx`, and the `stock` kind.
All run with Office's fonts installed. Some Arabic words do not match by
string, because PowerPoint's PDF stores them in presentation forms. See
ARCHITECTURE "Known limits" for what is still assumed.
