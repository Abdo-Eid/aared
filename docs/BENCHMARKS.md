# Benchmarks

> Measured, reproducible numbers for aared vs. the closest competitors.
> Qualitative landscape: [COMPETITORS.md](COMPETITORS.md).
> Method is the contract — numbers without it are marketing.

## Status

Speed, memory and search numbers come from `tools/benchmarks/measure.ts` driving the
bench's measurement page in a real browser; sizes come from its `--size` mode.
Every table below is **generated** from the latest results. Do not fill cells
by hand from vendor READMEs or single ad-hoc runs.

The run and search tables were measured on one machine and browser, so their
columns are directly comparable. aared's cells were re-measured in separate
passes on the same day: the two DOCX cells after the tab-stop change in
[Notes §37](Notes.md#note-37), the others before it (that change does not touch
PPTX or XLSX). The size table is merged from two dates and browsers; its caption
says which rows were measured when. One cell is **failed and disclosed rather
than omitted**: `docx-renderer` on `large-40x.docx` did not finish inside the 90 s
per-call budget. It is not broken: one open run once
outside the harness took about 110 s. Its render is synchronous, so the page
cannot answer the harness while it runs, and its time grows faster than the file
does (1.7 s for the coverage document, 16 s for 10 copies of it, 37 s for 20).
It works in the bench UI, which has no such budget; on the coverage document it
measured normally. Treat the large DOCX column as one package short, not as a
completed comparison.

`large-40x.docx` is the coverage body repeated 40 times with one table of
contents (`make-fixtures.py`): 2.1 MB, 409 laid-out pages, a stress file and not
a typical document. Cells measured on earlier versions of it (0.3 MB and 121
pages, then a version with 40 tables of contents and 721 pages) are not
comparable with these ([Notes §37](Notes.md#note-37)). On it aared opens in 1.3 s
warm, the lowest of the five renderers that finished, and in 1.7 s cold, third
behind docx-preview and omni (both 1.6 s). A repeat open holds 58 MB, the lowest.
After scrolling to the end it holds 112 MB, second only to extend (167 MB) and
against 68 MB for docx-preview, which does not paginate: every laid-out page
keeps its layout once shown. On the real `docx-coverage.docx` aared is the
fastest renderer measured warm (52 ms against 65 ms), not cold (304 ms against
172 to 183 ms).

## Method

- **Fixtures (pinned):** the real files in `fixtures/` (`pptx-coverage.pptx`,
  `post-production-review.pptx`, `post-production-tracker.xlsx`,
  `xlsx-coverage.xlsx`, `docx-coverage.docx`) for open, navigation and memory,
  plus three **large generated** files (`tools/benchmarks/search/make-fixtures.py`:
  300 slides, 13 sheets of ~50,000 rows, a 40× document: the coverage body
  repeated 40 times with one table of contents) for how cost grows with size
  and for search.
- **Stage:** one 1280×720 stage for every renderer, in Chromium headless
  (`Bun.WebView`). DOCX renders in layout mode (`pagination: { mode: "layout" }`).
  One renderer × fixture at a time; nothing is fetched but localhost files.
- **Aared integration:** format-local `preview()` owns opening, navigation and
  cleanup; XLSX sheet tabs are explicitly disabled, and DOCX uses layout mode.
  Navigating to the currently displayed unit is a no-op. Older recorded runs
  used low-level `open()` / `render()` (PPTX repainted even the current slide);
  generated tables retain those measurements until the harness is rerun.
- **Runs:** open — 5 cold runs (each a fresh page: modules not yet evaluated,
  browser cache warm) and 11 warm runs after 2 discarded; navigation — 5 passes;
  memory — 5 repeat-open runs, the 5 cold runs, 3 full reads; search — 5 fresh
  sessions. Reported as median and p95 (every memory sample is kept in the
  JSON).
- **Harness work stays out of the cells:** the bench's own side reads (a
  sheet's RTL flag for the SheetJS table, a document's widest section for the
  docx-preview lineage) use the browser's `DecompressionStream`, not a library
  under measurement, and run before anything is measured.
- **Fonts:** no font resolver is passed and nothing is downloaded, so text uses
  the system fonts for every renderer alike.
- **Metrics per cell:**
  - open — adapter entry to first paint, then two animation frames (the frame
    commit). Includes the renderer's dynamic import and, for aared, the Worker
    parse.
  - one move / all units — going to a slide or sheet, each move timed to the
    committed frame; and every slide or sheet once. DOCX has no comparable
    navigation in the baselines, so it is left out.
  - memory — Chrome DevTools `Runtime.getHeapUsage` (JS heap, embedder heap —
    DOM, layout, styles — and buffers, WASM memory included), read in the page
    **and in each of its Workers, each through its own DevTools session**
    (`tools/shared/cdp.ts`), after forced GCs once the page has settled.
    `--calibrate` checks the reading first: 10 MiB held on the page, then in a
    Worker, must each read as 10 MiB, counted once and in the right place.
    - *repeat open* — after-open minus before-open in a page that has opened
      the file before, so the library's code is already loaded. A Worker a
      renderer starts for each open counts in full. Split into page and
      Workers.
    - *first open* — the same in a fresh page, so loading the library's own
      code counts too, for every renderer alike.
    - *after a full read* — every slide or sheet visited, or the document
      scrolled to its end, a stage at a time (not on the large decks and
      workbooks, where a baseline's full pass takes minutes).
    - *left after destroy* — what the renderer keeps once it is gone.
    Pixels held by a canvas or the GPU are **not** counted, so canvas-based
    renderers (pptxviewjs, silurus) are under-reported.
  - elements — how many elements the open painted into the stage, the record of
    how much each renderer builds.
  - search — aared only (the others have no search): the first search reads and
    indexes every unit, a repeat scans what it kept.
  - size — **download size**: every file the browser fetched to open each real
    fixture, in a fresh page with caching off, served from the bench's Vite
    production build — the same build the timings use, so every library gets
    the same minifier, its dependencies bundled in, at the version
    `bench/package.json` installs (the version is recorded per row, and the run
    fails if a package resolves outside the workspace). Each file counts once;
    gzip level 6 and brotli level 11; split into main-thread code, Workers and
    WASM. Per format, the union over that format's fixtures. Code a library
    loads only later (another slide's chart, say) is not counted. "All three"
    is aared's three formats in one page; the stitched alternative is the
    lightest other library per format together, a shared file (jszip) counted
    once.
- **Machine disclosure:** every JSON header records CPU/RAM/OS/browser/Bun.
  Numbers from different machines are never mixed in one table.

## How to re-run

```bash
bun run --cwd bench build                    # builds the measurement page (bench/dist/measure.html)
bun run tools bench calibrate               # checks the memory reading against 10 MiB controls
bun run tools bench size                    # aared's download size, merged into the latest size-<stamp>.json
bun run tools bench run                     # aared's open / navigation / memory, merged into the latest <stamp>.json
bun run tools bench search                  # search on the large fixtures → tools/out/benchmarks/search-<stamp>.json
bun run tools bench report --md             # regenerate the tables below from the latest results
bun run tools bench report --site           # regenerate site/src/data/benchmarks.json (the home page's numbers)
```

**Competitors are measured once.** Their numbers change only with their
versions, so `--run` and `--size` measure aared alone and put its cells into
the latest result in place of its old ones, keeping every competitor's.
Each cell records when, and on which browser, it was measured, and each
table's header says so when they differ. `--run --all` and `--size --all`
measure every renderer into a new file: run them after bumping a
competitor's version in `bench/package.json`, or when a method changes.

**Still not run often.** An aared pass (calibration, size, `--run`,
`--search`) is a fraction of a full `--all` pass (about 30 minutes), and
both need an idle machine on mains power (on battery, small cells read twice
their time), so neither is part of the test suite, CI or a normal change.
Run it deliberately: before publishing numbers, or after a change meant to
move them. Between runs, the tables below and the home page's numbers are
the last run's, whatever the code has done since. For a quick check of one
change, re-run only the cells it touches (below) and compare.

`--run` also takes `fixture:renderer` pairs to re-measure just those cells,
merged the same way: the way to redo a failed cell. A hung page can leave the
browser unanswering DevTools, and every cell after it would then fail the
same way, so the pass stops at the first such error, keeps the cells measured
so far, and prints the `--run …` command that measures the rest in a new
process (a second browser cannot be started in the same one).
The large fixtures are optional; cells that need them are skipped until
`python tools/benchmarks/search/make-fixtures.py` has run. A full pass takes a while
and should be left alone — other work on the machine skews the timings.

The tables in this file are **generated** from the latest results by the
harness. Hand-editing a number without re-running is a docs bug.

Result files stay local: `tools/out/benchmarks/` is gitignored, and each
pass reads only the newest file of its kind. The committed artifacts are this
file's tables and `site/src/data/benchmarks.json`, so the run named in a
table's header may not exist in a fresh clone.

## Results (generated)

<!-- generated:open -->
> Run `20261014-154246.json` on AMD Ryzen 7 6800H with Radeon Graphics (16 threads), 15 GB, Windows_NT 10.0.26300, HeadlessChrome/155.0.0.0, Bun 1.4.3.

| Fixture | Renderer | cold open ms (med) | warm open ms (med / p95) | one move ms (med) | all units ms (med) | memory, repeat open MB (page + Workers) | memory, first open MB | memory after a full read MB | left after destroy MB | idle Worker kept MB | elements painted |
|---|---|---|---|---|---|---|---|---|---|---|---|
| pptx-coverage.pptx | aared | 58 | 32 / 33 | 33.3 | 751 | 1.0 (0.2 + 0.8) | 2.9 | 1.4 | 0.0 | 1.7 | 66 |
| pptx-coverage.pptx | pptxviewjs | 136 | 82 / 83 | 33.3 | 666 | 2.1 (2.1 + 0.0) | 6.9 | 2.0 | 2.0 | — | 1 |
| pptx-coverage.pptx | silurus | 250 | 166 / 200 | 33.2 | 666 | 2.1 (1.1 + 0.9) | 4.8 | 3.1 | 0.0 | — | 1 |
| pptx-coverage.pptx | aiden0z | 247 | 98 / 100 | 33.3 | 665 | 6.6 (6.6 + 0.0) | 11.1 | 6.9 | 0.0 | — | 976 |
| pptx-coverage.pptx | mdurbar | 75 | 38 / 66 | 33.3 | 658 | 1.2 (1.2 + 0.0) | 1.9 | 1.2 | -0.0 | — | 102 |
| pptx-coverage.pptx | pptx-browser | 29 | 32 / 34 | 33.2 | 667 | 1.0 (1.0 + 0.0) | 1.5 | 1.3 | 0.0 | — | 1 |
| pptx-coverage.pptx | extend | 128 | 82 / 83 | 33.4 | 666 | 0.1 (0.1 + 0.0) | 3.8 | 0.1 | 0.0 | — | 60 |
| pptx-coverage.pptx | omni | 251 | 98 / 101 | 33.3 | 632 | 5.5 (5.5 + 0.0) | 10.5 | 5.7 | 0.0 | — | 1222 |
| post-production-review.pptx | aared | 57 | 33 / 34 | 33.1 | 232 | 0.6 (0.1 + 0.5) | 2.4 | 0.8 | 0.0 | 1.5 | 39 |
| post-production-review.pptx | pptxviewjs | 101 | 49 / 50 | 33.2 | 233 | 1.0 (1.0 + 0.0) | 4.9 | 0.8 | 1.0 | — | 1 |
| post-production-review.pptx | silurus | 180 | 132 / 147 | 33.1 | 233 | 1.2 (0.3 + 0.9) | 3.9 | 2.7 | 0.0 | — | 1 |
| post-production-review.pptx | aiden0z | 246 | 61 / 83 | 33.3 | 222 | 4.4 (4.4 + 0.0) | 9.7 | 4.5 | 0.0 | — | 738 |
| post-production-review.pptx | mdurbar | 42 | 32 / 33 | 33.4 | 233 | 0.6 (0.6 + 0.0) | 1.3 | 0.6 | 0.0 | — | 42 |
| post-production-review.pptx | pptx-browser | 35 | 32 / 33 | 33.3 | 233 | 0.5 (0.5 + 0.0) | 0.9 | 0.6 | 0.0 | — | 1 |
| post-production-review.pptx | extend | 108 | 65 / 66 | 33.4 | 234 | 0.1 (0.1 + 0.0) | 3.1 | 0.1 | 0.0 | — | 35 |
| post-production-review.pptx | omni | 266 | 51 / 64 | 33.4 | 233 | 4.4 (4.4 + 0.0) | 9.9 | 4.4 | 0.9 | — | 916 |
| post-production-tracker.xlsx | aared | 110 | 28 / 29 | 33.1 | 127 | 1.3 (1.2 + 0.1) | 3.6 | 0.8 | 0.0 | 1.6 | 455 |
| post-production-tracker.xlsx | sheetjs | 52 | 33 / 34 | 33.1 | 133 | 0.6 (0.6 + 0.0) | 2.8 | 0.3 | -0.0 | — | 337 |
| post-production-tracker.xlsx | xlscalc | 267 | 100 / 117 | 33.1 | 133 | 0.9 (0.9 + 0.0) | 6.8 | 0.6 | 0.0 | — | 292 |
| post-production-tracker.xlsx | spreadsheet-preview | 540 | 382 / 399 | 33.3 | 133 | 1.1 (1.1 + 0.0) | 10.7 | 1.1 | 0.1 | — | 43 |
| post-production-tracker.xlsx | extend | 274 | 118 / 146 | 33.5 | 130 | 2.3 (1.1 + 1.2) | 14.3 | 2.2 | 0.0 | — | 73 |
| post-production-tracker.xlsx | omni | 68 | 32 / 34 | 32.2 | 32 | 1.1 (1.1 + 0.0) | 3.5 | 1.1 | -0.0 | — | 309 |
| xlsx-coverage.xlsx | aared | 73 | 28 / 29 | 33.1 | 196 | 0.7 (0.6 + 0.1) | 2.4 | 0.9 | 0.0 | 1.7 | 205 |
| xlsx-coverage.xlsx | sheetjs | 113 | 58 / 77 | 33.2 | 190 | 4.2 (4.2 + 0.0) | 6.4 | 4.1 | 0.0 | — | 69 |
| xlsx-coverage.xlsx | xlscalc | 356 | 168 / 183 | 33.1 | 200 | 4.7 (4.7 + 0.0) | 10.8 | 4.6 | 0.0 | — | 68 |
| xlsx-coverage.xlsx | spreadsheet-preview | 610 | 449 / 450 | 33.3 | 200 | 1.2 (1.2 + 0.0) | 11.1 | 1.2 | 0.2 | — | 47 |
| xlsx-coverage.xlsx | extend | 235 | 117 / 148 | 34.1 | 210 | 2.2 (1.1 + 1.1) | 14.0 | 2.2 | 0.0 | — | 38 |
| xlsx-coverage.xlsx | omni | 119 | 61 / 76 | 22.0 | 22 | 4.4 (4.4 + 0.0) | 6.8 | 4.4 | 0.0 | — | 83 |
| docx-coverage.docx | aared | 304 | 52 / 59 | — | — | 2.4 (2.1 + 0.3) | 6.5 | 3.5 | 0.0 | 1.9 | 1341 |
| docx-coverage.docx | docx-preview | 172 | 65 / 66 | — | — | 2.0 (2.0 + 0.0) | 3.1 | 2.0 | 0.0 | — | 1136 |
| docx-coverage.docx | docx-renderer | 1045 | 831 / 948 | — | — | 3.9 (3.9 + 0.0) | 6.6 | 4.0 | 2.0 | — | 1398 |
| docx-coverage.docx | apollo | 179 | 65 / 67 | — | — | 2.0 (2.0 + 0.0) | 3.1 | 2.0 | 0.0 | — | 1106 |
| docx-coverage.docx | extend | 966 | 768 / 785 | — | — | 5.0 (5.0 + 0.0) | 8.8 | 6.5 | -1.5 | — | 742 |
| docx-coverage.docx | omni | 183 | 65 / 66 | — | — | 2.7 (2.7 + 0.0) | 4.0 | 2.7 | 0.0 | — | 1080 |
| large-300-slides.pptx | aared | 69 | 32 / 34 | — | — | 1.3 (0.2 + 1.1) | 3.2 | — | 0.0 | 1.3 | 39 |
| large-300-slides.pptx | pptxviewjs | 1095 | 1072 / 1405 | — | — | 39.9 (39.9 + 0.0) | 61.4 | — | 38.4 | — | 1 |
| large-300-slides.pptx | silurus | 703 | 665 / 682 | — | — | 1.3 (0.4 + 0.9) | 4.1 | — | 0.0 | — | 1 |
| large-300-slides.pptx | aiden0z | 3064 | 2750 / 2893 | — | — | 152.0 (152.0 + 0.0) | 158.8 | — | 0.0 | — | 31706 |
| large-300-slides.pptx | mdurbar | 642 | 551 / 566 | — | — | 9.7 (9.7 + 0.0) | 10.6 | — | -0.0 | — | 42 |
| large-300-slides.pptx | pptx-browser | 74 | 49 / 66 | — | — | 6.3 (6.3 + 0.0) | 6.8 | — | 0.0 | — | 1 |
| large-300-slides.pptx | extend | 450 | 382 / 400 | — | — | 0.1 (0.1 + 0.0) | 8.1 | — | 0.0 | — | 35 |
| large-300-slides.pptx | omni | 1687 | 1489 / 1670 | — | — | 134.4 (134.4 + 0.0) | 142.8 | — | 30.2 | — | 39270 |
| large-13-sheets.xlsx | aared | 67 | 29 / 30 | — | — | 2.5 (0.6 + 1.8) | 4.0 | — | 0.0 | 1.3 | 205 |
| large-13-sheets.xlsx | sheetjs | 1366 | 1276 / 1319 | — | — | 55.0 (55.0 + 0.0) | 59.5 | — | 0.0 | — | 69 |
| large-13-sheets.xlsx | xlscalc | 12395 | 12000 / 12463 | — | — | 161.8 (161.8 + 0.0) | 168.1 | — | 0.0 | — | 68 |
| large-13-sheets.xlsx | spreadsheet-preview | 2932 | 2866 / 2916 | — | — | 115.9 (115.9 + 0.0) | 126.1 | — | 43.7 | — | 65 |
| large-13-sheets.xlsx | extend | 5117 | 4840 / 5126 | — | — | 2.9 (1.1 + 1.8) | 36.7 | — | 0.0 | — | 47 |
| large-13-sheets.xlsx | omni | 1334 | 1208 / 1243 | — | — | 57.0 (57.0 + 0.0) | 61.6 | — | 0.0 | — | 92 |
| large-40x.docx | aared | 1701 | 1326 / 1425 | — | — | 57.7 (55.5 + 2.2) | 50.8 | 112.0 | 0.0 | 2.1 | 53441 |
| large-40x.docx | docx-preview | 1624 | 1534 / 1663 | — | — | 66.9 (66.9 + 0.0) | 56.5 | 68.0 | 0.0 | — | 43412 |
| large-40x.docx | docx-renderer | *timed out after 90s: __m.open("docx", "docx-renderer", "large-40x.docx")* | | | | | | | | | |
| large-40x.docx | apollo | 5054 | 5221 / 6507 | — | — | 64.6 (64.6 + 0.0) | 54.6 | 65.5 | 0.0 | — | 41822 |
| large-40x.docx | extend | 28018 | 30362 / 63803 | — | — | 165.7 (165.7 + 0.0) | 157.7 | 167.4 | -1.5 | — | 27925 |
| large-40x.docx | omni | 1630 | 1426 / 1654 | — | — | 86.7 (86.7 + 0.0) | 77.0 | 87.7 | 0.0 | — | 40860 |
<!-- /generated:open -->

<!-- generated:search -->
| Fixture | Size | open ms | first search ms | repeat search ms | common word ms |
|---|---|---|---|---|---|
| large-300-slides.pptx | 0.9 MB | 44 | 229 | 0.3 | 1.0 (772 matches) |
| large-13-sheets.xlsx | 1.7 MB | 40 | 702 | 3.3 | 12.7 (22185 matches) |
| large-40x.docx | 2.1 MB | 501 | 28 | 0.4 | 1.9 (721 matches) |
<!-- /generated:search -->

<!-- generated:size -->
> Size `size-20261001-184013.json` from the bench's production build. Aared measured 2026-10-14 on HeadlessChrome/155.0.0.0; the other renderers measured 2026-10-01 on HeadlessChrome/154.0.0.0.

| Format | Renderer | version | gzip KiB | brotli KiB | main thread KiB (gzip) | Workers KiB (gzip) | WASM KiB (gzip) | files |
|---|---|---|---|---|---|---|---|---|
| pptx | aared | 0.1.0 | 134.5 | 113.7 | 53.6 | 80.9 | — | 10 |
| pptx | pptxviewjs | 1.1.9 | 253.2 | 204.5 | 253.2 | — | — | 4 |
| pptx | silurus | 0.87.0 | 922.3 | 674.3 | 255.9 | — | 666.4 | 2 |
| pptx | aiden0z | 1.3.0 | 381.7 | 315.8 | 381.7 | — | — | 5 |
| pptx | mdurbar | 0.2.2 | 28.7 | 24.5 | 28.7 | — | — | 1 |
| pptx | pptx-browser | 4.1.5 | 48.8 | 41.0 | 48.8 | — | — | 1 |
| pptx | extend | 0.2.1 | 494.8 | 406.8 | 269.0 | 2.1 | 223.8 | 13 |
| pptx | omni | 0.2.0 | 441.1 | 360.4 | 441.1 | — | — | 7 |
| xlsx | aared | 0.1.0 | 149.3 | 129.7 | 68.4 | 80.9 | — | 8 |
| xlsx | sheetjs | 0.20.3 | 160.3 | 132.1 | 160.3 | — | — | 1 |
| xlsx | xlscalc | 0.4.0 | 323.2 | 261.1 | 323.2 | — | — | 4 |
| xlsx | spreadsheet-preview | 0.1.0 | 428.1 | 293.1 | 428.1 | — | — | 1 |
| xlsx | extend | 0.16.6 | 2477.5 | 1855.9 | 787.8 | 38.3 | 1651.3 | 13 |
| xlsx | omni | 0.2.0 | 174.6 | 144.9 | 174.6 | — | — | 3 |
| docx | aared | 0.1.0 | 158.5 | 137.8 | 77.6 | 80.9 | — | 11 |
| docx | docx-preview | 0.4.0 | 50.3 | 44.5 | 50.3 | — | — | 4 |
| docx | docx-renderer | 0.2.2 | 132.4 | 113.2 | 132.4 | — | — | 4 |
| docx | apollo | 0.9.0 | 52.4 | 46.4 | 52.4 | — | — | 4 |
| docx | extend | 0.9.2 | 620.4 | 480.4 | 171.3 | 4.0 | 445.2 | 13 |
| docx | omni | 0.2.0 | 63.6 | 56.4 | 63.6 | — | — | 6 |
| all three | aared | — | 210.1 | 180.4 | 129.2 | 80.9 | — | 16 |
| all three | mdurbar (pptx) + sheetjs (xlsx) + docx-preview (docx) | — | 239.4 | 201.1 | 239.4 | — | — | 6 |
<!-- /generated:size -->

## Limits

- Human visual review still owns fidelity (see `test:visual:artifacts`
  pairs); these numbers cover size, speed, memory and search only. Automated
  RTL structure checks live in the unit and visual tests, not here.
- One machine, one browser engine, one run: read the numbers as the shape of
  the difference, not as universal constants.
- The baselines render less than aared does (aared draws the full sheet grid
  with virtualisation, real DOM text and layout paging; a canvas renderer draws
  pixels), so a faster open or a smaller heap is not a fidelity claim. Two
  adapters in particular do less work by design: the SheetJS column is a plain
  table of the used range capped at 200 × 52 cells with no fills, fonts,
  borders, chart or frozen panes (SheetJS is a parser, not a renderer), and
  docx-preview breaks pages only where the author did, so the 40× document
  has fewer, longer pages than aared's laid-out ones.
- Results taken before the Worker-aware memory reading (run files before
  version 2) and entry-file sizes (size files before version 1) counted the
  page twice and Workers never, and weighed competitors without their
  dependencies; the harness ignores them (Notes §33).
- Bench UI (`bun run dev`) is the visual side-by-side; this file is the
  deterministic record. Same fixtures, same adapters, different job.
