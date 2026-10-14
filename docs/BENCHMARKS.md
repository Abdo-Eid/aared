# Benchmarks

> Measured, reproducible numbers for aared vs. the closest competitors.
> Qualitative landscape: [COMPETITORS.md](COMPETITORS.md).
> Method is the contract — numbers without it are marketing.

## Status

Speed, memory and search numbers come from `tools/benchmarks/measure.ts` driving the
bench's measurement page in a real browser; sizes come from its `--size` mode.
Every table below is **generated** from the latest results. Do not fill cells
by hand from vendor READMEs or single ad-hoc runs.

The current refresh is partial: three aared coverage timing/memory cells and
all four aared download-size rows were re-measured. Other timing cells and search
retain their older results; captions disclose the split. This is not a controlled
before/after optimization comparison. A full benchmark refresh follows the next
editing pass.

## Method

- **Fixtures (pinned):** the real files in `fixtures/` (`pptx-coverage.pptx`,
  `post-production-review.pptx`, `post-production-tracker.xlsx`,
  `xlsx-coverage.xlsx`, `docx-coverage.docx`) for open, navigation and memory,
  plus three **large generated** files (`tools/benchmarks/search/make-fixtures.py`:
  300 slides, 13 sheets of ~50,000 rows, a 40× document) for how cost grows
  with size and for search.
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
> Run `20260924-235805.json` on AMD Ryzen 7 6800H with Radeon Graphics (16 threads), 15 GB, Windows_NT 10.0.26200, HeadlessChrome/153.0.0.0, Bun 1.4.2. Aared (3 of 8 cells) measured 2026-10-11 on HeadlessChrome/155.0.0.0; aared (5 of 8 cells), the other renderers measured 2026-09-24 on HeadlessChrome/153.0.0.0.

| Fixture | Renderer | cold open ms (med) | warm open ms (med / p95) | one move ms (med) | all units ms (med) | memory, repeat open MB (page + Workers) | memory, first open MB | memory after a full read MB | left after destroy MB | idle Worker kept MB | elements painted |
|---|---|---|---|---|---|---|---|---|---|---|---|
| pptx-coverage.pptx | aared | 128 | 32 / 33 | 33.5 | 949 | 1.0 (0.2 + 0.8) | 2.9 | 1.4 | 0.0 | 1.7 | 66 |
| pptx-coverage.pptx | pptxviewjs | 226 | 99 / 118 | 33.2 | 332 | 1.5 (1.5 + 0.0) | 5.8 | 1.5 | 1.5 | — | 1 |
| pptx-coverage.pptx | silurus | 420 | 250 / 312 | 33.5 | 333 | 2.0 (1.1 + 0.9) | 4.8 | 2.8 | 0.0 | — | 1 |
| pptx-coverage.pptx | aiden0z | 398 | 135 / 151 | 33.1 | 327 | 4.4 (4.4 + 0.0) | 8.9 | 5.0 | 0.0 | — | 575 |
| pptx-coverage.pptx | mdurbar | 147 | 77 / 102 | 33.3 | 321 | 1.2 (1.2 + 0.0) | 1.9 | 1.2 | 0.0 | — | 153 |
| pptx-coverage.pptx | omni | 422 | 116 / 145 | 33.2 | 320 | 3.7 (3.7 + 0.0) | 8.6 | 3.8 | 0.0 | — | 740 |
| post-production-review.pptx | aared | 98 | 32 / 33 | 33.4 | 233 | 0.6 (0.1 + 0.5) | 2.1 | 0.8 | 0.0 | 1.2 | 39 |
| post-production-review.pptx | pptxviewjs | 218 | 92 / 138 | 33.3 | 232 | 0.9 (0.9 + 0.0) | 4.9 | 0.9 | 0.9 | — | 1 |
| post-production-review.pptx | silurus | 370 | 248 / 340 | 33.3 | 233 | 1.2 (0.3 + 0.9) | 3.9 | 2.7 | 0.0 | — | 1 |
| post-production-review.pptx | aiden0z | 521 | 138 / 176 | 33.2 | 217 | 4.4 (4.4 + 0.0) | 9.6 | 4.9 | 0.0 | — | 738 |
| post-production-review.pptx | mdurbar | 100 | 51 / 65 | 33.3 | 232 | 0.6 (0.6 + 0.0) | 1.3 | 0.6 | 0.0 | — | 42 |
| post-production-review.pptx | omni | 518 | 103 / 123 | 33.1 | 226 | 4.3 (4.3 + 0.0) | 10.3 | 4.7 | 0.5 | — | 916 |
| post-production-tracker.xlsx | aared | 148 | 27 / 41 | 33.4 | 123 | 1.2 (1.1 + 0.1) | 2.6 | 0.8 | 0.0 | 1.2 | 453 |
| post-production-tracker.xlsx | sheetjs | 96 | 32 / 34 | 32.9 | 132 | 0.6 (0.6 + 0.0) | 2.8 | 0.3 | -0.0 | — | 337 |
| post-production-tracker.xlsx | xlscalc | 441 | 116 / 118 | 33.2 | 132 | 0.9 (0.9 + 0.0) | 6.7 | 0.6 | 0.0 | — | 292 |
| post-production-tracker.xlsx | spreadsheet-preview | 768 | 449 / 482 | 32.9 | 132 | 1.1 (1.1 + 0.0) | 10.7 | 1.1 | 0.1 | — | 43 |
| post-production-tracker.xlsx | omni | 128 | 32 / 34 | 31.5 | 32 | 1.1 (1.1 + 0.0) | 3.5 | 1.1 | 0.0 | — | 309 |
| xlsx-coverage.xlsx | aared | 148 | 28 / 38 | 41.8 | 242 | 0.7 (0.6 + 0.1) | 2.3 | 1.3 | -0.0 | 1.7 | 204 |
| xlsx-coverage.xlsx | sheetjs | 203 | 118 / 136 | 32.5 | 122 | 4.1 (4.1 + 0.0) | 6.3 | 4.2 | 0.0 | — | 69 |
| xlsx-coverage.xlsx | xlscalc | 540 | 169 / 200 | 33.5 | 149 | 4.5 (4.5 + 0.0) | 10.5 | 4.7 | 0.0 | — | 68 |
| xlsx-coverage.xlsx | spreadsheet-preview | 912 | 549 / 582 | 33.1 | 132 | 1.1 (1.1 + 0.0) | 10.8 | 1.1 | 0.1 | — | 43 |
| xlsx-coverage.xlsx | omni | 233 | 120 / 140 | 25.3 | 25 | 4.3 (4.3 + 0.0) | 6.7 | 4.3 | 0.0 | — | 81 |
| docx-coverage.docx | aared | 800 | 166 / 181 | — | — | 2.3 (2.0 + 0.3) | 6.5 | 3.4 | 0.0 | 1.9 | 1341 |
| docx-coverage.docx | docx-preview | 293 | 98 / 115 | — | — | 1.0 (1.0 + 0.0) | 2.1 | 1.0 | 0.0 | — | 647 |
| docx-coverage.docx | docx-renderer | 926 | 548 / 684 | — | — | 1.9 (1.9 + 0.0) | 4.4 | 1.9 | 1.0 | — | 710 |
| docx-coverage.docx | omni | 304 | 98 / 99 | — | — | 1.4 (1.4 + 0.0) | 2.7 | 1.4 | 0.0 | — | 599 |
| large-300-slides.pptx | aared | 118 | 32 / 45 | — | — | 1.3 (0.2 + 1.1) | 2.9 | — | 0.0 | 1.1 | 39 |
| large-300-slides.pptx | pptxviewjs | 2994 | 2500 / 2822 | — | — | 38.4 (38.4 + 0.0) | 61.4 | — | 38.4 | — | 1 |
| large-300-slides.pptx | silurus | 1527 | 1297 / 1364 | — | — | 1.3 (0.4 + 0.9) | 4.1 | — | 0.0 | — | 1 |
| large-300-slides.pptx | aiden0z | 5823 | 4886 / 5465 | — | — | 152.1 (152.1 + 0.0) | 159.6 | — | 0.0 | — | 31706 |
| large-300-slides.pptx | mdurbar | 1862 | 1596 / 1725 | — | — | 9.7 (9.7 + 0.0) | 10.6 | — | -0.0 | — | 42 |
| large-300-slides.pptx | omni | 3497 | 2919 / 3203 | — | — | 134.3 (134.3 + 0.0) | 142.7 | — | 30.2 | — | 39270 |
| large-13-sheets.xlsx | aared | 111 | 26 / 43 | — | — | 2.3 (0.6 + 1.8) | 3.5 | — | 0.0 | 1.0 | 202 |
| large-13-sheets.xlsx | sheetjs | 2978 | 2379 / 2525 | — | — | 55.0 (55.0 + 0.0) | 59.4 | — | 0.0 | — | 69 |
| large-13-sheets.xlsx | xlscalc | 13012 | 12748 / 13366 | — | — | 161.9 (161.9 + 0.0) | 168.0 | — | 0.0 | — | 68 |
| large-13-sheets.xlsx | spreadsheet-preview | 3019 | 2749 / 2832 | — | — | 115.7 (115.7 + 0.0) | 125.8 | — | 43.7 | — | 61 |
| large-13-sheets.xlsx | omni | 2798 | 2493 / 2619 | — | — | 56.9 (56.9 + 0.0) | 61.5 | — | 0.0 | — | 90 |
| large-40x.docx | aared | 1119 | 611 / 712 | — | — | 15.4 (15.1 + 0.3) | 15.3 | 33.0 | 14.9 | 1.2 | 19241 |
| large-40x.docx | docx-preview | 2021 | 1776 / 1940 | — | — | 24.4 (24.4 + 0.0) | 23.3 | 24.7 | 0.0 | — | 16940 |
| large-40x.docx | docx-renderer | *Cannot read properties of undefined (reading 'length')* | | | | | | | | | |
| large-40x.docx | omni | 2109 | 1845 / 2016 | — | — | 34.7 (34.7 + 0.0) | 33.9 | 35.1 | 0.0 | — | 17459 |
<!-- /generated:open -->

<!-- generated:search -->
| Fixture | Size | open ms | first search ms | repeat search ms | common word ms |
|---|---|---|---|---|---|
| large-300-slides.pptx | 0.9 MB | 30 | 217 | 0.3 | 0.3 (214 matches) |
| large-13-sheets.xlsx | 1.7 MB | 24 | 576 | 3.1 | 10.3 (22185 matches) |
| large-40x.docx | 0.3 MB | 145 | 7 | 0.2 | 0.5 (361 matches) |
<!-- /generated:search -->

<!-- generated:size -->
> Size `size-20261001-184013.json` from the bench's production build. Aared measured 2026-10-11 on HeadlessChrome/155.0.0.0; the other renderers measured 2026-10-01 on HeadlessChrome/154.0.0.0.

| Format | Renderer | version | gzip KiB | brotli KiB | main thread KiB (gzip) | Workers KiB (gzip) | WASM KiB (gzip) | files |
|---|---|---|---|---|---|---|---|---|
| pptx | aared | 0.0.0 | 134.3 | 113.7 | 52.7 | 81.6 | — | 13 |
| pptx | pptxviewjs | 1.1.9 | 253.2 | 204.5 | 253.2 | — | — | 4 |
| pptx | silurus | 0.87.0 | 922.3 | 674.3 | 255.9 | — | 666.4 | 2 |
| pptx | aiden0z | 1.3.0 | 381.7 | 315.8 | 381.7 | — | — | 5 |
| pptx | mdurbar | 0.2.2 | 28.7 | 24.5 | 28.7 | — | — | 1 |
| pptx | pptx-browser | 4.1.5 | 48.8 | 41.0 | 48.8 | — | — | 1 |
| pptx | extend | 0.2.1 | 494.8 | 406.8 | 269.0 | 2.1 | 223.8 | 13 |
| pptx | omni | 0.2.0 | 441.1 | 360.4 | 441.1 | — | — | 7 |
| xlsx | aared | 0.0.0 | 143.1 | 124.5 | 61.5 | 81.6 | — | 11 |
| xlsx | sheetjs | 0.20.3 | 160.3 | 132.1 | 160.3 | — | — | 1 |
| xlsx | xlscalc | 0.4.0 | 323.2 | 261.1 | 323.2 | — | — | 4 |
| xlsx | spreadsheet-preview | 0.1.0 | 428.1 | 293.1 | 428.1 | — | — | 1 |
| xlsx | extend | 0.16.6 | 2477.5 | 1855.9 | 787.8 | 38.3 | 1651.3 | 13 |
| xlsx | omni | 0.2.0 | 174.6 | 144.9 | 174.6 | — | — | 3 |
| docx | aared | 0.0.0 | 157.9 | 137.5 | 76.3 | 81.6 | — | 14 |
| docx | docx-preview | 0.4.0 | 50.3 | 44.5 | 50.3 | — | — | 4 |
| docx | docx-renderer | 0.2.2 | 132.4 | 113.2 | 132.4 | — | — | 4 |
| docx | apollo | 0.9.0 | 52.4 | 46.4 | 52.4 | — | — | 4 |
| docx | extend | 0.9.2 | 620.4 | 480.4 | 171.3 | 4.0 | 445.2 | 13 |
| docx | omni | 0.2.0 | 63.6 | 56.4 | 63.6 | — | — | 6 |
| all three | aared | — | 203.3 | 174.8 | 121.7 | 81.6 | — | 19 |
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
