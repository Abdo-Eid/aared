# Developer tools

Start with **`bun run tools --help`**. The dispatcher runs existing backends;
it is not another Office or browser harness. Run it from the repo root, or
`bun /path/to/aared/tools/cli.ts ...` from another directory (input and output
arguments are resolved against your current directory).

## Layout

Only `cli.ts`, `README.md` and `tsconfig.json` live at the root.

```text
tools/
  cli/                 command parsing, doctor and CLI tests
  benchmarks/          browser measurement, effects and search diagnostics/generator
  review/              shared export/capture/report code; pptx/, xlsx/, docx/ native adapters
  checks/              consumer-build and visual-test runner
  inspect/             headless model/RPC inspector
  shared/              Office/process/naming helpers, repository paths, page/Worker CDP
  chart-oracle/        specialized chart code and tracked test corpora
  font-calibration/    frozen scripts and pinned research evidence
  probe/               topic investigations and shared Office helpers
  out/                 ignored generated captures, runs and timing fixtures
```

`out/` separates generated data from source: `pptx/`, `xlsx/`, `docx/`,
`charts/`, `benchmarks/`, `search-fixtures/`. Existing local outputs were moved
without overwriting files and SHA-256 verified; the local migration manifest
records their original and new paths. Old script paths are not compatibility
wrappers: use the dispatcher or the grouped paths below.

## Common tasks

| Goal | Command | Prerequisites / output |
|---|---|---|
| Check your machine | `bun run tools doctor` | Reports Bun/WebView, builds, Python packages and Office registration; never launches Office or installs anything |
| Review a real Office file | `bun run tools review <file>` | Windows, PowerShell 7, matching Office app, Bun 1.4+; writes native/aared captures and `review.html` |
| Measure browser performance | `bun run tools bench run` | Package + bench production builds; aared by default, saved JSON in `tools/out/benchmarks/` |
| Measure download size | `bun run tools bench size` | Same production builds; main/Worker/WASM download sizes |
| Measure browser search | `bun run tools bench search` | Same builds + generated large fixtures |
| Verify memory measurement | `bun run tools bench calibrate` | Known-size page/Worker controls; run after touching `shared/cdp.ts` |
| Regenerate benchmark tables | `bun run tools bench report` | Saved runs; **writes tracked** `docs/BENCHMARKS.md` and `site/src/data/benchmarks.json` |
| Inspect parsed data | `bun run tools inspect <file> --out <json>` | Headless Bun; resolved units/layouts/media hashes, no Office |
| Diagnose parser cost | `bun run tools inspect <file> --bench` | Headless parser timing/heap/RSS; not browser performance |
| Check production packaging | `bun run tools check consumer-build` | Package build, installed bench dependencies, WebView; Vite/webpack production probes |

Each command accepts `--help`. Unknown flags, invalid selectors and
format-inappropriate options fail before launching a backend. Measurements
retain error artifacts but return a failing exit code for failed/empty passes.

## Review a file

```sh
bun run build
bun run tools review fixtures/pptx-coverage.pptx
bun run tools review fixtures/xlsx-coverage.xlsx --part "1:title block=A1:E2"
bun run tools review fixtures/docx-coverage.docx
```

The final output prints the `review.html` path and a `file://` link. Arbitrary
files outside `fixtures/` use the same routes. The source file is not modified.

- **PPTX:** native slide PNGs and same-size DOM/SVG captures. `--width 1280`
  sets the native/capture width. The visual suite uses the same capture function.
- **XLSX:** native A4 print and aared A4 fit, visible sheets only. Repeat
  `--part <sheet>:<name>=<range>` for curated row bands, numbered from 1 among
  visible sheets. Ranges start at column A, use uppercase references, and must
  fit one printable page without cutting a merge/drawing. Parts write cropped
  PNG pairs by default; `--pair-full` adds full-sheet PNG pairs. Names are
  sanitized identically on both sides; collisions are rejected.
- **DOCX:** native pages + exact span/word geometry, aared layout pages, token
  comparison and review page. Requires `bun run build` first: this route serves
  `dist/` because source bundling has a known Windows root-lookup problem.
  Page-count differences stay visible in the report, not hidden by an error.
  Token distances are diagnostic evidence, not a pass/fail fidelity threshold.
- **All formats:** `--absent-fonts` adds a third column using the frozen
  substitute resolver (network required). XLSX parts get that column too.
  DOCX absent-font tokens have their own file and never replace normal tokens.

XLSX/DOCX rasterization needs Python with PyMuPDF
(`pip install "PyMuPDF==1.27.2"`); XLSX PNG pairs additionally need Pillow
(`pip install Pillow`). Native adapters keep their existing print/export
policies. Reviews are for human typography/bidi/wrapping inspection, never
pixel equality. Format-specific details:
[PowerPoint](review/pptx/README.md), [Excel](review/xlsx/README.md),
[Word](review/docx/README.md).

**Outputs live in `tools/out/<format>/<file-stem>/`. Each export replaces that
file's folder.** Same-stem files share a folder, as before; keep derived work
elsewhere. There is no bulk cleanup command, and no automatic probe deletion.

`bun run test:visual` selects visual **files**, not titles, so borrowed-content,
search, lifecycle and theme tests run even without "visual" in their names.
Run it after changes to fonts, split wrapping or viewer-content lifetimes;
plain `bun test` deliberately skips browser tests. The normal and artifacts
commands use the same selection. Every browser test gets its gate, page runner,
static server and bundled entry from `packages/aared/test/harness.ts`
(`visual`, `withVisualPage`, `serveVisual`, `browserModule`), so a new test
should not hand-roll that setup; see `AGENTS.md` for the rules and
`Notes §127` for why.

`bun run test:visual:artifacts` runs renderer DOM tests and fixture-specific
PPTX/XLSX native reviews. It writes review pages, not the standalone XLSX
command's cropped `pair-*.png` files; use `tools review` for those. Word's
native review remains an explicit command, not part of the ordinary test suite.

## Benchmarks

```sh
bun run build
bun run --cwd bench build
bun run tools bench calibrate
bun run tools bench run pptx-coverage.pptx:aared
bun run tools bench size
bun run tools bench report --md       # only docs/BENCHMARKS.md
bun run tools bench report --site     # only site/src/data/benchmarks.json
```

`run` and `size` update aared's cells in the latest compatible same-machine
run, retaining competitor measurements and their dates. `--all` measures every
renderer into a new run. Explicit fixture/renderer selectors are validated;
missing explicitly selected fixtures fail rather than silently measuring none.
Optional ungenerated large files are reported as skips in a normal open pass.

Benchmarks are manual, on an idle machine on mains power. They are not routine
verification or CI. The method and interpretation live in
[docs/BENCHMARKS.md](../docs/BENCHMARKS.md).

## Specialist tools (kept separate deliberately)

| Tool | Role / status |
|---|---|
| `review/office-export.ts --format ...` | Native export only; callable driver and CLI used by review/test backends |
| `review/docx/capture.ts`, `compare-pagination.ts`, `review.ts` | Individual DOCX phases for debugging; capture defaults its output folder from the input filename |
| `review/xlsx.ts` | Individual XLSX review backend; same function used by the dispatcher |
| `checks/visual-tests.ts` | Sets the visual/artifact/absent-font test environment; rejects unknown flags |
| `benchmarks/search/make-fixtures.py` | Generates large timing files in `out/search-fixtures/` from real fixtures; not fidelity fixtures |
| `benchmarks/search/bench.ts` | In-process search diagnostics: abort latency and RSS as well as timings; **not** the browser benchmark feeding published tables |
| `benchmarks/group-effects.ts` | One-shot warm PPTX render/zoom comparison; results in `out/benchmarks/group-effects-performance*.json` |
| `chart-oracle/` | Specialized chart corpus, vector-PDF scoring and metrics generation; [its own workflow](chart-oracle/README.md), not A4 document review. Outputs go to `out/charts/`; refreshing tracked corpora/generated source is explicit |
| `font-calibration/` | **Frozen research** (Notes §53); scripts and pinned evidence remain available, no new calibration work; [details](font-calibration/README.md) |
| `probe/` | Ignored, topic-local Office investigations; [lifecycle and shared helpers](probe/README.md). Do not remove unfinished or another contributor's probes |

The chart oracle's tracked Office files are test inputs, not disposable output.
Font PDFs/JSON/TSV are pinned research evidence, not ordinary benchmark runs.
`probe/_common/office.py` and `strip_label.py` are live shared utilities.

## Verification and ownership

`bun run typecheck` includes `tools/tsconfig.json`; `bun run typecheck:tools`
checks tools alone. It covers tracked TypeScript backends/research, excludes
ignored probe/output trees, and does not execute any tool. Pure CLI/report
tests live alongside tools and run under `bun test`. Browser regression tests
under `tools/review/` run only with the visual test flag.

Common naming/process utilities live in `shared/office.ts`, stable roots in
`shared/paths.ts`, review chrome in `review/page.ts`, and the reused PPTX
capture in `review/pptx.ts`. `shared/cdp.ts` stays separate because page/Worker
DevTools routing is load-bearing. A layout regression test keeps the root free
of implementation scripts and checks that generated output remains ignored.
