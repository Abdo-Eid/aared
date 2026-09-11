# bench

Visual bench for aared: open real Office/PDF files and **see** them, with the
open speed on screen. Baseline third-party renderers today; aared's own
renderers drop into the second column as M1–M3 land.

```bash
bun install          # from the repo root (workspace)
bun run dev          # serves bench/ at http://localhost:4321
```

- Left rail: one-click sample buttons from canonical `../fixtures/` (Arabic deck
  first), file upload + drag-drop, per-format "what to look for" checks, and
  a run history with milliseconds.
- For PPTX, Single mode defaults to aared and exposes an aared/Baseline
  selector. Compare renders both routes side by side. XLSX and DOCX stay
  baseline-only until their aared renderers exist.
- Timing is `performance.now()` from bytes-in-hand to first meaningful
  paint, shown in the footer chip and the run history.

Known baseline gaps (not bench bugs): pptxviewjs paints canvas with no text
layer and drops paragraph-level RTL (guarded so it can't clobber
`window.FileReader` anymore); SheetJS CE drops fills/fonts/borders;
per-format notes live in the checks panel.
