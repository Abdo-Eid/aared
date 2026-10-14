# XLSX review backend

Exports a real `.xlsx` through the installed Microsoft Excel application,
not LibreOffice: every visible sheet prints to A4, one page per sheet.

```bash
bun tools/review/office-export.ts --format xlsx fixtures/xlsx-coverage.xlsx
```

Options: `--dpi <dpi>` sets the raster resolution (default 150), and `--part
<sheet>:<name>=<range>` (repeatable) also prints one named range of a sheet
as its own page. `<sheet>` is the 1-based position among the visible sheets,
so a hidden sheet is not counted.

## Paired capture of any workbook

To compare a workbook with aared outside the fixture-only visual suite, use
`bun run tools review <file.xlsx>`. It runs the same native Excel export, renders each
sheet through aared, and writes a `review.html` alongside the captures:

```bash
bun run tools review path/to/workbook.xlsx
```

Full sheets are captured at the review's A4 fit. For a readable comparison of
one area, request a part; its Excel and aared captures are trimmed to their
content and composed side by side into `pair-sheet-<sheet>-part-<name>.png`:

```bash
bun run tools review path/to/workbook.xlsx --part 1:top=A1:I17
```

Part sheet numbers are 1-based among visible sheets. Repeat `--part` to
capture multiple areas. Part pairs are produced by default; pass `--pair-full`
to also create side-by-side images of the complete sheets. Complete-sheet
pairs can be difficult to read for long sheets, so a curated part is usually
the useful comparison. `--absent-fonts` adds captures using the aared font
substitutes (network access required), including the curated parts.
The lower-level `tools/review/xlsx.ts` remains the same callable backend.
See [the tools catalog](../../README.md) for shared commands and prerequisites.

The capture requires Windows, Excel, PowerShell 7, and Python with PyMuPDF.
Composing cropped pairs requires Pillow (`pip install Pillow`). Results go to
the gitignored `tools/out/xlsx/<workbook-stem>/` directory, which is
cleared before each export. Keep derived images outside that directory if they
must survive a later capture.

The native export requires Windows, Excel, PowerShell 7, and Python with
PyMuPDF (`pip install "PyMuPDF==1.27.2"`). It opens the file read-only (link
updates off, macros force-disabled, dialogs suppressed, never saved), prints
into the gitignored `out/` directory, then closes and releases Excel. The
companion visual suite pairs each native page with an aared A4 fit in a
side-by-side review page:

```bash
bun run test:visual --artifacts
```

## Method

Print policy is hardcoded in `export.ps1` — one consistent A4 truth owned
by tooling, not read from the file: A4 portrait, fit sheet on one page
(shrink only, never stretch), gridlines plus row/column headings, normal
margins. Each sheet exports with `ExportAsFixedFormat` to PDF, then
`rasterize.py` renders it with PyMuPDF (MuPDF engine) at 150 DPI. The
metadata records the Excel version, the PyMuPDF version, and per-sheet
pixel dimensions.

An earlier revision captured the used range with `CopyPicture` instead.
That proved the screen grid but could never produce pages, needed a
visible Excel window (the screen pipeline captures blank when hidden),
and answered a different question than the A4 review. The print path
renders hidden, needs no window, and matches what the review compares.

## What the reference proves (and does not)

- Proves: merges, fills, borders, number formats, column widths/row heights,
  hidden row/column collapse, RTL mirroring, fit-to-page shrink — the print
  truth the aared A4 fit mimics.
- Does not prove: exact text raster (native pages are rasterized at
  150 DPI while browser text is shaped by the browser — expect
  antialiasing deltas; human review only, never pixel diffs), on-screen
  scrolling behavior, frozen-pane stickiness (freeze affects scrolling
  only; print captures the full range).
- The printed page header (e.g. "Public") has no aared counterpart yet;
  printed page footers would land the same way. Sheets with print titles
  (Shot Log repeats row 4) repeat that row on every native part page;
  the aared band shows the range alone.

## Determinism

Bytes are documented-nondeterministic across machines: Excel version,
PyMuPDF version, and fonts all intrude — all three are recorded in
`excel-export.json`. Assert presence and sanity, never hashes.
LibreOffice conversion remains a companion artifact only and is never the
XLSX oracle.
