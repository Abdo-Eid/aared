# tools/word-render

Exports a real `.docx` through the installed Microsoft Word, not LibreOffice,
and pairs each native page with aared's layout render at the same pixel size.

```bash
bun tools/word-render/index.ts fixtures/docx-coverage.docx
bun tools/word-render/capture.ts fixtures/docx-coverage.docx \
  tools/word-render/out/docx-coverage
```

The first command writes `word-page-N.png` (96 dpi, so CSS pixels),
`word-spans.json` and `word-export.json` into the gitignored `out/` directory.
The second writes `aared-page-N.png` beside them; review each pair
side by side. Requires Windows, Word, PowerShell 7 and Python with PyMuPDF
(`pip install "PyMuPDF==1.27.2"`).

## Method

`export.ps1` opens a private copy of the file read-only (no lock file beside
the fixture, macros force-disabled, dialogs suppressed, never saved), updates
fields so `PAGE`/`NUMPAGES`/TOC numbers print as Word computes them, and
exports a PDF. `rasterize.py` renders it with PyMuPDF and also dumps every
text span with its font, size and position in CSS pixels.

That second output is the point: **read a layout number from
`word-spans.json` rather than judging a screenshot.** A superscript's size and
raise, a heading's baseline, a marker's gap were all settled this way (for
example `PAL` is 6.96pt on a 10.56pt line, raised 4.7px; an appendix heading's
baseline is 24px below a post-break heading's).

## What the reference proves (and does not)

- Proves: page geometry, header/footer placement, line boxes, tab stops,
  table edges and padding, list marker geometry, paragraph spacing rules —
  the layout truth the renderer mimics.
- Does not prove: exact glyph raster (antialiasing differs between engines),
  fonts on a machine that lacks them, or how the page looks under a host
  page's CSS reset — check that in the bench.

## Determinism

Bytes are nondeterministic across machines (Word version, PyMuPDF version and
fonts all intrude; the first two are recorded in `word-export.json`). Assert
presence and sanity, never hashes. LibreOffice conversion is never the DOCX
oracle.
