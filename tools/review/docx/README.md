# DOCX review backend

Exports a real `.docx` through the installed Microsoft Word, not LibreOffice,
and pairs each native page with aared's layout render at the same pixel size.

```powershell
bun run build
bun run tools review fixtures/docx-coverage.docx
```

The command exports, captures, compares tokens and writes `review.html` into
`tools/out/docx/docx-coverage/`: Word's export beside
aared's capture, one row per physical page, at the same display width, with
`word-export.json` and `pagination-comparison.json` folded in when they are
there. Open it and judge the pairs. The PPTX and XLSX reviews write the same
page from inside `bun test`; this one is a hand-run command, because the DOCX
pipeline needs Word (AGENTS.md, Testing rule). `--absent-fonts` adds a third
capture while preserving the normal token file (`aared-absent-words.json` is
separate). Aared/Word page-count differences are shown rather than preventing
the report from being written. See [the tools catalog](../../README.md).

For individual phases while debugging:

```sh
bun tools/review/office-export.ts --format docx fixtures/docx-coverage.docx
bun tools/review/docx/capture.ts fixtures/docx-coverage.docx
bun tools/review/docx/compare-pagination.ts tools/out/docx/docx-coverage
bun tools/review/docx/review.ts tools/out/docx/docx-coverage
```

Capture defaults to the export's filename-derived directory; an explicit
output directory is still supported for custom reviews.

This sequence was run successfully from ordinary Windows PowerShell 7.6.6 on
October 6, 2026, with Word 16.0 and PyMuPDF 1.27.2.3. Bash is not required.

On the maintainer's Codex setup, Word export and browser capture need the
command runner's approved execution outside the agent sandbox: inside that
sandbox, Word COM returned `E_ACCESSDENIED` and the WebView browser closed
its pipe, while rasterization/comparison succeeded inside it (verified October
6). That sandbox caveat is specific to that setup — an agent running directly
on a Windows terminal runs the same native automation commands in place
(verified under OpenCode: WebView launch plus Word 16.0 COM launch/quit). Use
scoped command approval where a sandbox blocks them; changing shells or
disabling the whole sandbox is unnecessary.

The build command emits the renderer and shared chunks that the browser capture
serves directly. This avoids `Bun.build()`'s inferred-root lookup, which fails
with `EBADF` on this Windows checkout. The Word export writes
`word-page-N.png`, `word-spans.json`, `word-words.json` and `word-export.json`
into the gitignored `out/` directory. The aared capture writes
`aared-page-N.png` and `aared-words.json` beside them; review each image pair
side by side. The comparison command writes `pagination-comparison.json` and
reports matched-token vertical drift, page mismatches, missing/extra tokens,
and repeated-token groups. Matches use full normalized tokens and pair repeated
occurrences by page position instead of truncating text to a prefix.

Treat the comparison as a diagnostic report, not an automatic fidelity gate.
Repeated occurrences are paired without surrounding context. PDF extraction
and DOM text nodes can tokenize the same rendered text differently (spaced
header letters, Arabic strings, punctuation and runs), changing occurrence
counts and shifting pairs. Page mismatches and missing/extra groups therefore
need confirmation from the paired images and unambiguous text anchors; they
do not by themselves prove content moved or disappeared. The October 6 capture
reported 25 page mismatches, all involving repeated tokens, while all 435
tokens occurring exactly once in both exports stayed on the same page.

Add `--absent-fonts` to the capture command to write
`aared-absent-page-N.png` instead: the document as a device without Office's
fonts paints it, through the real `aared-fonts` resolver (network;
Notes §51). Requires Windows, Word, PowerShell 7 and Python with PyMuPDF
(`pip install "PyMuPDF==1.27.2"`).

## Method

`export.ps1` opens a private copy of the file read-only (no lock file beside
the fixture, macros force-disabled, dialogs suppressed, never saved), updates
fields so `PAGE`/`NUMPAGES`/TOC numbers print as Word computes them, and
exports a PDF. `rasterize.py` renders it with PyMuPDF and dumps every text
span and token box with its font, size and position in CSS pixels.

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
