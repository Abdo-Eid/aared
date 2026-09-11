# Fixtures

Real generated Office files, not synthetic mocks. See
[`AGENTS.md § Testing rule`](../AGENTS.md#testing-rule). The files are seeded
from the archived evaluation work or authored in PowerPoint; the archive remains
independent provenance, not the only copy.

| File | Format | Notes |
| --- | --- | --- |
| `arabic-pptx-rtl-coverage.pptx` | PPTX | **M1 gate fixture.** One real PowerPoint-authored 4-slide deck covering RTL text/bidi, an RTL table, common shapes plus future ink/image coverage, and gradients. See notes below. |
| `arabic-pptx-rtl-coverage.json` | Fixture manifest | Semantic selectors for the PPTX test suite. It locates coverage by unique authored text, not slide or shape indexes. |
| `post-production-review.pptx` | PPTX | Mixed English/Arabic production deck (7 slides, tables on 2) for broader PPTX variety. |
| `post-production-tracker.xlsx` | XLSX | Generic workbook, no Arabic content yet. |
| `arabic-xlsx-rtl-coverage.xlsx` | XLSX | **XLSX gate fixture.** Excel-assembled Arabic coverage workbook: RTL grid (styled cells, merges, frozen panes, hidden rows/columns, number formats, hyperlink), LTR summary, deep Shot Log (~100 rows for virtualization, Cost formulas live against the Notes rate), and Notes. Assembled in Excel from real Excel-owned sheets; keeps the reserved `fills[1]` (`gray125`) slot — a generated predecessor once put a custom fill there and Excel repaired the file for it. |
| `arabic-xlsx-rtl-coverage.json` | Fixture manifest | Semantic selectors for the XLSX test suite. It locates coverage sheets by unique authored text, not sheet indexes. |
| `arabic-docx-bidi-coverage.docx` | DOCX | **DOCX gate fixture.** Real Word-authored bilingual brief (Arabic + Latin): 17 mixed Arabic+Latin paragraphs, RTL paragraphs, a Heading 1–2 outline, 3 tables (fills, `gridSpan`/`vMerge`), bullet + decimal lists with nested levels, 2 inline images, 2 footnotes, headers/footers, internal/external hyperlinks. |
| `arabic-docx-bidi-coverage.json` | Fixture manifest | Semantic selectors for the DOCX test suite. It locates coverage by unique authored text, not block indexes. |
| `delivery-report.pdf` | PDF | Out of aared's scope; retained for LibreOffice sanity checks. |

**Coverage notes** (`arabic-pptx-rtl-coverage.pptx`, verified by unzipping):

- Slides 1 and 2 retain the PowerPoint-authored Arabic text/table gate. Table
  direction is a flag, not rewritten column order: `<a:tblPr rtl="1"
  firstRow="1" bandRow="1">`. Grid columns stay in logical XML order and
  display rightmost. Cells explicitly carry `rtl="0"`, while their table is
  RTL. Slide 1 includes RTL bullets, mixed bidi text, and an English line with
  `algn="l" rtl="1"`; bullet side follows direction, never alignment.
- Theme font references (`+mn-lt` / `+mn-cs`) and layout/master text defaults
  remain in the first two slides. They exercise theme lookup and placeholder
  inheritance rather than literal typefaces.
- Slide 3 contains common supported presets, solid fills/outlines, rotation,
  and authored adjustments. It intentionally also carries unsupported `star5`,
  arrow, not-equal, brace, custom freeform, pen ink, and an image for future
  coverage. These must remain visible fallbacks or explicit deferrals, never
  silently pass.
- Slide 4 contains a linear gradient and a radial/path gradient. The linear
  gradient is rendered; the radial/path gradient is deliberately marked
  unsupported until that fill path exists.
- Native PowerPoint PNGs and same-size WebView captures include every slide.
  Run `bun run test:visual:artifacts` and inspect its paired review page for
  text fidelity; LibreOffice conversion is available as a companion artifact
  but does not gate Arabic typography. Slides 3 and 4 are also covered by
  parser and DOM/SVG tests while their image, ink, custom-geometry, and
  path-gradient portions remain deferred.
- The deck closes/reopens cleanly in PowerPoint.
- The sidecar manifest is a test-only contract. When the deck is rebuilt, keep
  its selector text unique; tests reject missing or ambiguous selectors and
  include candidate summaries to make fixture drift actionable.

**Coverage notes** (`arabic-docx-bidi-coverage.docx`, verified by unzipping
and by the DOCX parser suite):

- Mixed-script paragraphs keep one run per script with the correct typeface
  slot: Latin runs resolve `w:ascii`, Arabic runs `w:cs` (e.g. Cambria +
  Arial inside one heading). Paragraph direction follows `w:bidi` only —
  mixed paragraphs without it stay LTR and let browser bidi place the runs.
- Lists resolve `w:numId`/`w:ilvl` through `numbering.xml`: em-dash bullets,
  decimal `1.`/`a)` levels, per-level indents, ordinals continuing across the
  section. Duplicate `Heading1`/`Heading2` style blocks resolve last-wins.
- Both images are `wp:inline` (no `wp:anchor` float sample exists in any
  Word-authored file here) and both footnotes are referenced in document
  order. Tracked changes have no Word-authored sample either: floats and
  `w:ins`/`w:del` are pinned by synthetic parser unit tests until one lands.
- `sectPr` carries landscape `pgSz` plus default header/footer references;
  flow mode paints the header/footer once, while layout mode resolves real
  page geometry and per-section headers/footers onto physical pages.
