# Fixtures

Real generated Office files, not synthetic mocks. See
[`AGENTS.md § Testing rule`](../AGENTS.md#testing-rule). The files are seeded
from the archived evaluation work or authored in PowerPoint; the archive remains
independent provenance, not the only copy.

| File | Format | Notes |
| --- | --- | --- |
| `pptx-coverage.pptx` | PPTX | **M1 gate fixture.** One real PowerPoint-authored 10-slide deck covering RTL text/bidi, an RTL table, common shapes plus ink/image coverage, gradients, bullets, WordArt, and shape/background/table-cell fills. See notes below. |
| `pptx-coverage.json` | Fixture manifest | Semantic selectors for the PPTX test suite. It locates coverage by unique authored text, not slide or shape indexes. |
| `post-production-review.pptx` | PPTX | Mixed English/Arabic production deck (7 slides, tables on 2) for broader PPTX variety. |
| `post-production-tracker.xlsx` | XLSX | Generic workbook, no Arabic content yet. |
| `xlsx-coverage.xlsx` | XLSX | **XLSX gate fixture.** Excel-assembled Arabic coverage workbook: RTL grid (styled cells, merges, frozen panes, hidden rows/columns, number formats, hyperlink), LTR summary, deep Shot Log (~100 rows for virtualization, Cost formulas live against the Notes rate), and Notes. Assembled in Excel from real Excel-owned sheets; keeps the reserved `fills[1]` (`gray125`) slot — a generated predecessor once put a custom fill there and Excel repaired the file for it. |
| `xlsx-coverage.json` | Fixture manifest | Semantic selectors for the XLSX test suite. It locates coverage sheets by unique authored text, not sheet indexes. |
| `docx-coverage.docx` | DOCX | **DOCX gate fixture.** Real Word-authored bilingual brief (Arabic + Latin): 17 mixed Arabic+Latin paragraphs, RTL paragraphs, a Heading 1–2 outline, 3 tables (fills, `gridSpan`/`vMerge`), bullet + decimal lists with nested levels, 2 inline images, 2 footnotes, headers/footers, internal/external hyperlinks. |
| `docx-coverage.json` | Fixture manifest | Semantic selectors for the DOCX test suite. It locates coverage by unique authored text, not block indexes. |
| `docx-outline.docx` | DOCX | Word-authored (COM automation) heading-outline fixture, bilingual: title property "Outline Coverage / تغطية المخطط"; body outline Introduction / مقدمة (H1), Details / التفاصيل (H2), Closing / الخاتمة (H1) across a page break; a Heading-styled paragraph in the header and a Heading 2 inside a table cell — both outside the outline. Pins the one heading index space (`headings.list()`, `goToHeading`, search) and `goToPage`. |
| `delivery-report.pdf` | PDF | Out of aared's scope; retained for LibreOffice sanity checks. |

**Coverage notes** (`pptx-coverage.pptx`, verified by unzipping):

- Slides 1 and 2 retain the PowerPoint-authored Arabic text/table gate. Table
  direction is a flag, not rewritten column order: `<a:tblPr rtl="1"
  firstRow="1" bandRow="1">`. Grid columns stay in logical XML order and
  display rightmost. Cells explicitly carry `rtl="0"`, while their table is
  RTL. Slide 1 includes RTL bullets, mixed bidi text, and an English line with
  `algn="l" rtl="1"`; bullet side follows direction, never alignment.
- Theme font references (`+mn-lt` / `+mn-cs`) and layout/master text defaults
  remain in the first two slides. They exercise theme lookup and placeholder
  inheritance rather than literal typefaces.
- Slide 3 contains common supported presets (ECMA-formula paths, star5
  included), the preset-alias `upArrow`, and a freeform `custGeom` scribble
  that the deferred-geometry engine renders (see DECISION_LOG §26). Every shape on the slide renders now (mathNotEqual and rightBrace resolve through the catalog); pen ink and the photo remain for their own coverage.
  These must remain visible fallbacks or explicit deferrals, never silently
  pass.
- Slide 4 contains a linear gradient and a centered radial gradient
  (`path="circle"`, `fillToRect` 50/50/50/50); both are rendered. Slide 6's
  gradient WordArt uses the same centered radial fill.
- Slides 7–10 were appended through PowerPoint's COM object model (real
  Office writing every fill, not hand-built XML) to cover fills no other
  slide carries: slide 7 shape fills (`pattFill` `dkDnDiag` and `smGrid`, a
  stretched picture fill, a tiled picture fill at 25 %, an off-center
  `path="rect"` gradient and a centered `path="rect"` gradient), slide 8 a
  `pattFill` background, slide 9 a tiled picture background, slide 10 an RTL
  table whose cells carry linear-gradient, `pattFill` and `path="rect"` fills.
  COM cannot author an off-center `path="circle"` or a `path="shape"`
  gradient; those are added by hand in PowerPoint.
- Native PowerPoint PNGs and same-size WebView captures include every slide.
  Run `bun run test:visual:artifacts` and inspect its paired review page for
  text fidelity; LibreOffice conversion is available as a companion artifact
  but does not gate Arabic typography. Slides 3 and 4 are also covered by
  parser and DOM/SVG tests while their image, ink, custom-geometry, and
  path-gradient portions remain deferred.
- The deck closes/reopens cleanly in PowerPoint.
- Strip the org sensitivity label (Purview "Public") before saving: the label
  pill is app chrome, but a saved label adds marking parts/metadata to the
  package and risks baked-in marking text. Slide 3's Group 2 is the
  PowerPoint-grouped rect pair (identity child space); Group 3 is the
  hexagon/octagon pair, rotated 20° and scaled ~1.4× (non-identity child
  space) — both pinned by `pptx-group.test.ts` and the group visual test, so
  re-grouping or moving them means re-baselining those tests.
- Slide 3's brace pair pins style-reference resolution: the right brace is
  unfilled (`fillRef idx="0"` with a vestigial accent1 child — PowerPoint
  paints no fill), the left brace carries a direct black fill and red
  outline. See DECISION_LOG §27; recoloring or moving them means
  re-baselining the idx-0 tests.
- The sidecar manifest is a test-only contract. When the deck is rebuilt, keep
  its selector text unique; tests reject missing or ambiguous selectors and
  include candidate summaries to make fixture drift actionable.

**Coverage notes** (`docx-coverage.docx`, verified by unzipping
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
