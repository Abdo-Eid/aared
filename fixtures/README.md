# Fixtures

Real Office-authored files, not synthetic mocks. See
[`AGENTS.md § Testing rule`](../AGENTS.md#testing-rule).

**Authoring rules.**

- Every file is written by real Office, never generated: a generated
  workbook once put a custom fill in the reserved `gray125` slot and Excel
  "repaired" the file for it.
- New coverage goes into the existing coverage fixtures through Office's COM
  object model; no new fixture file without sign-off.
- Scripts that drive Office close documents with `Close(0)` and `Quit(0)` in
  `finally` (Word's `Quit` takes `[ref]0` in Windows PowerShell), or a failed
  run leaves "Save your changes?" dialogs on screen.
- Every Office save re-applies the org sensitivity label (custom
  properties, a "Public" marking, and in Word even/first-page headers created
  to carry it): run `strip_label.py <file>` on the fixture afterwards. The
  script (and the `extend.py` probes the rows below name) live in the
  gitignored `tools/probe/`: local to the maintainer's machine, not in a clone.
- Freeze the fixture before comparing screenshots: save it, check the lock
  file is gone, and re-verify the page margins and key paragraphs. Unsaved
  Office state produces ghost differences.

| File | Format | Notes |
| --- | --- | --- |
| `pptx-coverage.pptx` | PPTX | One real PowerPoint-authored 18-slide deck covering RTL text/bidi, an RTL table, common shapes plus ink/image coverage, gradients, bullets, WordArt, shape/background/table-cell fills, line/connector shapes, merged table cells, an Office 2016 waterfall chart, and a second slide master (logo, picture background, own theme, two layouts). See notes below. |
| `pptx-coverage.json` | Fixture manifest | Semantic selectors for the PPTX test suite. It locates coverage by unique authored text, not slide or shape indexes. |
| `post-production-review.pptx` | PPTX | Mixed English/Arabic production deck (7 slides, tables on 2) for broader PPTX variety. |
| `post-production-tracker.xlsx` | XLSX | Bilingual English/Arabic workbook, four sheets: a Dashboard (KPI tiles, merges, a bar chart), a Shot Log (frozen rows and columns), a right-to-left budget sheet ("الميزانية", frozen header rows) and Notes; 59 Arabic cells in all. Also the bench and site's sample workbook. |
| `xlsx-coverage.xlsx` | XLSX | Excel-assembled Arabic coverage workbook: RTL grid (styled cells, merges, frozen panes, hidden rows/columns, number formats, hyperlink), LTR summary, deep Shot Log (~100 rows for virtualization, Cost formulas live against the Notes rate), and Notes. Assembled in Excel from real Excel-owned sheets; keeps the reserved `fills[1]` (`gray125`) slot — a generated predecessor once put a custom fill there and Excel repaired the file for it. |
| `xlsx-coverage.json` | Fixture manifest | Semantic selectors for the XLSX test suite. It locates coverage sheets by unique authored text, not sheet indexes. |
| `docx-coverage.docx` | DOCX | Real Word-authored bilingual brief (Arabic + Latin): 17 mixed Arabic+Latin paragraphs, RTL paragraphs, a Heading 1–2 outline, 3 tables (fills, `gridSpan`/`vMerge`), bullet + decimal lists with nested levels, 2 inline images, 2 footnotes, headers/footers, internal/external hyperlinks. Section 5 ("5. Delivery tracker / متابعة التسليم") was appended through Word's COM object model: a right-to-left table (`w:bidiVisual`) with a 3pt border set on its physical right (Word writes `w:left`, the start side), 0.5pt on its left (`w:right`), cell margins 0.3in left / 0.05in right (Word keeps them physical; DECISION_LOG §43), and a merged total row styled Heading 2; the appendix header carries a Heading 1 paragraph. A right-to-left paragraph before the tracker, "آخر تحديث 15 Oct 2026 بنسبة إنجاز 71%", has left-to-right runs for the date and percentage (no `w:rtl`; DECISION_LOG §44), set in Calibri beside the Arial Arabic at exactly 18pt line spacing (the exact-pitch case; `tools/probe/exact-pitch/extend.py`). Both headings sit outside the outline and pin the one heading index space (`headings.list()`, `goToHeading`). The org sensitivity label Word applied on save was stripped (custom properties, header marking, and the even/first-page headers Word adds only to carry it: `tools/probe/_common/strip_label.py`). The document's `compatibilityMode` is 15 (Word 2013+): two justified English paragraphs in the synopsis ("Sawt al-Nahr follows three families…", "Dates below are hard for the picture lock…") set `DocxParagraphLayout.justifyShrink` (DECISION_LOG §44). One portrait page after the landscape appendix holds seven anchored pictures on Arabic paragraphs, added through COM (`tools/probe/docx-float-wrap/extend.py`, label stripped): square at the left and right column edges, top and bottom and behind text centred, tight at the edge, and square and through centred with text down both sides (DECISION_LOG §55). The sixth, landscape page holds two charts on Arabic paragraphs, added through COM with Word's own chart (`AddChart2`, embedded workbook; `tools/probe/docx-charts/extend.py`, label stripped): "مخطط مضمن في سطره" with an inline 5 x 3in chart in a centred paragraph after it, and "مخطط بالتفاف مربع" with a 3.5 x 2.25in chart anchored at the left column edge, square wrap; both clustered columns with Arabic categories, series names and title (DECISION_LOG §71). A last portrait section ("8. Lists, headers and empty lines / القوائم والرؤوس والأسطر الفارغة", `tools/probe/docx-lists-headers/extend.py`, label stripped) turns on a different first page (`w:titlePg`) and writes only that first-page header, and has no footer; its single page holds three numbered lists (three items, a plain paragraph, two items that continue the first list, then a list Word restarts at 1 with a start override: markers 1-5, then 1-2), and a table whose second row has a cell holding only an empty paragraph (DECISION_LOG §76). The document is seven pages in Word after merging the picture-wrap examples onto portrait page 5 and the final lists/table content onto page 7. Square and tight edge wrapping remain separate examples because they exercise different wrap modes. A final anchored box was added through Word COM with Arabic text "صندوق ملاحظات: مراجعة التسليم" and the Latin paragraph "English note: confirm the final delivery.", a pale green fill, and a green outline. Word writes it as `mc:Choice` DrawingML plus a VML `mc:Fallback`, so it covers branch selection without duplicate text and keeps the mixed-script content searchable. Even-page headers are not in it (`w:evenAndOddHeaders` is document-wide and would blank the even pages of every earlier section); the tests build them synthetically. |
| `docx-coverage.json` | Fixture manifest | Semantic selectors for the DOCX test suite. It locates coverage by unique authored text, not block indexes. |
| `delivery-report.pdf` | PDF | Out of aared's scope; the bench's PDF control sample. |

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
  that the deferred-geometry engine renders (see DECISION_LOG §26). Every shape
  on the slide renders (mathNotEqual and rightBrace resolve through the
  catalog), pen ink and the photo included. A shape that could not paint must
  show as a marked fallback, never silently pass.
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
- Slide 11 ("خطوط وموصلات") was appended the same way (a duplicate of
  slide 10 with the table removed): four line shapes PowerPoint writes as
  `p:cxnSp` — a straight connector (`straightConnector1`, zero height), a
  dashed vertical line (zero width), a dotted line and a `flipV` diagonal; plus an RTL text box "الإيرادات دون الموازنة بنسبة 2.1%" whose number is an en-US run (DECISION_LOG §44),
  spaced exactly 30pt with the number in Calibri (the exact-pitch case).
  PowerPoint applied the org sensitivity label on save; its custom
  properties and master "Public" marking were stripped from the package.
- Slide 12 ("تباعد الفقرات") was appended the same way (a duplicate of slide
  11 with everything but the title and slide number removed): an RTL text
  box and a one-cell table whose paragraphs are spaced in lines
  (`a:spcBef`/`a:spcAft` with `a:spcPct`): a line before the first
  paragraph and after the last (PowerPoint shows neither), half a line
  before, a line before a paragraph holding a 28pt word in 18pt text (the
  line counts 28pt), and a line after one paragraph plus a line before the
  next (they add up). DECISION_LOG §47.
- Slides 13 and 14 were appended through PowerPoint's object model
  (`tools/probe/pptx-merged-chartex/extend.py`, label stripped). Slide 13
  ("خلايا مدمجة", a duplicate of slide 12 with its body removed) holds a 4 x 4
  table whose first two header cells are merged across (`gridSpan="2"` on the
  cell, `hMerge="1"` on the one it covers) and whose third-column cell in row 2
  is merged down two rows (`rowSpan="2"`, `vMerge="1"`). Slide 14
  ("الشلال: الميزانية") holds Excel's own waterfall chart from
  `tools/chart-oracle/chartex-oracle.xlsx`, copied and pasted into PowerPoint
  (its `AddChart2` refuses the 2016 chart types), so the deck carries the
  `chartEx` part PowerPoint writes, with its data cached as `cx:lvl` (five
  Arabic categories, values 100, 60, -40, -15, 105).
- Slides 15 to 18 sit on a second slide master added through PowerPoint
  (`tools/probe/pptx-second-master/extend.py`, label stripped), so its graphics
  show on these slides only. The master has its own theme (accent 1 a dark
  green `1E6B3C`, the first master's is blue `156082`), a "Master logo"
  rectangle ("شعار الاستوديو"), and a picture background (a pale sand PNG). It
  has two layouts, each with one graphic of its own and only its title and
  slide-number placeholders: layout A ("رسم التخطيط أ", shows the master's
  graphics) and layout B ("رسم التخطيط ب", `showMasterSp="0"`). Slide 15 and
  slide 16 are on A (slide 15 also has an "Accent box" filled with theme accent
  1: "لون التمييز الأول"), slide 17 is on B, and slide 18 is on A with its own
  `showMasterSp="0"` (neither graphic). `pptx-layout.test.ts` and
  `pptx-layout.visual.test.ts` pin them.
- Native PowerPoint PNGs and same-size WebView captures include every slide.
  Run `bun run test:visual:artifacts` and inspect its paired review page for
  text fidelity; LibreOffice conversion is available as a companion artifact
  but does not gate Arabic typography. Slides 3 and 4 are also covered by
  parser and DOM/SVG tests.
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
- The two original images are `wp:inline`, and both footnotes are referenced
  in document order. The seven anchored pictures and the two charts are
  described in the fixture's table row. Tracked changes have no Word-authored
  sample: `w:ins`/`w:del` are pinned by synthetic parser unit tests until one
  lands.
- `sectPr` carries landscape `pgSz` plus default header/footer references;
  flow mode paints the header/footer once, while layout mode resolves real
  page geometry and per-section headers/footers onto physical pages.
