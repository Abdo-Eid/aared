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
| `pptx-coverage.pptx` | PPTX | One real PowerPoint-authored 20-slide deck covering RTL text/bidi, an RTL table, common shapes plus ink/image coverage, gradients, bullets, WordArt, shape/background/table-cell fills, line/connector shapes, merged table cells, an Office 2016 waterfall chart, a second slide master (logo, picture background, own theme, two layouts), and group appearance. See notes below. |
| `pptx-coverage.json` | Fixture manifest | Semantic selectors for the PPTX test suite. It locates coverage by unique authored text, not slide or shape indexes. |
| `post-production-review.pptx` | PPTX | Mixed English/Arabic production deck (7 slides, tables on 2) for broader PPTX variety. |
| `post-production-tracker.xlsx` | XLSX | Bilingual English/Arabic workbook, four sheets: a Dashboard (KPI tiles, merges, a bar chart), a Shot Log (frozen rows and columns), a right-to-left budget sheet ("الميزانية", frozen header rows) and Notes; 59 Arabic cells in all. Also the bench and site's sample workbook. |
| `xlsx-coverage.xlsx` | XLSX | Excel-assembled Arabic coverage workbook: RTL grid (styled cells, merges, frozen panes, hidden rows/columns, number formats, hyperlink), an RTL bilingual sales Summary (redesigned through Excel COM, completed `xlsx-summary-redesign` authoring probe, label stripped; Notes §116): a merged two-row teal title band over the table at A4:G9, frozen under its header, in Segoe UI with SKUs in Consolas, every cell explicitly aligned, single top/bottom rules only, banded rows, gridlines off; same-sheet formulas (`E5=D5/C5-1`, `F5` status, `C9=SUM(C5:C8)`) exercise formula-reference highlighting; positive/negative change values use native conditional font colors and the status column three `cellIs`-equal chip fills; the total row's centred status beside an empty G9 pins centred overflow, and a merged footnote sits at A11, deep Shot Log (~100 rows for virtualization, Cost formulas live against the Notes rate), Notes, and a COM-authored `CF Rules` sheet (3Arrows icons on 1–10, a today rule on 2020 dates, 1-deviation above-average on 10,9,8,8,8,7,6,5,4,3; Notes §107, label stripped; three-colour scale on E2:E12, §109; expression `=MOD($F2,2)=0` on F2:F12 over 1..11 painting evens, §110; TODAY/COUNTIF/SEARCH columns plus an x14 cross-sheet rule on G2:J13, §111). The Notes sheet carries a COM-inserted SVG floating picture at E2 ("SVG picture", outside the A-C content so no grid test moves): Excel embeds `xl/media/image2.svg` plus a rasterized `xl/media/image1.png` fallback (`a:blip` + `asvg:svgBlip`; completed `svg-pictures` probe, label stripped). A second COM pass gave it a natively-authored outer shadow (gray, 8pt blur, 6pt offset, `rotWithShape="0"`) and 12° rotation (second pass of the same completed probe, label stripped), proving spreadsheet picture effects through the shared effect reader and painter. A sixth sheet, "Fills" (an RTL Arabic table, authored through Excel's COM object model, label stripped; Notes §130), gives its value cells the fills Excel writes outside a solid colour: a 90° and a 0° `gradientFill`, a from-centre `gradientFill type="path"`, and the `darkGrid`, `gray125` and `lightDown` patterns. Assembled in Excel from real Excel-owned sheets; keeps the reserved `fills[1]` (`gray125`) slot — a generated predecessor once put a custom fill there and Excel repaired the file for it. |
| `xlsx-coverage.json` | Fixture manifest | Semantic selectors for the XLSX test suite. It locates coverage sheets by unique authored text, not sheet indexes. |
| `docx-coverage.docx` | DOCX | Real Word-authored bilingual brief (Arabic + Latin): 17 mixed Arabic+Latin paragraphs, RTL paragraphs, a Heading 1–2 outline, 3 tables (fills, `gridSpan`/`vMerge`), bullet + decimal lists with nested levels, 2 inline images, 2 footnotes, headers/footers, internal/external hyperlinks. Section 5 ("5. Delivery tracker / متابعة التسليم") was appended through Word's COM object model: a right-to-left table (`w:bidiVisual`) with a 3pt border set on its physical right (Word writes `w:left`, the start side), 0.5pt on its left (`w:right`), cell margins 0.3in left / 0.05in right (Word keeps them physical; Notes §43), and a merged total row styled Heading 2; the appendix header carries a Heading 1 paragraph. A right-to-left paragraph before the tracker, "آخر تحديث 15 Oct 2026 بنسبة إنجاز 71%", has left-to-right runs for the date and percentage (no `w:rtl`; Notes §44), set in Calibri beside the Arial Arabic at exactly 18pt line spacing (the exact-pitch case; `tools/probe/exact-pitch/extend.py`). Both headings sit outside the outline and pin the one heading index space (`headings.list()`, `goToHeading`). The org sensitivity label Word applied on save was stripped (custom properties, header marking, and the even/first-page headers Word adds only to carry it: `tools/probe/_common/strip_label.py`). The document's `compatibilityMode` is 15 (Word 2013+): two justified English paragraphs in the synopsis ("Sawt al-Nahr follows three families…", "Dates below are hard for the picture lock…") set `DocxParagraphLayout.justifyShrink` (Notes §44). One portrait page after the landscape appendix holds seven anchored pictures on Arabic paragraphs, added through COM (`tools/probe/docx-float-wrap/extend.py`, label stripped): square at the left and right column edges, top and bottom and behind text centred, tight at the edge, and square and through centred with text down both sides (Notes §55). The sixth, landscape page holds two charts on Arabic paragraphs, added through COM with Word's own chart (`AddChart2`, embedded workbook; `tools/probe/docx-charts/extend.py`, label stripped): "مخطط مضمن في سطره" with an inline 5 x 3in chart in a centred paragraph after it, and "مخطط بالتفاف مربع" with a 3.5 x 2.25in chart anchored at the left column edge, square wrap; both clustered columns with Arabic categories, series names and title (Notes §71). A last portrait section ("8. Lists, headers and empty lines / القوائم والرؤوس والأسطر الفارغة", `tools/probe/docx-lists-headers/extend.py`, label stripped) turns on a different first page (`w:titlePg`) and writes only that first-page header, and has no footer; its single page holds three numbered lists (three items, a plain paragraph, two items that continue the first list, then a list Word restarts at 1 with a start override: markers 1-5, then 1-2), and a table whose second row has a cell holding only an empty paragraph (Notes §76). The document was seven pages in Word after merging the picture-wrap examples onto portrait page 5 and the final lists/table content onto page 7. Square and tight edge wrapping remain separate examples because they exercise different wrap modes. A final anchored box was added through Word COM with Arabic text "صندوق ملاحظات: مراجعة التسليم" and the Latin paragraph "English note: confirm the final delivery.", a pale green fill, and a green outline. Word writes it as `mc:Choice` DrawingML plus a VML `mc:Fallback`, so it covers branch selection without duplicate text and keeps the mixed-script content searchable. Even-page headers are not in it (`w:evenAndOddHeaders` is document-wide and would blank the even pages of every earlier section); the tests build them synthetically. A second box was added the same way, with the fill and the outline both turned off (Arabic "صندوق بلا تعبئة ولا حد: مراجعة التسليم", Latin "No fill, no outline: confirm the archive copy before the handover.", Word's COM object model): Word writes `a:noFill` for the fill, `a:noFill` in the outline, and the hidden outline's black in `a:extLst/a:ext/a14:hiddenLine`, which is the case aared once painted as a solid black box (Notes §94). It is anchored with no wrap so it cannot reflow the pages around it. A final table was added through COM with `Grid Table 4 - Accent 1` (`tools/probe/docx-table-style/extend.py`, label stripped), which is the fixture's only table Word styles through `w:tblStyle`: Word authors no per-cell shading for it, so its header row, banded body rows and style borders all live in `styles.xml` (Notes §99). It carries a header row, banded rows, a vertically merged cell (rows 4-5, column one) and one cell with its own direct fill on a banded row, so direct formatting is pinned against the style's. Its cells inherit the document's right-to-left, justified, bold context, so Word flattens that into direct run properties on save: the header cell's `w:b w:val="0"` is what keeps the region's bold off. |
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
  that the deferred-geometry engine renders (see Notes §26). Every shape
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
  dashed vertical line (zero width), a dotted line and a `flipV` diagonal; plus an RTL text box "الإيرادات دون الموازنة بنسبة 2.1%" whose number is an en-US run (Notes §44),
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
  next (they add up). Notes §47.
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
  Slide 19 ("Group appearance / مظهر المجموعة") adds solid group inheritance,
  a direct child override, a nested group inheriting a continuous gradient,
  a native `noFill` child, outer/inner shadows, glow, reflection and soft edges
  (§101). Python drove real PowerPoint to author the shapes, fills and effect
  values. Because COM distributes group effects and gradients onto children,
  its authored payloads were relocated into `grpSpPr` for the group-level
  cases, then reopened, saved and exported by PowerPoint; the inner-shadow
  case retains the native distributed form. This blank-layout slide carries
  no slide-number placeholder. The red child's text has native-resaved `grpFill`:
  PowerPoint paints it unfilled (PDF trace opacity zero), distinct from shape
  inheritance. Parser and actual DOM/SVG assertions live in
  `pptx-group-appearance.test.ts` and `appearance-resolved.visual.test.ts`.
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
  outline. See Notes §27; recoloring or moving them means
  re-baselining the idx-0 tests.
- The sidecar manifest is a test-only contract. When the deck is rebuilt, keep
  its selector text unique; tests reject missing or ambiguous selectors and
  include candidate summaries to make fixture drift actionable.
- Slide 20 ("SVG picture / صورة SVG") was appended through PowerPoint's COM
  object model (completed `svg-pictures` probe, label stripped): a
  caption text box plus an SVG picture inserted via `Shapes.AddPicture`.
  PowerPoint embeds only the vector part (`ppt/media/image13.svg`, svgBlip-only
  with no PNG beside it — the same shape as the slide 5 picture bullet), so
  this slide proves the SVG-primary path while the XLSX Notes picture proves
  SVG-primary with a raster fallback.

**Coverage notes** (`xlsx-coverage.xlsx`):

- `CF Rules!E2:E12` adds Excel-COM-authored three-colour-scale coverage
  (§109): 0 through 100 in steps of ten, min / 50th percentile / max, red
  `F8696B`, yellow `FFEB84`, green `63BE7B`. `CF Rules!F2:F12` adds an
  expression rule (`=MOD($F2,2)=0` over 1..11, §110). `CF Rules!G2:J13`
  closes the formula gap (§111): frozen 2020/2099 serials under `=$G2<TODAY()`,
  fruit duplicates under `=COUNTIF($H$2:$H$13,H2)>1`, text under
  `=ISNUMBER(SEARCH("apple",$I2))`, and hundreds under the x14
  `=$J2>Notes!$B$6` cross-sheet rule (red inline `dxf`). The existing A-D
  columns and rules are retained, and the sensitivity label was stripped
  after save. `xlsx-conditional.test.ts` compares the resolved fills with
  Excel's measured `DisplayFormat` colours; `xlsx-capture.visual.test.ts`
  re-reads actual DOM backgrounds and captures the curated
  `three-colour-scale` part (`A1:J13`). Recreate the reviewed pair with
  `bun run tools review fixtures/xlsx-coverage.xlsx --part 5:three-colour-scale=A1:J13 --pair-full`.

**Coverage notes** (`docx-coverage.docx`, verified by unzipping
and by the DOCX parser suite):

- Five bilingual "Theme ink" paragraphs were appended through Word's COM
  `TextColor.TintAndShade` (§100): base accent1, -25%, +40%, -50%, and +80%.
  Native Word PDF colours are `156082`, `0F4761`, `45B0E1`, `0A2F40`, and
  `C1E4F5`. `docx-theme-ink.test.ts` checks the resolved Latin/Arabic runs and
  recalculation against a changed theme; `appearance-resolved.visual.test.ts`
  re-reads the actual browser ink. The document's geometry is not altered.

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
- Pages 8–10 extend the original seven-page document with seventeen native
  Word-authored text-bearing shapes (nineteen text boxes total then; page 11
  adds the six effect shapes, twenty-five text boxes and eleven pages in all).
  Three manual page breaks and page-relative anchors keep them clear of the
  original content. Word's COM object model authored every shape, including
  `TextFrame.AutoSize = -1` for resize-to-text; no shape XML is injected. All
  earlier text runs and original media were checked for retention, and the
  sensitivity label was stripped after save with
  `tools/probe/_common/strip_label.py`. The new cases use unique visible `TB`
  tokens and matching shape names/alternative text:
  - Page 8: `TBGradient`, `TBPattern`, `TBPicture` (the fixture's existing colour
    bars image as a stretched fill), `TBCombined` (rounded gradient, middle
    alignment, adjustment 0.35), `TBRoundWide` (220×120pt, adjustment 0.10),
    `TBRoundTall` (110×140pt, adjustment 0.45).
  - Page 9: `TBOvalWide` (220×120pt), `TBOvalTall` (110×140pt), `TBRotated`
    (25° rotation with mixed Arabic/English text), `TBMiddle`, `TBBottom`,
    `TBDash` (1.5pt dashed outline).
  - Page 10: `TBNoWrap`, `TBFixedOverflow`, `TBAutoLong`, `TBAutoShort`,
    `TBAutoArabic`. Native Word reopens the resize cases at 132.9pt, 36.3pt,
    and 91.5pt high, although `wp:extent` and `a:xfrm/a:ext` retain their
    authored 40pt, 140pt and 40pt heights. The no-wrap box has `noAutofit`:
    Word widens it from the saved 150pt to 174.7pt. The fixed 40pt box clips
    its remaining lines even with `vertOverflow="overflow"`.
  `packages/aared/test/docx-text-box.test.ts` re-reads these cases from the
  parsed fixture, and `packages/aared/test/docx-text-box.visual.test.ts` re-reads
  what they paint (Notes §95). The fixture, its semantic tests and the
  paired native review in `tools/out/docx/docx-coverage/` are the durable
  record.

- Page 11 (a portrait section of its own, 612×792pt with 54pt margins) holds the
  shared drawing effects, authored through Word's COM `Shape.Shadow`,
  `Shape.Glow`, `Shape.SoftEdge` and `Shape.Reflection` (§102):
  - Six anchored shapes named `Word effect outer shadow`, `... inner shadow`,
    `... glow`, `... soft edge`, `... reflection` and `... rotated shadow` (12°
    rotation, so its shadow's offset no longer rotates with the shape). Each is
    a filled box with mixed Latin/Arabic text, all in front of text
    (`wdWrapNone`), positioned relative to the column and the paragraph.
  - Three pictures through the three existing routes: an anchored one
    (`wp:anchor`, outer shadow), a picture-only paragraph (`wp:inline` in its own
    block, outer shadow), and one riding a text line beside
    "Word effect inline badge" (glow). Word's own `wp:effectExtent` reserves the
    room each effect takes.
  - The two picture paragraphs are `w:bidi` with `w:jc="both"`: Word paints their
    pictures at the right edge of the column, which is what aared now does too.
  `docx-drawing-effects.test.ts` reads the resolved effects and insets back out
  of the parsed model, and `docx-drawing-effects.visual.test.ts` re-reads the
  filters, the reflection and its fallback, and the unchanged pagination.
