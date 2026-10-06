# tools/font-calibration

Ground-truth probes for the substitution table in `packages/aared-fonts/src/
index.ts` (`FONT_SUBSTITUTIONS`, `FONT_CALIBRATION`, `ARABIC_SCRIPT_FALLBACKS`).
Not OOXML probes: they measure real fonts installed on the calibration machine
(Windows with Office: `simpo.ttf` Simplified Arabic, `trado.ttf` Traditional
Arabic, `majalla.ttf` Sakkal Majalla, `arabtype.ttf` Arabic Typesetting,
`arial.ttf`, `times.ttf`, `calibri.ttf`, … under `C:\Windows\Fonts`) against
candidate open fonts, on real shaped Arabic text (Arabic's contextual letterforms
change width by context, so summing an unshaped glyph table is wrong).

Re-run whenever the table changes or the calibration machine's fonts do: the
numbers are pinned in code, not derived at build or run time.

## The Arabic survey (what the table rests on)

- `arabic-survey.ts` — every Arabic-only run in the three fixtures (59 samples),
  in each reference font's real file (served locally, loaded through `FontFace`)
  against each candidate loaded from the same Google Fonts files the resolver
  delivers (private family names, so a local install can never stand in), with
  every load verified. Reports per pair: mean width ratio (what `size-adjust`
  removes), sd of per-sample ratios (what line breaks follow), worst-case
  deviation, the `size-adjust` needed, and alef/seen height after adjusting.
  `--bold` for weight 700, `--refs`/`--cands` to narrow, `--calib "Ref=Cand,…"`
  to print the exact descriptors. Writes `arabic-survey.last.json` (with the
  samples, for the Word probe).
- `word-arabic-fallback.ps1` + `.py` — what Word itself draws: one Arabic sample
  per page per font through Word COM, exported to PDF, each page's real text
  width read with PyMuPDF (Word's `Range.Information()` caret positions are
  unreliable for bidi and were discarded). Validates that canvas measurement of
  the installed fonts agrees with Word (within about 3%), and shows which font
  Word substitutes for one with no Arabic glyphs: Cambria's Arabic is drawn in
  Times New Roman. The script quits Word only if it started that instance.
- `alias-composition.ts` — proves two faces under one alias compose: Arimo for
  Latin plus a calibrated, `unicode-range`-limited Amiri for Arabic. Latin and
  the punctuation both Google subsets claim measure as Arimo alone; Arabic as
  calibrated Amiri alone; vertical metrics land at Arial's; add order does not
  matter.
- `arabic-candidates.ts` — the first, narrow pass (three candidates, three
  paragraphs) that chose Amiri for all four fonts. Superseded by the survey.
- `arabic-fallback-fit.ts` — the numbers for `ARABIC_SCRIPT_FALLBACKS`. Such a
  face paints only Arabic letters; the spaces, digits and punctuation around
  them come from the alias's Latin face (Arimo, Tinos, Caladea, Carlito). So it
  is fitted through the real alias (two `size-adjust` probes; width is linear
  in it), never with `arabic-survey.ts --calib`, which measures the face alone
  and left Arial's Arabic 3% too wide (DECISION_LOG §46). Then checks line
  breaks: every paragraph with Arabic laid out at 11pt across seven widths in
  the real font vs the shipped and the refitted alias.
- `vertical-check.ts` — every calibrated face's line box at `line-height:
  normal` and 2000px against the real font, with the shipped numbers and the
  ones before §46. Line boxes round to whole pixels, so smaller sizes hide
  errors. Arabic Typesetting stays 0.7% short there: its `normal` box includes
  a line gap the overrides do not set, which aared never paints with.
- `vazirmatn-bench.ts` — Vazirmatn as a one-font Arial substitute (fitted per
  script, and as-is) against Arimo + Lateef, by line breaks for English,
  Arabic and mixed paragraphs. Rejected (§46).

Vertical metrics are read at 10000px (canvas returns whole pixels; a 40px
reading put the overrides up to 2.5% of the em off). They match each font
file's `usWin` ascent/descent, which is what Chromium on Windows paints with.
The shipped overrides divide those by the shipped (rounded) `size-adjust`;
`--calib` divides by the unrounded one, so it can print 0.1 differently.

### Result (regular weight; bold in `--bold`)

| Original | Substitute | size-adjust | ascent / descent override | sd of ratios | vs previous best |
|---|---|---|---|---|---|
| Simplified Arabic | Lateef | 130.3% | 90.5% / 36.7% | 0.023 | Amiri 0.044 (47% less) |
| Traditional Arabic | Amiri | 80.0% | 124.3% / 62.6% | 0.050 | Lateef 0.044 (12% less; below the 25% bar) |
| Sakkal Majalla | Markazi Text | 91.0% | 97.1% / 56.3% | 0.036 | Amiri 0.060 (41% less) |
| Arabic Typesetting | Amiri | 67.1% | 104.4% / 65.4% | 0.044 | Lateef 0.048 (10% worse) |
| Arial (Arabic text) | Lateef | 126.9% | 71.3% / 16.7% | 0.035 (through the alias) | Amiri 0.038 (alone) |
| Times New Roman (Arabic) | Lateef | 126.9% | 70.2% / 17.0% | 0.035 (through the alias) | Amiri 0.038 (alone) |
| Cambria (Arabic, measured against Times New Roman's) | Lateef | 128.9% | 69.1% / 16.8% | 0.035 (through the alias) | — |
| Calibri (Arabic) | Markazi Text | 104.3% | 91.3% / 25.8% | 0.020 (through the alias) | Amiri 0.046 (alone) |
| Calibri Light (Arabic) | Markazi Text | 101.9% | 93.4% / 26.4% | 0.021 (through the alias) | Amiri 0.046 (alone) |

Bold differs from regular by up to 5.6%, so Traditional Arabic and Calibri carry
their own bold numbers; the others are within about 1.5% and use one set.
Courier New's Arabic is monospace: worst-case per-sample error is 45–54% for every
proportional candidate, so no fallback is offered.

Facts established along the way: Arial and Times New Roman's Arabic match
Simplified Arabic within 1.3%; Calibri 6.27 (Windows 11) carries Arabic; Cambria
carries none; Arimo and Tinos have no Arabic subset (Google Fonts metadata). Arabic
Typesetting is a modern reading of Naskh (Microsoft Typography), not
"decorative/Diwani" as an earlier comment claimed.

## Aptos (DECISION_LOG §49)

The real Aptos, Aptos Display and Aptos Narrow come from Office's cloud-font
cache (`%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts`, filled once a document
using them is opened; files named by hash, found by name table in
`cloud-fonts.ts`). Never copied or shipped.

- `latin-survey.ts` — every Latin run in the five fixtures (194 samples, plus 77
  digit strings reported apart) in the real fonts against 61 open candidates
  loaded through the resolver's own `fontCssUrl` request. Mean, sd, worst, and
  x-height / cap height after size-adjust. `--bold`, `--italic`, `--refs`,
  `--cands`, `--calib "Aptos=Radio Canada,…"`.
- `latin-fit.ts` — the acceptance: size-adjust fitted on total width, every
  delivered subset under one alias with the overrides, real paragraphs laid out
  at eight widths in both weights; counts paragraphs broken exactly like the
  original, and the line box at `line-height: normal`. `--shipped` checks
  `FONT_CALIBRATION`'s numbers. "Arimo as-is" is what painted before (the
  generic sans-serif is Arial on Windows).
- `aptos-arabic-fallback.ps1` + `.py` — what Excel and Word draw for Arabic set
  in Aptos (Arial, at Arial's widths, in both apps and both Word slots).
  `arabic-fallback-fit.ts` then fits Lateef through each Aptos alias, with the
  Latin face registered with its own calibration.

| Original | Substitute | size-adjust / ascent / descent (regular; bold) | paragraphs broken like the original (regular/bold) | before |
|---|---|---|---|---|
| Aptos | Radio Canada | 92.2 / 101.8 / 30.6; 93.2 / 100.7 / 30.2 | 96% / 93% | 75% / 63% |
| Aptos Display | Roboto | 90.7 / 103.5 / 31.1; 93.7 / 100.2 / 30.1 | 95% / 94% | 42% / 38% |
| Aptos Narrow | IBM Plex Sans Condensed | 96.9 / 96.9 / 29.1; 95.7 / 98.1 / 29.4 | 96% / 95% | 56% / 51% |
| Arabic in any Aptos | Lateef (Arial's Arabic, through the alias) | 130.9 / 69.1 / 16.2 (Aptos), 131.4 / 68.9 / 16.1 (Display), 132.6 / 68.3 / 16.0 (Narrow) | Arabic line starts 96-100% | — |

Aptos's vertical metrics are its typo ones (USE_TYPO_METRICS): 0.939/0.282, not
the win 1.010/0.275. Italics measure 2-3% off the upright numbers (1% for
Narrow); not carried.

## Other scripts

- `metric-clone-check.py` — confirms Carlito/Caladea/Arimo/Tinos/Cousine's advance
  widths against the real Calibri/Cambria/Arial/Times New Roman/Courier New
  installed here. Network for the Google Fonts woff2 (converted with fontTools),
  the Office fonts read locally.
- `size-adjust-interaction.ts` — the empirical check of how `size-adjust`
  interacts with `ascent-override`/`descent-override` on a `FontFace`: a browser
  multiplies an override percentage by `size-adjust`'s own factor, so the numbers
  in `FONT_CALIBRATION` are `reference metric / size-adjust`, not the metric alone.
- `fonts-load-behavior.ts` — `document.fonts.load()` resolves near-instantly for a
  family with no registered `FontFace` at all, and correctly awaits a real
  registered face to `"loaded"`. Backs `awaitDocumentFonts` (`text-style.ts`).
- `fonts-check-probe.ts` — `document.fonts.check()` returns `true` even for a name
  that exists nowhere, so `detectFontFidelity` measures widths against a control
  instead. Note the reference caveat: on Windows the generic `sans-serif`/`serif`
  ARE Arial/Times New Roman, so a reference measuring like the generic is not
  evidence it failed to load; check `FontFace.status` instead.
