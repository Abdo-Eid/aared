# Changelog

Changes after the October 14, 2026 development baseline are recorded here.
The baseline is not a published release: the package remains private,
unpublished and at version `0.0.0`. Do not treat it as `0.1.0` or as an npm
release.

Earlier development is recorded in the frozen
[pre-baseline changelog](docs/archive/CHANGELOG-pre-baseline.md).
Decisions, Office findings, rejected alternatives and engineering experience
live in [Notes](docs/Notes.md); current behavior, gaps and release readiness
live in [ARCHITECTURE](docs/ARCHITECTURE.md).

## Unreleased

- **XLSX gradient and pattern cell fills paint.** `gradientFill` (linear and
  from-centre/corner) and SpreadsheetML's 17 non-solid `patternFill` types
  used to be dropped while parsing, leaving those cells blank. They now read
  into the shared fill model and paint behind the cell's text through the
  shared fill painter, as Excel draws them (Notes §130). The cell style's
  `fill` is a `ShapeFillModel` instead of a CSS colour string.
- **Built-in number format 47 is `mm:ss.0`**, as Excel resolves it, not
  ECMA-376's misprinted `mmss.0`. Time formats show fractional seconds
  (`ss.0`, `ss.00`, `ss.000`).
- `xlsx-coverage.xlsx` gains a "Fills" sheet authored in Excel.
