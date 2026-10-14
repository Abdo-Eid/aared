# PPTX review backend

Exports a real `.pptx` through the installed Microsoft PowerPoint application,
not LibreOffice. The output is a PowerPoint-rendered PNG for every slide.

For an end-to-end native/aared review of any deck, use
`bun run tools review <file.pptx> [--width 1280] [--absent-fonts]`.
The command and visual suite share the same DOM/SVG capture implementation.
See [the tools catalog](../../README.md). The command below exports native PNGs only.

```bash
bun tools/review/office-export.ts --format pptx fixtures/pptx-coverage.pptx --width 1280
```

The tool requires Windows, PowerPoint, and PowerShell 7. It opens the file
read-only without a PowerPoint window, exports into the gitignored `out/`
directory, then closes and releases PowerPoint. The companion visual suite uses
the same export to create same-size aared captures and a side-by-side review
page:

```bash
bun run test:visual --artifacts
```
