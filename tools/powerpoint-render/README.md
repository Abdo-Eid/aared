# tools/powerpoint-render

Exports a real `.pptx` through the installed Microsoft PowerPoint application,
not LibreOffice. The output is a PowerPoint-rendered PNG for every slide.

```bash
bun tools/powerpoint-render/index.ts fixtures/pptx-coverage.pptx --width 1280
```

The tool requires Windows, PowerPoint, and PowerShell 7. It opens the file
read-only without a PowerPoint window, exports into the gitignored `out/`
directory, then closes and releases PowerPoint. The companion visual suite uses
the same export to create same-size aared captures and a side-by-side review
page:

```bash
bun run test:visual --artifacts
```
