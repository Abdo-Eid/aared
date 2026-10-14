<p align="center">
  <img src="site/public/aared-mark.svg" alt="aared logo" width="120">
</p>

# aared (عارض)

**Preview real PowerPoint, Excel, and Word files in the browser — correctly,
in Arabic.**

See [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) for how it works, current
gaps, and the differences it accepts on purpose.

Some lightweight client-side Office previewers paint to `<canvas>` for speed.
In the ones we reviewed, that dropped paragraph-level direction for Arabic and
other RTL content, and once text is pixels there's no CSS left to fix it with.
`aared` renders through the browser's own DOM/SVG text stack instead, so
shaping, bidi, and RTL come from the browser, the way they do in any other web page.

```text
  .pptx / .xlsx / .docx  ──parse (Worker)──►  resolved model  ──render──►  DOM/SVG
                                                                             (never canvas)
```

View-only, by design — not an editor, not a converter. PDF is deliberately
out of scope; [`pdf.js`](https://github.com/mozilla/pdf.js) already covers
it well, so `aared` pairs with it instead of duplicating it.

---

## Compared to other libraries

Full table and reasoning: [`docs/COMPETITORS.md`](docs/COMPETITORS.md).
Measured numbers (method + generated tables): [`docs/BENCHMARKS.md`](docs/BENCHMARKS.md).

| Library | Formats | Client-side | Paragraph-level RTL |
|---|---|---|---|
| **aared** (this) | PPTX, XLSX, DOCX | ✅ | ✅ (design goal) |
| Extend `@extend-ai/react-*` | PPTX, DOCX, XLSX (React components) | ✅ | PPTX only — DOCX paragraphs stay LTR, XLSX grid isn't mirrored; 6–40× aared's download |
| `docx-preview` | DOCX | ✅ | ✅ |
| `pptxviewjs` | PPTX | ✅ | ❌ — canvas output, no CSS fix |
| SheetJS (Community Edition) | XLSX (parse only) | ✅ | N/A — styling silently dropped |
| Univer / LibreOffice-WASM | Full suite | ✅ | ✅ (but a full office-suite runtime, not an embeddable previewer) |

Nothing else spans all three formats at this weight with RTL that actually
works. That's the gap this project exists to close — the concrete evidence
is in [`docs/Notes.md §1`](docs/Notes.md#note-1).

## Install

```bash
npm install @aared/viewer
```

*(Not published yet: `@aared/viewer` on npm is still an empty name-holding
placeholder. Until the first release, work from this checkout: see
[Develop](#develop).)*

## Quick start

Import the element once from your app's bundled code (Vite, webpack,
Next.js; bare package subpaths need a bundler or equivalent import map):

```ts
import "@aared/viewer/element";
```

```html
<aared-viewer src="quarterly-report.pptx" controls></aared-viewer>
```

When your app accepts only one format, use that entry's complete viewer:

```ts
import { preview } from "@aared/viewer/pptx";

const viewer = await preview({ source: file, target: container });
await viewer.next();
const matches = await viewer.search("budget");
viewer.destroy(); // Removes the render and closes the document
```

For thumbnails or data beside a viewer, use its restricted `viewer.content`
with the format's `render()`/reads: the same open file, without exposing its
`close()` or session. Search stays on `viewer.search()`. Destroy additional
renders before the viewer; use `open()` below when you need to own the document.

`@aared/viewer/xlsx` and `@aared/viewer/docx` expose the same object-form `preview()` with
format-specific handles. Their builds omit sibling renderers; the shared Worker
still contains all parsers. Use `@aared/viewer/preview` to detect an arbitrary file type.

Or use the low-level entrypoints to own the document and render lifetimes
separately. Render from its handles; data reads do not expose painting models:

```ts
import { open, render } from "@aared/viewer/pptx";

const deck = await open(file);
const viewer = await render({
  slide: deck.slides.get(0),
  target: document.querySelector<HTMLElement>("#stage")!,
});
// On unmount: viewer.destroy(); deck.close(); (destroy leaves the painted DOM in #stage: remove it yourself)
```

```ts
import { open, render } from "@aared/viewer/xlsx";

const workbook = await open(file);
const sheet = workbook.sheets.get("Revenue");
const data = await sheet.read({ range: "A1:M200" }); // cell values, no rendering
const viewer = await render({ sheet, target: document.querySelector<HTMLElement>("#stage")! });
// On unmount: viewer.destroy(); workbook.close();
```

```ts
import { open, render } from "@aared/viewer/docx";

const doc = await open(file); // not `document`: that would shadow the page's
const viewer = await render({ document: doc, target: document.querySelector<HTMLElement>("#stage")! });
// On unmount: viewer.destroy(); doc.close();
```

Full API: the API reference page in the docs site — `bun run docs:dev`,
open `/api` (source: `site/src/content/docs/api.mdx`). New to it: `/start`
covers installing, picking an entry, recipes and troubleshooting.

Format entries export supported data types: XLSX `SheetData`/`RowData`/`CellData`
and `CellRange`, PPTX `SlideText` (paragraphs and optional notes), DOCX
`SectionInfo` (page metadata in inches). Rendering models, cell styles and rich
Word nodes stay internal; detached-model rendering is not a public capability.

### Search

The shared viewer API works with PPTX, XLSX and DOCX:

```ts
import { preview } from "@aared/viewer/preview";

const viewer = await preview({ source: file, target: container });
const matches = await viewer.search("nile"); // Results and highlights together
await viewer.search("nile", { highlight: false }); // Results only
await viewer.search(""); // Clear highlights
```

Search covers the whole presentation or document, and the current spreadsheet
sheet. PPTX results count text occurrences, so one slide can have several
results. Highlights follow navigation and newly painted rows; matching DOCX
headers and footers highlight on every displayed page. The app supplies the
input and results list. Low-level document and sheet searches return data only.
For a compact find bar, `viewer.searchNavigation` exposes the active `index`,
total `count`, and wrapping `next()` / `previous()` methods. The active match
has a stronger outlined highlight; `searchchange` updates the counter.
See the [search recipe](site/src/content/docs/start.mdx#add-a-search-box) for
an input example and the [API reference](site/src/content/docs/api.mdx#auto-detect-preview-aaredviewerpreview)
for result navigation, cancellation and browser support.

## Fonts

Office files name fonts nobody has installed (Aptos, Calibri, Sakkal
Majalla…). aared asks for the font the file names, and where the reader
doesn't have it, the browser's own fonts paint instead: nothing downloads,
and on a phone the system's Arabic and Latin faces are both faster and
easier to read than any substitute. Some lines break differently from Office.

An app that needs Office's line breaks exactly can pass its own
`fontResolver` to `open()` (see the API reference). The repo's
`aared-fonts` package holds calibrated look-alikes for that, but it is
frozen and not published: on a real phone it made the first page 7 s
slower and the Arabic harder to read (`docs/Notes.md §53`).

Fonts embedded in the file itself are not read yet.

## How it works

Three formats — PPTX, XLSX, DOCX — share one packaging format underneath
(OOXML: a ZIP of XML parts) and a surprising amount of internal machinery
(DrawingML shapes/pictures/text, chart XML, theme resolution). `aared`
parses that shared substrate once, in a Worker, into a resolved model; each
format then has its own renderer that paints that model straight to DOM/SVG.
Nothing is rasterized by `aared` itself — text shaping, bidi, and font
fallback are the browser's job, which is exactly what makes Arabic/RTL
content come out correct without a special case.

Full design: [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md). Why it's
shaped this way, with native findings and rejected alternatives:
[`docs/Notes.md`](docs/Notes.md).

## What you can preview today

The entrypoint-backed route opens and renders PPTX slides, XLSX sheets, and DOCX
documents through `@aared/viewer/pptx`, `@aared/viewer/xlsx`, `@aared/viewer/docx`, or
`<aared-viewer>`. Each format has real Arabic/RTL fixture coverage, including
tables. The supported surface, the gaps still open and the differences kept on
purpose are recorded in [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md).

DOCX renders flow-style by default; `pagination: { mode: "layout" }` opts into
physical pages with structural (not pixel-identical) Word fidelity: orientations,
hard breaks, and per-page headers/footers/footnotes. Page counts can differ
from Word's.

## Docs

- [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) — how it's designed to work, plus the gap/limits inventory, release readiness and proposals
- [Getting started](https://abdo-eid.github.io/aared/start/) (`bun run docs:dev`, open `/aared/start/`) — installation, integration choice and recipes
- [API reference](https://abdo-eid.github.io/aared/api/) (`bun run docs:dev`, open `/aared/api/`) — the public API
- [`docs/Notes.md`](docs/Notes.md) — decisions and engineering experience, by topic
- [`CHANGELOG.md`](CHANGELOG.md) — changes after the development baseline
- [Archived development history](docs/archive/CHANGELOG-pre-baseline.md) — what led to the baseline, not published releases
- [`docs/COMPETITORS.md`](docs/COMPETITORS.md) — the landscape
- [Unscheduled proposals](docs/ARCHITECTURE.md#unscheduled-proposals) — undecided capabilities and consumer tradeoffs, not promised APIs
- [`AGENTS.md`](AGENTS.md) — working charter for anyone (human or agent)
  contributing code; start here

## Develop

```bash
bun install            # workspace install
bun run build          # tsdown: the package exports point at dist/, so build before anything else
bun run dev            # bench/ visual harness at http://localhost:5173
bun run typecheck      # tsc --noEmit across packages and tools
bun test               # bun:test — unit tests
bun run test:visual    # WebView DOM-output + capture tests (captures need Office)
```

Developer tooling has one front door: `bun run tools --help`. For example,
`bun run tools review <file.pptx|xlsx|docx>` writes the native/aared comparison
and `review.html`; `bun run tools doctor` checks local prerequisites without
launching Office. See [`tools/README.md`](tools/README.md) for reviews,
benchmarks, model inspection and specialist research tools.

## License

[Apache-2.0](LICENSE) © 2026 [Abdulhamed Eid](https://github.com/Abdo-Eid)
