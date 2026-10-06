<p align="center">
  <img src="site/public/aared-mark.svg" alt="aared logo" width="120">
</p>

# aared (عارض)

**Preview real PowerPoint, Excel, and Word files in the browser — correctly,
in Arabic.**

> ⚠️ Pre-release. Nothing is published yet. See
> [`docs/ROADMAP.md`](docs/ROADMAP.md) for current status.

Every lightweight client-side Office previewer makes the same trade to stay
fast: paint to `<canvas>` instead of building a real DOM. That trade quietly
breaks Arabic and other RTL content — paragraph-level direction gets dropped,
and once text is pixels there's no CSS left to fix it with. `aared` makes
the other trade: render through the browser's own DOM/SVG text stack, so
shaping, bidi, and RTL come for free, the way they do in any other web page.

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
| `pptxviewjs` / `pptx-preview` | PPTX | ✅ | ❌ — canvas output, no CSS fix |
| SheetJS (Community Edition) | XLSX (parse only) | ✅ | N/A — styling silently dropped |
| Univer / LibreOffice-WASM | Full suite | ✅ | ✅ (but a full office-suite runtime, not an embeddable previewer) |

Nothing else spans all three formats at this weight with RTL that actually
works. That's the gap this project exists to close — the concrete evidence
is in [`docs/DECISION_LOG.md §1`](docs/DECISION_LOG.md#1-why-this-project-exists-at-all).

## Install

```bash
npm install @aared/viewer
```

*(Not released yet — `@aared/viewer` on npm is a name-holding placeholder
with no code, and the snippets below still say `aared/...` until the package
is renamed for its first release at `0.1.0`. Until then, build it from
source: see [Develop](#develop).)*

## Quick start

Import the element once from your app's bundled code (Vite, webpack,
Next.js; a plain `<script type="module">` cannot resolve `aared/element`):

```ts
import "aared/element";
```

```html
<aared-viewer src="quarterly-report.pptx" controls></aared-viewer>
```

When your app accepts only one format, use that entry's complete viewer:

```ts
import { preview } from "aared/pptx";

const viewer = await preview({ source: file, target: container });
await viewer.next();
const matches = await viewer.search("budget");
viewer.destroy(); // Removes the render and closes the document
```

`aared/xlsx` and `aared/docx` expose the same object-form `preview()` with
format-specific handles. Their builds omit sibling renderers; the shared Worker
still contains all parsers. Use `aared/preview` to detect an arbitrary file type.

Or build from the file's models on the headless entrypoints — one per format,
same shape. Here the render and document have separate lifetimes:

```ts
import { open, render } from "aared/pptx";

const deck = await open(file);
const viewer = await render({
  slide: deck.slides.get(0),
  target: document.querySelector<HTMLElement>("#stage")!,
});
```

```ts
import { open, render } from "aared/xlsx";

const workbook = await open(file);
const sheet = workbook.sheets.get("Revenue");
const data = await sheet.read({ range: "A1:M200" }); // cell values, no rendering
const viewer = await render({ sheet, target: document.querySelector<HTMLElement>("#stage")! });
```

```ts
import { open, render } from "aared/docx";

const doc = await open(file); // not `document`: that would shadow the page's
const viewer = await render({ document: doc, target: document.querySelector<HTMLElement>("#stage")! });
```

Full API: the API reference page in the docs site — `bun run docs:dev`,
open `/api` (source: `site/src/content/docs/api.mdx`). New to it: `/start`
covers installing, picking an entry, recipes and troubleshooting.

### Search

The shared viewer API works with PPTX, XLSX and DOCX:

```ts
import { preview } from "aared/preview";

const viewer = await preview({ source: file, target: container });
const matches = await viewer.search("nile"); // Results and highlights together
await viewer.search("nile", { highlight: false }); // Results only
await viewer.search(""); // Clear highlights
```

Search covers the whole presentation or document, and the current spreadsheet
sheet. PPTX results count text occurrences, so one slide can have several
results. Highlights follow navigation and newly painted rows; matching DOCX
headers and footers highlight on every displayed page. The app supplies the
input and results list. Headless document and sheet searches return data only.
For a compact find bar, `viewer.searchNavigation` exposes the active `index`,
total `count`, and wrapping `next()` / `previous()` methods. The active match
has a stronger outlined highlight; `searchchange` updates the counter.
See the [search recipe](site/src/content/docs/start.mdx#add-a-search-box) for
an input example and the [API reference](site/src/content/docs/api.mdx#methods)
for result navigation, cancellation and browser support.

## Fonts

Office files name fonts nobody has installed (Aptos, Calibri, Sakkal
Majalla…). aared asks for the font the file names, and where the reader
doesn't have it, the browser's own fonts paint instead: nothing downloads,
and on a phone the system's Arabic and Latin faces are both faster and
easier to read than any substitute. Some lines break a word differently from
Office; nothing overlaps.

An app that needs Office's line breaks exactly can pass its own
`fontResolver` to `open()` (see the API reference). The repo's
`aared-fonts` package holds calibrated look-alikes for that, but it is
frozen and not published: on a real phone it made the first page 7 s
slower and the Arabic harder to read (`docs/DECISION_LOG.md §53`).

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
shaped this way, chronologically: [`docs/DECISION_LOG.md`](docs/DECISION_LOG.md).

## What you can preview today

The entrypoint-backed route opens and renders PPTX slides, XLSX sheets, and DOCX
documents through `aared/pptx`, `aared/xlsx`, `aared/docx`, or
`<aared-viewer>`. Each format has real Arabic/RTL fixture coverage, including
tables. Fidelity is still in progress; inspect [`docs/ROADMAP.md`](docs/ROADMAP.md)
for supported features and deliberate gaps.

DOCX renders flow-style by default; `pagination: { mode: "layout" }` opts into
physical pages with structural (not pixel-identical) Word fidelity: correct
page count, orientations, hard breaks, and per-page headers/footers/footnotes.

## Docs

- [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) — how it's designed to work
- API reference (`bun run docs:dev`, open `/api`) — the public API
- [`docs/DECISION_LOG.md`](docs/DECISION_LOG.md) — why, chronologically
- [`docs/COMPETITORS.md`](docs/COMPETITORS.md) — the landscape
- [`docs/ROADMAP.md`](docs/ROADMAP.md) — current status, open work, what's next
- [`AGENTS.md`](AGENTS.md) — working charter for anyone (human or agent)
  contributing code; start here

## Develop

```bash
bun install            # workspace install
bun run dev            # bench/ visual harness at http://localhost:5173
bun run typecheck      # tsc --noEmit across all packages
bun test               # bun:test — unit tests
bun run test:visual    # WebView DOM-output + capture tests (captures need Office)
bun run build          # tsdown for the two packages — real dist output
```

## License

[Apache-2.0](LICENSE) © 2026 [Abdulhamed Eid](https://github.com/Abdo-Eid)
