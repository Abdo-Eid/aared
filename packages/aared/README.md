# aared (عارض) — @aared/viewer

aared is a TypeScript Office viewer for PPTX, XLSX and DOCX in the browser,
published on npm as `@aared/viewer`. View-only, with Arabic and other
RTL text rendered as real DOM/SVG rather than canvas pixels.

Fidelity is checked against native PowerPoint, Excel and Word output, including
real Arabic fixtures with tables in every format. See the
[evaluated browser-library comparisons](https://github.com/Abdo-Eid/aared/blob/main/docs/COMPETITORS.md)
and [supported features and limits](https://abdo-eid.github.io/aared/api/#fidelity-is-part-of-the-contract).

```sh
npm install @aared/viewer
```

Parsing runs in a Web Worker. Rendering requires a browser; there is no server
mode and no headless renderer.

## Pick your layer

| You want… | Import from |
|---|---|
| A viewer with built-in paging/zoom controls | `@aared/viewer/element` |
| A JavaScript viewer whose UI you own | `@aared/viewer/preview` |
| A viewer in React that tears itself down | `@aared/viewer/react` |
| Semantic data reads, or document lifetime you own | `@aared/viewer/pptx`, `@aared/viewer/xlsx` or `@aared/viewer/docx` |

Every layer uses the same renderers, so a simpler entry does not cost fidelity.
Import a documented subpath, not the package root. Known format? Import
`preview()` from that format's entry to skip detection and skip the other
formats' code. React is an optional peer, needed only by `@aared/viewer/react`.

## 1. HTML element

Built-in controls, no lifecycle code. Import once in your bundled app code.

```ts
import "@aared/viewer/element";
```

```html
<aared-viewer src="./report.xlsx" controls style="display: block; height: 480px"></aared-viewer>
```

`src` takes a URL; the `file` property takes a `File`/`Blob`/`ArrayBuffer`. Of
the two, whichever was set last opens. Removing the element is its destroy.

## 2. JavaScript viewer

Detects the format, loads only that format's code, and returns one handle.

```ts
import { preview } from "@aared/viewer/preview";

const target = document.querySelector<HTMLElement>("#preview")!;
target.style.cssText = "height: 480px; overflow: auto"; // a sheet or document needs a height

const viewer = await preview({ source: "./report.xlsx", target });
await viewer.goTo(1);
const matches = await viewer.search("budget");
viewer.destroy(); // when the view goes away
```

`viewer.format` narrows the handle, so `viewer.xlsx.goToCell("D14")` is
available on a workbook. `viewer.content` gives semantic reads of the same open
document for thumbnails or range reads, without reopening the file.

## 3. React

```tsx
import { usePreview } from "@aared/viewer/react";
import { messageFor } from "@aared/viewer/core";

export function Preview({ file }: { file: File }) {
  const { ref, status, viewer, index, error, retry } = usePreview({ source: file });
  return (
    <>
      <div ref={ref} style={{ height: 480 }} />
      {status === "loading" && <p>Loading…</p>}
      {status === "error" && <p>{messageFor(error)} <button onClick={retry}>Retry</button></p>}
      {status === "ready" && <p>{index + 1} of {viewer.info.count}</p>}
    </>
  );
}
```

The hook owns the `AbortController`, the loading/ready/error state and teardown
on unmount or source change — there is no `destroy()` to call. Use `usePptx()`,
`useXlsx()` or `useDocx()` to assert one format.

## 4. Manual document lifetime

For data reads without a viewer, or several renders of one document.

```ts
import { open, render } from "@aared/viewer/pptx";

const deck = await open("./report.pptx");          // or an ArrayBuffer/File
const slide = deck.slides.get(0);
const text = await slide.read();                   // semantic data, no rendering
const found = await deck.search("budget");

const target = document.querySelector<HTMLElement>("#preview")!;
target.style.cssText = "height: 480px; overflow: auto";
const view = await render({ slide, target });      // a registered handle, not a data copy

view.destroy();                                     // each render handle
deck.close();                                       // then the document
```

XLSX uses `sheet = deck.sheets.get(0)` with `render({ sheet, target })`; DOCX
uses `render({ document, target })`. `render()` never accepts projected data or
a detached model — only handles or content borrowed from an open document.

## Boundaries worth knowing before you rely on it

- **View-only.** No editing, no format conversion, no formula recalculation.
- **XLSX shows the values stored in the file.** Formula results are read as
  saved; nothing recalculates them.
- **DOCX pagination is structural, not Word-identical.** Page counts can differ
  from Word's, and layout mode is opt-in via `pagination: { mode: "layout" }`.
- **Wrapping follows the fonts installed in the browser.** Office-only fonts
  will reflow the text unless you supply a `fontResolver`.
- **Supported features are a contract, not a best effort.** Check the fidelity
  section of the API reference before depending on an exact Office behaviour.

## Docs

- Getting started: <https://abdo-eid.github.io/aared/start/>
- API reference and fidelity limits: <https://abdo-eid.github.io/aared/api/>
- Plain Markdown for agents: <https://abdo-eid.github.io/aared/llms.txt>
- Comparison evidence: <https://github.com/Abdo-Eid/aared/blob/main/docs/COMPETITORS.md>
- Dated performance measurements: <https://github.com/Abdo-Eid/aared/blob/main/docs/BENCHMARKS.md>

Exact signatures are in the shipped declarations (`dist/**/*.d.mts`) — read the
ones in **your** installed version rather than trusting this README for types.
Source and issues: <https://github.com/Abdo-Eid/aared>

Apache-2.0. See LICENSE.
