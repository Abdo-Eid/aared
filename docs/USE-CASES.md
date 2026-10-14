# Use cases: consumer-friction audit

This file is a working tool for evaluating how a consumer lands on aared's API.
Read each scenario as a new app developer: can you find the entry, understand the
ownership, and finish the job without learning unnecessary machinery?

The purpose is to expose complexity, not explain it away. A shorter example is
not evidence of a simpler experience if it omits state, cleanup or extra opens.
Keep straightforward scenarios too: they show what should remain simple.

## How to use this audit

Each scenario records **goal → current path → code → friction → cause → possible
improvement → disposition**. Code shows the API available today; proposed APIs
are described in prose, not presented as working imports. `file`, `container`
and named app DOM nodes are supplied by the application. Cleanup comments mark
work to run on unmount, not immediately after opening.

Findings below come from walking the current examples, documented contracts and
source. They are not measured usability-study results. "Investigate" marks a
candidate, not an approved change or an open-gap commitment.

| Disposition | Meaning |
|---|---|
| **Keep** | Current path is direct and its guidance is documented (the entry links where); do not add abstraction merely for symmetry |
| **Investigate** | Friction is visible; compare alternatives before changing the API |
| **Clarify contract** | A documented distinction needs a deliberate consumer-facing explanation |
| **Resolved** | The API changed to remove the friction; the decision is in Notes |

The disposition says whether the API is defensible. A separate **UX cost**
says what the journey still costs a first-time developer, so a defensible
**Keep** is not read as "no friction":

| UX cost | Meaning |
|---|---|
| **Low** | One documented call or recipe; little to learn beyond the goal itself |
| **Medium** | A second concept, lifetime or format branch to learn, documented next to the goal |
| **High** | The developer changes abstraction level or writes their own bookkeeping to finish |

### Document ownership

- [Start](../site/src/content/docs/start.mdx) owns onboarding and usable recipes.
- [API reference](../site/src/content/docs/api.mdx) owns signatures and semantics.
- This audit owns current consumer journeys and their friction assessments.
- [Architecture](ARCHITECTURE.md) owns mechanisms, accepted limits, and the
  [gap inventory](ARCHITECTURE.md#gaps-limits-and-fidelity-gaps) whose OPEN
  rows are committed work; [Engineering notes](Notes.md) owns adopted
  design decisions.

Do not paste each discussion back into Start. Resolve the cause, update the
owning document/API, then update the disposition here. A new user-facing
capability should gain a scenario; a proposal becomes a commitment only after
an explicit decision.

Current imports use `aared/...`. The planned first release uses
`@aared/viewer/...`; that rename does not change the journeys evaluated here.

## Scan the journeys

| Journey | Main friction | Disposition | UX cost |
|---|---|---|---|
| [Pick a route](#pick-a-route) | Format subpaths mistaken for exclusively low-level entries | Keep | Medium |
| [HTML embed](#html-embed) | "HTML" still requires JS registration, a build and a host height | Keep | Low |
| [JavaScript viewer](#javascript-viewer) | Generic versus fixed-format choice before the first call | Keep | Low |
| [React viewer](#react-viewer) | Hooks provide lifecycle, not built-in controls; host must stay mounted | Keep | Medium |
| [SSR](#server-rendering) | Import safety mistaken for server rendering support | Keep | Low |
| [File sources](#url-picker-drop-and-bytes) | Element uses `src`/`file`, functions use `source` | Keep | Low |
| [Authenticated fetch](#authenticated-fetch) | Download and viewer are two separate operations | Keep | Low |
| [Type detection](#detect-before-opening) | Separate URL calls can download twice | Clarify contract | Medium |
| [Toolbar](#custom-toolbar) | Hand-written UI and event subscriptions | Keep | Medium |
| [Sheets and cells](#sheet-tabs-and-cell-navigation) | Sheet-level versus cell-level moves; no title selection on the viewer | Investigate | Medium |
| [Word navigation](#word-outline-and-pages) | One document unit is not one physical page | Clarify contract | Medium |
| [Find UI](#find-bar-and-results) | Result count versus painted-occurrence count; XLSX results are per sheet | Resolved | Low |
| [Data reads](#read-cells-slides-and-document-content) | Different read shapes and sparse data | Clarify contract | Medium |
| [Media](#pictures-and-detached-models) | Model alone does not carry media bytes | Keep | Medium |
| [Multiple views](#main-view-with-thumbnails-or-several-units) | Views from `viewer.content` must be destroyed before the viewer | Resolved | Low |
| [Whole-workbook search](#search-without-rendering) | No workbook-level search operation | Investigate | Medium |
| [Appearance](#theme-and-spreadsheet-appearance) | Element exposes only the theme | Resolved | Low |
| [CSS](#element-css-and-host-styling) | Shadow DOM and renderer defaults limit host overrides | Clarify contract | Medium |
| [Cleanup](#cleanup-and-ownership) | One teardown for viewers, two for low-level | Keep | Low |
| [Errors](#errors-and-events) | App-owned calls must be caught; built-in actions report through events | Resolved | Medium |
| [Cancellation](#rapid-source-or-unit-changes) | Low-level paging needs controller/handle bookkeeping | Keep | High |
| [Deployment](#worker-and-csp) | Special hosting still needs explicit CSP/CORS permissions | Resolved | Medium |
| [Fonts and sizing](#fonts-sizing-and-browser-support) | Successful load does not guarantee Office-identical layout | Clarify contract | Medium |

## Landing and first render

### Pick a route

**Goal:** Find the smallest API for showing a file, without studying package internals.

**Current path:** Choose viewer versus data access, then known versus unknown
format. A format entry carries both depths:

```ts
import { preview } from "aared/pptx";      // viewer for a known format
import { open, render } from "aared/pptx"; // document and manual rendering
// For unknown format: import { preview } from "aared/preview";
```

- **Friction:** Labeling `aared/pptx` "low-level" hides its `preview()` export.
  Multiple import names look like competing systems before their jobs are clear.
- **Cause:** Landing terminology mixes format and ownership. The exports
  themselves support both jobs, as the [package table](../site/src/content/docs/api.mdx#packages) shows.
- **Possible improvement:** Lead with the job, then format; show both functions
  from one entry instead of teaching subpath architecture first. Extra subpaths
  would add choices without removing this decision.
- **Disposition:** **Keep.** The [layer guide](../site/src/partials/choosing-a-layer.mdx)
  makes the viewer-versus-data choice before explaining format imports.

### HTML embed

**Goal:** Put a file on a page with built-in paging and zoom controls.

**Current path:** Install the package, register the element in the app entry,
write the tag and give it a height.

```ts
// App entry, built by Vite/webpack/etc.
import "aared/element";
```

```html
<aared-viewer src="report.pptx" controls style="height: 600px"></aared-viewer>
```

- **Friction:** "No UI code" can be mistaken for "only an HTML tag". Without
  registration the tag does nothing; without a bounded host some layouts surprise.
- **Cause:** Native Custom Elements need registration. Bare package imports need
  bundler resolution or an import map; sizing belongs to the embedding page.
- **Possible improvement:** Present registration, markup and height together.
  A no-build distribution would be a separate packaging decision, not something
  documentation can imply already exists.
- **Disposition:** **Keep.** Start gives registration, markup and sizing
  together; no no-build distribution is implied.

### JavaScript viewer

**Goal:** Open an uploaded Office file and build custom UI around it.

**Current path:** Use generic `preview()` for any format; a format entry for a
fixed format. Both accept one object and return a viewer.

```ts
import { preview } from "aared/preview";

const viewer = await preview({ source: file, target: container });
console.log(viewer.format, viewer.info.count);
// On unmount: viewer.destroy();
```

- **Friction:** The consumer meets two imports for `preview()`. A `type` assertion
  on the generic entry can also look equivalent to importing the format entry.
- **Cause:** Same viewer contract, different route selection and build reachability.
  Assertion/narrowing is not renderer bundle isolation.
- **Possible improvement:** Default unknown uploads to the generic entry; introduce
  format-local preview only for a fixed-format app. Put the build trade alongside
  that choice, not in a naming glossary.
- **Disposition:** **Keep.** Start gives a minimal generic recipe before the
  known-format route; keep the one-object call.

### React viewer

**Goal:** Show a file, loading/error state and retry inside a React component.

**Current path:** Attach the hook's ref to a persistent host, inspect `status`,
and use its ready viewer. Hooks do not draw the element's controls.

```tsx
import { usePreview } from "aared/react";

function Preview({ file }: { file: File }) {
  const { ref, status, error, retry } = usePreview({ source: file });
  return <>
    <div ref={ref} style={{ height: 600 }} />
    {status === "loading" && <p>Loading...</p>}
    {status === "error" && <p>{error.message} <button onClick={retry}>Retry</button></p>}
  </>;
}
```

- **Friction:** "Same, in React" can suggest a ready-made controlled widget.
  Returning only an error panel removes the host needed for retry. `index`/`zoom`
  look controlled but are start values; `appearance` is reactive.
- **Cause:** A lifecycle hook rather than a component; documented mount-versus-update
  semantics ([hook contract](../site/src/content/docs/api.mdx#react-hooks-aaredreact),
  [implementation](../packages/aared/src/react/index.ts)).
- **Possible improvement:** Show host, state and retry together. Evaluate whether
  start-value naming should be more explicit before adding controlled-state APIs;
  controlled props would need loop/synchronization rules.
- **Disposition:** **Keep** the mount/retry guidance. Start-value naming remains
  a design question, not an approved rename.

### Server rendering

**Goal:** Use the viewer in Next.js without browser-global errors on the server.

**Current path:** Import safely, mount from browser/client code.

```tsx
"use client";
import { usePreview } from "aared/react";

export function Preview({ file }: { file: File }) {
  const { ref } = usePreview({ source: file });
  return <div ref={ref} style={{ height: 600 }} />;
}
```

- **Friction:** SSR-safe import can be read as server-side rendering support.
- **Cause:** Rendering requires live DOM, fonts and measurement; imports do not.
- **Possible improvement:** Keep the browser boundary next to the framework
  example; no custom wrapper is necessary.
- **Disposition:** **Keep.** Start already supplies this path. No server rendering
  or no-browser thumbnail capability follows from import safety.

## Getting the file into the viewer

### URL, picker, drop and bytes

**Goal:** Accept a file from whichever input the app already has.

**Current path:** Functions/hooks use `source`; the element uses `src` for URLs
and `file` for files, blobs or bytes. Picker/drop handlers pass their `File`.

```ts
import "aared/element";

const element = document.querySelector("aared-viewer")!;
const input = document.querySelector<HTMLInputElement>('input[type="file"]')!;
input.addEventListener("change", () => {
  element.file = input.files?.[0] ?? null;
});
```

- **Friction:** Same source concept has attribute and property names on the
  element. Consumers must know the last-set source wins and how clearing works.
- **Cause:** URL strings fit HTML attributes; binary objects do not. This is not
  a different viewer per input source ([element contract](../site/src/content/docs/api.mdx#attributes--properties)).
- **Possible improvement:** Show source selection once, independently of API
  depth. Keep picker/drop examples as thin native event handlers.
- **Disposition:** **Keep.** Start now states the shared source choices and links
  to drop usage. No upload service or file-picker abstraction is needed.

### Authenticated fetch

**Goal:** Preview a file behind custom authentication headers.

**Current path:** Fetch using the app's authentication, check the response,
then pass a blob to the viewer.

```ts
import { preview } from "aared/preview";

const response = await fetch("/files/report", { headers: { Authorization: "Bearer token" } });
if (!response.ok) throw new Error(`Download failed: ${response.status}`);
const viewer = await preview({ source: await response.blob(), target: container });
// On unmount: viewer.destroy();
```

- **Friction:** The input is called a URL source, but app-specific headers require
  a separate download operation and its error/cancellation handling.
- **Cause:** The app owns authentication; native `fetch()` already handles it.
- **Possible improvement:** Keep the fetch-then-blob recipe adjacent to URL usage.
  A custom downloader API would duplicate native fetch and need a demonstrated
  benefit beyond reducing two lines.
- **Disposition:** **Keep.** The authenticated fetch recipe already lives on Start.

### Detect before opening

**Goal:** Route a file by format, then open it without downloading twice.

**Current path:** Detection is a separate operation. Reuse fetched bytes if both
the app and viewer need them.

```ts
import { detect, preview } from "aared/preview";

const response = await fetch("/files/report");
if (!response.ok) throw new Error(`Download failed: ${response.status}`);
const bytes = await response.arrayBuffer();
const type = await detect(bytes);
const viewer = await preview({ source: bytes, type, target: container });
// On unmount: viewer.destroy();
```

- **Friction:** `detect(url)` followed by `preview({ source: url, ... })` looks
  reusable but makes separate requests. Generic preview already detects internally.
- **Cause:** Stateless calls do not share a downloaded resource. See
  [detect implementation](../packages/aared/src/preview/index.ts) and
  [session opening](../packages/aared/src/embed-viewer.ts).
- **Possible improvement:** Warn against redundant detection in the normal open
  recipe. Keep byte reuse explicit for routing instead of introducing hidden caching.
- **Disposition:** **Clarify contract.** Start advises ordinary callers not to
  pre-detect. Generic preview gets the format from its Worker open and adopts
  that session; an explicit `detect()` is still a separate probe.

## Navigation and custom UI

### Custom toolbar

**Goal:** Add paging and zoom using the app's own design.

**Current path:** Buttons call viewer methods; counters read state and subscribe
to change events.

```ts
const update = (): void => { position.textContent = `${viewer.index + 1} / ${viewer.info.count}`; };
viewer.addEventListener("indexchange", update);
update();
await viewer.next();
viewer.zoomIn(0.1);
viewer.resetZoom();
```

- **Friction:** Wiring is application code; failures need promise/event handling.
  Numeric zoom is scale, not percent, and reset means format default rather than 100%.
- **Cause:** UI-independent handle plus explicit units. Native `EventTarget`
  supplies subscription; a custom observable would add a second state surface.
- **Possible improvement:** Preserve one complete toolbar recipe with initial
  update and error handling, as Start already does.
- **Disposition:** **Keep.** No wrapper or separate toolbar API is needed.

### Sheet tabs and cell navigation

**Goal:** Use custom tabs and an address box instead of spreadsheet defaults.

**Current path:** Hide built-in tabs; use top-level `goTo()` for sheets and XLSX
controls for cells. Subscribe to `sheetchange` for custom active-tab state.

```ts
import { preview } from "aared/xlsx";

const viewer = await preview({
  source: file, target: container,
  appearance: { spreadsheet: { sheetTabs: false } },
});
await viewer.goTo(1);
await viewer.xlsx.goToCell("D14");
viewer.xlsx.scrollTo({ range: "F120:J140", align: "center" });
// On unmount: viewer.destroy();
```

- **Friction:** Unit and cell moves live at different depths; low-level sheet
  renders also have `goToSheet(title)`, which is not a viewer method. (Hidden
  tabs now survive a theme switch: `setAppearance()` merges, §119.)
- **Cause:** Shared unit navigation versus format-native places. See
  [controls](../packages/aared/src/preview/index.ts) and
  [XLSX rendering](../site/src/content/docs/api.mdx#rendering-a-sheet).
- **Possible improvement:** Compare whether title-based sheet selection should be
  reachable from viewer controls. It would improve a normal sheet selector but
  add overlap with `goTo()`. Do not imply that it exists today.
- **Disposition:** **Investigate.** Basic index/cell navigation is already direct;
  title selection is the remaining seam. **UX cost:** Medium.

### Word outline and pages

**Goal:** Navigate a Word document by its outline or physical pages.

**Current path:** Headings work in either mode; pages require layout mode.

```ts
import { preview } from "aared/docx";

const viewer = await preview({ source: file, target: container, pagination: { mode: "layout" } });
const [heading] = viewer.docx.headings;
if (heading) await viewer.docx.goToHeading(heading.index);
await viewer.docx.goToPage(0);
// On unmount: viewer.destroy();
```

- **Friction:** `viewer.info.count` is one, not page count. `viewer.goTo(1)` cannot
  select page two; page count is zero in flow mode. Body headings exclude tables
  and text boxes even when those contain heading-styled text.
- **Cause:** Document structure and browser page layout are different concepts.
  [DOCX controls](../site/src/content/docs/api.mdx#auto-detect-preview-aaredpreview)
  reflect that distinction.
- **Possible improvement:** Show `viewer.docx.pageCount` and `goToPage()` beside
  the layout option. Keep section, heading and page terminology separate; adding
  a fake universal page collection would obscure source structure.
- **Disposition:** **Clarify contract.** Start links to page/heading examples;
  defaults and fidelity must be judged against a real reading task, not symmetry.

### Find bar and results

**Goal:** Offer a compact find bar or a clickable result list for any format.

**Current path:** Search the open viewer. Compact find uses `searchNavigation`;
a clickable results list hands each result to `goToSearchMatch()`, whatever
the format.

```ts
await viewer.search("budget");
await viewer.searchNavigation.next();
const { index, count } = viewer.searchNavigation;
counter.textContent = index === null ? "No active match" : `${index + 1} of ${count}`;
viewer.addEventListener("searchchange", updateCounter);
```

```ts
for (const match of await viewer.search("budget")) {
  const item = list.appendChild(document.createElement("li"));
  item.textContent = match.text;
  item.onclick = () => void viewer.goToSearchMatch(match);
}
```

- **Friction:** XLSX results cover the current sheet and need refreshing on a
  sheet change; a result from another sheet moves there but is no longer the
  active one. DOCX painted-occurrence count can exceed source-result count
  because headers/footers repeat; a clicked header result becomes the copy on
  the page being read.
- **Cause:** Source matches and rendered occurrences differ
  ([search contract](../site/src/content/docs/api.mdx#auto-detect-preview-aaredpreview)).
- **Resolution:** `viewer.goToSearchMatch(match)` takes `PreviewSearchMatch`
  on every handle, keeps the native match data, and makes a listed result the
  active one (Notes §119). `viewer.docx.goToSearchMatch` was removed.
- **Disposition:** **Resolved.** **UX cost:** Low. Workbook-wide results remain
  a separate question.

## Data access and combined views

### Read cells, slides and document content

**Goal:** Build an index, data inspection panel or custom visualization.

**Current path:** Open the known format, then use its native data access. An
app that already shows the file reads the same way from `viewer.content`,
without opening it again.

```ts
import { open } from "aared/xlsx";

const workbook = await open(file);
try {
  const data = await workbook.sheets.get(0).read({ range: "A1:D20" });
  for (const row of data.rows) {
    for (const { col, cell } of row.cells) console.log(row.index, col, cell.formatted);
  }
} finally { workbook.close(); }
```

```ts
// With opened documents from their matching format entries:
const slide = await deck.slides.get(0).read();
const sections = doc.sections.list(); // resolved eagerly; no doc.read()
const headings = doc.headings.list();
```

- **Friction:** One imagined universal `read()` does not exist. Sparse XLSX rows
  require `{ col, cell }` iteration; cells have no address property. PPTX metadata
  titles are generic until actual text is read; DOCX content is already resolved.
- **Cause:** Format-native model contracts and lazy versus eager parsing. See
  [range reads](../site/src/content/docs/api.mdx#sheets-and-range-reading),
  [slides](../site/src/content/docs/api.mdx#slides), and [DOCX](../site/src/content/docs/api.mdx#docx-aareddocx).
- **Possible improvement:** Lead data scenarios with the wanted output, including
  sparse iteration. If ordinary cell access is repeatedly cumbersome, investigate
  a helper using the existing range model before changing the model itself.
- **Disposition:** **Clarify contract.** Start now includes a real range-read
  recipe. Do not add universal reads merely to make the three formats look alike.

### Pictures and detached models

**Goal:** Build a picture gallery or render a retained resolved model.

**Current path:** Get media bytes from an open document. A detached model render
needs a resolver; rendering a unit handle supplies session-backed media itself.

```ts
const bytes = await deck.media.get(mediaId);
const model = await deck.slides.get(0).read();
const view = await render({ slide: model, target: container, mediaResolver: (id) => deck.media.get(id) });
// On unmount: view.destroy(); deck.close();
```

- **Friction:** Serializable model data does not mean a self-contained document.
  Closing the session before resolving pictures makes a detached model incomplete.
- **Cause:** Media bytes and resolved model lifetimes are intentionally separate;
  see [Media](../site/src/content/docs/api.mdx#media).
- **Possible improvement:** Prefer handle-based rendering in first examples.
  For retained models, explicitly retain needed media; compare an asset snapshot
  capability only if that consumer job warrants its memory/API cost.
- **Disposition:** **Keep.** Resolver plumbing belongs to the detached-model job,
  not every ordinary preview.

### Main view with thumbnails or several units

**Goal:** Show a main slide and clickable thumbnails using one parsed deck.

**Current path:** The viewer lends restricted read/render content, not its
opened document. Thumbnails and reads use `viewer.content`, with one parse.

```ts
import { preview, render } from "aared/pptx";

const viewer = await preview({ source: file, target: mainHost });
const thumbnail = await render({ slide: viewer.content.slides.get(0), target: thumbHost });
// On unmount: thumbnail.destroy(); viewer.destroy(); (the viewer closes the document)
```

- **Friction:** The viewer owns the document, so its views must be destroyed
  before the viewer, and reads after `destroy()` reject `closed`. There is no
  exposed document `close()`; metadata and already-returned read-only data remain
  inspectable, but new secondary renders depend on the viewer's lifetime.
- **Cause:** One owner per document; [fixed-format preview contract](../site/src/content/docs/api.mdx#single-format-preview).
- **Resolution:** Restricted `viewer.content` on every preview handle, owned by
  the viewer (Notes §122 corrects §119's full-document exposure and keeps
  §69's lifecycle boundary). Reads/metadata are intentional content capabilities;
  search stays on the viewer. DOCX renders from the facade with images/headings
  intact. A viewer over a caller-owned document was rejected as a
  second lifetime story on the one-call path. The low-level route remains for an
  app that pages itself or keeps a document past any viewer.
- **Disposition:** **Resolved.** **UX cost:** Low.

### Search without rendering

**Goal:** Search an entire workbook for an audit/exported data report.

**Current path:** Sheets own data-only search, so the app loops through them.

```ts
for (const sheet of workbook.sheets.list()) {
  for (const hit of await sheet.search("revenue")) {
    console.log(sheet.info.title, hit.address);
  }
}
// PPTX: deck.search(query); DOCX: doc.search(query)
```

- **Friction:** Whole-deck and whole-document search are direct calls, but a
  whole-workbook job needs a loop and a decision about hidden sheets and cancellation.
- **Cause:** Sheet-local Find semantics also define the low-level XLSX surface;
  see [Searching a workbook](../site/src/content/docs/api.mdx#searching-a-workbook).
- **Possible improvement:** Evaluate a workbook-wide search with explicit sheet
  scope, preserving the sheet search used by viewers. Balance common audit jobs
  against another callable and potentially large first-search cost.
- **Disposition:** **Investigate.** Keep cached-value/formatted-text semantics;
  adding formula recalculation is not a search simplification.

## Customization and integration

### Theme and spreadsheet appearance

**Goal:** Change theme while preserving custom tabs, gridlines and selection.

**Current path:** Pass only what changes; the element only exposes theme.

```ts
viewer.setAppearance({ theme: "dark" }); // hidden tabs stay hidden
// React: usePreview({ source: file, appearance: { theme, spreadsheet: { sheetTabs: false } } })
// Element: element.theme = "dark"
```

- **Friction:** A knob goes back to its default only when the default is passed.
  The React option is whole state (a removed knob resets), while the imperative
  call merges. Consumers switching from element to JS meet a larger settings surface.
- **Cause:** One merge rule for the imperative call, declarative state for the
  hook; limited element settings; [Appearance](../site/src/content/docs/api.mdx#appearance).
- **Resolution:** Every `setAppearance()`, viewer and render handles alike,
  merges (Notes §119). No second patch method, so there is one update story.
- **Disposition:** **Resolved.** **UX cost:** Low. Theme still affects chrome,
  never authored content.

### Element CSS and host styling

**Goal:** Make the viewer match the app without depending on internal markup.

**Current path:** Use shadow parts/tokens for the element; use appearance or a
low-level host for spreadsheet internals.

```css
aared-viewer { --aared-page-shadow: 0 1px 5px rgb(0 0 0 / 0.25); }
aared-viewer::part(controls) { font-family: system-ui; }
```

- **Friction:** Host CSS cannot reach every shadow-internal node. Renderer
  defaults also have deliberate specificity; copying `outerHTML` loses defaults.
- **Cause:** Encapsulation and CSSOM stylesheet delivery;
  [Styling and CSP](../site/src/content/docs/api.mdx#styling-and-csp).
- **Possible improvement:** Judge hooks against concrete styling jobs rather than
  exporting every class. A new part/token should solve a visible app requirement.
- **Disposition:** **Clarify contract.** Keep stable hooks distinct from internal
  classes; this audit does not treat style serialization as HTML export support.

### Cleanup and ownership

**Goal:** Unmount cleanly or retain data after destroying one view.

**Current path:** Viewer cleanup owns everything; low-level cleanup is split.

```ts
viewer.destroy(); // plain JS preview
view.destroy();   // low-level render only
doc.close();      // low-level document/session
// Element: remove it. React hook: unmount; no public destroy().
```

- **Friction:** Moving to data access introduces a second lifetime. Hook and
  element handles intentionally omit methods their wrappers own.
- **Cause:** A retained document may serve multiple views/read operations;
  [lifecycle contracts](../site/src/content/docs/api.mdx#single-format-preview).
- **Possible improvement:** Keep teardown next to each creation example. Do not
  collapse low-level destroy/close and lose the ability to keep reading.
- **Disposition:** **Keep.** Two lifetimes are justified for this job; the audit
  should expose them, not omit cleanup to make the code shorter.

### Errors and events

**Goal:** Report one useful failure and ignore intentional cancellation.

**Current path:** Every call the app makes throws/rejects only. Built-in actions
(sheet tabs, element attributes and controls) catch their own failures and
report one `error` event. React has one `error` field for mount/built-in failures;
an imperative action's catch belongs to the application.

```ts
import { isAbort, messageFor } from "aared/core";

const report = (error: unknown): void => {
  if (!isAbort(error)) status.textContent = messageFor(error);
};
next.onclick = () => void viewer.next().catch(report);
find.onclick = () => void viewer.search("budget").catch(report);
viewer.addEventListener("error", (event) => {
  report(event.detail.cause); // built-in actions, never duplicate the calls above
});
```

- **Friction:** An app-owned button must catch its promise (`void` alone is not
  handling); a listener only covers built-in actions, not everything on a viewer.
- **Cause:** Reporting follows action ownership, not which method failed.
- **Resolution:** One rule across navigation, search and all embed layers
  (Notes §120); sheet-tab failures no longer disappear. React's former
  `navigationError` field is removed. The [error recipe](../site/src/content/docs/api.mdx#errors)
  handles both owners without duplicate notification.
- **Disposition:** **Resolved.** **UX cost:** Medium: app callbacks still own
  their error presentation. No error-policy option or catch-detection magic.

### Rapid source or unit changes

**Goal:** Prevent a slower old file/slide from replacing the latest selection.

**Current path:** Element/hooks own source replacement; viewer owns navigation;
manual low-level paging uses signals and handles.

```ts
import { open, render, type PptxRenderHandle } from "aared/pptx";
import { isAbort } from "aared/core";

const deck = await open(file);
let latest: AbortController | undefined;
let shown: PptxRenderHandle | undefined;

async function show(index: number): Promise<void> {
  latest?.abort();
  const mine = (latest = new AbortController());
  try {
    shown = await render({
      slide: deck.slides.get(index), target: container, signal: mine.signal,
      ...(shown === undefined ? {} : { zoom: shown.zoom }),
    });
  } catch (error) {
    if (!isAbort(error)) throw error;
  }
}

// On unmount:
function cleanup(): void {
  latest?.abort();
  shown?.destroy();
  deck.close();
}
```

- **Friction:** The complete low-level pattern is larger than one render call:
  keep zoom, catch cancellation and clean up pending and mounted work.
- **Cause:** Consumer deliberately owns paging, while viewers already supply
  newest-wins behavior ([Cancellation](../site/src/content/docs/api.mdx#cancellation)).
- **Possible improvement:** Default a single reading surface to preview. A
  gallery beside a viewer now renders from `viewer.content` (see Multiple
  views), so the bookkeeping above is left to apps that page with no viewer.
- **Disposition:** **Keep** the existing viewer behavior. The extra low-level
  work must remain visible when judging the combined-view scenario.

### Worker and CSP

**Goal:** Deploy under the app's bundler, CDN and security policy.

**Current path:** Default bundler setup needs no override. One explicit Worker
input handles an emitted URL, inline source, or a native factory; aared owns
generated Blob URLs and their cleanup.

```ts
// worker.ts
import "aared/worker";
```

```ts
import AaredWorker from "./worker?worker"; // Vite
import { open } from "aared/pptx";
const deck = await open(file, { worker: { create: () => new AaredWorker() } });
// On completion: deck.close();
```

- **Friction:** CDN and single-file deployments still need the appropriate
  browser security policy; a package cannot grant CSP/CORS permissions.
- **Cause:** Browser origin/security and bundler asset rules, not format parsing.
  Canonical recipes: [Worker override](../site/src/content/docs/api.mdx#worker-override).
- **Resolution:** `worker: { url, crossOrigin: "blob" }` for a CDN,
  `{ source }` for a single-file build, and `{ create }` for an advanced native
  factory (Notes §121). No standalone helper, implicit security fallback,
  or application-owned Blob cleanup. The existing production-build tool checks
  all three formats with Vite and webpack, including refused configurations.
- **Disposition:** **Resolved.** **UX cost:** Medium: explicit host permissions
  remain, but browser resource bookkeeping is package-owned.

### Fonts, sizing and browser support

**Goal:** Show a readable, correctly scrolling document with predictable fidelity.

**Current path:** Bound the host, accept browser font fallback, or provide an
aliased FontFace resolver for fonts the app delivers.

```ts
import { fontAlias, type FontResolver } from "aared/core";

const resolver: FontResolver = async (family) => {
  if (family !== "Calibri") return undefined;
  const face = new FontFace(fontAlias(family), "url(/fonts/carlito.woff2)");
  document.fonts.add(face);
  return face;
};
// Pass fontResolver to open(), preview() or the hook. The element takes no resolver.
```

- **Friction:** A successful open can still wrap differently from Office. Unbounded
  sheets/documents scroll with the page; missing fonts are not a parsing failure.
  Browser support for highlighting is separate from basic viewing.
- **Cause:** Browser text/layout and embedding geometry;
  [fonts](../site/src/content/docs/api.mdx#sources-and-options),
  [sizing](../site/src/content/docs/start.mdx#sizing),
  [browser support](../site/src/content/docs/api.mdx#browser-support).
- **Possible improvement:** Include height in first embeds and explain fallback
  at the fidelity boundary. Keep native FontFace delivery; do not promise exact
  Office line breaks merely because a substitute was supplied.
- **Disposition:** **Clarify contract.** No-font-download remains the default.

## Scope checks

The audit evaluates delivered view-only journeys. Editing, write-back, formula
recalculation and conversion are not missing shortcuts to propose. Printing,
CSV/TSV input and no-browser thumbnails are not available capabilities.

Use [known limits](ARCHITECTURE.md#known-limits) for accepted fidelity differences
and the [gap inventory](ARCHITECTURE.md#gaps-limits-and-fidelity-gaps) for open
capability work. A fidelity defect needs its own
source/model/DOM/native-Office investigation, not an API naming workaround.
