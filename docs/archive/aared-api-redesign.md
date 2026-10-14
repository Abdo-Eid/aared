> ⚠️ **Superseded — provenance only.** This proposal predates the API as
> built. The current public API is the API reference page
> (`site/src/content/docs/api.mdx`); why it took the shape it did is in
> [`docs/Notes.md`](../Notes.md). Nothing here is a
> contract.

# aared Public API Redesign Proposal

## Status

Design proposal for the public API before the package surface is frozen.

This proposal assumes the existing architectural direction remains valid:

- `aared` is a **client-side previewer**, not an authoring SDK.
- PPTX, XLSX, and DOCX share a headless OOXML/core layer.
- Parsing/resolution happens in a Worker.
- Headless DOCX pagination/layout work also happens in the Worker.
- Rendering happens on the main thread using real DOM/SVG.
- Text remains browser-rendered.
- Each Office format has materially different layout semantics.
- Consumers may want to ship support for **only one format**, so format-specific entrypoints must remain first-class.
- DOCX renders as continuous flow by default and supports opt-in structural physical-page layout.
- DOCX sections are source/document structure; physical pages are derived layout output.

The goal here is not to force every format into one abstraction.

The goal is to make the package feel like **one coherent library**.

---

# Executive Recommendation

Keep the format entrypoints:

```text
aared/pptx
aared/xlsx
aared/docx
```

and make them share the same **API grammar**, lifecycle rules, naming conventions, error behavior, styling philosophy, and interaction patterns.

Do **not** make them literally identical APIs.

Use format-native terminology:

```ts
deck.slides.get(3);
workbook.sheets.get("Revenue");
document.sections.get(0);
```

rather than forcing everything into generic concepts such as:

```ts
doc.views.get(...)
```

just for consistency.

The core principle should be:

> **Same API grammar, format-native vocabulary.**

That gives us consistency without sacrificing clarity.

---

# Why Keep Separate Format Entrypoints

A consumer may only need one file type.

For example:

```ts
import { open, render } from "aared/pptx";
```

should not force that application to ship the XLSX and DOCX renderers.

Likewise:

```ts
import { open, render } from "aared/xlsx";
```

should not pull in PPTX or DOCX rendering code.

The subpath separation is therefore not merely organization.

It is part of the package's bundle-size and deployment contract.

Recommended first-class public entrypoints:

```text
aared/pptx
aared/xlsx
aared/docx
aared/core
aared/element
aared-fonts
```

Potentially later:

```text
aared
```

could exist as an optional convenience entrypoint for consumers who explicitly want automatic format detection and are comfortable supporting multiple formats.

That all-format entrypoint should be optional, not the architectural foundation.

---

# The Main Design Principle

There are two extremes we should avoid.

## Bad extreme 1: three unrelated libraries

```text
aared/pptx
aared/xlsx
aared/docx
```

with unrelated naming, lifecycle, errors, rendering signatures, styling, and events.

This forces developers and AI agents to relearn the library three times.

## Bad extreme 2: one fake universal Office model

Trying to pretend these are the same thing:

```text
PowerPoint slide
Excel sheet
Word section
```

would produce vague concepts such as:

```ts
document.views.get(...)
```

and eventually an overly generic model that hides useful format semantics.

The better middle ground is:

```text
same library conventions
+
native format concepts
```

---

# API Grammar

A developer should be able to learn one pattern:

```text
open()
  ↓
document.<native collection>.list()
document.<native collection>.get(...)
  ↓
read() when needed
  ↓
render(...)
  ↓
destroy()
  ↓
close()
```

Across formats:

```ts
// PPTX
const deck = await open(file);
const slide = deck.slides.get(0);

// XLSX
const workbook = await open(file);
const sheet = workbook.sheets.get("Revenue");

// DOCX
const document = await open(file);
const sections = document.sections.list();
```

The nouns differ because the file formats differ.

The grammar remains familiar.

---

# Shared Naming Conventions

The following concepts should behave consistently across entrypoints where they make sense:

```text
open()
render()

info
capabilities

list()
get()
read()

media.get()

goTo()
scrollTo()
focus()
search()
setZoom()

destroy()
close()

AbortSignal
typed errors
progress
events
appearance
```

These conventions are what should be unified.

---

# What Should Not Be Unified

Do not force format-specific concepts behind generic names.

Prefer:

```ts
deck.slides
workbook.sheets
document.sections
document.headings
```

over:

```ts
doc.units
doc.views
doc.items
```

Prefer:

```ts
sheet.read("A1:F200");
```

over a generic abstraction that hides the fact that spreadsheets have cell ranges.

Prefer:

```ts
deck.slides.get(4);
```

over:

```ts
deck.views.get(4);
```

Prefer:

```ts
document.headings.list();
```

over pretending headings are equivalent to sheets or slides.

---

# Recommended Package Surface

## PPTX

```ts
import {
  open,
  render,
} from "aared/pptx";
```

Possible object model:

```ts
const deck = await open(file);

deck.info;
deck.slides;
deck.media;
deck.capabilities;

deck.close();
```

Slide access:

```ts
const slide = deck.slides.get(0);
```

Listing:

```ts
for (const slide of deck.slides.list()) {
  console.log(slide.index, slide.title);
}
```

Reading:

```ts
const model = await slide.read();
```

Rendering:

```ts
const rendered = await render({
  slide,
  target: container,
});

rendered.destroy();
```

---

# XLSX

```ts
import {
  open,
  render,
} from "aared/xlsx";
```

Possible object model:

```ts
const workbook = await open(file);

workbook.info;
workbook.sheets;
workbook.media;
workbook.capabilities;

workbook.close();
```

Sheet access:

```ts
const sheet = workbook.sheets.get("Revenue");
```

or:

```ts
const sheet = workbook.sheets.get(1);
```

Listing:

```ts
for (const sheet of workbook.sheets.list()) {
  console.log(sheet.index, sheet.name);
}
```

Range reading:

```ts
const range = await sheet.read("A1:M200");
```

or:

```ts
const range = await sheet.read({
  start: { row: 0, col: 0 },
  end: { row: 199, col: 12 },
});
```

Rendering:

```ts
const rendered = await render({
  sheet,
  target: container,
});

rendered.destroy();
```

---

# DOCX

```ts
import {
  open,
  render,
} from "aared/docx";
```

Possible object model:

```ts
const document = await open(file);

document.info;
document.sections;
document.headings;
document.media;
document.capabilities;

document.close();
```

Section access:

```ts
const section = document.sections.get(0);
```

Listing:

```ts
const sections = document.sections.list();
```

Outline:

```ts
const headings = document.headings.list();
```

A DOCX section is source/document structure. It is not a physical page and should not be treated as the Word equivalent of a PowerPoint slide.

The normal DOCX rendering input is the ordered section set. Flow is the default:

Flow mode renders the ordered sections directly as continuous content and does not run the physical pagination/layout step. Pagination is only performed when pagination.mode is "layout".

That makes the pipeline explicit:
```
DocumentSection[]
   ├─ flow   → render continuously
   └─ layout → paginate → render pages
```

flow concatenates the section roots, while the paginator is the separate layout path.

```ts
const rendered = render(sections, container);

rendered.destroy();
```

Applications that need physical pages can opt into layout:

```ts
const rendered = render(sections, container, {
  pagination: {
    mode: "layout",
    softBreaks: "prefer", // "ignore" | "prefer" | "force"
  },
});
```

The soft-break policy is DOCX-specific and should remain format-native rather than being generalized across PPTX/XLSX.

Physical pages are derived by layout and belong to the render result rather than the source document model. Do not introduce a source collection such as:

```ts
document.pages.get(...);
```

The render handle exposes `pageCount` and `pages`. In flow mode, `pageCount` is `1` and `pages` is empty. In layout mode, `pages` contains the derived physical-page metadata. `document.sections` remains the authored/source structure.

The fidelity boundary should be explicit:

> **DOCX physical pagination is supported, but pixel-identical Microsoft Word pagination is not part of the contract.**

This keeps DOCX native to its actual structure without forcing it into the single-item rendering model used naturally by PPTX and XLSX.

---

# Shared Collection Convention

Where a format exposes an addressable collection, use the same collection grammar.

For example:

```ts
interface Collection<TItem, TKey = number> {
  list(): readonly TItem[];
  get(key: TKey): TItem;
}
```

Examples:

```ts
deck.slides.list();
deck.slides.get(3);

workbook.sheets.list();
workbook.sheets.get("Revenue");

document.sections.list();
document.sections.get(0);
```

The implementation can use generic collection helpers internally.

The public API should expose domain-specific names.

---

# Lightweight Handles

Where incremental parsing materially improves performance, collection entries should ideally be lightweight handles rather than eagerly parsed full models.

PPTX example:

```ts
interface SlideHandle {
  readonly info: SlideInfo;
  read(options?: ReadOptions): Promise<SlideModel>;
}
```

XLSX example:

```ts
interface SheetHandle {
  readonly info: SheetInfo;

  read(
    range?: CellRange | A1Range,
    options?: ReadOptions
  ): Promise<SheetRangeModel>;
}
```

This preserves the existing performance principle:

> index first, parse the requested content second.

Do not turn that optimization into a universal public abstraction. DOCX sections are meaningful resolved source structures consumed together by rendering/layout, so they do not need an artificial `section.read()` method merely to match PPTX/XLSX.

---

# `open()` Should Be Consistent

Each entrypoint should expose the same high-level opening convention:

```ts
const document = await open(source, options);
```

Recommended source types:

```ts
type Source =
  | File
  | Blob
  | ArrayBuffer
  | string;
```

Recommended common options:

```ts
interface OpenOptions {
  signal?: AbortSignal;
  fontResolver?: FontResolver;
}
```

Format-specific entrypoints already know the expected format.

Therefore:

```ts
open()
```

in `aared/xlsx` does not need to return a PPTX or DOCX document.

If the file is not actually XLSX, it should fail with a typed format error.

This improves type inference and simplifies integration code.

---

# `render()` Should Be Consistent

Each entrypoint should expose `render()` with predictable lifecycle and naming, but the call shape does not need to be literally identical when the format's natural rendering unit differs.

Keep the proposed PPTX and XLSX higher-level signatures:

```ts
render({
  slide,
  target,
  appearance,
});
```

```ts
render({
  sheet,
  target,
  appearance,
});
```

DOCX now has a document-oriented render signature because its normal input is the ordered section set:

```ts
render(sections, target, {
  pagination: {
    mode: "flow",
  },
});
```

This is a format-specific difference, not a reason to redesign PPTX/XLSX into positional render APIs.

For APIs that use a structured options object, keep it extensible because it:

- scales better,
- is easier for AI agents to generate correctly,
- is easier to extend without breaking call sites,
- avoids ambiguous positional parameters.

The consistency target should be shared concepts and lifecycle semantics, not forced signature symmetry.

---

# Should `preview()` Exist?

This should remain an **open design question**, not a forced recommendation.

An earlier direction was:

```ts
preview({
  source,
  target,
});
```

as a one-call convenience API.

There is value in this:

- very easy README example,
- ideal for simple application usage,
- easy for AI agents,
- can create the document + renderer + controller together.

However, it also introduces another concept beside:

```text
open()
render()
<aared-viewer>
```

and `aared` is already, by definition, a previewing library.

We should not add `preview()` just because it sounds convenient.

The decision should come from real integration usage.

## Keep `preview()` only if

the common application flow repeatedly looks like:

```ts
const doc = await open(file);
const item = ...;
const viewer = await render(...);
```

and most consumers immediately need a controller wrapping all three.

## Do not add `preview()` if

`render()` plus the web component already cover the simple path cleanly.

For now, treat `preview()` as:

```text
candidate convenience API
not yet part of the recommended frozen surface
```

---

# Optional All-Formats Entrypoint

A future optional API could be:

```ts
import { open } from "aared";
```

with automatic format detection.

Internally it could lazy-import the correct format implementation.

However:

- it should not replace `aared/pptx`,
- it should not replace `aared/xlsx`,
- it should not replace `aared/docx`,
- it should not make one-format consumers pay for all formats,
- it should not become the only documented API.

Think of it as an explicit convenience layer for applications that genuinely support multiple Office file types.

---

# Public Types Should Follow the Same Philosophy

Use shared structural types where the concepts really are shared.

Examples:

```text
Size
Bounds
Transform
MediaResolver
FontResolver
RenderHandle
Appearance
AaredError
ProgressEvent
```

Use format-specific types where meaning differs.

Examples:

```text
SlideModel
ShapeModel
TableModel

SheetModel
XlsxCell
CellStyle
CellRange

DocumentSection
SectionProperties
SectionPageGeometry
FlowNode
DocxPageBreakNode
HeadingInfo
DocxRenderOptions
```

---

# XLSX Needs the Biggest Public API Improvement

The XLSX API should become a first-class readonly spreadsheet-viewing API.

It should not merely expose a rendered HTML table.

The important requirements are:

```text
range reading
virtualization
navigation
selection
keyboard interaction
sheet tabs
frozen panes
RTL
authored styles
viewer styling
programmatic cell targeting
```

---

# XLSX Range-Based Reading

A large spreadsheet must not imply:

```ts
const wholeSheet = await sheet.read();
```

as the only API.

The public model should support range reads directly.

Preferred syntax:

```ts
const data = await sheet.read("A1:M200");
```

Structured alternative:

```ts
const data = await sheet.read({
  start: { row: 0, col: 0 },
  end: { row: 199, col: 12 },
});
```

A1 syntax should be first-class because it is:

- familiar to spreadsheet users,
- concise,
- easy to document,
- highly recognizable to coding agents.

---

# XLSX Virtualization

The renderer should own virtualization.

Consumers should not manually compute visible ranges.

Expected internal flow:

```text
viewport changes
      ↓
calculate visible + overscan range
      ↓
sheet.read(range, { signal })
      ↓
render current window
      ↓
reuse DOM nodes
      ↓
evict distant cached ranges
```

The renderer should support both:

```text
row virtualization
column virtualization
```

for genuinely large sheets.

---

# XLSX Cancellation During Scroll

Range loading must integrate with cancellation.

Example:

```text
viewport at row 200
      ↓
range request starts
      ↓
user jumps to row 30,000
      ↓
old request is cancelled
      ↓
new range request gets priority
```

API:

```ts
await sheet.read("A1:Z500", {
  signal,
});
```

---

# XLSX Navigation

The readonly viewer should expose spreadsheet-native navigation.

Examples:

```ts
viewer.goTo("D14");
```

or:

```ts
viewer.goTo({
  cell: "D14",
});
```

Scroll without activating:

```ts
viewer.scrollTo({
  cell: "F120",
  align: "center",
});
```

Range navigation:

```ts
viewer.scrollTo({
  range: "F120:J140",
  align: "center",
});
```

Focus:

```ts
viewer.focus("F120");
```

The exact overloads can be decided later, but the concepts should exist.

---

# XLSX Selection

Readonly selection belongs in a viewer.

Clicking:

```text
D14
```

can select the cell without enabling editing.

Keyboard navigation should support at least:

```text
ArrowLeft
ArrowRight
ArrowUp
ArrowDown
Home
End
PageUp
PageDown
```

Potential later support:

```text
Ctrl/Cmd + Home
Ctrl/Cmd + End
Shift + Arrow
```

for range selection.

None of this implies authoring.

---

# XLSX Interaction Events

Useful events:

```text
cellactivate
selectionchange
linkactivate
```

Example:

```ts
viewer.addEventListener("cellactivate", event => {
  console.log(event.detail);
});
```

Possible detail:

```ts
{
  sheet: "Revenue",
  address: "D14",
  row: 13,
  column: 3,
  cell: {
    ...
  },
}
```

This is especially useful for agent-driven interfaces that need to point the UI at specific spreadsheet locations.

---

# XLSX Default Viewer Experience

A polished spreadsheet viewer should look like a spreadsheet by default.

Recommended default viewer behavior:

```ts
{
  headers: true,
  sheetTabs: true,
  selection: true,
  gridlines: "source",
}
```

A lower-level renderer can remain more minimal.

Do not require consumers to manually recreate standard spreadsheet chrome just to get a useful preview.

---

# Authored Styling vs Viewer Styling

This distinction is critical.

There are two completely different styling systems.

---

## Authored document styling

This comes from the Office file.

Examples:

```text
cell fill
font
font size
font weight
font style
font color
border
alignment
reading order
number format

PowerPoint fill
gradient
stroke
text layout
font
table style
```

These are part of document fidelity.

They must remain in the parsed/resolved models.

They are **not UI theme settings**.

---

## Viewer/UI styling

This belongs to the embedding application.

Examples:

```text
viewer background
toolbar
buttons
sheet tabs
row headers
column headers
gridlines
selection
focus ring
loading state
error state
scrollbars
status UI
controls
```

These should primarily be controlled through:

```text
good defaults
CSS variables
shadow DOM parts
small appearance configuration
```

---

# `appearance`

Keep the JavaScript appearance API small.

Example:

```ts
render({
  sheet,
  target,

  appearance: {
    theme: "system",
    preset: "default",
  },
});
```

Possible initial themes:

```ts
type Theme =
  | "light"
  | "dark"
  | "system";
```

Possible presets:

```ts
type AppearancePreset =
  | "default"
  | "minimal";
```

Do not build a large theme framework before real consumers need it.

---

# CSS Variables

Prefer CSS variables for visual customization.

Example:

```css
aared-viewer {
  --aared-background: #f7f7f8;
  --aared-surface: #ffffff;

  --aared-border: #e5e7eb;
  --aared-text: #111827;
  --aared-muted-text: #6b7280;

  --aared-accent: #2563eb;

  --aared-sheet-header-bg: #f4f4f5;
  --aared-sheet-header-text: #52525b;
  --aared-gridline: #e4e4e7;

  --aared-selection: #2563eb;
  --aared-focus-ring: #2563eb;
}
```

This is preferable to a giant JavaScript object with every color and spacing value.

---

# Shadow DOM Parts

Expose structural parts for consumers that need deeper theming.

Recommended examples:

```text
::part(stage)
::part(content)
::part(controls)
::part(toolbar)

::part(sheet-tabs)
::part(sheet-tab)
::part(row-header)
::part(column-header)

::part(status)
::part(error)
```

The exact part list should follow real DOM structure rather than being invented purely for API symmetry.

---

# Do Not Mix Authored Style Overrides Into the Main API

Avoid APIs like:

```ts
render({
  sheet,

  cellStyle(cell) {
    return {
      background: "...",
    };
  },
});
```

as part of the core public surface.

That starts turning `aared` into a programmable spreadsheet component rather than a faithful previewer.

It also creates difficult interactions with:

- conditional formatting,
- source styles,
- virtualization,
- caching,
- selection,
- rendering performance.

---

# Viewer Decorations

If applications need visual annotations, introduce them separately from Office styles.

Example concept:

```ts
decorations: [
  {
    range: "C4:C10",
    kind: "highlight",
  },
];
```

Decorations should:

- be viewer-only,
- never mutate the Office file,
- never pretend to be authored formatting,
- render as overlays or clearly separate annotations.

This preserves the boundary:

```text
Office fidelity
!=
application annotation
```

---

# XLSX Canonical Cell Model

There should be one canonical cell model across code and documentation.

Recommended direction:

```ts
interface XlsxCell {
  value:
    | string
    | number
    | boolean
    | null;

  formatted: string;

  type: CellType;

  formula?: string;

  hyperlink?: {
    target: string;
    display?: string;
  };

  style?: CellStyle;

  conditional?: ConditionalStyle;
}
```

---

# XLSX Cell Style

Example:

```ts
interface CellStyle {
  fill?: CellFill;

  font?: {
    family?: string;
    size?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
  };

  alignment?: {
    horizontal?: string;
    vertical?: string;
    wrap?: boolean;
    readingOrder?: "context" | "ltr" | "rtl";
    rotation?: number;
    indent?: number;
  };

  border?: CellBorder;
}
```

This is readonly fidelity data.

Consumers may inspect it.

The renderer consumes it.

There should be no corresponding setters.

---

# XLSX Grid Styling

Viewer-level spreadsheet chrome should be configurable without modifying cell styles.

Example:

```ts
appearance: {
  spreadsheet: {
    headers: true,
    gridlines: "source",
  },
}
```

Potential gridline behavior:

```ts
type GridlineMode =
  | "source"
  | "show"
  | "hide";
```

Headers:

```ts
headers: boolean;
```

Do not expose benchmark/test-only concepts such as print comparison framing as normal viewer options.

---

# Sheet Tabs

The built-in XLSX viewer should support sheet tabs.

Expected behavior:

- visible sheet names,
- active sheet state,
- hidden sheets excluded by default,
- RTL-safe layout,
- keyboard accessibility,
- clicking a tab changes the active sheet,
- programmatic sheet navigation stays synchronized.

The headless API should remain separate from this UI.

---

# Frozen Panes

Frozen panes should be treated as document fidelity, not as optional viewer decoration.

The viewer should preserve:

```text
frozen rows
frozen columns
```

and make headers interact correctly with them.

Navigation APIs must account for frozen panes when scrolling or focusing a target cell.

---

# RTL

RTL is part of the core value proposition and must stay explicit.

For XLSX:

```text
sheet RTL controls visual column order
cell reading order controls cell text direction
```

Do not let sheet-level RTL force all cell text to RTL.

For PPTX:

- table direction and cell paragraph direction remain distinct,
- paragraph-level bidi must remain independent.

For DOCX:

- text and flow direction should remain semantic and browser-rendered in both flow and layout modes.

---

# Capabilities

Each format document should expose capabilities.

Example XLSX:

```ts
workbook.capabilities;
```

Possible result:

```ts
{
  textSelection: true,
  media: true,

  spreadsheet: {
    rangeRead: true,
    cellNavigation: true,
    selection: true,
    formulas: "cached-only",
  },

  editing: false,
}
```

Example PPTX:

```ts
{
  textSelection: true,
  media: true,

  presentation: {
    notes: true,
    charts: true,
    animations: false,
  },

  editing: false,
}
```

Capabilities help consumers avoid relying on format-name conditionals for every behavior.

---

# Typed Errors

All entrypoints should share one error model.

Example:

```ts
class AaredError extends Error {
  code: AaredErrorCode;
  cause?: unknown;
}
```

Possible codes:

```ts
type AaredErrorCode =
  | "unsupported-format"
  | "invalid-document"
  | "encrypted-document"
  | "unsupported-feature"
  | "network-error"
  | "media-error"
  | "font-error"
  | "aborted";
```

Example:

```ts
try {
  const workbook = await open(file);
} catch (error) {
  if (
    error instanceof AaredError &&
    error.code === "encrypted-document"
  ) {
    // show a password-protected-file message
  }
}
```

The same error codes should mean the same thing in all format packages.

---

# Cancellation

Cancellation should remain consistent everywhere.

Opening:

```ts
const document = await open(source, {
  signal,
});
```

XLSX range read:

```ts
await sheet.read("A1:Z500", {
  signal,
});
```

Media:

```ts
await workbook.media.get(mediaId, {
  signal,
});
```

The Worker should stop stale work at safe checkpoints.

Renderers should cancel obsolete range/media work when navigation changes.

---

# Media API

Use a collection-like API:

```ts
const bytes = await document.media.get(mediaId);
```

rather than:

```ts
document.media(mediaId);
```

This creates a discoverable pattern:

```ts
deck.slides.get(...);
workbook.sheets.get(...);
document.media.get(...);
```

A format may expose no media collection if that format's public API genuinely has no use for it, but the shared convention should be reused where applicable.

---

# Cleanup and Lifecycle

Keep the semantic distinction:

```ts
document.close();
```

for parsed/open document resources.

And:

```ts
rendered.destroy();
viewer.destroy();
```

for mounted UI resources.

Do not rename these purely for symmetry.

They represent different lifecycle boundaries.

`close()` should clean up:

```text
Worker
RPC channel
pending reads
parsed model cache
document-owned media state
```

`destroy()` should clean up:

```text
DOM nodes
event listeners
observers
object URLs
scroll listeners
virtualization pools
selection state
temporary render state
```

---

# Viewer Controllers

If `render()` returns an interactive controller rather than a minimal render handle, its shared concepts should remain predictable.

Potential shared methods:

```ts
viewer.setZoom(...);
viewer.search(...);
viewer.destroy();
```

Format-specific methods should remain native.

PPTX:

```ts
viewer.goToSlide(4);
```

or possibly the more generic:

```ts
viewer.goTo(4);
```

XLSX:

```ts
viewer.goTo("D14");
viewer.scrollTo("F120");
viewer.focus("F120");
```

DOCX:

```ts
viewer.goToHeading("Introduction");
```

When DOCX is rendered in physical layout mode, page metadata belongs to the rendered viewer/render handle rather than to the opened document model.

We should not force one navigation signature if doing so makes the call ambiguous.

---

# Events

Shared lifecycle events should use shared names where semantics really match.

Examples:

```text
load
progress
zoomchange
error
```

Format-native navigation events can differ:

```text
slidechange
sheetchange
sectionchange
cellactivate
selectionchange
```

This is preferable to forcing everything through:

```text
viewchange
```

if consumers lose useful meaning.

Again:

> same grammar does not mean identical vocabulary.

---

# Web Component

The built-in web component is a slightly different design problem because it may intentionally support multiple formats.

Current direction:

```html
<aared-viewer
  src="report.xlsx"
  controls>
</aared-viewer>
```

That can remain useful as the **generic convenience UI**.

However, its packaging should be explicit.

If importing:

```ts
import "aared/element";
```

makes the consumer support all three formats, that should be clearly documented.

A future optimization could provide format-specific element entrypoints:

```text
aared/pptx/element
aared/xlsx/element
aared/docx/element
```

if bundle analysis proves the generic element pulls too much code.

Do not add those extra exports preemptively unless measurement shows they are useful.

---

# Generic Element Vocabulary

Because `<aared-viewer>` spans formats, some generic terminology is acceptable at the element level.

But avoid using terms such as:

```text
page
```

as universal cross-format vocabulary. A physical page is meaningful for DOCX only when layout mode is active, and it still does not describe XLSX or replace PPTX's native `slide` terminology.

Potential generic concepts:

```text
active
index
zoom
controls
```

Format-specific state can remain available through properties/events.

For example XLSX may expose:

```ts
viewer.sheet;
viewer.cell;
```

rather than pretending these are pages.

---

# Accessibility

Accessibility must be part of the renderer contract.

## XLSX

Use appropriate grid semantics:

```text
role="grid"
role="row"
role="gridcell"
aria-rowcount
aria-colcount
aria-rowindex
aria-colindex
roving tabindex
keyboard navigation
```

## PPTX / DOCX

Preserve:

- selectable browser-native text,
- semantic table structure where possible,
- accessible links,
- predictable focus behavior,
- headings where source semantics support them.

---

# Search

Search is useful and remains preview-only.

Potential format-native results:

## PPTX

```ts
const matches = await deck.search("revenue");
```

```ts
[
  {
    slide: 3,
    text: "Revenue Growth",
  },
]
```

## XLSX

```ts
const matches = await workbook.search("revenue");
```

```ts
[
  {
    sheet: "Summary",
    cell: "D18",
  },
]
```

## DOCX

```ts
const matches = await document.search("revenue");
```

```ts
[
  {
    section: 1,
    heading: "Revenue",
  },
]
```

Search result shapes do not need to be identical when format-native locations are more useful.

---

# External Link Safety

The viewer must never silently follow or fetch external relationships.

Hyperlinks should:

- be represented in the model,
- render visibly,
- notify the embedding application on activation,
- leave navigation policy to the host application.

Example:

```ts
viewer.addEventListener("linkactivate", event => {
  // host decides what to do
});
```

External workbook relationships should not automatically trigger network fetches.

---

# ZIP / Decompression Safety

OOXML is a ZIP container, so defensive resource limits should be part of the implementation.

Potential safeguards:

```text
maximum entry count
maximum decompressed size per entry
maximum total decompressed size
maximum XML part size
maximum image dimensions
maximum relationship depth
maximum chart/data-point count
```

Limit failures should use typed `AaredError` codes.

---

# SSR-Safe Imports

Package imports should not unexpectedly touch browser globals.

These should be safe to import in SSR code:

```ts
import { open } from "aared/xlsx";
import { open } from "aared/pptx";
import { open } from "aared/docx";
```

The module should not touch:

```text
window
document
HTMLElement
customElements
```

at import time unless the consumer explicitly imports the element registration entrypoint.

---

# AI-Agent-Friendly API Design

Since coding agents are an important consumer of modern APIs, naming should optimize for semantic obviousness.

Agents already understand concepts such as:

```ts
workbook.sheets.get("Revenue");
sheet.read("A1:F20");

deck.slides.get(3);

document.headings.list();

viewer.scrollTo(...);
viewer.focus(...);

media.get(...);

destroy();
close();
```

These names are much easier to infer than implementation terms such as:

```text
unit
part
OPC node
renderer bridge
worker model
document proxy
```

Those implementation terms should remain internal unless they are genuinely useful to advanced consumers.

---

# Why Native Vocabulary Helps AI Agents

An AI coding agent asked:

> "Open the spreadsheet and inspect the Revenue sheet from A1:F20."

can naturally generate:

```ts
const workbook = await open(file);
const sheet = workbook.sheets.get("Revenue");
const data = await sheet.read("A1:F20");
```

An agent asked:

> "Render slide 4."

can naturally generate:

```ts
const deck = await open(file);
const slide = deck.slides.get(3);

render({
  slide,
  target,
});
```

The API becomes predictable because it mirrors the domain itself.

This is better than teaching the agent that:

```text
unit 3
```

sometimes means a slide, sometimes a sheet, and sometimes a Word section.

---

# Format-Specific Recommendations

## PPTX

Recommended public vocabulary:

```text
deck
slides
slide
notes
shapes
media
render
zoom
```

Potential API:

```ts
const deck = await open(file);

const slide = deck.slides.get(3);
const model = await slide.read();

const viewer = await render({
  slide,
  target,
});

viewer.destroy();
deck.close();
```

Do not generalize `slide` into `view`.

---

## XLSX

Recommended public vocabulary:

```text
workbook
sheets
sheet
range
cell
selection
frozen panes
headers
gridlines
media
render
```

Potential API:

```ts
const workbook = await open(file);

const sheet = workbook.sheets.get("Revenue");

const data = await sheet.read("A1:M200");

const viewer = await render({
  sheet,
  target,
});

viewer.goTo("D14");

viewer.destroy();
workbook.close();
```

This format deserves the richest interaction API because spreadsheet navigation is fundamental to preview usability.

---

## DOCX

Recommended public vocabulary:

```text
document
sections
section
headings
paragraphs
tables
media
render
flow
pagination
pages (render output only)
```

Potential API:

```ts
const document = await open(file);
const sections = document.sections.list();

const viewer = render(sections, target);

viewer.destroy();
document.close();
```

Physical layout is opt-in:

```ts
const viewer = render(
  document.sections.list(),
  target,
  {
    pagination: {
      mode: "layout",
    },
  },
);
```

Keep `document.sections` as the source-level collection. Do not create `document.pages` as though physical pages were authored source objects; page metadata belongs to the layout render result.

The DOCX fidelity statement should be explicit:

> **Flow is the default. Physical pagination is supported as an opt-in structural layout mode. Pixel-identical Word pagination is not promised.**

Do not keep a generic abstraction merely for symmetry.

---

# Recommended API Comparison

```text
PPTX
open()
deck.slides.list()
deck.slides.get()
slide.read()
render({ slide, target })
deck.media.get()
deck.close()

XLSX
open()
workbook.sheets.list()
workbook.sheets.get()
sheet.read("A1:F20")
render({ sheet, target })
workbook.media.get()
workbook.close()

DOCX
open()
document.sections.list()
document.sections.get()
render(document.sections.list(), target, options?)
render handle -> pageCount / pages
document.media.get()
document.close()
```

Notice the repeated grammar:

```text
open
collection
list/get
read when applicable
render
media
close
```

while the domain nouns and rendering granularity remain accurate. DOCX does not need a ceremonial `section.read()` step when its resolved sections are already the rendering input.

---

# Things We Should Explicitly Avoid

## Avoid `unit`

It is an implementation abstraction, not a domain abstraction.

## Avoid forcing `view`

`view` is better than `unit`, but still less clear than:

```text
slide
sheet
section
```

for headless APIs.

It may still be useful internally.

## Avoid a universal Office model

Do not make a giant:

```ts
UniversalDocumentElement
```

that tries to model:

```text
shape
cell
paragraph
```

with one interface.

## Avoid editor-style setters

No:

```ts
cell.setValue(...)
shape.setText(...)
paragraph.setStyle(...)
workbook.save(...)
```

## Avoid arbitrary styling callbacks

Do not turn XLSX into a programmable data-grid styling engine.

## Avoid unnecessary convenience APIs

Do not add `preview()` or a top-level all-formats API until a real usage pattern justifies them.

---

# Current `preview()` Recommendation

For the moment:

```text
open()   -> recommended
render() -> recommended
preview() -> defer / validate first
```

The built-in `<aared-viewer>` already provides a convenience path.

Before adding `preview()`, test whether users actually need a second convenience layer.

This is a deliberate pushback against increasing API surface too early.

---

# Recommended Architectural Layers

```text
shared OOXML/core machinery
        │
        ├── PPTX parser/model
        │       ↓
        │   aared/pptx
        │       ↓
        │   PPTX renderer/controller
        │
        ├── XLSX parser/model
        │       ↓
        │   aared/xlsx
        │       ↓
        │   XLSX renderer/controller
        │
        └── DOCX parser/model
                ↓
            aared/docx
                ↓
            DOCX renderer/controller

optional:
aared/element
    generic multi-format viewer

possible later:
aared
    optional auto-detect convenience entrypoint
```

---

# XLSX Detailed Checklist

Before freezing the XLSX public API, explicitly cover:

- sheet listing,
- lookup by index,
- lookup by name,
- used range / extent,
- A1 range addressing,
- structured range addressing,
- range-based Worker RPC,
- virtualized rows,
- virtualized columns,
- overscan,
- stale request cancellation,
- merged cells,
- frozen rows,
- frozen columns,
- hidden rows,
- hidden columns,
- row heights,
- column widths,
- sheet-level RTL,
- cell reading order,
- authored fill,
- authored font,
- authored borders,
- authored alignment,
- number formats,
- cached formula values,
- formula metadata,
- conditional-formatting results,
- hyperlinks,
- images,
- charts,
- sheet tabs,
- row headers,
- column headers,
- source gridlines,
- viewer gridline overrides,
- readonly cell selection,
- keyboard navigation,
- A1 `goTo`,
- programmatic `scrollTo`,
- `focus`,
- copy behavior,
- scroll position behavior,
- zoom behavior,
- accessible grid semantics,
- viewer appearance tokens,
- cleanup behavior.

---

# PPTX Detailed Checklist

Before freezing the PPTX public API, explicitly cover:

- slide listing,
- slide lookup,
- slide title metadata,
- slide size,
- notes,
- shapes,
- geometry,
- fills,
- gradients,
- strokes,
- transforms,
- text bodies,
- paragraph RTL,
- tables,
- table RTL,
- pictures,
- charts,
- z-order,
- theme colors,
- theme fonts,
- embedded fonts when supported,
- searchable/selectable text,
- fit-page,
- fit-width,
- explicit zoom,
- unsupported geometry fallback,
- unsupported fill fallback,
- media lifecycle,
- cleanup behavior.

---

# DOCX Detailed Checklist

Before freezing the DOCX public API, explicitly cover:

- flow model,
- `DocumentSection` as the source-level section model,
- section listing and lookup,
- headings,
- paragraphs,
- text runs,
- lists,
- tables,
- images,
- footnotes,
- headers/footers according to the fidelity contract,
- paragraph RTL,
- search,
- heading navigation,
- readonly text selection,
- basic link behavior,
- flow virtualization where needed,
- flow as the default render mode,
- opt-in physical layout mode,
- section-driven page geometry and orientation,
- mandatory vs soft page-break policy,
- derived `pageCount` / `pages` metadata on the render result,
- `PAGE` / `NUMPAGES` behavior in layout mode,
- an explicit fidelity contract: structural physical pagination is supported, Word-identical pagination is not promised,
- no source-level `document.pages` abstraction,
- cleanup behavior.

---

# Shared Detailed Checklist

Across all entrypoints:

- same `open()` grammar,
- same source types,
- same `AbortSignal` behavior,
- same error model,
- same progress semantics,
- same `close()` semantics,
- same media collection conventions,
- shared render option concepts where applicable, while allowing format-native signatures and options,
- same `destroy()` semantics,
- SSR-safe imports,
- external-link safety,
- ZIP/decompression limits,
- accessibility,
- CSS design tokens,
- shadow DOM styling contract,
- capabilities,
- documentation parity,
- type-level parity between documentation and implementation,
- format-specific entrypoints remain independently consumable.

---

# Migration From the Current API

Current generic concepts such as:

```ts
doc.unit(i)
```

should move toward native collections.

PPTX:

```ts
deck.slides.get(i)
```

XLSX:

```ts
workbook.sheets.get(i)
```

DOCX:

```ts
const sections = document.sections.list();
const section = document.sections.get(i);
```

The ordered section list is the normal DOCX render input; individual lookup remains useful for inspection and navigation.

Current generic:

```text
page
```

should be removed from any context where it means:

```text
slide OR sheet OR section
```

The generic web component may still need an index-oriented navigation concept, but it should not claim all Office content is a page.

DOCX physical layout produces real derived pages, but those pages belong to the rendered layout rather than replacing `document.sections` as the source model.

---

# Documentation Strategy

The docs should teach shared grammar first.

For example:

## Common pattern

```text
1. import a format entrypoint
2. open the document
3. access the format-native collection
4. read lazily addressed content when needed
5. render the format-native content
6. destroy UI
7. close document
```

Then show one example per format.

The DOCX example should make the important difference visible:

```text
PPTX -> render a slide
XLSX -> render a sheet
DOCX -> render the ordered document sections, with flow default and layout opt-in
```

This is better than documenting each format as if it were a completely separate library, and better than pretending all three have the same renderable unit.

---

# README Example Set

## PPTX

```ts
import {
  open,
  render,
} from "aared/pptx";

const deck = await open(file);
const slide = deck.slides.get(0);

const viewer = await render({
  slide,
  target: container,
});

viewer.destroy();
deck.close();
```

## XLSX

```ts
import {
  open,
  render,
} from "aared/xlsx";

const workbook = await open(file);
const sheet = workbook.sheets.get("Revenue");

const viewer = await render({
  sheet,
  target: container,
});

viewer.goTo("D14");

viewer.destroy();
workbook.close();
```

## DOCX

```ts
import {
  open,
  render,
} from "aared/docx";

const document = await open(file);
const sections = document.sections.list();

const viewer = render(sections, container);

viewer.destroy();
document.close();
```

For a physical-page preview:

```ts
const viewer = render(
  document.sections.list(),
  container,
  {
    pagination: {
      mode: "layout",
    },
  },
);
```

---

# Final Recommendation

The package should follow five rules.

## 1. Keep format-specific entrypoints

```text
aared/pptx
aared/xlsx
aared/docx
```

This preserves explicit bundle control.

## 2. Unify API grammar, not nouns

Use:

```text
open
list
get
render
media
close
destroy
```

consistently.

Use:

```text
read
```

consistently where lazy content access is actually part of the format model; do not require it ceremonially for DOCX sections.

But keep:

```text
slides
sheets
sections
cells
ranges
headings
```

where they are naturally correct.

## 3. Keep models format-specific

Do not create a universal Office document model.

## 4. Treat XLSX as a readonly virtualized spreadsheet

Not a static HTML table.

Not an editor.

## 5. Keep authored styles separate from viewer styles

Office styles are fidelity data.

Viewer styling belongs to CSS variables, shadow parts, and a small appearance API.

---

# Final Design Principle

> **aared should feel like one library without pretending Office formats are one data model.**

More concretely:

> **Separate format entrypoints, shared API grammar, native terminology, readonly interaction, strong defaults, and format-specific models underneath.**

That is the ground I would recommend freezing toward.

# Content Color Modes / Dark Mode

Dark mode should be implemented by `aared` itself, but **content should not behave identically across PPTX, XLSX, and DOCX**.

That difference is intentional.

The formats have different visual semantics:

```text
PPTX -> authored visual canvas
XLSX -> authored data grid with many automatic/default visual values
DOCX -> authored flow document with optional section-driven physical layout, commonly consumed as reading content
```

Trying to apply one dark-mode rule to all three would either:

- damage fidelity,
- make dark mode ineffective,
- or force consumers to rebuild format-specific behavior outside `aared`.

The design should therefore separate:

```text
viewer theme
content color mode
```

---

## Why These Must Be Separate

A host application may want any of the following:

```text
dark app + original light document
dark app + adaptive dark document
light app + original document
```

The application theme therefore cannot implicitly dictate document rendering.

For example, a user may run the host application in dark mode but still want to inspect an XLSX exactly as it appears in Excel.

Likewise, a DOCX reader may intentionally offer dark reading mode even when the original document was authored on a white page.

So:

```ts
appearance: {
  theme: "dark",
  contentMode: "source",
}
```

and:

```ts
appearance: {
  theme: "dark",
  contentMode: "adaptive",
}
```

must mean different things.

Recommended common values:

```ts
type ViewerTheme =
  | "light"
  | "dark"
  | "system";

type ContentMode =
  | "source"
  | "adaptive";
```

The exact valid `ContentMode` should be format-specific where necessary.

---

# Viewer Theme

`theme` controls only UI owned by `aared`.

Examples:

```text
viewer background
toolbar
controls
sheet tabs
row/column headers
selection
focus state
loading state
error state
status UI
scrollbars where styled
```

Example:

```ts
appearance: {
  theme: "system",
}
```

`"system"` should follow the host/browser color preference where practical.

It should not by itself recolor authored Office content.

---

# Content Mode

`contentMode` controls whether Office content remains source-faithful or may adapt for dark reading.

The important rule is:

> **Content adaptation is format-specific because fidelity means different things for each Office format.**

---

# PPTX Behavior

## Recommendation

PPTX should always preserve the slide as authored.

Conceptually:

```ts
type PptxContentMode = "source";
```

Example:

```ts
render({
  slide,
  target,

  appearance: {
    theme: "dark",
  },
});
```

The surrounding stage and controls become dark.

The slide itself does not change.

---

## Reasoning

A PowerPoint slide is fundamentally a **designed canvas**.

Its visual composition depends on intentional relationships between:

```text
backgrounds
fills
text colors
images
charts
gradients
shapes
contrast
brand colors
z-order
```

Changing "default white" to dark or "default black" to light is not reliably safe because those values may be part of an intentional composition.

A slide can also contain:

- screenshots,
- photography,
- logos,
- diagrams,
- branded templates,
- contrast-sensitive overlays.

There is no robust generic rule for converting these into dark mode without redesigning the presentation.

That would violate `aared`'s fidelity contract.

Therefore:

> **PPTX dark mode applies to the viewer, never to the authored slide.**

This is not a limitation to work around later. It is the correct default product behavior for a fidelity-oriented presentation viewer.

If adaptive presentation recoloring is ever added in the future, it should be treated as a separate transformation feature, not normal dark mode.

---

# XLSX Behavior

## Recommendation

XLSX should support both:

```text
source
adaptive
```

Example source mode:

```ts
appearance: {
  theme: "dark",
  contentMode: "source",
}
```

Example adaptive mode:

```ts
appearance: {
  theme: "dark",
  contentMode: "adaptive",
}
```

---

## XLSX `source`

`source` preserves workbook appearance as authored.

This means a workbook can remain visually light inside a dark `aared` viewer.

That behavior is necessary for:

```text
visual fidelity review
financial/reporting documents
screenshots
design comparisons
documents where color carries semantic meaning
```

The host application's dark mode should never silently destroy the user's ability to inspect the original spreadsheet appearance.

---

## XLSX `adaptive`

`adaptive` provides a dark spreadsheet reading/viewing experience without blindly recoloring authored styles.

In adaptive mode, `aared` may adapt values that are genuinely default or automatic:

```text
automatic/default sheet background -> dark
automatic/default text             -> light
default gridlines                  -> dark-mode gridlines
row/column headers                 -> dark
blank cells                        -> dark
selection/focus UI                 -> dark viewer colors
```

Explicit authored values remain source-faithful:

```text
explicit cell fill           -> preserved
explicit font color          -> preserved
table/header fill            -> preserved
conditional formatting       -> preserved
images                       -> preserved
charts                       -> preserved
explicit borders             -> preserved unless proven automatic
```

Do not use global inversion.

Do not apply a generic CSS filter to the rendered sheet.

---

## Reasoning

XLSX is neither a fixed visual canvas like PPTX nor a semantic document-flow format like DOCX.

It is primarily a **data grid**.

Many spreadsheet cells rely on implicit defaults:

```text
white background
automatic black text
default gridlines
default empty cells
```

Those defaults can safely participate in dark mode because they are often not intentional document design decisions.

At the same time, spreadsheets frequently use explicit colors to communicate meaning:

```text
red = loss/error
green = success
yellow = warning
blue = category/header
conditional formatting = data semantics
```

Those must not be silently transformed.

Therefore XLSX needs a semantic distinction between:

```text
automatic/default appearance
```

and:

```text
authored appearance
```

This is why XLSX can support adaptive dark mode without abandoning fidelity.

---

# Important XLSX Model Requirement

Adaptive mode requires the model to preserve **style provenance**.

The renderer must be able to distinguish:

```text
explicitly authored black
```

from:

```text
automatic/default black
```

even though both may appear as the same CSS color in light mode.

For example:

```text
explicit black text   -> stays black in adaptive mode
automatic black text  -> may become light
```

The same reasoning applies to values such as:

```text
fills
font colors
borders where relevant
default sheet/background colors
```

Therefore the Worker must not always resolve automatic/default styling into indistinguishable final CSS values.

The resolved model needs enough information for the renderer to know whether adaptation is permitted.

Possible implementation directions include:

```ts
color: {
  value: "#000000",
  source: "automatic",
}
```

or:

```ts
color?: string;
automaticColor?: boolean;
```

The exact model shape can be decided separately.

The important contract is:

> **Resolution must not erase information that dark-mode rendering needs.**

---

# DOCX Behavior

## Recommendation

DOCX should support both:

```text
source
adaptive
```

Example:

```ts
appearance: {
  theme: "dark",
  contentMode: "adaptive",
}
```

---

## DOCX `source`

`source` renders authored document colors faithfully.

A document that was authored as black text on white remains black text on white, even inside a dark viewer.

This mode is useful for:

```text
visual inspection
layout review
brand/style review
source fidelity
screenshots
```

---

## DOCX `adaptive`

`adaptive` acts as a reading mode.

Typical adaptations:

```text
default page/background -> dark
automatic/default text -> light
default separators/borders -> dark-mode equivalents where safe
viewer-generated surfaces -> dark
```

Explicit authored styling remains preserved where practical:

```text
explicit text colors -> preserved
explicit highlights -> preserved
explicit paragraph shading -> preserved
images -> preserved
tables with explicit fills -> preserved
```

The implementation may eventually need contrast safeguards for explicit authored colors that become unreadable against an adaptive background, but those rules should be conservative and separately specified.

---

## Reasoning

DOCX differs fundamentally from PPTX.

A Word document is primarily a **flow-reading format**.

Users commonly consume Word-like documents for:

```text
reading
reviewing
searching
selecting
copying
long-form navigation
```

Dark reading modes are therefore useful and expected.

`aared` keeps DOCX content browser-rendered and semantic in both continuous flow and physical layout modes. Layout mode changes page arrangement; it does not turn the document into a fixed authored canvas like PPTX.

That makes adaptive dark mode appropriate in either rendering mode.

At the same time, DOCX can contain authored colors, tables, shading, images, and branded material.

Therefore dark mode should adapt defaults, not blindly invert the document.

The intended principle is:

> **DOCX adaptive mode improves reading comfort while preserving authored visual intent whenever that intent is explicit.**

---

# Why PPTX, XLSX, and DOCX Intentionally Differ

The behavior can be summarized as:

| Format | Source mode | Adaptive mode | Reason |
|---|---|---|---|
| PPTX | Yes | No by default | A slide is a designed visual canvas; recoloring breaks composition and fidelity |
| XLSX | Yes | Yes | A grid contains many safe automatic defaults, while explicit styles often carry data meaning |
| DOCX | Yes | Yes | DOCX remains semantic reading content in both flow and physical layout modes, so adaptive reading mode is useful |

This difference should be documented explicitly so future maintainers do not "simplify" the API by forcing identical behavior.

The consistent principle is not:

> every format supports the same dark-mode transformation.

The consistent principle is:

> **`aared` owns viewer theming, preserves source fidelity, and only adapts document content where the format semantics make adaptation defensible.**

---

# Recommended Defaults

Recommended default content behavior:

```text
PPTX -> source
XLSX -> source
DOCX -> source
```

This keeps fidelity as the default contract.

Applications can opt into:

```text
XLSX -> adaptive
DOCX -> adaptive
```

when they want an integrated dark reading experience.

This is preferable to automatically linking `contentMode` to `theme`, because a dark host application does not necessarily mean the user wants the Office document recolored.

---

# `system` Theme and Content Mode

A host can implement:

```ts
appearance: {
  theme: "system",
  contentMode: "source",
}
```

to let viewer chrome follow the OS/app theme while always keeping original document colors.

Or:

```ts
appearance: {
  theme: "system",
  contentMode: "adaptive",
}
```

for XLSX/DOCX if the desired behavior is for both chrome and readable content to adapt.

`aared` should own the actual application of those modes.

Consumers should not need to manually inject dark spreadsheet or document CSS.

---

# CSS Token Behavior

Viewer tokens should change with `theme`.

For example:

```css
--aared-background
--aared-surface
--aared-border
--aared-text
--aared-muted-text
--aared-accent
--aared-selection
--aared-focus-ring
```

XLSX adaptive content may additionally use semantic tokens such as:

```css
--aared-sheet-background
--aared-sheet-text
--aared-gridline
--aared-sheet-header-bg
--aared-sheet-header-text
```

DOCX adaptive content may use:

```css
--aared-document-background
--aared-document-text
--aared-document-muted-text
--aared-document-border
```

These tokens are fallback/adaptive rendering values.

They must not overwrite explicit authored styles.

---

# Dark-Mode Design Rule

The final rule should be easy to remember:

> **Theme changes `aared`. Content mode may change document defaults. Authored visual intent remains protected.**

Per format:

```text
PPTX:
dark viewer + unchanged slide

XLSX:
dark viewer + source spreadsheet
or
dark viewer + adaptive defaults with authored styles preserved

DOCX:
dark viewer + source document
or
dark viewer + adaptive reading mode with authored styles preserved
```

This distinction should remain part of the public design rationale, not merely an implementation detail.

# Deferred: Browser Print / Save as PDF

> **Status: Deferred to the end of the roadmap.**
>
> This should not block the core viewer, fidelity, XLSX interaction, styling, dark-mode, or API-unification work.

`aared` may eventually support a browser-driven **Print / Save as PDF** workflow.

This should **not** become a PDF-generation engine inside `aared`.

The intended model is:

```text
aared renderer
    ↓
print-optimized DOM/SVG
    ↓
browser print flow
    ↓
window.print()
    ↓
browser / OS print dialog
    ↓
Print or Save as PDF
```

The browser remains responsible for producing the PDF or sending the document to a printer.

---

## Why This Should Not Be Called `exportPdf()`

Avoid an API such as:

```ts
await viewer.exportPdf();
```

because that strongly implies that `aared` will:

- generate PDF bytes itself,
- return a `Blob` or `ArrayBuffer`,
- work headlessly,
- provide deterministic PDF output,
- own a second PDF-specific pagination system,
- own the PDF rendering engine.

That is not the intended contract.

Prefer a browser-native API such as:

```ts
viewer.print();
```

or, if a preparation step becomes necessary:

```ts
viewer.openPrintView();
```

The simpler API should be preferred unless implementation experience proves otherwise.

---

## Public Product Language

Describe the feature as:

> **Print / Save as PDF** — prepares the current document for browser printing. PDF creation is handled by the browser's native print dialog.

This is accurate and sets the correct expectations.

---

# Why This Fits `aared`

`aared` already renders Office content using browser-native DOM/SVG.

A browser-print workflow can reuse that same rendering pipeline rather than introducing another document renderer.

This has several advantages:

- text shaping remains browser-native,
- Arabic/RTL behavior stays aligned with the normal viewer,
- fonts use the same resolution path,
- tables and shapes use the same rendered model,
- there is no second PDF-specific fidelity implementation,
- `aared` does not expand into a full PDF-generation library.

This also preserves the existing architectural boundary:

> `aared` previews Office documents; it does not become a general-purpose PDF engine.

---

# Format-Specific Print Behavior

The printing contract should differ by format because the formats have different layout semantics.

---

## PPTX

PPTX is the cleanest case.

Recommended behavior:

```text
1 slide = 1 printed page
```

Each slide should preserve:

- authored aspect ratio,
- authored colors,
- authored positioning,
- authored typography,
- source fidelity.

The print view should remove viewer chrome and print only the slide content.

Potential future API:

```ts
viewer.print();
```

or later:

```ts
viewer.print({
  slides: "current",
});
```

Possible future options may include:

```text
current slide
all slides
selected slide range
```

but these should not be added before there is a demonstrated need.

---

## XLSX

XLSX requires more deliberate print policy.

The renderer must **not** print only the currently virtualized viewport.

Before invoking the browser print flow, `aared` must construct the appropriate printable sheet representation.

Possible future behavior:

```ts
viewer.print({
  sheets: "current",
  fit: "width",
});
```

Potential policy options:

```text
current sheet
all visible sheets

actual size
fit width
fit page
```

However, the first version should remain intentionally small.

A reasonable initial contract could simply be:

```text
current sheet
fit to available print width
```

The print implementation must account for:

- virtualized rows,
- virtualized columns,
- merged cells,
- frozen panes,
- hidden rows/columns,
- row/column headers where enabled,
- RTL,
- charts,
- images,
- authored cell styles.

Print-specific framing used only by benchmark/testing tools should not automatically become part of the public API.

---

## DOCX

DOCX printing is useful, and physical layout mode provides a stronger foundation for page-oriented print preparation.

However, the fidelity limit remains explicit:

> **`aared` does not promise Word-identical pagination.**

When page-oriented output is desired, print preparation should reuse the DOCX physical layout rather than introduce a second DOCX paginator. A flow render can still be printed as browser flow when that is the selected viewing model.

In either case, it must not claim:

- identical page breaks to Microsoft Word,
- identical line breaking or glyph metrics,
- identical pagination-dependent layout in every environment,
- a Word-compatible pagination engine.

The correct contract is:

> **DOCX printing prints the browser-rendered `aared` document/layout. Physical layout can preserve `aared`'s structural pages, but Microsoft Word remains a different layout engine.**

This distinction must remain visible in public documentation.

---

# Print Styling

Printing should use a dedicated print stylesheet rather than ordinary viewer chrome.

Example responsibilities:

```text
hide controls
hide toolbar
hide sheet tabs unless intentionally included
remove interactive selection states
remove hover/focus effects
normalize print backgrounds
preserve authored Office colors
use print-appropriate margins
ensure content is not clipped by viewport containers
disable viewport-only overflow constraints
```

The implementation should rely on:

```css
@media print
```

where possible.

If temporary print DOM is required, it should still use the same resolved document models and renderer logic.

---

# Dark Mode and Printing

Printing should default to **source document appearance**, not adaptive dark content.

For example, if the viewer is currently:

```ts
appearance: {
  theme: "dark",
  contentMode: "adaptive",
}
```

the print output should generally use:

```text
contentMode: source
```

unless a future explicit option allows otherwise.

Reasoning:

- PDF/print output is usually expected to reflect the source document,
- printing dark-mode reading adaptations can consume unnecessary ink,
- adaptive dark content is a viewing convenience rather than authored document state,
- PPTX already remains source-faithful regardless of viewer theme.

A future explicit option could allow:

```ts
viewer.print({
  contentMode: "current",
});
```

but this should not be part of the initial version.

---

# Why This Is Deferred

This feature should come **after** the core viewer work because it depends on several foundations being correct first:

```text
stable public API
PPTX fidelity
XLSX range loading and virtualization
XLSX full-sheet render path
DOCX flow and layout fidelity
media lifecycle
font resolution
RTL correctness
styling architecture
dark-mode/content-mode separation
cleanup behavior
```

Implementing print earlier risks creating a second set of workarounds around incomplete rendering behavior.

The right sequence is:

```text
1. make normal viewing correct
2. make the public API stable
3. make styling/theme behavior stable
4. make large-document rendering reliable
5. then reuse that foundation for printing
```

This is why browser print / Save as PDF should be treated as an **end-of-roadmap capability**, not part of the initial API freeze.

---

# Deferred Recommendation

For now, document only the intended direction:

```ts
// future / deferred
viewer.print();
```

with this contract:

> `aared` prepares the currently rendered Office document for browser printing. The browser or operating system owns printing and Save as PDF.

Do not add the API until the underlying format renderers are stable enough that print can reuse them cleanly rather than introducing format-specific hacks.
