/**
 * What a developer writes, per playground tab and per file format: the
 * copy-paste snippets the playground shows. The demo files under `src/demos`
 * run the same calls with page plumbing (status line, teardown) around them.
 * Keep the two in step: an option a demo passes (like the low-level
 * `sheetTabs`; the embed layers show the tabs by default) must appear here too.
 */
export type SnippetFormat = "pptx" | "xlsx" | "docx";

export interface PlaygroundSnippet {
  title: string;
  lang: "html" | "tsx" | "ts";
  code: string;
}

export const SNIPPET_FORMATS: SnippetFormat[] = ["pptx", "xlsx", "docx"];

const FIXTURE_FILE: Record<SnippetFormat, string> = {
  pptx: "post-production-review.pptx",
  xlsx: "post-production-tracker.xlsx",
  docx: "docx-coverage.docx",
};

const SHEET_TABS = "appearance: { spreadsheet: { sheetTabs: true } }";

/** Slides size themselves; a sheet or a document needs a height to scroll inside (the API page's Sizing). */
const NEEDS_HEIGHT = "// el needs a height: a sheet or document scrolls inside it\n";

export const htmlSnippets: Record<SnippetFormat, PlaygroundSnippet> = Object.fromEntries(
  SNIPPET_FORMATS.map((format) => [
    format,
    {
      title: "index.html",
      lang: "html",
      // A plain module script cannot resolve a bare "@aared/viewer/element": the
      // import lives in the app code a bundler builds.
      code: `<!-- once, in your bundled app code: import "@aared/viewer/element"; -->
<aared-viewer
  src="./${FIXTURE_FILE[format]}"
  style="height: 480px"
  controls>
</aared-viewer>`,
    },
  ]),
) as Record<SnippetFormat, PlaygroundSnippet>;

const reactSnippet = (): PlaygroundSnippet => {
  // A bounded host works for every format and keeps the first example uniform.
  const style = ' style={{ height: 480 }}';
  return {
    title: "Viewer.tsx",
    lang: "tsx",
    code: `import { usePreview } from "@aared/viewer/react";
import { messageFor } from "@aared/viewer/core";

function Viewer({ file }: { file: string | File }) {
  // No built-in bar. The hook owns cleanup; add app-owned controls if needed.
  const { ref, status, error, retry } = usePreview({ source: file });
  return <>
    <div ref={ref}${style} aria-busy={status === "loading"} />
    {status === "loading" && <p role="status">Loading...</p>}
    {status === "error" && <p role="alert">{messageFor(error)} <button onClick={retry}>Retry</button></p>}
  </>;
}`,
  };
};

export const reactSnippets: Record<SnippetFormat, PlaygroundSnippet> = {
  pptx: reactSnippet(),
  xlsx: reactSnippet(),
  docx: reactSnippet(),
};

const facadeSnippet = (format: SnippetFormat): PlaygroundSnippet => {
  return {
    title: "app.ts",
    lang: "ts",
    code: `import { preview } from "@aared/viewer/preview";

${format === "pptx" ? "" : NEEDS_HEIGHT}const viewer = await preview({ source: file, target: el });

viewer.format;      // "pptx" | "xlsx" | "docx"
viewer.info.count;  // slides, sheets, or 1 for a document

// Call when the feature unmounts.
function cleanup() { viewer.destroy(); }`,
  };
};

export const facadeSnippets: Record<SnippetFormat, PlaygroundSnippet> = {
  pptx: facadeSnippet("pptx"),
  xlsx: facadeSnippet("xlsx"),
  docx: facadeSnippet("docx"),
};

export const headlessSnippets: Record<SnippetFormat, PlaygroundSnippet> = {
  pptx: {
    title: "app.ts",
    lang: "ts",
    code: `import { open, render } from "@aared/viewer/pptx";

const deck = await open(file);
const slideView = await render({ slide: deck.slides.get(0), target: el });

// Call when the feature unmounts: renders first, then the document.
function cleanup() { slideView.destroy(); deck.close(); }`,
  },
  xlsx: {
    title: "app.ts",
    lang: "ts",
    code: `import { open, render } from "@aared/viewer/xlsx";

const workbook = await open(file);
${NEEDS_HEIGHT}const sheetView = await render({
  sheet: workbook.sheets.get(0),
  target: el,
  ${SHEET_TABS},
});

// Call when the feature unmounts: renders first, then the document.
function cleanup() { sheetView.destroy(); workbook.close(); }`,
  },
  docx: {
    title: "app.ts",
    lang: "ts",
    code: `import { open, render } from "@aared/viewer/docx";

const doc = await open(file); // not \`document\`: that would shadow the page's
${NEEDS_HEIGHT}const docView = await render({ document: doc, target: el });

// Call when the feature unmounts: renders first, then the document.
function cleanup() { docView.destroy(); doc.close(); }`,
  },
};
