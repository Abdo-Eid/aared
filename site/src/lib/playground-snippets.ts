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
      // A plain module script cannot resolve a bare "aared/element": the
      // import lives in the app code a bundler builds.
      code: `<!-- once, in your bundled app code: import "aared/element"; -->
<aared-viewer
  src="./${FIXTURE_FILE[format]}"${format === "pptx" ? "" : `\n  style="height: 480px"`}
  controls>
</aared-viewer>`,
    },
  ]),
) as Record<SnippetFormat, PlaygroundSnippet>;

const reactSnippet = (format: SnippetFormat): PlaygroundSnippet => {
  // Slides size themselves. A sheet scrolls inside the height it is given;
  // a document is as tall as its pages, so a capped container must scroll.
  const style = format === "xlsx" ? " style={{ height: 480 }}" : format === "docx" ? ' style={{ height: 480, overflow: "auto" }}' : "";
  return {
    title: "Viewer.tsx",
    lang: "tsx",
    code: `import { usePreview } from "aared/react";

function Viewer({ file }: { file: string | File }) {
  // Also returned, once status is "ready": viewer (viewer.info,
  // viewer.next(), viewer.zoomIn(), viewer.resetZoom(), ...), index and
  // zoom: enough for your own toolbar. No built-in bar; use <aared-viewer controls>.
  const { ref, status } = usePreview({ source: file });
  return <div ref={ref}${style} aria-busy={status === "loading"} />;
}`,
  };
};

export const reactSnippets: Record<SnippetFormat, PlaygroundSnippet> = {
  pptx: reactSnippet("pptx"),
  xlsx: reactSnippet("xlsx"),
  docx: reactSnippet("docx"),
};

const facadeSnippet = (format: SnippetFormat): PlaygroundSnippet => {
  return {
    title: "app.ts",
    lang: "ts",
    code: `import { preview } from "aared/preview";

${format === "pptx" ? "" : NEEDS_HEIGHT}const viewer = await preview({ source: file, target: el });

viewer.format;      // "pptx" | "xlsx" | "docx"
viewer.info.count;  // slides, sheets, or 1 for a document

// when you're done
viewer.destroy();`,
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
    code: `import { open, render } from "aared/pptx";

const deck = await open(file);
const slideView = await render({ slide: deck.slides.get(0), target: el });

// when you're done
slideView.destroy();
deck.close();`,
  },
  xlsx: {
    title: "app.ts",
    lang: "ts",
    code: `import { open, render } from "aared/xlsx";

const workbook = await open(file);
${NEEDS_HEIGHT}const sheetView = await render({
  sheet: workbook.sheets.get(0),
  target: el,
  ${SHEET_TABS},
});

// when you're done
sheetView.destroy();
workbook.close();`,
  },
  docx: {
    title: "app.ts",
    lang: "ts",
    code: `import { open, render } from "aared/docx";

const doc = await open(file); // not \`document\`: that would shadow the page's
${NEEDS_HEIGHT}const docView = await render({ document: doc, target: el });

// when you're done
docView.destroy();
doc.close();`,
  },
};
