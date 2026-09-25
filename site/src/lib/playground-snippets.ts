/**
 * What a developer writes, per playground tab and per file format: the
 * copy-paste snippets the playground shows. The demo files under `src/demos`
 * run the same calls with page plumbing (status line, teardown) around them.
 * Keep the two in step: an option a demo passes (like `sheetTabs`) must
 * appear here too.
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

export const htmlSnippets: Record<SnippetFormat, PlaygroundSnippet> = Object.fromEntries(
  SNIPPET_FORMATS.map((format) => [
    format,
    {
      title: "index.html",
      lang: "html",
      code: `<script type="module">
  import "aared/element";
</script>

<aared-viewer
  src="./${FIXTURE_FILE[format]}"
  controls>
</aared-viewer>`,
    },
  ]),
) as Record<SnippetFormat, PlaygroundSnippet>;

const reactSnippet = (format: SnippetFormat): PlaygroundSnippet => {
  const tabs = format === "xlsx";
  // Slides size themselves. A sheet scrolls inside the height it is given;
  // a document is as tall as its pages, so a capped container must scroll.
  const style = format === "xlsx" ? " style={{ height: 480 }}" : format === "docx" ? ' style={{ height: 480, overflow: "auto" }}' : "";
  return {
    title: "Viewer.tsx",
    lang: "tsx",
    code: `import { usePreview } from "aared/react";
${tabs ? `\n// Defined outside the component so it stays referentially stable.\nconst appearance = { spreadsheet: { sheetTabs: true } };\n` : ""}
function Viewer({ file }: { file: string | File }) {
  // Also returned: info, index, goTo(i), setZoom(zoom) — enough to build
  // your own navigation. No built-in bar; use <aared-viewer controls> for that.
  const { ref, status } = usePreview({ source: file${tabs ? ", appearance" : ""} });
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
  const tabs = format === "xlsx";
  return {
    title: "app.ts",
    lang: "ts",
    code: `import { preview } from "aared/preview";

const viewer = await preview(file, el${tabs ? `, {\n  ${SHEET_TABS},\n}` : ""});

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
const viewer = await render({ slide: deck.slides.get(0), target: el });

// when you're done
viewer.destroy();
deck.close();`,
  },
  xlsx: {
    title: "app.ts",
    lang: "ts",
    code: `import { open, render } from "aared/xlsx";

const workbook = await open(file);
const viewer = await render({
  sheet: workbook.sheets.get(0),
  target: el,
  ${SHEET_TABS},
});

// when you're done
viewer.destroy();
workbook.close();`,
  },
  docx: {
    title: "app.ts",
    lang: "ts",
    code: `import { open, render } from "aared/docx";

const document = await open(file);
const viewer = await render({ document, target: el });

// when you're done
viewer.destroy();
document.close();`,
  },
};
