import { expect, test } from "bun:test";

const API_PATH = `${import.meta.dir}/../../../site/src/content/docs/api.mdx`;
const MODULES = {
  "aared/core": `${import.meta.dir}/../src/core/index.ts`,
  "aared/pptx": `${import.meta.dir}/../src/pptx/index.ts`,
  "aared/xlsx": `${import.meta.dir}/../src/xlsx/index.ts`,
  "aared/docx": `${import.meta.dir}/../src/docx/index.ts`,
  "aared/preview": `${import.meta.dir}/../src/preview/index.ts`,
  "aared/react": `${import.meta.dir}/../src/react/index.ts`,
  "aared/element": `${import.meta.dir}/../src/element/index.ts`,
} as const;

function exportedNames(source: string): string[] {
  const names = new Set<string>();
  for (const match of source.matchAll(/^export\s+(?:async\s+)?(?:type|interface|class|function|const|enum)\s+(\w+)/gm)) {
    names.add(match[1]!);
  }
  for (const match of source.matchAll(/export\s+type\s*{([\s\S]*?)}\s+from/g)) {
    for (const entry of match[1]!.split(",")) {
      const name = entry.trim().match(/^(\w+)/)?.[1];
      if (name !== undefined) names.add(name);
    }
  }
  // Value re-exports (`export { AaredError } from ...`) are public API too;
  // `export type { ... }` blocks are covered above and skipped here.
  for (const match of source.matchAll(/export\s*(type\s*)?{([\s\S]*?)}\s+from/g)) {
    if (match[1] !== undefined) continue;
    for (const entry of match[2]!.split(",")) {
      const name = entry.trim().match(/^(\w+)/)?.[1];
      if (name !== undefined) names.add(name);
    }
  }
  return [...names].sort();
}

test("API reference names every public package export", async () => {
  const api = await Bun.file(API_PATH).text();
  for (const [module, path] of Object.entries(MODULES)) {
    const exported = exportedNames(await Bun.file(path).text());
    for (const name of exported) {
      expect(api, `${module} export ${name}`).toMatch(new RegExp(`\\b${name}\\b`));
    }
  }
});

test("API reference callable examples resolve to public implementations", async () => {
  const [api, core, pptx, xlsx, docx, preview, react, element] = await Promise.all([
    Bun.file(API_PATH).text(),
    Bun.file(MODULES["aared/core"]).text(),
    Bun.file(MODULES["aared/pptx"]).text(),
    Bun.file(MODULES["aared/xlsx"]).text(),
    Bun.file(MODULES["aared/docx"]).text(),
    Bun.file(MODULES["aared/preview"]).text(),
    Bun.file(MODULES["aared/react"]).text(),
    Bun.file(MODULES["aared/element"]).text(),
  ]);
  for (const callable of [
    "open",
    "preview",
    "detect",
    "slides.get",
    "sheets.get",
    "sections.list",
    "headings.list",
    "slide.read",
    "sheet.read",
    "media.get",
    "deck.search",
    "sheet.search",
    "doc.search",
    // The borrowed-content grammar (§122) and the viewer-level search steps.
    "content.slides.list",
    "content.slides.get",
    "content.sheets.get",
    "content.sections.list",
    "content.media.get",
    "searchNavigation.next",
    "viewer.scale",
    "render({ slide",
    "render({ sheet",
    "render({ document",
    "deck.close",
    "render",
    "viewer.setZoom",
    "viewer.setAppearance",
    "viewer.goTo",
    "viewer.next",
    "viewer.previous",
    "viewer.zoomIn",
    "viewer.zoomOut",
    "viewer.resetZoom",
    "viewer.defaultZoom",
    "viewer.search",
    "viewer.goToSearchMatch",
    "viewer.content",
    "viewer.xlsx.goToCell",
    "viewer.xlsx.goToSheet",
    "viewer.docx.goToHeading",
    "viewer.docx.goToPage",
    "viewer.xlsx.scrollTo",
    "viewer.xlsx.focus",
    "viewer.destroy",
    "element.open",
    "paintSlide",
  ]) {
    expect(api, `documented callable ${callable}`).toContain(callable);
  }
  expect(core).toContain("export type { AaredErrorCode }");
  expect(core).toContain("export { AaredError, isAaredError, isAbort, messageFor }");
  expect(core).toContain("WorkerInit");
  expect(core).not.toContain("export async function open");
  expect(pptx).toContain("export async function open");
  expect(pptx).toContain("export async function render");
  expect(pptx).toContain("read(options?: ReadOptions): Promise<SlideModel>");
  expect(pptx).toContain("close(): void");
  expect(xlsx).toContain("export async function open");
  expect(xlsx).toContain("export async function render");
  expect(xlsx).toContain("read(options?: SheetReadOptions): Promise<SheetRangeModel>");
  expect(xlsx).toContain("goToCell(cell: string): Promise<void>");
  expect(docx).toContain("export async function open");
  expect(docx).toContain("export async function render");
  expect(docx).toContain("search(query: string, options?: ReadOptions): Promise<DocxSearchMatch[]>");
  expect(docx).toContain("goToHeading(target: string | number): Promise<void>");
  expect(docx).toContain("goToPage(index: number): Promise<void>");
  expect(pptx).toContain("export function paintSlide(input: PaintSlideInput)");
  expect(preview).toContain("export async function preview(");
  expect(preview).toContain("export async function detect(");
  expect(preview).toContain("goTo(index: number): Promise<void>");
  expect(preview).toContain("next(): Promise<void>");
  expect(preview).toContain("previous(): Promise<void>");
  expect(preview).toContain("zoomIn(step?: number): void");
  expect(preview).toContain("zoomOut(step?: number): void");
  expect(preview).toContain("resetZoom(): void");
  expect(preview).toContain("readonly defaultZoom: ZoomValue");
  expect(preview).toContain("search(query: string, options?: PreviewSearchOptions): Promise<PptxSearchMatch[]>");
  expect(preview).toContain("search(query: string, options?: PreviewSearchOptions): Promise<XlsxSearchMatch[]>");
  expect(preview).toContain("search(query: string, options?: PreviewSearchOptions): Promise<DocxSearchMatch[]>");
  expect(preview).toContain("setAppearance(appearance: Appearance): void");
  expect(pptx).toContain("setAppearance(");
  expect(xlsx).toContain("setAppearance(");
  expect(docx).toContain("setAppearance(");
  // The hook's result changes nothing itself: changes go through `viewer`.
  expect(react).toContain("retry: () => void;");
  expect(react).not.toContain("goTo: (index: number) => Promise<void>;");
  expect(react).toContain("export function usePreview(");
  expect(react).toContain("export function usePptx(");
  expect(react).toContain("export function useXlsx(");
  expect(react).toContain("export function useDocx(");
  expect(react).toContain("export interface UsePreviewOptions");
  expect(react).toContain("export type UsePreviewResult");
  expect(element).toContain("async open(): Promise<void>");
  expect(element).toContain("get viewer(): AaredViewerHandle | null");
  expect(element).not.toContain("async goTo(");
});

/**
 * The other half of the gate. Export parity does not stop a page from
 * reintroducing an API that was deliberately removed, so the surfaces that
 * outlived each correction are checked for absence across every consumer page
 * the charter treats as one contract.
 */
test("removed APIs stay out of every consumer page", async () => {
  const pages = await Promise.all([
    Bun.file(API_PATH).text(),
    Bun.file(`${import.meta.dir}/../../../site/src/content/docs/start.mdx`).text(),
    Bun.file(`${import.meta.dir}/../../../site/src/partials/choosing-a-layer.mdx`).text(),
  ]);
  for (const [index, page] of pages.entries()) {
    const where = ["api", "start", "choosing-a-layer"][index]!;
    // §122: viewers lend restricted content, never the document's lifecycle.
    expect(page, where).not.toContain("viewer.document");
    expect(page, where).not.toMatch(/viewer\.document\b/);
    // §120: one `error` field on the hook result, one `error` event.
    expect(page, where).not.toContain("navigationError");
    // §119: any format's match is activated on the viewer itself.
    expect(page, where).not.toMatch(/docx\.goToSearchMatch/);
    // §121: the Worker takes exactly one object-form input.
    expect(page, where).not.toMatch(/worker:\s*(?:create|async|\(\))/);
  }
});
