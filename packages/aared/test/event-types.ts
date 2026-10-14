import type { AaredViewer } from "../src/element/index.js";
import type { PreviewHandle } from "../src/preview/index.js";
import type { PptxRenderHandle } from "../src/pptx/index.js";
import type { XlsxRenderHandle } from "../src/xlsx/index.js";

/**
 * Compile-time checks, run by the typecheck (`tsc` over `test/`): each event
 * name types its `detail`, and `<aared-viewer>` is typed by its tag. Never
 * called at runtime.
 */
function typedListeners(preview: PreviewHandle, slide: PptxRenderHandle, sheet: XlsxRenderHandle): void {
  const viewer: AaredViewer | null = document.querySelector("aared-viewer");
  const created: AaredViewer = document.createElement("aared-viewer");
  created.file = new Uint8Array(0);
  viewer?.addEventListener("load", (event) => event.detail.count satisfies number);
  viewer?.addEventListener("error", (event) => event.detail.cause.code satisfies string);
  viewer?.addEventListener("progress", (event) => event.detail.loaded satisfies number);
  // The page's own element events keep their types.
  viewer?.addEventListener("click", (event) => event.clientX satisfies number);
  // @ts-expect-error `load` carries DocumentInfo, which has no `nope`.
  viewer?.addEventListener("load", (event) => event.detail.nope);
  preview.addEventListener("indexchange", (event) => event.detail.index satisfies number);
  preview.addEventListener("error", (event) => event.detail.cause.code satisfies string);
  // search() returns the format's own matches once `format` is checked.
  if (preview.format === "xlsx") void preview.search("x").then((matches) => matches[0]?.address satisfies string | undefined);
  if (preview.format === "pptx") void preview.search("x").then((matches) => matches[0]?.slide.index satisfies number | undefined);
  if (preview.format === "docx") void preview.search("x").then((matches) => matches[0]?.heading?.index satisfies number | undefined);
  slide.addEventListener("linkactivate", (event) => event.detail.target satisfies string);
  sheet.addEventListener("cellactivate", (event) => event.detail.address satisfies string);
  sheet.addEventListener("selectionchange", (event) => event.detail.selection?.row satisfies number | undefined);
  // Any other name still takes a plain listener.
  sheet.addEventListener("custom", (event: Event) => event.type);
}

/**
 * `el.viewer`: the preview handle without what the element owns. Checking
 * `format` still narrows it, its events stay typed, and `setAppearance` /
 * `destroy` go through the element (`theme`, removing it).
 */
function elementViewer(element: AaredViewer): void {
  const viewer = element.viewer;
  if (viewer === null) return;
  viewer.addEventListener("indexchange", (event) => event.detail.index satisfies number);
  element.pagination = "layout";
  // @ts-expect-error a DOCX is laid out as "flow" or "layout".
  element.pagination = "pages";
  void viewer.goTo(1);
  viewer.setZoom("fit-width");
  // The toolbar methods reach the element's viewer too.
  void viewer.next();
  void viewer.previous();
  viewer.zoomIn(0.1);
  viewer.zoomOut();
  viewer.resetZoom();
  viewer.setZoom(viewer.defaultZoom);
  if (viewer.format === "xlsx") void viewer.search("x").then((matches) => matches[0]?.sheet.index satisfies number | undefined);
  if (viewer.format === "xlsx") void viewer.xlsx.goToCell("D14");
  if (viewer.format === "docx") void viewer.docx.goToHeading(0);
  // @ts-expect-error the element's `theme` attribute owns the appearance.
  viewer.setAppearance({ theme: "dark" });
  // @ts-expect-error removing the element is its teardown.
  viewer.destroy();
}

void typedListeners;
void elementViewer;
