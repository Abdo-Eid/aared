// One consumer module: generic preview and format renderers share their
// internal capability registry, just as they do in a production build.
export { preview as autoPreview } from "../src/preview/index.js";
export { preview as pptxPreview, render as pptxRender } from "../src/pptx/index.js";
export { preview as xlsxPreview, render as xlsxRender } from "../src/xlsx/index.js";
export { preview as docxPreview, render as docxRender } from "../src/docx/index.js";
import "../src/element/index.js";

import type { PreviewHandle } from "../src/preview/index.js";
// This function is type-checked, never executed: the facade is not the owner.
function contentContract(viewer: PreviewHandle): void {
  // @ts-expect-error the full document is not exposed.
  void viewer.document;
  // @ts-expect-error lifecycle stays with the viewer.
  void viewer.content.close;
  // @ts-expect-error search stays on the viewer.
  void viewer.content.search;
  if (viewer.format === "xlsx") {
    const sheet = viewer.content.sheets.get(0);
    // @ts-expect-error no owner escape on a borrowed sheet.
    void sheet.sheetDocument;
    // @ts-expect-error no session field on a borrowed sheet.
    void sheet.session;
  }
}
void contentContract;
