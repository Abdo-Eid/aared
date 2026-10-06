/** Fixed-format preview: this route has no dependency on sibling renderers. */
import { open, render } from "./index.js";
import type { DocxDocument } from "./index.js";
import { openViewer } from "../embed-viewer.js";
import type { ViewerRoute } from "../format-route.js";
import type { PreviewInput, DocxPreview } from "../preview/index.js";

/** DOCX preview input; the import determines the format. */
export interface DocxPreviewInput extends Omit<PreviewInput, "type"> {}

/** Internal adapter reused by the auto-detect layer after choosing this format. */
export const route: ViewerRoute = {
  format: "docx",
  open,
  render: ({ document, index: _index, ...input }) => render({ ...input, document: document as DocxDocument }),
};

/** Open, render and navigate a DOCX file; destroy closes the document too. */
export async function preview(input: DocxPreviewInput): Promise<DocxPreview> {
  const { source, target, ...options } = input;
  return await openViewer({ scope: "aared/docx", source, target, options, route }) as DocxPreview;
}
