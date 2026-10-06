/** Fixed-format preview: this route has no dependency on sibling renderers. */
import { open, render } from "./index.js";
import type { PptxDocument } from "./index.js";
import { openViewer } from "../embed-viewer.js";
import type { ViewerRoute } from "../format-route.js";
import type { PreviewInput, PptxPreview } from "../preview/index.js";

/** PPTX preview input; the import determines the format. */
export interface PptxPreviewInput extends Omit<PreviewInput, "type" | "pagination"> {}

/** Internal adapter reused by the auto-detect layer after choosing this format. */
export const route: ViewerRoute = {
  format: "pptx",
  open,
  render: ({ document, index, ...input }) => render({ ...input, slide: (document as PptxDocument).slides.get(index) }),
};

/** Open, render and navigate a PPTX file; destroy closes the document too. */
export async function preview(input: PptxPreviewInput): Promise<PptxPreview> {
  const { source, target, ...options } = input;
  return await openViewer({ scope: "aared/pptx", source, target, options, route }) as PptxPreview;
}
