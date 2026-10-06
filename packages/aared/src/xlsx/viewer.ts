/** Fixed-format preview: this route has no dependency on sibling renderers. */
import { open, render } from "./index.js";
import type { XlsxDocument } from "./index.js";
import { openViewer } from "../embed-viewer.js";
import type { ViewerRoute } from "../format-route.js";
import type { PreviewInput, XlsxPreview } from "../preview/index.js";

/** XLSX preview input; the import determines the format. */
export interface XlsxPreviewInput extends Omit<PreviewInput, "type" | "pagination"> {}

/** Internal adapter reused by the auto-detect layer after choosing this format. */
export const route: ViewerRoute = {
  format: "xlsx",
  open,
  render: ({ document, index, ...input }) => render({ ...input, sheet: (document as XlsxDocument).sheets.get(index) }),
};

/** Open, render and navigate a XLSX file; destroy closes the document too. */
export async function preview(input: XlsxPreviewInput): Promise<XlsxPreview> {
  const { source, target, ...options } = input;
  return await openViewer({ scope: "aared/xlsx", source, target, options, route }) as XlsxPreview;
}
