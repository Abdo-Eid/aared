/** Fixed-format preview: this route has no dependency on sibling renderers. */
import { render, XlsxDocument } from "./index.js";
import { openViewer } from "../embed-viewer.js";
import type { ViewerRoute } from "../format-route.js";
import type { PreviewInput, XlsxPreview } from "../preview/index.js";
import { borrowContent, contentCollection, contentMedia } from "../content-view.js";
import type { SheetReadOptions } from "./index.js";

/** XLSX preview input; the import determines the format. */
export interface XlsxPreviewInput extends Omit<PreviewInput, "type" | "pagination"> {}

/** Internal adapter reused by the auto-detect layer after choosing this format. */
export const route: ViewerRoute = {
  format: "xlsx",
  adopt: async (session) => new XlsxDocument(session),
  content: (source, assertOpen, signal) => {
    const document = source as XlsxDocument;
    return Object.freeze({ info: document.info,
      sheets: contentCollection(document.sheets, (sheet) => borrowContent(sheet, {
        info: sheet.info,
        read: async (options?: SheetReadOptions) => { assertOpen(); return sheet.read(options); },
      }, assertOpen, signal), assertOpen),
      media: contentMedia(document.media, assertOpen),
    });
  },
  render: ({ document, index, ...input }) => render({ ...input, sheet: (document as XlsxDocument).sheets.get(index) }),
};

/** Open, render and navigate a XLSX file; destroy closes the document too. */
export async function preview(input: XlsxPreviewInput): Promise<XlsxPreview> {
  const { source, target, ...options } = input;
  return await openViewer({ scope: "aared/xlsx", source, target, options, route }) as XlsxPreview;
}
