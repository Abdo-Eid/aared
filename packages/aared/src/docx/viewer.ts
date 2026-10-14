/** Fixed-format preview: this route has no dependency on sibling renderers. */
import { render } from "./renderer.js";
import type { DocxDocument } from "./renderer.js";
import { documentFromSession } from "./session-document.js";
import { openViewer } from "../embed-viewer.js";
import type { ViewerRoute } from "../format-route.js";
import type { PreviewInput, DocxPreview } from "../preview/index.js";
import { borrowContent, contentCollection, contentMedia } from "../content-view.js";

/** DOCX preview input; the import determines the format. */
export interface DocxPreviewInput extends Omit<PreviewInput, "type"> {}

/** Internal adapter reused by the auto-detect layer after choosing this format. */
export const route: ViewerRoute = {
  format: "docx",
  adopt: documentFromSession,
  content: (source, assertOpen, signal) => {
    const document = source as DocxDocument;
    return borrowContent(document, { info: document.info,
      // Section metadata is freshly projected, so do not retain each read in
      // contentCollection's identity cache (which is for stable handles).
      sections: Object.freeze({
        list: () => { assertOpen(); return document.sections.list(); },
        get: (index: number) => { assertOpen(); return document.sections.get(index); },
      }),
      headings: contentCollection(document.headings, (heading) => heading, assertOpen),
      media: contentMedia(document.media, assertOpen),
    }, assertOpen, signal);
  },
  render: ({ document, index: _index, ...input }) => render({ ...input, document: document as DocxDocument }),
};

/** Open, render and navigate a DOCX file; destroy closes the document too. */
export async function preview(input: DocxPreviewInput): Promise<DocxPreview> {
  const { source, target, ...options } = input;
  return await openViewer({ scope: "@aared/viewer/docx", source, target, options, route }) as DocxPreview;
}
