/** Fixed-format preview: this route has no dependency on sibling renderers. */
import { render, PptxDocument } from "./renderer.js";
import { openViewer } from "../embed-viewer.js";
import type { ViewerRoute } from "../format-route.js";
import type { PreviewInput, PptxPreview } from "../preview/index.js";
import { borrowContent, contentCollection, contentMedia } from "../content-view.js";
import type { ReadOptions } from "../core/index.js";

/** PPTX preview input; the import determines the format. */
export interface PptxPreviewInput extends Omit<PreviewInput, "type" | "pagination"> {}

/** Internal adapter reused by the auto-detect layer after choosing this format. */
export const route: ViewerRoute = {
  format: "pptx",
  adopt: async (session) => new PptxDocument(session),
  content: (source, assertOpen, signal) => {
    const document = source as PptxDocument;
    return Object.freeze({ info: document.info,
      slides: contentCollection(document.slides, (slide) => borrowContent(slide, {
        info: slide.info,
        read: async (options?: ReadOptions) => { assertOpen(); return slide.read(options); },
      }, assertOpen, signal), assertOpen),
      media: contentMedia(document.media, assertOpen),
    });
  },
  render: ({ document, index, ...input }) => render({ ...input, slide: (document as PptxDocument).slides.get(index) }),
};

/** Open, render and navigate a PPTX file; destroy closes the document too. */
export async function preview(input: PptxPreviewInput): Promise<PptxPreview> {
  const { source, target, ...options } = input;
  return await openViewer({ scope: "@aared/viewer/pptx", source, target, options, route }) as PptxPreview;
}
