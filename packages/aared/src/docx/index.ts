/** Public Word operations, section metadata and outline; layout models are private. */
import type { Appearance, Collection, MediaCollection, MediaResolver, OpenOptions, ReadOptions, Source, ZoomValue } from "../core/index.js";
import type { DocxContent, DocxInfo, DocxPagination, DocxRenderHandle, DocxSearchMatch, HeadingInfo } from "./renderer.js";
import type { SectionInfo } from "./data.js";
import { open as openDocument, render as renderDocument, isDocumentInput } from "./renderer.js";
import { AaredError, throwIfAborted } from "../core/errors.js";

export type { SectionInfo } from "./data.js";
export type { DocxInfo, HeadingInfo, DocxPageInfo, DocxSearchMatch, DocxContent, DocxPagination, SoftBreakPolicy, DocxRenderEventMap, DocxRenderHandle } from "./renderer.js";

export interface SectionCollection extends Collection<SectionInfo, number> {}
export interface HeadingCollection extends Collection<HeadingInfo, number | string> {}
export interface DocxDocument {
  readonly info: DocxInfo;
  readonly sections: SectionCollection;
  readonly headings: HeadingCollection;
  readonly media: MediaCollection;
  search(query: string, options?: ReadOptions): Promise<DocxSearchMatch[]>;
  close(): void;
}
export interface DocxRenderInput {
  document: DocxDocument | DocxContent;
  target: HTMLElement;
  pagination?: DocxPagination;
  appearance?: Appearance;
  zoom?: ZoomValue;
  mediaResolver?: MediaResolver;
  signal?: AbortSignal;
}
export async function open(source: Source, options?: OpenOptions): Promise<DocxDocument> {
  return openDocument(source, options);
}
export async function render(input: DocxRenderInput): Promise<DocxRenderHandle> {
  if (typeof input !== "object" || input === null) throw new AaredError("invalid-argument", "@aared/viewer/docx: render() takes an input object");
  throwIfAborted(input.signal);
  if (!isDocumentInput(input.document)) throw new AaredError("invalid-argument", "@aared/viewer/docx: render() takes a document from open() or viewer.content");
  return renderDocument(input);
}
export async function preview(input: import("./viewer.js").DocxPreviewInput): Promise<import("../preview/index.js").DocxPreview> {
  return (await import("./viewer.js")).preview(input);
}
export type { DocxPreviewInput } from "./viewer.js";
export type { DocxPreview, PreviewSearchMatch, PreviewSearchOptions, PreviewEventMap, SearchNavigation, SearchNavigationState } from "../preview/index.js";
