/** Public presentation operations and semantic data; rendering models are private. */
import type { Appearance, Collection, MediaCollection, MediaOptions, MediaResolver, OpenOptions, ReadOptions, Source, ZoomValue } from "../core/index.js";
import type { PptxInfo, PptxRenderHandle, PptxSearchMatch, SlideContent, SlideInfo } from "./renderer.js";
import type { SlideText } from "./data.js";
import { open as openDocument, render as renderSlide, isSlideInput } from "./renderer.js";
import { AaredError, throwIfAborted } from "../core/errors.js";

export type { SlideText } from "./data.js";
export type { PptxInfo, SlideInfo, SlideContent, PptxContent, PptxSearchMatch, PptxRenderEventMap, PptxRenderHandle } from "./renderer.js";

export interface SlideHandle extends SlideContent {
  read(options?: ReadOptions): Promise<SlideText>;
  media(mediaId: string, options?: MediaOptions): Promise<ArrayBuffer>;
}
export interface SlideCollection extends Collection<SlideHandle, number> {}
export interface PptxDocument {
  readonly info: PptxInfo;
  readonly slides: SlideCollection;
  readonly media: MediaCollection;
  search(query: string, options?: ReadOptions): Promise<PptxSearchMatch[]>;
  close(): void;
}
export interface PptxRenderInput {
  slide: SlideHandle | SlideContent;
  target: HTMLElement;
  appearance?: Appearance;
  zoom?: ZoomValue;
  mediaResolver?: MediaResolver;
  signal?: AbortSignal;
}
export async function open(source: Source, options?: OpenOptions): Promise<PptxDocument> {
  return openDocument(source, options);
}
export async function render(input: PptxRenderInput): Promise<PptxRenderHandle> {
  if (typeof input !== "object" || input === null) throw new AaredError("invalid-argument", "@aared/viewer/pptx: render() takes an input object");
  throwIfAborted(input.signal);
  if (!isSlideInput(input.slide)) throw new AaredError("invalid-argument", "@aared/viewer/pptx: render() takes a slide from open() or viewer.content.slides");
  return renderSlide(input);
}
export async function preview(input: import("./viewer.js").PptxPreviewInput): Promise<import("../preview/index.js").PptxPreview> {
  return (await import("./viewer.js")).preview(input);
}
export type { PptxPreviewInput } from "./viewer.js";
export type { PptxPreview, PreviewSearchMatch, PreviewSearchOptions, PreviewEventMap, SearchNavigation, SearchNavigationState } from "../preview/index.js";
