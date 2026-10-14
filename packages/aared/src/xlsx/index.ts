/** Public spreadsheet operations and cached cell data; painting models are private. */
import type { Appearance, Collection, MediaCollection, MediaOptions, MediaResolver, OpenOptions, ReadOptions, Source, ZoomValue } from "../core/index.js";
import type { SheetContent, SheetReadOptions, XlsxInfo, XlsxRenderHandle, XlsxSearchMatch } from "./renderer.js";
import type { SheetData } from "./data.js";
import { open as openDocument, render as renderSheet, isSheetInput } from "./renderer.js";
import { AaredError, throwIfAborted } from "../core/errors.js";

export type { CellData, RowData, SheetData, CellRange } from "./data.js";
export type { XlsxInfo, SheetInfo, XlsxSearchMatch, SheetReadOptions, SheetContent, XlsxContent, XlsxRenderEventMap, XlsxRenderHandle } from "./renderer.js";

export interface SheetHandle extends SheetContent {
  read(options?: SheetReadOptions): Promise<SheetData>;
  search(query: string, options?: ReadOptions): Promise<XlsxSearchMatch[]>;
  media(mediaId: string, options?: MediaOptions): Promise<ArrayBuffer>;
}
export interface SheetCollection extends Collection<SheetHandle, number | string> {}
export interface XlsxDocument {
  readonly info: XlsxInfo;
  readonly sheets: SheetCollection;
  readonly media: MediaCollection;
  close(): void;
}
export interface XlsxRenderInput {
  sheet: SheetHandle | SheetContent;
  target: HTMLElement;
  appearance?: Appearance;
  zoom?: ZoomValue;
  mediaResolver?: MediaResolver;
  signal?: AbortSignal;
}
export async function open(source: Source, options?: OpenOptions): Promise<XlsxDocument> {
  return openDocument(source, options);
}
export async function render(input: XlsxRenderInput): Promise<XlsxRenderHandle> {
  if (typeof input !== "object" || input === null) throw new AaredError("invalid-argument", "@aared/viewer/xlsx: render() takes an input object");
  throwIfAborted(input.signal);
  if (!isSheetInput(input.sheet)) throw new AaredError("invalid-argument", "@aared/viewer/xlsx: render() takes a sheet from open() or viewer.content.sheets");
  return renderSheet(input);
}
export async function preview(input: import("./viewer.js").XlsxPreviewInput): Promise<import("../preview/index.js").XlsxPreview> {
  return (await import("./viewer.js")).preview(input);
}
export type { XlsxPreviewInput } from "./viewer.js";
export type { XlsxPreview, PreviewSearchMatch, PreviewSearchOptions, PreviewEventMap, SearchNavigation, SearchNavigationState } from "../preview/index.js";
