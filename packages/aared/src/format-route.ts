/**
 * How `aared/preview` and `<aared-viewer>` reach the format entrypoints: the
 * one place that lazily imports the detected format, plus the defaults and
 * events both facades share. Renderer-side. Each facade still owns its own
 * render call (the element adds sheet tabs and its theme attribute;
 * `preview()` forwards the caller's options), so nothing here paints.
 *
 * The format modules are reached only through dynamic `import()`, so a
 * consumer of one format never loads the other two.
 */

import { AaredError } from "./core/index.js";
import type { Appearance, DocumentFormat, OpenOptions, UnitRef, ZoomValue } from "./core/index.js";
import type { DocxDocument } from "./docx/index.js";
import type { PptxDocument } from "./pptx/index.js";
import type { XlsxDocument } from "./xlsx/index.js";

export type AnyDocument = PptxDocument | XlsxDocument | DocxDocument;

/**
 * Each format's starting zoom when the caller names none: a slide opens whole,
 * a sheet at 100% like Excel, a document to the width of its container.
 */
export const DEFAULT_ZOOM: Readonly<Record<DocumentFormat, ZoomValue>> = {
  pptx: "fit-page",
  xlsx: 1,
  docx: "fit-width",
};

/** What every format's render handle offers a facade. */
export interface RenderHandle extends EventTarget {
  destroy(): void;
  readonly destroyed: boolean;
  setZoom(zoom: ZoomValue): void;
  readonly zoom: ZoomValue;
  readonly scale: number;
  setAppearance(appearance: Appearance): void;
}

/** Open `bytes` as `format`, loading only that format's entrypoint. */
export async function openFormat(format: DocumentFormat, bytes: ArrayBuffer, options: OpenOptions | undefined): Promise<AnyDocument> {
  if (format === "pptx") return (await import("./pptx/index.js")).open(bytes, options);
  if (format === "xlsx") return (await import("./xlsx/index.js")).open(bytes, options);
  return (await import("./docx/index.js")).open(bytes, options);
}

/**
 * A facade never shows a document with nothing in it: zero slides, sheets or
 * sections is `invalid-document` in the facade's scope. Shared so the element
 * and `preview()` cannot drift apart.
 */
export function assertHasUnits(scope: string, document: AnyDocument): void {
  if (document.info.count === 0) {
    throw new AaredError("invalid-document", `${scope}: this document has no slides, sheets, or sections to show`);
  }
}

/** The unit reference (`{ index, title }`) for `index` in `document`. */
export function unitRef(document: AnyDocument, index: number): UnitRef {
  return { index, title: document.info.titles[index] ?? "" };
}

/**
 * Announce a navigation to another unit: `indexchange` always, plus the
 * format-native sibling (`slidechange`, `sheetchange`) with the same
 * `{ index, title }` detail. Never called for the starting unit — the ready
 * signal (`load`, `preview()` resolving) carries that.
 */
export function announceUnit(target: EventTarget, format: DocumentFormat, document: AnyDocument, index: number): void {
  target.dispatchEvent(new CustomEvent("indexchange", { detail: { index } }));
  const ref = unitRef(document, index);
  if (format === "pptx") target.dispatchEvent(new CustomEvent("slidechange", { detail: ref }));
  else if (format === "xlsx") target.dispatchEvent(new CustomEvent("sheetchange", { detail: ref }));
}

/**
 * Events a render handle raises that a facade re-raises on itself. A handle's
 * own `sheetchange` (a sheet-tab click) is not here: the facade turns it into
 * its own unit announcement so its `index` follows (see `followSheet`).
 */
const FORWARDED_EVENTS = ["zoomchange", "linkactivate", "cellactivate", "selectionchange"] as const;

/** Re-raise the handle's events on `to` until `signal` aborts. */
export function forwardRenderEvents(from: EventTarget, to: EventTarget, signal: AbortSignal): void {
  for (const type of FORWARDED_EVENTS) {
    from.addEventListener(type, (event) => {
      to.dispatchEvent(new CustomEvent(type, { detail: (event as CustomEvent).detail }));
    }, { signal });
  }
}

/**
 * Keep a facade's unit index in step with the XLSX tab strip: when a tab click
 * switches sheets inside the render handle, `onSheet` receives the new index
 * (the facade updates itself, then announces it). Nothing for PPTX or DOCX,
 * whose handles never change unit on their own.
 */
export function followSheet(handle: EventTarget, signal: AbortSignal, onSheet: (index: number) => void): void {
  handle.addEventListener("sheetchange", (event) => {
    onSheet((event as CustomEvent<UnitRef>).detail.index);
  }, { signal });
}
