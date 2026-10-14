/**
 * How the embed viewer (`embed-viewer.ts`, behind `@aared/viewer/preview`,
 * `<aared-viewer>` and the hooks) reaches the format entrypoints: the one
 * defaults, route contracts, checks and events it shares. Renderer-side;
 * nothing here paints or imports a format implementation at runtime.
 */

import { AaredError } from "./core/index.js";
import type { Appearance, DocumentFormat, MediaResolver, OpenOptions, UnitRef, ZoomValue } from "./core/index.js";
import type { DocxPagination, PreviewEventMap } from "./preview/index.js";
import type { DocxContent, DocxDocument } from "./docx/index.js";
import type { PptxContent, PptxDocument } from "./pptx/index.js";
import type { XlsxContent, XlsxDocument } from "./xlsx/index.js";
import type { WorkerSession } from "./core/session.js";

export type AnyDocument = PptxDocument | XlsxDocument | DocxDocument;
export type AnyContent = PptxContent | XlsxContent | DocxContent;

/** Listed as a record so adding a public event requires forwarding it too. */
const VIEWER_EVENT_NAMES: Readonly<Record<keyof PreviewEventMap, true>> = {
  searchchange: true,
  indexchange: true,
  slidechange: true,
  sheetchange: true,
  zoomchange: true,
  error: true,
  linkactivate: true,
  cellactivate: true,
  selectionchange: true,
};

/** Event forwarding needs metadata, never an eager import of the controller. */
export const VIEWER_EVENTS = Object.keys(VIEWER_EVENT_NAMES) as ReadonlyArray<keyof PreviewEventMap>;

/**
 * The appearance an embed layer renders with: the caller's, with the embed
 * defaults under whatever it leaves out. A workbook shows its sheet tabs,
 * since an embed has no other way to switch sheets; the headless `render()`
 * default stays off, for an app that draws its own. Applied at every render
 * and every appearance change the viewer forwards.
 */
export function embedAppearance(format: DocumentFormat, appearance: Appearance): Appearance;
export function embedAppearance(format: DocumentFormat, appearance: Appearance | undefined): Appearance | undefined;
export function embedAppearance(format: DocumentFormat, appearance: Appearance | undefined): Appearance | undefined {
  if (format !== "xlsx" || appearance?.spreadsheet?.sheetTabs !== undefined) return appearance;
  return { ...appearance, spreadsheet: { ...appearance?.spreadsheet, sheetTabs: true } };
}

/**
 * The viewer's internal "replace the whole appearance" call, for the React
 * hook alone: its `appearance` option is the full state it wants, so a knob
 * removed from the option must go back to its default, which the public
 * merging `setAppearance()` never does. A registry symbol, so the hook and
 * the lazily loaded viewer chunk agree on it without sharing a module.
 */
export const REPLACE_APPEARANCE: unique symbol = Symbol.for("aared.replaceAppearance");

/** What every format's render handle offers a facade. */
export interface RenderHandle extends EventTarget {
  destroy(): void;
  readonly destroyed: boolean;
  setZoom(zoom: ZoomValue): void;
  readonly zoom: ZoomValue;
  readonly scale: number;
  setAppearance(appearance: Appearance): void;
}

/** Resolved renderer inputs shared by statically bound format routes. */
export interface ViewerRenderInput {
  document: AnyDocument;
  index: number;
  target: HTMLElement;
  signal: AbortSignal;
  appearance?: Appearance;
  zoom?: ZoomValue;
  mediaResolver?: MediaResolver;
  pagination?: DocxPagination;
}

/** A format supplies parsing and rendering; the shared viewer owns lifecycle. */
export interface ViewerRoute {
  readonly format: DocumentFormat;
  adopt(session: WorkerSession, options: OpenOptions | undefined): Promise<AnyDocument>;
  content(document: AnyDocument, assertOpen: () => void, signal: AbortSignal): AnyContent;
  render(input: ViewerRenderInput): Promise<RenderHandle>;
}

/**
 * A facade never shows a document with nothing in it: zero slides, sheets or
 * sections is `invalid-document` in the facade's scope.
 */
export function assertHasUnits(scope: string, document: AnyDocument): void {
  if (document.info.count === 0) {
    throw new AaredError("invalid-document", `${scope}: this document has no slides, sheets, or sections to show`);
  }
}

/**
 * A facade's unit index (`goTo()`, the starting `index`): a whole number in
 * `0 <= index < count`. A whole number outside is `not-found`; anything else
 * is not an index at all, `invalid-argument`. The one index check, so every
 * layer raises the same code for the same index (two hand-written checks
 * drifted apart once, §59).
 */
export function assertUnitIndex(scope: string, index: number, count: number): void {
  if (!Number.isInteger(index)) {
    throw new AaredError("invalid-argument", `${scope}: a unit index must be a whole number; received ${String(index)}`);
  }
  if (index < 0 || index >= count) {
    throw new AaredError("not-found", `${scope}: unit index ${index} is outside this ${count}-unit document`);
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
const FORWARDED_EVENTS = ["zoomchange", "linkactivate", "cellactivate", "selectionchange", "error"] as const;

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
