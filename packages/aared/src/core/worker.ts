import { OpcPackage, officeFormatOf } from "./opc.js";
import type { OfficeFormat } from "./opc.js";
import type { PptxDeck } from "./pptx.js";
import type { XlsxWorkbook, XlsxWorkbookIndex } from "./xlsx/index.js";
import type { DocxIndex } from "./docx/document.js";
import type { DocumentInfo } from "./index.js";
import type { SheetModel, SheetRangeModel, SlideLayoutModel, SlideModel, UnitModel } from "./model.js";
import type { SerializedError, UnitRequest, WorkerRequest, WorkerResponse, WorkerResult } from "./rpc.js";
import { AaredError, damagedDocument, isAbort, throwIfAborted } from "./errors.js";
import { ZipError } from "./zip.js";
import { XmlError } from "./xml.js";
import { collectStrings } from "./collect.js";
import { isInsideGrid } from "./xlsx/limits.js";
import { PPTX_SLIDE_CACHE, UnitCache, XLSX_SHEET_CACHE } from "./unit-cache.js";

interface WorkerScope {
  onmessage: ((event: MessageEvent<WorkerRequest>) => void) | null;
  postMessage(message: WorkerResponse, transfer?: Transferable[]): void;
}

interface FormatSession {
  format: OfficeFormat;
  pkg: OpcPackage;
  deck?: PptxDeck;
  /** The most recently read slides, so re-reading one (a re-render, a neighbour) skips the parse; bounded. */
  pptxSlides?: UnitCache<SlideModel>;
  xlsxIndex?: XlsxWorkbookIndex;
  /** Set only once a read finishes: an aborted read must not leave a rejected result behind. */
  xlsxWorkbook?: XlsxWorkbook;
  /** The most recently read sheets, so range reads slice without re-parsing; bounded. */
  xlsxSheets?: UnitCache<SheetModel>;
  /** A DOCX session's light index; the body and notes are read, and dropped, per section parse. */
  docxIndex?: DocxIndex;
  /** Parts the media route may serve: seeded from package media/font dirs, extended with emitted mediaIds. */
  mediaParts: Set<string>;
}

/**
 * Format parsers load on first use through dynamic `import()`. The published
 * `dist/core/worker.mjs` is bundled flat (`codeSplitting: false`), so all
 * three resolve in place inside the one file — the worker needs no sibling
 * chunks and also backs Blob-URL workers directly.
 */
const formats = {
  pptx: () => import("./pptx.js"),
  xlsx: () => import("./xlsx/index.js"),
  docx: () => import("./docx/document.js"),
};

const scope = globalThis as unknown as WorkerScope;
const sessions = new Map<number, FormatSession>();
const sessionOpenRequests = new Map<number, number>();
const cancelled = new Set<string>();
const queue: Array<Exclude<WorkerRequest, { type: "cancel" | "close" }>> = [];
let draining = false;
let activeRequest: string | undefined;
let activeController: AbortController | undefined;

function requestKey(sessionId: number, requestId: number): string {
  return `${sessionId}:${requestId}`;
}

function serializeError(error: unknown): SerializedError {
  if (error instanceof AaredError) return { name: error.name, message: error.message, code: error.code };
  if (isAbort(error)) return { name: error.name, message: error.message, code: "aborted" };
  // ZIP failures are document problems from the consumer's view: a file
  // that is not, or not cleanly, a ZIP archive is not a previewable
  // document. Preserve the `ZipError` message text; only the code changes.
  if (error instanceof ZipError) {
    return { name: "AaredError", message: error.message, code: "invalid-document" };
  }
  // A part that is not well-formed XML is the file's fault, like the explicit
  // structural checks (`damagedDocument`, thrown where they are made).
  if (error instanceof XmlError) {
    const typed = damagedDocument(error.message.replace(/^aared\/core: /, ""));
    return { name: typed.name, message: typed.message, code: typed.code };
  }
  // Anything else is aared failing on this file, not the file being broken:
  // `internal`, never "re-save it in Office". The original error travels as
  // `cause` (name, message, stack) for the bug report.
  const original = error instanceof Error ? error : new Error(String(error));
  return {
    name: "AaredError",
    message: `aared/core: the viewer failed while reading this file (${original.name}: ${original.message})`,
    code: "internal",
    cause: { name: original.name, message: original.message, ...(original.stack === undefined ? {} : { stack: original.stack }) },
  };
}

/**
 * A request for a session this Worker no longer holds (closed, or its open was
 * cancelled). A usage error, never a cancellation: `"aborted"` here let
 * `isAbort()` swallow it silently.
 */
function sessionClosed(): AaredError {
  return new AaredError("closed", "aared/core: document session is closed");
}

function formatOf(pkg: OpcPackage, signal?: AbortSignal): OfficeFormat | undefined {
  for (const part of pkg.partNames()) {
    throwIfAborted(signal);
    const format = officeFormatOf(pkg.contentTypeOf(part));
    if (format !== undefined) return format;
  }
  return undefined;
}

function themeInfo(theme: PptxDeck["theme"]): DocumentInfo["theme"] {
  return {
    colors: { ...theme.colors },
    fonts: {
      major: theme.fonts.major.latin || theme.fonts.major.complexScript,
      minor: theme.fonts.minor.latin || theme.fonts.minor.complexScript,
    },
  };
}

function infoForPptx(deck: PptxDeck): DocumentInfo {
  return {
    format: "pptx",
    count: deck.slideParts.length,
    titles: deck.slideParts.map((_, index) => `Slide ${index + 1}`),
    sizes: deck.slideParts.map(() => ({ ...deck.size })),
    theme: themeInfo(deck.theme),
  };
}

function infoForXlsx(index: XlsxWorkbookIndex): DocumentInfo {
  return {
    format: "xlsx",
    count: index.sheets.length,
    titles: index.sheets.map((sheet) => sheet.name),
    sheetStates: index.sheets.map((sheet) => sheet.state),
    theme: themeInfo(index.theme),
  };
}

/**
 * A Word document is one unit (one flow; pages are a render result), so it has
 * one title and no `sizes`: flow mode has no size, layout mode one per
 * section, and neither is a per-unit fact. The outline is `document.headings`.
 */
function infoForDocx(index: DocxIndex): DocumentInfo {
  return {
    format: "docx",
    count: 1,
    titles: [index.title ?? "Document"],
    theme: themeInfo(index.theme),
  };
}

function seedMediaParts(pkg: OpcPackage): Set<string> {
  return new Set(
    pkg.partNames().filter((part) =>
      part.startsWith("ppt/media/")
      || part.startsWith("ppt/fonts/")
      || part.startsWith("xl/media/")
      || part.startsWith("word/media/")),
  );
}

/** Units whose `mediaId`s are already recorded; keyed by object, so a cache hit skips the walk. */
const mediaCollected = new WeakSet<object>();

/** Record every `mediaId` a resolved unit references so the media route can serve it. */
function collectMediaIds(unit: UnitModel | SlideLayoutModel, into: Set<string>): void {
  collectStrings(unit, "mediaId", into);
}

function deleteSession(sessionId: number, openRequestId?: number): void {
  if (openRequestId !== undefined && sessionOpenRequests.get(sessionId) !== openRequestId) return;
  sessions.delete(sessionId);
  sessionOpenRequests.delete(sessionId);
}

function removeQueuedSessionRequests(sessionId: number): void {
  for (let index = queue.length - 1; index >= 0; index--) {
    if (queue[index]!.sessionId === sessionId) queue.splice(index, 1);
  }
}

function postResponse(sessionId: number, requestId: number, result: WorkerResult, transfer?: Transferable[]): void {
  const key = requestKey(sessionId, requestId);
  if (cancelled.delete(key)) return;
  scope.postMessage({ type: "response", sessionId, requestId, ok: true, result }, transfer);
}

function postFailure(sessionId: number, requestId: number, error: unknown): void {
  const key = requestKey(sessionId, requestId);
  if (cancelled.delete(key)) return;
  scope.postMessage({ type: "response", sessionId, requestId, ok: false, error: serializeError(error) });
}

async function xlsxWorkbookOf(session: FormatSession, signal: AbortSignal): Promise<XlsxWorkbook> {
  if (session.xlsxIndex === undefined) throw sessionClosed();
  // The worker handles one request at a time, so there is never a concurrent
  // first read to share; caching the finished value (not the promise) keeps a
  // cancelled first read from failing every later sheet read.
  session.xlsxWorkbook ??= await (await formats.xlsx()).readWorkbook(session.pkg, signal, session.xlsxIndex);
  return session.xlsxWorkbook;
}

/**
 * A parsed slide: from the session's bounded cache when it is there. A read
 * that won't paint (`retain` false: search) neither refreshes nor fills the
 * cache, so a first search over every slide doesn't evict the one on screen.
 */
async function pptxSlideOf(session: FormatSession, deck: PptxDeck, index: number, retain: boolean, signal: AbortSignal): Promise<SlideModel> {
  session.pptxSlides ??= new UnitCache(PPTX_SLIDE_CACHE);
  const cached = retain ? session.pptxSlides.get(index) : session.pptxSlides.peek(index);
  if (cached !== undefined) return cached;
  const slide = await (await formats.pptx()).parsePptxSlide(session.pkg, deck, index, signal);
  if (retain) session.pptxSlides.set(index, slide);
  return slide;
}

async function xlsxSheetOf(session: FormatSession, index: number, retain: boolean, signal: AbortSignal): Promise<SheetModel> {
  session.xlsxSheets ??= new UnitCache(XLSX_SHEET_CACHE);
  const cached = retain ? session.xlsxSheets.get(index) : session.xlsxSheets.peek(index);
  if (cached !== undefined) return cached;
  const workbook = await xlsxWorkbookOf(session, signal);
  throwIfAborted(signal);
  const sheet = await (await formats.xlsx()).parseXlsxSheet(session.pkg, workbook, index, signal);
  if (retain) session.xlsxSheets.set(index, sheet);
  return sheet;
}

function checkRange(range: NonNullable<UnitRequest["range"]>): void {
  const { start, end } = range;
  if (
    !isInsideGrid(start.row, start.col) || !isInsideGrid(end.row, end.col)
    || start.row > end.row || start.col > end.col
  ) {
    throw new AaredError(
      "invalid-range",
      `aared/core: invalid sheet window (${start.row}:${start.col} to ${end.row}:${end.col}); rows and columns are zero-based with start at or before end`,
    );
  }
}

async function handle(request: Exclude<WorkerRequest, { type: "cancel" | "close" }>, signal: AbortSignal): Promise<void> {
  let installedSession = false;
  try {
    throwIfAborted(signal);
    if (request.type === "detect") {
      const pkg = await OpcPackage.open(request.bytes, signal);
      throwIfAborted(signal);
      const format = formatOf(pkg, signal);
      if (format === undefined) {
        throw new AaredError("unsupported-format", "aared/core: unsupported Office document format");
      }
      throwIfAborted(signal);
      postResponse(request.sessionId, request.requestId, format);
      return;
    }
    if (request.type === "open") {
      const pkg = await OpcPackage.open(request.bytes, signal);
      throwIfAborted(signal);
      const format = formatOf(pkg, signal);
      if (format === undefined) throw new AaredError("unsupported-format", "aared/core: unsupported Office document format");
      const mediaParts = seedMediaParts(pkg);
      if (format === "pptx") {
        const deck = await (await formats.pptx()).indexPptx(pkg, signal);
        throwIfAborted(signal);
        sessions.set(request.sessionId, { format, pkg, deck, mediaParts });
        sessionOpenRequests.set(request.sessionId, request.requestId);
        installedSession = true;
        throwIfAborted(signal);
        postResponse(request.sessionId, request.requestId, infoForPptx(deck));
        return;
      }
      if (format === "xlsx") {
        const xlsxIndex = await (await formats.xlsx()).indexXlsx(pkg, signal);
        throwIfAborted(signal);
        sessions.set(request.sessionId, { format, pkg, xlsxIndex, mediaParts });
        sessionOpenRequests.set(request.sessionId, request.requestId);
        installedSession = true;
        throwIfAborted(signal);
        postResponse(request.sessionId, request.requestId, infoForXlsx(xlsxIndex));
        return;
      }
      const docxIndex = await (await formats.docx()).indexDocx(pkg, signal);
      throwIfAborted(signal);
      sessions.set(request.sessionId, { format, pkg, docxIndex, mediaParts });
      sessionOpenRequests.set(request.sessionId, request.requestId);
      installedSession = true;
      throwIfAborted(signal);
      postResponse(request.sessionId, request.requestId, infoForDocx(docxIndex));
      return;
    }

    const session = sessions.get(request.sessionId);
    if (session === undefined) throw sessionClosed();
    if (request.type === "unit") {
      let unit: UnitModel;
      if (session.format === "pptx") {
        if (request.range !== undefined) {
          throw new AaredError(
            "unsupported-feature",
            "aared/core: range reads are an XLSX-only feature; a slide always resolves whole",
          );
        }
        if (session.deck === undefined) throw sessionClosed();
        unit = await pptxSlideOf(session, session.deck, request.index, request.retain !== false, signal);
      } else if (session.format === "xlsx") {
        const sheet = await xlsxSheetOf(session, request.index, request.retain !== false, signal);
        throwIfAborted(signal);
        let windowed: SheetModel | SheetRangeModel = sheet;
        if (request.range !== undefined) {
          checkRange(request.range);
          windowed = (await formats.xlsx()).sliceSheetToRange(sheet, request.range);
        }
        unit = windowed;
      } else {
        if (request.range !== undefined) {
          throw new AaredError(
            "unsupported-feature",
            "aared/core: range reads are an XLSX-only feature; a DOCX document always resolves whole",
          );
        }
        if (session.docxIndex === undefined) throw sessionClosed();
        if (request.index !== 0) {
          throw new AaredError(
            "not-found",
            `aared/core: unit index ${request.index} is outside this single-document file`,
          );
        }
        // Not cached: `aared/docx` reads unit 0 once at open and the page keeps
        // the sections, so the Worker keeps nothing the page already holds.
        unit = await (await formats.docx()).parseDocxSections(session.pkg, session.docxIndex, signal);
      }
      throwIfAborted(signal);
      // Cached units (PPTX slides, XLSX sheets) come back as the same object; walk each one only once.
      if (!mediaCollected.has(unit)) {
        collectMediaIds(unit, session.mediaParts);
        mediaCollected.add(unit);
      }
      postResponse(request.sessionId, request.requestId, unit);
      return;
    }

    if (request.type === "layout") {
      if (session.format !== "pptx" || session.deck === undefined) {
        throw new AaredError("unsupported-feature", "aared/core: layouts are a PPTX-only feature");
      }
      if (!session.pkg.hasPart(request.layoutId) || !session.pkg.contentTypeOf(request.layoutId)?.endsWith("slideLayout+xml")) {
        throw new AaredError("not-found", `aared/core: "${request.layoutId}" is not a slide layout in this presentation`);
      }
      const layout = await (await formats.pptx()).parsePptxLayout(session.pkg, session.deck, request.layoutId, signal);
      throwIfAborted(signal);
      if (!mediaCollected.has(layout)) {
        collectMediaIds(layout, session.mediaParts);
        mediaCollected.add(layout);
      }
      postResponse(request.sessionId, request.requestId, layout);
      return;
    }

    if (!session.mediaParts.has(request.mediaId) || !session.pkg.hasPart(request.mediaId)) {
      throw new AaredError(
        "media-error",
        `aared/core: media "${request.mediaId}" is not available in this ${session.format.toUpperCase()} document`,
      );
    }
    // `readPart` returns a buffer this worker owns outright: transfer it as is.
    const result = (await session.pkg.readPart(request.mediaId, signal)).buffer as ArrayBuffer;
    throwIfAborted(signal);
    postResponse(request.sessionId, request.requestId, result, [result]);
  } catch (error) {
    if (request.type === "open" && installedSession) deleteSession(request.sessionId, request.requestId);
    postFailure(request.sessionId, request.requestId, error);
  }
}

async function drain(): Promise<void> {
  if (draining) return;
  draining = true;
  while (queue.length > 0) {
    const request = queue.shift()!;
    const key = requestKey(request.sessionId, request.requestId);
    if (cancelled.delete(key)) continue;
    activeRequest = key;
    activeController = new AbortController();
    try {
      await handle(request, activeController.signal);
    } finally {
      activeRequest = undefined;
      activeController = undefined;
    }
  }
  draining = false;
}

scope.onmessage = (event) => {
  const request = event.data;
  if (request.type === "cancel") {
    const key = requestKey(request.sessionId, request.requestId);
    const queued = queue.findIndex((candidate) => requestKey(candidate.sessionId, candidate.requestId) === key);
    if (queued !== -1) {
      queue.splice(queued, 1);
      return;
    }
    if (sessionOpenRequests.get(request.sessionId) === request.requestId) {
      deleteSession(request.sessionId, request.requestId);
      removeQueuedSessionRequests(request.sessionId);
    }
    if (activeRequest === key) {
      cancelled.add(key);
      const controller = activeController;
      activeController = undefined;
      controller?.abort();
    }
    return;
  }
  if (request.type === "close") {
    deleteSession(request.sessionId);
    removeQueuedSessionRequests(request.sessionId);
    if (activeRequest?.startsWith(`${request.sessionId}:`) === true) {
      cancelled.add(activeRequest);
      const controller = activeController;
      activeController = undefined;
      controller?.abort();
    }
    return;
  }
  queue.push(request);
  void drain();
};

scope.postMessage({ type: "ready" });
