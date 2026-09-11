import { OpcPackage } from "./opc.js";
import { indexPptx, parsePptxSlide } from "./pptx.js";
import type { PptxDeck } from "./pptx.js";
import { indexXlsx, parseXlsxSheet, readWorkbook, sliceSheetToRange } from "./xlsx/index.js";
import type { XlsxWorkbook, XlsxWorkbookIndex } from "./xlsx/index.js";
import { indexDocx, parseDocxSections } from "./docx/document.js";
import type { DocxIndex } from "./docx/document.js";
import { attr, firstChild, textContent } from "./xml.js";
import type { XmlElement } from "./xml.js";
import type { DocumentInfo } from "./index.js";
import type { SheetModel, SheetRangeModel, UnitModel } from "./model.js";
import type { SerializedError, UnitRequest, WorkerRequest, WorkerResponse, WorkerResult } from "./rpc.js";
import { AaredError } from "./errors.js";
import { ZipError, throwIfAborted } from "./zip.js";

interface WorkerScope {
  onmessage: ((event: MessageEvent<WorkerRequest>) => void) | null;
  postMessage(message: WorkerResponse, transfer?: Transferable[]): void;
}

interface FormatSession {
  format: "pptx" | "xlsx" | "docx";
  pkg: OpcPackage;
  deck?: PptxDeck;
  xlsxIndex?: XlsxWorkbookIndex;
  xlsxWorkbook?: Promise<XlsxWorkbook>;
  /** Parsed sheets, cached so range reads slice without re-parsing. */
  xlsxSheets?: Map<number, SheetModel>;
  docxIndex?: DocxIndex;
  /** Parts the media route may serve: seeded from package media/font dirs, extended with emitted mediaIds. */
  mediaParts: Set<string>;
}

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
  if (error instanceof Error && error.name === "AbortError") {
    return { name: error.name, message: error.message, code: "aborted" };
  }
  // ZIP failures are document problems from the consumer's view: a file
  // that is not, or not cleanly, a ZIP archive is not a previewable
  // document. Preserve the `ZipError` message text; only the code changes.
  if (error instanceof ZipError) {
    return { name: "AaredError", message: error.message, code: "invalid-document" };
  }
  if (error instanceof Error) return { name: error.name, message: error.message };
  return { name: "Error", message: `aared/core: worker failed: ${String(error)}` };
}

function sessionClosed(): AaredError {
  return new AaredError("aborted", "aared/core: document session is closed");
}

function formatOf(pkg: OpcPackage, signal?: AbortSignal): "pptx" | "xlsx" | "docx" | undefined {
  for (const part of pkg.partNames()) {
    throwIfAborted(signal);
    switch (pkg.contentTypeOf(part)) {
      case "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml":
        return "pptx";
      case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml":
      case "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml":
        return "xlsx";
      case "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml":
      case "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml":
        return "docx";
    }
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

function headingTitles(index: DocxIndex): string[] {
  const titles: string[] = [];
  const visit = (children: XmlElement[]): void => {
    for (const child of children) {
      if (child.local === "p") {
        const properties = firstChild(child, "pPr");
        const style = properties === undefined ? undefined : firstChild(properties, "pStyle");
        const styleVal = style === undefined ? undefined : attr(style, "val");
        const outline = properties !== undefined && firstChild(properties, "outlineLvl") !== undefined;
        // Style ids read `Heading1`..`Heading9`; style names read `heading 1`.
        if ((styleVal !== undefined && /^heading\s?\d$/i.test(styleVal)) || outline) {
          const text = textContent(child).trim();
          if (text !== "") titles.push(text);
        }
      } else if (child.local === "sdt") {
        for (const wrap of child.children) {
          if (wrap.local === "sdtContent") visit(wrap.children);
        }
      }
    }
  };
  visit(index.body.children);
  return titles.length > 0 ? titles.slice(0, 50) : ["Document"];
}

function infoForDocx(index: DocxIndex): DocumentInfo {
  return {
    format: "docx",
    count: 1,
    titles: headingTitles(index),
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

/** Record every `mediaId` a resolved unit references so the media route can serve it. */
function collectMediaIds(unit: UnitModel, into: Set<string>): void {
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    if (value === null || typeof value !== "object") return;
    for (const [key, child] of Object.entries(value)) {
      if (key === "mediaId" && typeof child === "string" && child !== "") into.add(child);
      else visit(child);
    }
  };
  visit(unit);
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
  if (session.xlsxWorkbook === undefined) session.xlsxWorkbook = readWorkbook(session.pkg, signal);
  return session.xlsxWorkbook;
}

async function xlsxSheetOf(session: FormatSession, index: number, signal: AbortSignal): Promise<SheetModel> {
  if (session.xlsxSheets === undefined) session.xlsxSheets = new Map();
  let sheet = session.xlsxSheets.get(index);
  if (sheet === undefined) {
    const workbook = await xlsxWorkbookOf(session, signal);
    throwIfAborted(signal);
    sheet = await parseXlsxSheet(session.pkg, workbook, index, signal);
    session.xlsxSheets.set(index, sheet);
  }
  return sheet;
}

function checkRange(range: NonNullable<UnitRequest["range"]>): void {
  const { start, end } = range;
  const valid = (value: number, max: number): boolean =>
    Number.isInteger(value) && value >= 0 && value <= max;
  if (
    !valid(start.row, 1_048_575) || !valid(end.row, 1_048_575)
    || !valid(start.col, 16_383) || !valid(end.col, 16_383)
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
        const deck = await indexPptx(pkg, signal);
        throwIfAborted(signal);
        sessions.set(request.sessionId, { format, pkg, deck, mediaParts });
        sessionOpenRequests.set(request.sessionId, request.requestId);
        installedSession = true;
        throwIfAborted(signal);
        postResponse(request.sessionId, request.requestId, infoForPptx(deck));
        return;
      }
      if (format === "xlsx") {
        const xlsxIndex = await indexXlsx(pkg, signal);
        throwIfAborted(signal);
        sessions.set(request.sessionId, { format, pkg, xlsxIndex, mediaParts });
        sessionOpenRequests.set(request.sessionId, request.requestId);
        installedSession = true;
        throwIfAborted(signal);
        postResponse(request.sessionId, request.requestId, infoForXlsx(xlsxIndex));
        return;
      }
      const docxIndex = await indexDocx(pkg, signal);
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
        unit = await parsePptxSlide(session.pkg, session.deck, request.index, signal);
      } else if (session.format === "xlsx") {
        const sheet = await xlsxSheetOf(session, request.index, signal);
        throwIfAborted(signal);
        let windowed: SheetModel | SheetRangeModel = sheet;
        if (request.range !== undefined) {
          checkRange(request.range);
          windowed = sliceSheetToRange(sheet, request.range);
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
        unit = await parseDocxSections(session.pkg, session.docxIndex, signal);
      }
      throwIfAborted(signal);
      collectMediaIds(unit, session.mediaParts);
      postResponse(request.sessionId, request.requestId, unit);
      return;
    }

    if (!session.mediaParts.has(request.mediaId) || !session.pkg.hasPart(request.mediaId)) {
      throw new AaredError(
        "media-error",
        `aared/core: media "${request.mediaId}" is not available in this ${session.format.toUpperCase()} document`,
      );
    }
    const bytes = await session.pkg.readPart(request.mediaId, signal);
    throwIfAborted(signal);
    const result = new Uint8Array(bytes).buffer;
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
