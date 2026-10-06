import type { DocumentInfo } from "./index.js";
import type { SlideLayoutModel, UnitModel } from "./model.js";
import type { AaredErrorCode } from "./errors.js";

/** Narrow numeric window (zero-based, inclusive) for an XLSX `unit` request; other formats reject it. */
export interface UnitRange {
  start: { row: number; col: number };
  end: { row: number; col: number };
}

export interface SerializedError {
  message: string;
  /** Always set: the Worker types every failure before posting it; restored on the main thread. */
  code: AaredErrorCode;
  /** For `internal` failures, the original error, rebuilt as `cause` so a bug report has its name and stack. */
  cause?: { name: string; message: string; stack?: string };
}

export interface OpenRequest {
  type: "open";
  sessionId: number;
  requestId: number;
  bytes: ArrayBuffer;
}

export interface UnitRequest {
  type: "unit";
  sessionId: number;
  requestId: number;
  index: number;
  /**
   * XLSX-only. The worker slices the parsed sheet to this window; other
   * formats reject the request with `unsupported-feature`.
   */
  range?: UnitRange;
}

/**
 * Search's read of one unit: its words, never its model. PPTX answers with a
 * slide's paragraphs (`string[]`) read straight from the slide XML; XLSX with a
 * sheet's cell text (`SheetText`). Nothing it reads enters the Worker's unit
 * cache, so a first search over every unit doesn't push out the ones being
 * viewed. DOCX rejects it: its page already holds the whole document.
 */
export interface TextRequest {
  type: "text";
  sessionId: number;
  requestId: number;
  index: number;
}

/**
 * A sheet's searchable text, in parallel arrays (row-major cells, then each
 * chart's words at its anchor cell). `chart` names the chart a word came from,
 * `null` for a cell.
 */
export interface SheetText {
  rows: number[];
  cols: number[];
  text: string[];
  chart: Array<string | null>;
}

/** Format sniffing for `<aared-viewer>`: resolves the format without keeping a session. */
export interface DetectRequest {
  type: "detect";
  sessionId: number;
  requestId: number;
  bytes: ArrayBuffer;
}

export interface MediaRequest {
  type: "media";
  sessionId: number;
  requestId: number;
  mediaId: string;
}

/**
 * PPTX only: the shared graphics of one layout (and its master), by the
 * `layoutId` a slide carries. Requested once per layout per session.
 */
export interface LayoutRequest {
  type: "layout";
  sessionId: number;
  requestId: number;
  layoutId: string;
}

export interface CancelRequest {
  type: "cancel";
  sessionId: number;
  requestId: number;
}

export interface CloseRequest {
  type: "close";
  sessionId: number;
}

export type WorkerRequest = OpenRequest | UnitRequest | TextRequest | DetectRequest | MediaRequest | LayoutRequest | CancelRequest | CloseRequest;

export type WorkerCall =
  | Omit<OpenRequest, "sessionId" | "requestId">
  | Omit<UnitRequest, "sessionId" | "requestId">
  | Omit<TextRequest, "sessionId" | "requestId">
  | Omit<DetectRequest, "sessionId" | "requestId">
  | Omit<MediaRequest, "sessionId" | "requestId">
  | Omit<LayoutRequest, "sessionId" | "requestId">;

export interface ReadyResponse {
  type: "ready";
}

export interface SuccessResponse<T> {
  type: "response";
  sessionId: number;
  requestId: number;
  ok: true;
  result: T;
}

export interface FailureResponse {
  type: "response";
  sessionId: number;
  requestId: number;
  ok: false;
  error: SerializedError;
}

export type WorkerResponse = ReadyResponse | SuccessResponse<WorkerResult> | FailureResponse;

export type WorkerResult = DocumentInfo | UnitModel | SlideLayoutModel | ArrayBuffer | DocumentInfo["format"] | string[] | SheetText;
