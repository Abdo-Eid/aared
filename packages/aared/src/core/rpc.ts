import type { DocumentInfo } from "./index.js";
import type { SlideLayoutModel, UnitModel } from "./model.js";
import type { AaredErrorCode } from "./errors.js";

export interface SerializedError {
  name: string;
  message: string;
  /** Present when the worker threw an `AaredError`; restored on the main thread. */
  code?: AaredErrorCode;
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
   * XLSX-only numeric window (zero-based, inclusive). The worker slices the
   * parsed sheet to this window; other formats reject the request with
   * `unsupported-feature`.
   */
  range?: { start: { row: number; col: number }; end: { row: number; col: number } };
  /**
   * `false`: a read that won't be painted (search). A cached unit is still
   * served, but the result neither enters nor refreshes the session's
   * bounded unit cache.
   */
  retain?: false;
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

export type WorkerRequest = OpenRequest | UnitRequest | DetectRequest | MediaRequest | LayoutRequest | CancelRequest | CloseRequest;

export type WorkerCall =
  | Omit<OpenRequest, "sessionId" | "requestId">
  | Omit<UnitRequest, "sessionId" | "requestId">
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

export type WorkerResult = DocumentInfo | UnitModel | SlideLayoutModel | ArrayBuffer | DocumentInfo["format"];
