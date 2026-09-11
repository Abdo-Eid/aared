import type { DocumentInfo } from "./index.js";
import type { UnitModel } from "./model.js";
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

export interface CancelRequest {
  type: "cancel";
  sessionId: number;
  requestId: number;
}

export interface CloseRequest {
  type: "close";
  sessionId: number;
}

export type WorkerRequest = OpenRequest | UnitRequest | DetectRequest | MediaRequest | CancelRequest | CloseRequest;

export type WorkerCall =
  | Omit<OpenRequest, "sessionId" | "requestId">
  | Omit<UnitRequest, "sessionId" | "requestId">
  | Omit<DetectRequest, "sessionId" | "requestId">
  | Omit<MediaRequest, "sessionId" | "requestId">;

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

export type WorkerResponse = ReadyResponse | SuccessResponse<DocumentInfo | UnitModel | ArrayBuffer | DocumentInfo["format"]> | FailureResponse;

export type WorkerResult = DocumentInfo | UnitModel | ArrayBuffer | DocumentInfo["format"];
