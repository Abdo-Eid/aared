/**
 * Shared typed errors: one error model across `aared/pptx`, `aared/xlsx`
 * and `aared/docx`, so the same code means the same thing in every
 * format entrypoint. Headless-safe (no DOM, no Worker globals).
 */

export type AaredErrorCode =
  | "unsupported-format"
  | "invalid-document"
  | "encrypted-document"
  | "unsupported-feature"
  | "network-error"
  | "media-error"
  | "font-error"
  | "not-found"
  | "invalid-range"
  | "aborted";

/**
 * The single error class thrown by every public `aared` API. `code` is
 * stable and documented; `message` stays human-readable and always
 * carries an `aared/<entrypoint>:` prefix naming the layer that failed.
 *
 * Cancellation surfaces as an `AaredError` with `code: "aborted"` and
 * `name: "AbortError"`, so existing `AbortSignal` handling (`instanceof
 * DOMException` checks aside) keeps working while gaining a typed code.
 */
export class AaredError extends Error {
  readonly code: AaredErrorCode;

  constructor(code: AaredErrorCode, message: string, options?: { cause?: unknown }) {
    super(message, options === undefined ? undefined : { cause: options.cause });
    this.name = code === "aborted" ? "AbortError" : "AaredError";
    this.code = code;
  }
}

/** Rebuild a worker-serialized `AaredError` on the main thread. */
export function aaredErrorFrom(code: AaredErrorCode, message: string): AaredError {
  return new AaredError(code, message);
}

/** Type guard for `catch` sites that branch on `error.code`. */
export function isAaredError(error: unknown): error is AaredError {
  return error instanceof AaredError;
}
