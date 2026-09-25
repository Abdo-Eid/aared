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
  | "worker-error"
  | "closed"
  | "aborted";

/**
 * The single error class thrown by every public `aared` API. `code` is
 * stable and documented; `message` stays human-readable and always
 * carries an `aared/<entrypoint>:` prefix naming the layer that failed.
 *
 * Cancellation surfaces as an `AaredError` with `code: "aborted"` and
 * `name: "AbortError"`, so existing `AbortSignal` handling (`instanceof
 * DOMException` checks aside) keeps working while gaining a typed code.
 * Using a document after `close()` is a usage error, not a cancellation:
 * it is `code: "closed"`, which `isAbort()` never swallows.
 */
export class AaredError extends Error {
  readonly code: AaredErrorCode;

  constructor(code: AaredErrorCode, message: string, options?: { cause?: unknown }) {
    super(message, options === undefined ? undefined : { cause: options.cause });
    this.name = code === "aborted" ? "AbortError" : "AaredError";
    this.code = code;
  }
}

/** Internal cancellation error shared by the worker and the main-thread session. */
export function abortError(): AaredError {
  return new AaredError("aborted", "aared/core: request was cancelled");
}

/** Stop between synchronous parsing/decompression stages. */
export function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted === true) throw abortError();
}

/**
 * `list[index]` for a collection's `get()`: an integer inside the list, else a
 * `not-found` error carrying the caller's own wording (each unit names itself).
 */
export function itemAt<T>(list: readonly T[], index: number, outside: () => string): T {
  const item = list[index];
  if (!Number.isInteger(index) || item === undefined) throw new AaredError("not-found", outside());
  return item;
}

/** Type guard for `catch` sites that branch on `error.code`. */
export function isAaredError(error: unknown): error is AaredError {
  return error instanceof AaredError;
}

/**
 * Canonical abort check: returns `true` for any error representing
 * cancellation — `DOMException` `AbortError`, typed `AaredError("aborted")`,
 * or a plain `Error` named `"AbortError"` (the form `aared/core` emits).
 *
 * Use this instead of `error instanceof DOMException && error.name === "AbortError"`,
 * which misses typed `AaredError` aborts on the main thread.
 */
export function isAbort(error: unknown): error is Error {
  return error instanceof Error && error.name === "AbortError";
}

/**
 * Human-readable message for a known `AaredError` code. Returns a
 * fallback string for unknown codes. Intended for UI display or logging;
 * the returned string never includes the `aared/...` scope prefix.
 */
const MESSAGES: Record<AaredErrorCode, string> = {
  "unsupported-format": "This file format is not supported",
  "invalid-document": "This file is not a valid or supported document",
  "encrypted-document": "This file is password-protected and cannot be opened",
  "unsupported-feature": "This feature is not supported",
  "network-error": "A network error occurred while loading the file",
  "media-error": "A media resource could not be loaded",
  "font-error": "A required font could not be loaded",
  "not-found": "The requested item was not found",
  "invalid-range": "The requested range is invalid",
  "worker-error": "The document worker failed to start or stopped unexpectedly",
  "closed": "The document was closed",
  "aborted": "The operation was cancelled",
};

export function messageFor(error: unknown): string {
  if (isAaredError(error)) return MESSAGES[error.code];
  return "An unexpected error occurred";
}
