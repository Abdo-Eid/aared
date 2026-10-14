/**
 * Shared typed errors: one error model across `@aared/viewer/pptx`, `@aared/viewer/xlsx`
 * and `@aared/viewer/docx`, so the same code means the same thing in every
 * format entrypoint. Headless-safe (no DOM, no Worker globals).
 */

export type AaredErrorCode =
  | "unsupported-format"
  | "invalid-document"
  | "encrypted-document"
  | "unsupported-feature"
  | "network-error"
  | "media-error"
  | "not-found"
  | "invalid-range"
  | "invalid-argument"
  | "invalid-state"
  | "worker-error"
  | "internal"
  | "closed"
  | "aborted";

/**
 * The single error class thrown by every public `aared` API. `code` is
 * stable and documented; `message` stays human-readable and always
 * carries an `@aared/viewer/<entrypoint>:` prefix naming the layer that failed.
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

/**
 * The file itself is broken (a required part, element or id missing, a value
 * that is not what the format allows): `invalid-document`, naming what failed,
 * plus the one repair a reader can try. Thrown where the check is, so the code
 * says what happened; an unexpected throw is `internal` instead.
 */
export function damagedDocument(detail: string): AaredError {
  return new AaredError(
    "invalid-document",
    `@aared/viewer/core: ${detail}. The file may be damaged: opening and re-saving it in Office usually repairs it`,
  );
}

/** Internal cancellation error shared by the worker and the main-thread session. */
export function abortError(): AaredError {
  return new AaredError("aborted", "@aared/viewer/core: request was cancelled");
}

/** Stop between synchronous parsing/decompression stages. */
export function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted === true) throw abortError();
}

/**
 * `list[index]` for a collection's `get()`: an integer inside the list. A
 * whole number outside it is `not-found`, in the caller's own wording (each
 * unit names itself); anything that is not a whole number is the call itself
 * being wrong, `invalid-argument`, in `scope` (`@aared/viewer/<entrypoint>`).
 */
export function itemAt<T>(scope: string, list: readonly T[], index: number, outside: () => string): T {
  if (!Number.isInteger(index)) {
    throw new AaredError("invalid-argument", `${scope}: an index must be a whole number; received ${String(index)}`);
  }
  const item = list[index];
  if (item === undefined) throw new AaredError("not-found", outside());
  return item;
}

/** Type guard for `catch` sites that branch on `error.code`. */
export function isAaredError(error: unknown): error is AaredError {
  return error instanceof AaredError;
}

/**
 * Canonical abort check: returns `true` for any error representing
 * cancellation — `DOMException` `AbortError`, typed `AaredError("aborted")`,
 * or a plain `Error` named `"AbortError"` (the form `@aared/viewer/core` emits).
 *
 * Use this instead of `error instanceof DOMException && error.name === "AbortError"`,
 * which misses typed `AaredError` aborts on the main thread.
 */
export function isAbort(error: unknown): error is Error {
  return error instanceof Error && error.name === "AbortError";
}

const MESSAGES: Record<AaredErrorCode, string> = {
  "unsupported-format": "This file format is not supported",
  "invalid-document": "This file is not a valid or supported document",
  "encrypted-document": "This file is password-protected and cannot be opened",
  "unsupported-feature": "This feature is not supported",
  "network-error": "A network error occurred while loading the file",
  "media-error": "A media resource could not be loaded",
  "not-found": "The requested item was not found",
  "invalid-range": "The requested range is invalid",
  "invalid-argument": "The viewer was called with an invalid value",
  "invalid-state": "The viewer is not ready for this yet",
  "worker-error": "The document worker failed to start or stopped unexpectedly",
  "internal": "Something went wrong in the viewer while reading this file",
  "closed": "The document was closed",
  "aborted": "The operation was cancelled",
};

/**
 * A short, generic English line per code, for a UI that wants one sentence
 * without the `@aared/viewer/...` scope prefix; a fallback for anything else. It
 * deliberately drops the detail (`error.message` has it: which index, which
 * part, which HTTP status) and is not localized: a host that shows errors in
 * another language maps `error.code` to its own strings.
 */
export function messageFor(error: unknown): string {
  if (isAaredError(error)) return MESSAGES[error.code];
  return "An unexpected error occurred";
}
