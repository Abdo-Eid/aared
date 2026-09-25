import { expect, test } from "bun:test";
import { AaredError, isAbort, isAaredError, messageFor } from "../src/core/index.js";
import { itemAt } from "../src/core/errors.js";

test("isAbort recognizes AaredError with code 'aborted'", () => {
  const error = new AaredError("aborted", "test: cancelled");
  expect(isAbort(error)).toBe(true);
  expect(error.name).toBe("AbortError");
});

test("isAbort recognizes a plain Error named AbortError", () => {
  const error = new Error("cancelled");
  error.name = "AbortError";
  expect(isAbort(error)).toBe(true);
});

test("isAbort recognizes DOMException AbortError", () => {
  const error = new DOMException("aborted", "AbortError");
  expect(isAbort(error)).toBe(true);
});

test("isAbort returns false for non-abort errors", () => {
  expect(isAbort(new AaredError("invalid-document", "bad"))).toBe(false);
  expect(isAbort(new Error("something"))).toBe(false);
  expect(isAbort(null)).toBe(false);
  expect(isAbort(undefined)).toBe(false);
  expect(isAbort("string")).toBe(false);
});

test("messageFor maps known codes to human-readable strings", () => {
  expect(messageFor(new AaredError("encrypted-document", "x"))).toBe("This file is password-protected and cannot be opened");
  expect(messageFor(new AaredError("aborted", "x"))).toBe("The operation was cancelled");
  expect(messageFor(new AaredError("not-found", "x"))).toBe("The requested item was not found");
  expect(messageFor(new AaredError("network-error", "x"))).toBe("A network error occurred while loading the file");
  expect(messageFor(new AaredError("unsupported-format", "x"))).toBe("This file format is not supported");
  expect(messageFor(new AaredError("invalid-document", "x"))).toBe("This file is not a valid or supported document");
  expect(messageFor(new AaredError("media-error", "x"))).toBe("A media resource could not be loaded");
  expect(messageFor(new AaredError("font-error", "x"))).toBe("A required font could not be loaded");
  expect(messageFor(new AaredError("invalid-range", "x"))).toBe("The requested range is invalid");
  expect(messageFor(new AaredError("unsupported-feature", "x"))).toBe("This feature is not supported");
});

test("messageFor returns fallback for unknown errors", () => {
  expect(messageFor(new Error("oops"))).toBe("An unexpected error occurred");
  expect(messageFor(null)).toBe("An unexpected error occurred");
  expect(messageFor(42)).toBe("An unexpected error occurred");
});

test("isAbort narrows to Error so callers read name and message without a cast", () => {
  const error: unknown = new AaredError("aborted", "test: cancelled");
  if (!isAbort(error)) throw new Error("expected an abort");
  expect(error.name).toBe("AbortError");
  expect(error.message).toBe("test: cancelled");
});

test("itemAt returns an in-range item and refuses everything else with the caller's wording", () => {
  const list = ["a", "b"];
  expect(itemAt(list, 0, () => "unused")).toBe("a");
  expect(itemAt(list, 1, () => "unused")).toBe("b");
  for (const index of [-1, 2, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
    let caught: unknown;
    try {
      itemAt(list, index, () => `item ${String(index)} is outside this ${list.length}-item list`);
    } catch (error) {
      caught = error;
    }
    expect(caught).toBeInstanceOf(AaredError);
    expect((caught as AaredError).code).toBe("not-found");
    expect((caught as AaredError).message).toBe(`item ${String(index)} is outside this 2-item list`);
  }
  // An undefined item is out of range too, even at a valid integer index.
  expect(() => itemAt([undefined, "b"], 0, () => "hole")).toThrow("hole");
});

test("isAaredError type guard works", () => {
  const error = new AaredError("invalid-document", "bad");
  expect(isAaredError(error)).toBe(true);
  expect(isAaredError(new Error("bad"))).toBe(false);
});
