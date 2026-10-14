import { expect, test } from "bun:test";
import { AaredError, isAbort, isAaredError, messageFor } from "../src/core/index.js";
import { itemAt } from "../src/core/errors.js";

test("isAbort recognizes every abort shape, and nothing else", () => {
  const aared = new AaredError("aborted", "test: cancelled");
  expect(isAbort(aared)).toBe(true);
  expect(aared.name).toBe("AbortError");
  const named = new Error("cancelled");
  named.name = "AbortError";
  expect(isAbort(named)).toBe(true);
  expect(isAbort(new DOMException("aborted", "AbortError"))).toBe(true);
  // Once narrowed, a caller reads name and message without a cast.
  const unknown: unknown = aared;
  if (!isAbort(unknown)) throw new Error("expected an abort");
  expect(`${unknown.name}: ${unknown.message}`).toBe("AbortError: test: cancelled");
  for (const other of [new AaredError("invalid-document", "bad"), new Error("something"), null, undefined, "string"]) {
    expect(isAbort(other)).toBe(false);
  }
});

test("messageFor maps codes to human-readable strings", () => {
  expect(messageFor(new AaredError("encrypted-document", "x"))).toBe("This file is password-protected and cannot be opened");
  expect(messageFor(new AaredError("aborted", "x"))).toBe("The operation was cancelled");
  expect(messageFor(new AaredError("invalid-document", "x"))).toBe("This file is not a valid or supported document");
  // Every code has a message: `MESSAGES` is an exhaustive Record, enforced at compile time.
});

test("messageFor returns fallback for unknown errors", () => {
  expect(messageFor(new Error("oops"))).toBe("An unexpected error occurred");
  expect(messageFor(null)).toBe("An unexpected error occurred");
  expect(messageFor(42)).toBe("An unexpected error occurred");
});

test("itemAt returns an in-range item: a whole number outside is not-found, anything else invalid-argument", () => {
  const list = ["a", "b"];
  expect(itemAt("test", list, 0, () => "unused")).toBe("a");
  expect(itemAt("test", list, 1, () => "unused")).toBe("b");
  const failure = (index: number): AaredError => {
    try {
      itemAt("test", list, index, () => `test: item ${String(index)} is outside this ${list.length}-item list`);
    } catch (error) {
      return error as AaredError;
    }
    throw new Error("expected itemAt to throw");
  };
  for (const index of [-1, 2]) {
    const caught = failure(index);
    expect(caught).toBeInstanceOf(AaredError);
    expect(caught.code).toBe("not-found");
    expect(caught.message).toBe(`test: item ${index} is outside this 2-item list`);
  }
  // Not an index at all: the call is wrong, in the caller's scope.
  for (const index of [1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
    const caught = failure(index);
    expect(caught.code).toBe("invalid-argument");
    expect(caught.message).toBe(`test: an index must be a whole number; received ${String(index)}`);
  }
  // An undefined item is out of range too, even at a valid integer index.
  expect(() => itemAt("test", [undefined, "b"], 0, () => "hole")).toThrow("hole");
});

test("isAaredError type guard works", () => {
  const error = new AaredError("invalid-document", "bad");
  expect(isAaredError(error)).toBe(true);
  expect(isAaredError(new Error("bad"))).toBe(false);
});
