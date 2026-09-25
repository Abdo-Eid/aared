import { expect, test } from "bun:test";
import { UnitCache } from "../src/core/unit-cache.js";

test("a unit cache keeps the most recently read units and evicts the oldest", () => {
  const cache = new UnitCache<string>(2);
  cache.set(0, "a");
  cache.set(1, "b");
  expect(cache.get(0)).toBe("a");
  cache.set(2, "c");
  // 1 was the least recent once 0 was read again.
  expect(cache.peek(1)).toBeUndefined();
  expect([cache.peek(0), cache.peek(2), cache.size]).toEqual(["a", "c", 2]);
});

test("peek serves a cached unit without refreshing it", () => {
  const cache = new UnitCache<string>(2);
  cache.set(0, "a");
  cache.set(1, "b");
  expect(cache.peek(0)).toBe("a");
  cache.set(2, "c");
  expect(cache.peek(0)).toBeUndefined();
  expect(cache.peek(1)).toBe("b");
});
