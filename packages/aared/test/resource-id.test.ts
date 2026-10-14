import { expect, test } from "bun:test";
import { resourceId } from "../src/resource-id.js";

test("resource names are CSS-safe, distinct and independent of source Office IDs", () => {
  const names = Array.from({ length: 1000 }, () => resourceId("gradient"));
  expect(new Set(names).size).toBe(names.length);
  expect(names.every((name) => /^aared-gradient-[a-z0-9-]+$/.test(name))).toBe(true);
  expect(resourceId("search")).not.toBe(names[0]);
});
