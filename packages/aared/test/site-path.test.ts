import { expect, test } from "bun:test";
import { withSiteBase } from "../../../site/src/lib/site-path.js";

test("site URLs preserve fragments, external links and already-prefixed paths", () => {
  for (const [path, expected] of [
    ["/", "/aared/"], ["/start/#sizing", "/aared/start/#sizing"],
    ["/aared/api/", "/aared/api/"], ["#navigation", "#navigation"],
    ["https://example.com/", "https://example.com/"], ["//example.com/logo.svg", "//example.com/logo.svg"],
    ["./file.xlsx", "./file.xlsx"],
  ]) expect(withSiteBase(path!, "/aared/")).toBe(expected!);
  expect(withSiteBase("/api/", "/")).toBe("/api/");
});
