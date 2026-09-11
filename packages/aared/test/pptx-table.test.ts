import { expect, test } from "bun:test";
import { tableCellJustifyContent } from "../src/pptx/table.js";

test("maps table cell vertical alignment to flex justification", () => {
  expect(tableCellJustifyContent("top")).toBe("flex-start");
  expect(tableCellJustifyContent("middle")).toBe("center");
  expect(tableCellJustifyContent("bottom")).toBe("flex-end");
});
