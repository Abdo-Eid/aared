import { expect, test } from "bun:test";
import { useDocx, usePreview, usePptx, useXlsx } from "../src/react/index.js";

test("react entrypoint exposes the preview hooks", () => {
  for (const hook of [usePreview, usePptx, useXlsx, useDocx]) {
    expect(typeof hook).toBe("function");
  }
});

test("react entrypoint is import-time safe (no DOM or Worker globals touched)", async () => {
  expect(typeof document).toBe("undefined");
  const module = await import("../src/react/index.js");
  expect(typeof module.usePreview).toBe("function");
  expect(typeof module.usePptx).toBe("function");
  expect(typeof module.useXlsx).toBe("function");
  expect(typeof module.useDocx).toBe("function");
});
