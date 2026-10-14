import { expect, test } from "bun:test";
test("react entrypoint exposes the preview hooks without a DOM", async () => {
  expect(typeof document).toBe("undefined");
  const module = await import("../src/react/index.js");
  for (const hook of [module.usePreview, module.usePptx, module.useXlsx, module.useDocx]) {
    expect(typeof hook).toBe("function");
  }
});
