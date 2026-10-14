import { expect, test } from "bun:test";

/**
 * Every public entry loads where there is no DOM (a server render, Node, a
 * Next.js client component prerendered on the server): importing touches no
 * `document`, `HTMLElement` or `customElements`. The element registers itself
 * only where the page has a custom-element registry.
 */
test("every public entry imports with no DOM globals (SSR-safe)", async () => {
  // Bun's test runtime has none of these. Checked first: a global leaked from
  // another test file would make the imports below pass without testing anything.
  expect(typeof document).toBe("undefined");
  expect(typeof HTMLElement).toBe("undefined");
  expect(typeof customElements).toBe("undefined");
  for (const entry of ["core", "pptx", "xlsx", "docx", "preview", "react", "element"]) {
    await import(`../src/${entry}/index.js`);
  }
  const { AaredViewer } = await import("../src/element/index.js");
  expect(typeof AaredViewer).toBe("function");
});
