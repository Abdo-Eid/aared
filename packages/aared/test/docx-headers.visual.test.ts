import { expect } from "bun:test";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { visual, withVisualPage } from "./harness.js";
import { serveDocx } from "./docx-harness.js";
import { fourSectionDocument } from "./docx-headers-doc.js";

// The layout-mode DOM shows, on each physical page, the header Word shows there
// (the table `docx-headers.test.ts` pins), and `PAGE` counts the printed number.
visual("visual DOM output: first-page, even-page and inherited headers paint on the pages Word shows them on", async () => {
  const dir = await mkdtemp(join(tmpdir(), "aared-headers-"));
  const path = join(dir, "headers.docx");
  await writeFile(path, fourSectionDocument());
  const { server, sections } = await serveDocx(path, `<!doctype html><style>html,body{margin:0}#host{position:absolute;left:0;top:0}</style><main id="host"></main><script type="module">
      import { render } from "/docx.mjs";
      window.paint = async (model) => {
        await render({ document: model, target: document.querySelector("#host"), pagination: { mode: "layout" } });
        await document.fonts.ready;
        return [...document.querySelectorAll('[data-aared-element="page"]')].map((p) => (p.querySelector('[data-aared-element="page-header"]')?.textContent ?? "").trim());
      };
    </script>`);
  try {
    await withVisualPage({ server, width: 1000, height: 1200 }, async (view) => {
    const headers = (await view.evaluate(`paint(${JSON.stringify(sections)})`)) as string[];
    expect(headers).toEqual(["A-FIRST", "A-EVEN", "A-DEFAULT", "A-EVEN", "A-DEFAULT", "A-EVEN", "A-FIRST", "A-EVEN", "D-DEFAULT", "A-EVEN", "D-DEFAULT"]);
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}, 60_000);
