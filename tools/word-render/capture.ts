/**
 * Render a DOCX through aared's layout mode in a headless WebView and write one
 * 96 dpi PNG per physical page — the same pixel size `index.ts` gives Word's
 * pages, so the two pair directly. Reads the source, so no build is needed.
 *
 *   bun tools/word-render/capture.ts fixtures/docx-coverage.docx tools/word-render/out/bidi
 *
 * Writes aared-page-N.png beside Word's word-page-N.png when pointed at the
 * same directory `index.ts` used. The page here has no host CSS reset; a
 * consumer's reset (Tailwind's `img { display: block }` did bite) is a
 * separate thing to check in the bench.
 */
import { mkdir, writeFile } from "node:fs/promises";
import { join, resolve } from "node:path";
import { serveDocx } from "../../packages/aared/test/docx-harness.js";

const [fixtureArg, outputArg] = process.argv.slice(2);
if (fixtureArg === undefined || outputArg === undefined) {
  throw new Error("Usage: bun tools/word-render/capture.ts <fixture.docx> <output-dir>");
}
const outputDir = resolve(outputArg);
await mkdir(outputDir, { recursive: true });

const { server, sections } = await serveDocx(resolve(fixtureArg), `<!doctype html><style>html,body{margin:0;background:#fff;overflow:hidden}#host{position:absolute;left:0;top:0}</style><main id="host"></main><script type="module">
      import { render } from "/docx.mjs";
      window.paint = async (model) => {
        const handle = await render({
          document: model, target: document.querySelector("#host"), pagination: { mode: "layout" },
          mediaResolver: async (id) => (await fetch("/media/" + id.split("/").at(-1))).arrayBuffer(),
        });
        await document.fonts.ready;
        await new Promise((resolve) => setTimeout(resolve, 400));
        for (const page of document.querySelectorAll('[data-aared-element="page"]')) page.style.boxShadow = "none";
        return handle.pageCount;
      };
      // Slide the host so page n sits at the viewport origin, and report its size.
      window.showPage = (n) => {
        const page = [...document.querySelectorAll('[data-aared-element="page"]')][n];
        const box = page.getBoundingClientRect();
        const host = document.querySelector("#host");
        host.style.left = (parseFloat(host.style.left || "0") - box.left) + "px";
        host.style.top = (parseFloat(host.style.top || "0") - box.top) + "px";
        return { width: Math.round(box.width), height: Math.round(box.height) };
      };
    </script>`);

try {
  await using view = new Bun.WebView({ width: 1200, height: 1200 });
  await view.navigate(server.url.href);
  const pageCount = (await view.evaluate(`paint(${JSON.stringify(sections)})`)) as number;
  for (let index = 0; index < pageCount; index++) {
    const size = (await view.evaluate(`showPage(${index})`)) as { width: number; height: number };
    await view.resize(size.width, size.height);
    await new Promise((resolve) => setTimeout(resolve, 150));
    await writeFile(join(outputDir, `aared-page-${index + 1}.png`), await view.screenshot({ encoding: "buffer", format: "png" }));
    console.log(`page ${index + 1}: ${size.width}x${size.height}`);
  }
} finally {
  server.stop();
}
