/**
 * Render a DOCX through aared's layout mode in a headless WebView and write one
 * 96 dpi PNG per physical page — the same pixel size `index.ts` gives Word's
 * pages, so the two pair directly. Reads the source model and serves the
 * normal package build; run `bun run build` first.
 *
 *   bun tools/word-render/capture.ts fixtures/docx-coverage.docx tools/word-render/out/bidi
 *   bun tools/word-render/capture.ts fixtures/docx-coverage.docx tools/word-render/out/bidi --absent-fonts
 *
 * Writes aared-page-N.png beside Word's word-page-N.png when pointed at the
 * same directory `index.ts` used. With `--absent-fonts` it writes
 * aared-absent-page-N.png instead: the document as a device without Office's
 * fonts paints it, through the real aared-fonts resolver (network; see
 * packages/aared/test/absent-fonts.ts). The page here has no host CSS reset; a
 * consumer's reset (Tailwind's `img { display: block }` did bite) is a
 * separate thing to check in the bench.
 */
import { mkdir, writeFile } from "node:fs/promises";
import { join, resolve } from "node:path";
import { absentFontsBrowserModule, serveDocxFromDist } from "./browser-server.js";

const [fixtureArg, outputArg] = process.argv.slice(2).filter((arg) => !arg.startsWith("--"));
const absent = process.argv.includes("--absent-fonts");
if (fixtureArg === undefined || outputArg === undefined) {
  throw new Error("Usage: bun tools/word-render/capture.ts <fixture.docx> <output-dir> [--absent-fonts]");
}
const outputDir = resolve(outputArg);
await mkdir(outputDir, { recursive: true });

const extraModules = absent ? { "/absent-fonts.mjs": await absentFontsBrowserModule() } : {};
const { server, sections } = await serveDocxFromDist(resolve(fixtureArg), `<!doctype html><style>html,body{margin:0;background:#fff;overflow:hidden}#host{position:absolute;left:0;top:0}</style><main id="host"></main><script type="module">
      import { render } from "/docx.mjs";
      ${absent ? 'import "/absent-fonts.mjs";' : ""}
      window.paint = async (model) => {
        ${absent ? "model = await window.simulateAbsentFonts(model);" : ""}
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
      // Export one position for each whitespace-delimited text token. The
      // matching Word export is tokenized by PyMuPDF's word boxes, so wrapped
      // runs and repeated words remain individually observable.
      window.collectWords = () => {
        const out = [];
        [...document.querySelectorAll('[data-aared-element="page"]')].forEach((page, i) => {
          const walker = document.createTreeWalker(page, NodeFilter.SHOW_TEXT);
          for (let node; (node = walker.nextNode());) {
            const text = node.textContent;
            if (!text) continue;
            const tokens = /\\S+/gu;
            for (let match; (match = tokens.exec(text));) {
              const range = document.createRange();
              range.setStart(node, match.index);
              range.setEnd(node, match.index + match[0].length);
              const box = range.getBoundingClientRect();
              if (box.width === 0 || box.height === 0) continue;
              const pageBox = page.getBoundingClientRect();
              out.push({ page: i + 1, text: match[0], x0: box.left - pageBox.left, y0: box.top - pageBox.top, x1: box.right - pageBox.left, y1: box.bottom - pageBox.top });
            }
          }
        });
        return out;
      };
    </script>`, extraModules);

try {
  await using view = new Bun.WebView({ width: 1200, height: 1200 });
  await view.navigate(server.url.href);
  const pageCount = (await view.evaluate(`paint(${JSON.stringify(sections)})`)) as number;
  const words = (await view.evaluate("collectWords()")) as Array<{ page: number; text: string; x0: number; y0: number; x1: number; y1: number }>;
  await writeFile(join(outputDir, "aared-words.json"), JSON.stringify(words));
  for (let index = 0; index < pageCount; index++) {
    const size = (await view.evaluate(`showPage(${index})`)) as { width: number; height: number };
    await view.resize(size.width, size.height);
    await new Promise((resolve) => setTimeout(resolve, 150));
    await writeFile(join(outputDir, `aared-${absent ? "absent-" : ""}page-${index + 1}.png`), await view.screenshot({ encoding: "buffer", format: "png" }));
    console.log(`page ${index + 1}: ${size.width}x${size.height}`);
  }
} finally {
  server.stop();
}
