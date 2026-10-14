/**
 * Render a DOCX through aared's layout mode in a headless WebView and write one
 * 96 dpi PNG per physical page — the same pixel size the native driver gives Word's
 * pages, so the two pair directly. Reads the source model and serves the
 * normal package build; run `bun run build` first.
 *
 *   bun tools/review/docx/capture.ts fixtures/docx-coverage.docx tools/out/docx/bidi
 *   bun tools/review/docx/capture.ts fixtures/docx-coverage.docx tools/out/docx/bidi --absent-fonts
 *
 * Writes aared-page-N.png beside Word's word-page-N.png when pointed at the
 * same directory `office-export.ts` used. With `--absent-fonts` it writes
 * aared-absent-page-N.png instead: the document as a device without Office's
 * fonts paints it, through the real aared-fonts resolver (network; see
 * packages/aared/test/absent-fonts.ts). The page here has no host CSS reset; a
 * consumer's reset (Tailwind's `img { display: block }` did bite) is a
 * separate thing to check in the bench.
 */
import { mkdir, writeFile } from "node:fs/promises";
import { join, resolve } from "node:path";
import { absentFontsBrowserModule, serveDocxFromDist } from "./browser-server.js";
import { outputDirectory } from "../../shared/office.js";
import { reviewOutputRoot } from "../../shared/paths.js";
import { within } from "../../../packages/aared/test/harness.js";

const argv = process.argv.slice(2);
for (const arg of argv) if (arg.startsWith("--") && arg !== "--absent-fonts") throw new Error(`Unknown Word capture option: ${arg}`);
const [fixtureArg, outputArg, extra] = argv.filter((arg) => !arg.startsWith("--"));
const absent = process.argv.includes("--absent-fonts");
if (fixtureArg === undefined || extra !== undefined) {
  throw new Error("Usage: bun tools/review/docx/capture.ts <fixture.docx> [output-dir] [--absent-fonts]");
}
const outputDir = outputArg === undefined ? outputDirectory(fixtureArg, reviewOutputRoot("docx")) : resolve(outputArg);
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
  await within("loading the Word review host", view.navigate(server.url.href), 30_000);
  const pageCount = (await within("rendering the Word document", view.evaluate(`paint(${JSON.stringify(sections)})`), absent ? 120_000 : 60_000)) as number;
  const words = (await within("reading Word capture tokens", view.evaluate("collectWords()"), 30_000)) as Array<{ page: number; text: string; x0: number; y0: number; x1: number; y1: number }>;
  await writeFile(join(outputDir, `aared-${absent ? "absent-" : ""}words.json`), JSON.stringify(words));
  for (let index = 0; index < pageCount; index++) {
    const size = (await within(`measuring Word page ${index + 1}`, view.evaluate(`showPage(${index})`), 10_000)) as { width: number; height: number };
    await within("sizing the Word viewport", view.resize(size.width, size.height), 10_000);
    await new Promise((resolve) => setTimeout(resolve, 150));
    await writeFile(join(outputDir, `aared-${absent ? "absent-" : ""}page-${index + 1}.png`), await within(`capturing Word page ${index + 1}`, view.screenshot({ encoding: "buffer", format: "png" }), 10_000));
    console.log(`page ${index + 1}: ${size.width}x${size.height}`);
  }
} finally {
  server.stop();
}
