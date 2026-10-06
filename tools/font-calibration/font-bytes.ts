// Substitute font bytes a real render fetches: every face registered by the resolver is recorded
// with its file, and the ones that left "unloaded" are what the browser downloaded. Four runs per
// document: this machine's fonts or none of Office's (the absent-fonts simulation), and faces
// loaded lazily (now) or all up front (before).
import { OpcPackage } from "../../../packages/aared/src/core/opc.ts";
import { indexPptx, parsePptxSlide } from "../../../packages/aared/src/core/pptx.ts";
import { indexDocx, parseDocxSections } from "../../../packages/aared/src/core/docx/document.ts";
import { readWorkbook } from "../../../packages/aared/src/core/xlsx/workbook.ts";
import { parseXlsxSheet } from "../../../packages/aared/src/core/xlsx/sheet.ts";

const F = `${import.meta.dir}/../../../fixtures`;
const open = async (name: string) => OpcPackage.open(await Bun.file(`${F}/${name}`).arrayBuffer());
const deck = await open("pptx-coverage.pptx");
const deckIndex = await indexPptx(deck);
const review = await open("post-production-review.pptx");
const reviewIndex = await indexPptx(review);
const doc = await open("docx-coverage.docx");
const book = await open("xlsx-coverage.xlsx");
const workbook = await readWorkbook(book);
const CASES: Array<{ name: string; format: string; model: unknown }> = [
  { name: "pptx-coverage slide 1 (Aptos + Arial)", format: "pptx", model: await parsePptxSlide(deck, deckIndex, 0) },
  { name: "post-production-review slide 1", format: "pptx", model: await parsePptxSlide(review, reviewIndex, 0) },
  { name: "docx-coverage (whole document)", format: "docx", model: await parseDocxSections(doc, await indexDocx(doc)) },
  { name: "xlsx-coverage sheet 1", format: "xlsx", model: await parseXlsxSheet(book, workbook, 0) },
];

const build = await Bun.build({ entrypoints: [`${import.meta.dir}/font-bytes-entry.ts`], target: "browser", format: "esm" });
if (!build.success) throw new Error(build.logs.map((log) => log.message).join("\n"));
const bundle = await build.outputs[0]!.text();
const server = Bun.serve({ port: 0, fetch(request) {
  if (new URL(request.url).pathname === "/entry.mjs") return new Response(bundle, { headers: { "content-type": "text/javascript" } });
  return new Response(`<!doctype html><meta charset=utf-8><style>html,body{margin:0}#host{width:1280px;height:720px}</style><main id="host"></main><script type="module" src="/entry.mjs"></script>`, { headers: { "content-type": "text/html" } });
}});
const sizes = new Map<string, number>();
const bytes = async (urls: string[]) => {
  let total = 0;
  for (const url of urls) {
    if (!sizes.has(url)) sizes.set(url, (await (await fetch(url)).arrayBuffer()).byteLength);
    total += sizes.get(url)!;
  }
  return total;
};
try {
  for (const c of CASES) {
    const row: string[] = [];
    for (const absent of [false, true]) {
      for (const eager of [true, false]) {
        await using view = new Bun.WebView({ width: 1300, height: 800 });
        await view.navigate(server.url.href);
        for (let i = 0; i < 50 && !(await view.evaluate("typeof window.run === 'function'")); i++) await Bun.sleep(100);
        const urls = JSON.parse(await view.evaluate(`run(${JSON.stringify(c.format)}, ${JSON.stringify(c.model)}, ${absent}, ${eager})`) as string) as string[];
        row.push(`${absent ? "no Office fonts" : "this machine"} ${eager ? "up front" : "lazy"}: ${urls.length} files ${(await bytes(urls) / 1024).toFixed(0)} KB`);
      }
    }
    console.log(`== ${c.name}\n   ${row.join("\n   ")}`);
  }
} finally {
  server.stop();
}
