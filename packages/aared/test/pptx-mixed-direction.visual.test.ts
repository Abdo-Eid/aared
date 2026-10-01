import { expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide, parseTextBody } from "../src/core/pptx.js";
import { parseXml } from "../src/core/xml.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const COVERAGE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const PRODUCTION = `${import.meta.dir}/../../../fixtures/post-production-review.pptx`;

const FIRST = "واحد";
const SECOND = "اثنان";
const THIRD = "ثالثة";

// Ground truth, PowerPoint's own PDF (probe en-us-arabic-run), left to right on
// the page: an en-US run in a left-to-right paragraph keeps `W1 + W2` and
// `W1 (W2) W3` in written order; the same text tagged ar-SA, `W1 W2 W3` in any
// language, and every right-to-left paragraph read right to left.
visual("visual DOM output: PPTX orders Arabic words around ASCII punctuation as PowerPoint does", async () => {
  const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
  const { theme } = await indexPptx(pkg);
  const cases = [
    { id: "en-plus", text: `${FIRST} + ${SECOND}`, lang: "en-US", rtl: false, order: [FIRST, SECOND] },
    { id: "en-paren", text: `${FIRST} (${SECOND}) ${THIRD}`, lang: "en-US", rtl: false, order: [FIRST, SECOND, THIRD] },
    { id: "en-space", text: `${FIRST} ${SECOND} ${THIRD}`, lang: "en-US", rtl: false, order: [THIRD, SECOND, FIRST] },
    { id: "ar-plus", text: `${FIRST} + ${SECOND}`, lang: "ar-SA", rtl: false, order: [SECOND, FIRST] },
    { id: "en-rtl-plus", text: `${FIRST} + ${SECOND}`, lang: "en-US", rtl: true, order: [SECOND, FIRST] },
  ];
  const model = {
    size: { width: 13.333, height: 7.5 },
    elements: cases.map((item, index) => ({
      kind: "shape",
      id: item.id,
      name: item.id,
      bounds: { x: 0.5, y: 0.3 + index * 0.9, width: 8, height: 0.8 },
      text: parseTextBody(
        parseXml(`<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr${item.rtl ? ` rtl="1" algn="l"` : ""}/><a:r><a:rPr lang="${item.lang}" sz="2400"/><a:t>${item.text}</a:t></a:r></a:p></a:txBody>`),
        theme,
      ),
      geometry: { path: "M 0 0 H 100000 V 100000 H 0 Z", viewBox: { width: 100000, height: 100000 } },
    })),
  };
  const build = await Bun.build({ entrypoints: [`${import.meta.dir}/../src/pptx/index.ts`], target: "browser", format: "esm" });
  if (!build.success) throw new Error(build.logs.map((log) => log.message).join("\n"));
  const module = await build.outputs[0]!.text();
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      if (new URL(request.url).pathname === "/pptx.mjs") return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      return new Response(`<!doctype html><style>html,body{margin:0}#host{width:1280px}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        window.renderFixture = async (model) => {
          await render({ slide: model, target: document.querySelector("#host"), zoom: 1 });
          await document.fonts.ready;
        };
        // Where the text node holding exactly this token starts, anywhere on the slide.
        window.tokenX = (token) => {
          const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
          for (let node = walker.nextNode(); node; node = walker.nextNode()) {
            if (node.data.trim() !== token) continue;
            const range = document.createRange();
            range.selectNodeContents(node);
            return range.getBoundingClientRect().left;
          }
          return null;
        };
        // Where each word starts on the page, read from the painted text nodes.
        window.wordX = (id, words) => {
          const shape = document.querySelector('[data-aared-shape-id="' + id + '"]');
          const walker = document.createTreeWalker(shape, NodeFilter.SHOW_TEXT);
          const found = {};
          for (let node = walker.nextNode(); node; node = walker.nextNode()) {
            for (const word of words) {
              const at = node.data.indexOf(word);
              if (at < 0) continue;
              const range = document.createRange();
              range.setStart(node, at); range.setEnd(node, at + word.length);
              found[word] = range.getBoundingClientRect().left;
            }
          }
          return found;
        };
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1280, height: 720 });
    await view.navigate(server.url.href);
    await view.evaluate(`renderFixture(${JSON.stringify(model)})`);
    for (const item of cases) {
      const x = await view.evaluate(`wordX(${JSON.stringify(item.id)}, ${JSON.stringify([FIRST, SECOND, THIRD])})`) as Record<string, number>;
      const leftToRight = item.order.map((word) => x[word]!);
      expect(leftToRight.every((value) => Number.isFinite(value)), item.id).toBe(true);
      expect([...leftToRight].sort((a, b) => a - b), item.id).toEqual(leftToRight);
    }

    // The case that motivated it, on the real deck: slide 5's table cell
    // "عربي + إنجليزي" in a left-to-right en-US paragraph reads Arabic, plus, English.
    const production = await OpcPackage.open(await Bun.file(PRODUCTION).arrayBuffer());
    const deck = await indexPptx(production);
    const slide = await parsePptxSlide(production, deck, 4);
    await view.evaluate(`renderFixture(${JSON.stringify(slide)})`);
    const arabic = await view.evaluate(`tokenX("عربي")`) as number | null;
    const english = await view.evaluate(`tokenX("إنجليزي")`) as number | null;
    expect(arabic).not.toBeNull();
    expect(english).not.toBeNull();
    expect(arabic!).toBeLessThan(english!);
  } finally {
    server.stop();
  }
});
