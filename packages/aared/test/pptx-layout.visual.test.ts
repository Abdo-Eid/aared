import { expect } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxLayout, parsePptxSlide } from "../src/core/pptx.js";
import { browserModule, visual, withVisualPage } from "./harness.js";

const COVERAGE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

/**
 * Slides 15 to 18 of the coverage deck sit on a second master (a logo, a picture background) with layout A
 * (shows the master's graphics) and layout B (hides them). The page must paint, back to front: the
 * background, the master's graphics, the layout's, then the slide's own, and leave the shared graphics out
 * of a slide that hides them. Read back from the DOM by each graphic's text.
 */
visual("visual DOM output: master and layout graphics paint under the slide, and showMasterSp=0 leaves them out", async () => {
  const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
  const deck = await indexPptx(pkg);
  const models = await Promise.all([14, 15, 16, 17].map(async (index) => {
    const slide = await parsePptxSlide(pkg, deck, index);
    return { ...slide, layout: slide.layoutId === undefined ? undefined : await parsePptxLayout(pkg, deck, slide.layoutId) };
  }));
  const renderer = await browserModule(`${import.meta.dir}/../src/pptx/index.ts`);
  const page = `<!doctype html><style>html,body{margin:0}#host{width:1000px;height:600px}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        window.paint = async (model) => {
          window.handle?.destroy();
          const mediaResolver = async (id) => (await fetch("/media/" + id)).arrayBuffer();
          window.handle = await render({ slide: model, target: document.querySelector("#host"), mediaResolver });
          const order = [];
          for (const el of document.querySelectorAll("#host *")) {
            const own = [...el.childNodes].filter((n) => n.nodeType === 3).map((n) => n.textContent.trim()).join("");
            if (own !== "") order.push(own);
          }
          const viewport = document.querySelector("#host [data-aared-background]");
          return { order, background: viewport?.dataset.aaredBackground ?? "none", mediaId: viewport?.dataset.aaredMediaId ?? "" };
        };
      </script>`;
  await withVisualPage({
    modules: { "/pptx.mjs": renderer },
    page,
    onRequest: async (url) => {
      if (url.pathname.startsWith("/media/")) {
        return new Response((await pkg.readPart(decodeURIComponent(url.pathname.slice("/media/".length)))).slice().buffer as ArrayBuffer, { headers: { "content-type": "image/png" } });
      }
      return undefined;
    },
    width: 1000,
    height: 700,
  }, async (view) => {
    type Painted = { order: string[]; background: string; mediaId: string };
    const paint = async (index: number): Promise<Painted> => (await view.evaluate(`paint(${JSON.stringify(models[index])})`)) as Painted;
    const first = await paint(0);
    // The master's picture background paints under everything, on every slide on that master (also the hidden one).
    expect(first.background).toBe("picture");
    expect(first.mediaId).toMatch(/^ppt\/media\//);
    // Master logo, then layout A's graphic, then the slide's title: back to front.
    expect(first.order.indexOf("شعار الاستوديو")).toBeGreaterThanOrEqual(0);
    expect(first.order.indexOf("شعار الاستوديو")).toBeLessThan(first.order.indexOf("رسم التخطيط أ"));
    expect(first.order.indexOf("رسم التخطيط أ")).toBeLessThan(first.order.indexOf("شريحة على التخطيط أ"));
    const second = await paint(1);
    expect(second.order).toEqual(expect.arrayContaining(["شعار الاستوديو", "رسم التخطيط أ"]));
    // Layout B hides the master's graphics: its own only.
    const onB = await paint(2);
    expect(onB.order).toContain("رسم التخطيط ب");
    expect(onB.order).not.toContain("شعار الاستوديو");
    // A slide that hides background graphics shows neither layout's nor the master's.
    const onHidden = await paint(3);
    expect(onHidden.order).toContain("شريحة تخفي الرسومات");
    expect(onHidden.order).not.toContain("شعار الاستوديو");
    expect(onHidden.order).not.toContain("رسم التخطيط أ");
    // Hiding background graphics leaves the background itself: the same inherited picture.
    expect(onHidden.background).toBe("picture");
    expect(onHidden.mediaId).toBe(first.mediaId);
  });
}, 30000);
