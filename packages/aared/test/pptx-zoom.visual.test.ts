import { expect, test } from "bun:test";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { OpcPackage } from "../src/core/opc.js";
import { browserModule } from "./harness.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

visual("visual DOM output: a slide opens whole — fit-page by default, width when the container has no height", async () => {
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const slide = await parsePptxSlide(pkg, await indexPptx(pkg), 0);
  const module = await browserModule(`${import.meta.dir}/../src/pptx/index.ts`);
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      if (new URL(request.url).pathname === "/pptx.mjs") {
        return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      return new Response(`<!doctype html><style>html,body{margin:0}</style>
        <div id="wide" style="width:1000px;height:300px"></div>
        <div id="auto" style="width:600px"></div>
        <script type="module">
          import { render } from "/pptx.mjs";
          window.measure = async (model, id, zoom) => {
            const target = document.getElementById(id);
            const handle = await render({ slide: model, target });
            if (zoom !== undefined) handle.setZoom(zoom);
            const slide = target.querySelector('[data-aared-element="slide"]').getBoundingClientRect();
            const box = target.getBoundingClientRect();
            const result = {
              zoom: handle.zoom,
              width: Math.round(slide.width),
              height: Math.round(slide.height),
              left: Math.round(slide.left - box.left),
              boxHeight: Math.round(box.height),
            };
            handle.destroy();
            return result;
          };
        </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1280, height: 900 });
    await view.navigate(server.url.href);
    const model = JSON.stringify(slide);
    // A wide, short container: the 16:9 slide fits its 300px height and sits centred.
    const wide = await view.evaluate(`measure(${model}, "wide")`) as { zoom: string; width: number; height: number; left: number };
    expect(wide.zoom).toBe("fit-page");
    expect(wide.height).toBe(300);
    expect(wide.width).toBe(Math.round(300 * 16 / 9));
    expect(Math.abs(wide.left - (1000 - wide.width) / 2)).toBeLessThanOrEqual(1);
    // A container with no height of its own offers nothing to fit: fill its width.
    const auto = await view.evaluate(`measure(${model}, "auto")`) as { width: number; height: number; boxHeight: number };
    expect(auto.width).toBe(600);
    expect(auto.height).toBe(Math.round(600 * 9 / 16));
    // fit-width stays available, and overflows a short container as asked.
    const widthFit = await view.evaluate(`measure(${model}, "wide", "fit-width")`) as { zoom: string; width: number };
    expect(widthFit).toMatchObject({ zoom: "fit-width", width: 1000 });
  } finally {
    server.stop();
  }
});
