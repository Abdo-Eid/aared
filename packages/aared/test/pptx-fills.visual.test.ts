import { expect, test } from "bun:test";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { OpcPackage } from "../src/core/opc.js";
import type { ShapeModel, SlideModel } from "../src/core/model.js";
import { findCoverageSlide } from "./fixture.js";
import { browserModule } from "./harness.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

/** Every media id a model names, with its bytes, so the page can resolve picture fills. */
async function mediaOf(pkg: OpcPackage, models: unknown[]): Promise<Record<string, string>> {
  const media: Record<string, string> = {};
  for (const match of JSON.stringify(models).matchAll(/"mediaId":"([^"]+)"/g)) {
    media[match[1]!] ??= Buffer.from(await pkg.readPart(match[1]!)).toString("base64");
  }
  return media;
}

visual("visual DOM output: pattern, picture and rect-gradient fills paint on shapes, backgrounds and table cells", async () => {
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const deck = await indexPptx(pkg);
  const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
  const [{ slide: shapes }, { slide: patternBackground }, { slide: tiledBackground }, { slide: table }] = await Promise.all([
    findCoverageSlide(slides, "shape-fills"),
    findCoverageSlide(slides, "pattern-background"),
    findCoverageSlide(slides, "tiled-background"),
    findCoverageSlide(slides, "table-fills"),
  ]);
  const media = await mediaOf(pkg, [shapes, tiledBackground]);
  const module = await browserModule(`${import.meta.dir}/../src/pptx/index.ts`);
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      if (new URL(request.url).pathname === "/pptx.mjs") {
        return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      return new Response(`<!doctype html><style>html,body{margin:0}#host{width:1280px;height:720px}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        const decode = (b64) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)).buffer;
        window.renderFixture = async (model, media) => {
          window.rendered?.destroy();
          window.rendered = await render({
            slide: model,
            target: document.querySelector("#host"),
            ...(media === undefined ? {} : { mediaResolver: async (id) => decode(media[id]) }),
          });
          const deadline = Date.now() + 8000;
          while (document.querySelector("[data-aared-fill-loading]") !== null && Date.now() < deadline) {
            await new Promise((resolve) => setTimeout(resolve, 25));
          }
        };
        // The paint server a shape's main path references, and its size in CSS px.
        window.serverOf = (svg) => {
          const path = [...svg.querySelectorAll(":scope > path")].find((p) => p.getAttribute("fill")?.startsWith("url("));
          const id = path?.getAttribute("fill")?.slice(5, -1);
          const server = id === undefined ? null : svg.querySelector("#" + CSS.escape(id));
          const [, , vbWidth, vbHeight] = (svg.getAttribute("viewBox") ?? "0 0 1 1").split(" ").map(Number);
          const box = svg.getBoundingClientRect();
          return { server, unitsPerPx: vbWidth / box.width, unitsPerPy: vbHeight / box.height };
        };
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1280, height: 720 });
    await view.navigate(server.url.href);

    await view.evaluate(`renderFixture(${JSON.stringify(shapes)}, ${JSON.stringify(media)})`);
    const idOf = (name: string): string => {
      const shape = shapes.elements.find((element) => element.kind === "shape" && element.name === name);
      if (shape === undefined) throw new Error(`fixture structure changed: expected shape "${name}"`);
      return shape.id;
    };
    const shapeFills = await view.evaluate(`((ids) => Object.fromEntries(Object.entries(ids).map(([name, id]) => {
      const svg = document.querySelector('[data-aared-shape-id="' + id + '"] svg[data-aared-element="shape-geometry"]');
      const { server, unitsPerPx, unitsPerPy } = serverOf(svg);
      const rects = server ? [...server.querySelectorAll(":scope > rect")] : [];
      const image = server?.querySelector("image");
      return [name, {
        tag: server?.tagName ?? null,
        widthPx: server ? Number(server.getAttribute("width")) / unitsPerPx : null,
        heightPx: server ? Number(server.getAttribute("height")) / unitsPerPy : null,
        fills: [...new Set(rects.map((rect) => rect.getAttribute("fill")))],
        imageLoaded: image ? image.getAttribute("href")?.startsWith("blob:") === true : null,
        polygons: server ? server.querySelectorAll("polygon").length : 0,
      }];
    })))(${JSON.stringify({
      dkDnDiag: idOf("Pattern DkDnDiag"),
      smGrid: idOf("Pattern SmGrid"),
      stretch: idOf("Picture Stretch"),
      tile: idOf("Picture Tile"),
      corner: idOf("Gradient Rect Corner"),
      center: idOf("Gradient Rect Center"),
    })})`) as Record<string, { tag: string | null; widthPx: number | null; heightPx: number | null; fills: string[]; imageLoaded: boolean | null; polygons: number }>;
    // Patterns: an 8×8 px tile in the authored colours.
    expect(shapeFills.dkDnDiag).toMatchObject({ tag: "pattern", fills: ["#FFFFFF", "#156082"] });
    expect(shapeFills.dkDnDiag!.widthPx).toBeCloseTo(8, 1);
    expect(shapeFills.dkDnDiag!.heightPx).toBeCloseTo(8, 1);
    expect(shapeFills.smGrid).toMatchObject({ tag: "pattern", fills: ["#FFF4E0", "#E8A33D"] });
    // Pictures: a loaded image; the tile is the 1400 px image at 25 %.
    expect(shapeFills.stretch).toMatchObject({ tag: "pattern", imageLoaded: true });
    expect(shapeFills.tile).toMatchObject({ tag: "pattern", imageLoaded: true });
    expect(shapeFills.tile!.widthPx).toBeCloseTo(1400 * 0.25, 0);
    // Rect gradients: the corner focus leaves two sides, the centred one four.
    expect(shapeFills.corner).toMatchObject({ tag: "pattern", polygons: 2 });
    expect(shapeFills.center).toMatchObject({ tag: "pattern", polygons: 4 });
    expect(await view.evaluate(`document.querySelectorAll("[data-aared-fill-fallback]").length`) as number).toBe(0);

    // Backgrounds paint beneath the scene, never over it.
    await view.evaluate(`renderFixture(${JSON.stringify(patternBackground)})`);
    const patterned = await view.evaluate(`(() => {
      const viewport = document.querySelector('[data-aared-element="slide"]');
      const svg = viewport.querySelector('[data-aared-element="slide-background"]');
      const pattern = svg?.querySelector("pattern");
      return {
        kind: viewport.dataset.aaredBackground,
        first: viewport.firstElementChild === svg,
        cellPx: pattern ? Number(pattern.getAttribute("width")) : null,
      };
    })()`);
    expect(patterned).toEqual({ kind: "pattern", first: true, cellPx: 8 });
    await view.evaluate(`renderFixture(${JSON.stringify(tiledBackground)}, ${JSON.stringify(media)})`);
    const tiled = await view.evaluate(`(() => {
      const viewport = document.querySelector('[data-aared-element="slide"]');
      const pattern = viewport.querySelector('[data-aared-element="slide-background"] pattern');
      return { kind: viewport.dataset.aaredBackground, tilePx: pattern ? Math.round(Number(pattern.getAttribute("width"))) : null };
    })()`);
    expect(tiled).toEqual({ kind: "picture", tilePx: Math.round(1400 * 0.3) });

    // Table cells: non-solid fills paint in a layer under the cell text.
    await view.evaluate(`renderFixture(${JSON.stringify(table)})`);
    const cells = await view.evaluate(`[...document.querySelectorAll('[data-aared-element="table-cell"]')].map((cell) => {
      const svg = cell.querySelector('[data-aared-element="cell-fill"]');
      const id = svg?.querySelector(":scope > rect")?.getAttribute("fill")?.slice(5, -1);
      return {
        text: cell.textContent.trim(),
        server: id ? svg.querySelector("#" + CSS.escape(id))?.tagName ?? null : null,
        solid: getComputedStyle(cell).backgroundColor,
        textAbove: svg ? getComputedStyle(svg).zIndex === "-1" && getComputedStyle(cell).isolation === "isolate" : null,
      };
    })`) as Array<{ text: string; server: string | null; solid: string; textAbove: boolean | null }>;
    const cell = (text: string) => cells.find((candidate) => candidate.text === text);
    expect(["البند", "الحالة", "الملاحظات"].map((text) => cell(text)?.server)).toEqual(["linearGradient", "linearGradient", "linearGradient"]);
    expect(cell("نقش")).toMatchObject({ server: "pattern", textAbove: true });
    expect(cell("تدرج")).toMatchObject({ server: "pattern", textAbove: true });
    expect(cell("تصميم")).toMatchObject({ server: null, solid: "rgb(204, 210, 216)" });
  } finally {
    server.stop();
  }
});

visual("visual DOM output: fills that cannot paint are marked, never silently dropped", async () => {
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const deck = await indexPptx(pkg);
  const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
  const { slide } = await findCoverageSlide(slides, "shape-fills");
  const { slide: background } = await findCoverageSlide(slides, "pattern-background");
  // Real fixture shapes with their fill swapped for one the renderer cannot paint.
  const pattern = slide.elements.find((element): element is ShapeModel => element.kind === "shape" && element.name === "Pattern SmGrid")!;
  pattern.fill = { kind: "unsupported", source: "gradient-path" };
  const unmarkedBackground: SlideModel = { ...background, background: { kind: "unsupported", source: "gradient-path" } };
  const module = await browserModule(`${import.meta.dir}/../src/pptx/index.ts`);
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      if (new URL(request.url).pathname === "/pptx.mjs") {
        return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      return new Response(`<!doctype html><main id="host" style="width:1280px;height:720px"></main><script type="module">
        import { render } from "/pptx.mjs";
        window.renderFixture = async (model) => {
          window.rendered?.destroy();
          window.rendered = await render({ slide: model, target: document.querySelector("#host") });
        };
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1280, height: 720 });
    await view.navigate(server.url.href);
    // No media resolver: the picture fills cannot paint either.
    await view.evaluate(`renderFixture(${JSON.stringify(slide)})`);
    const marked = await view.evaluate(`((ids) => ids.map((id) => {
      const host = document.querySelector('[data-aared-shape-id="' + id + '"]');
      const path = host.querySelector('svg[data-aared-element="shape-geometry"] > path');
      return { reason: host.dataset.aaredFillFallback ?? null, fill: path.getAttribute("fill"), dashed: path.hasAttribute("stroke-dasharray") };
    }))(${JSON.stringify([pattern.id, slide.elements.find((element) => element.kind === "shape" && element.name === "Picture Stretch")!.id])})`);
    expect(marked).toEqual([
      { reason: "gradient-path", fill: "none", dashed: true },
      { reason: "missing-resolver", fill: "none", dashed: true },
    ]);
    await view.evaluate(`renderFixture(${JSON.stringify(unmarkedBackground)})`);
    const backgroundMark = await view.evaluate(`(() => {
      const viewport = document.querySelector('[data-aared-element="slide"]');
      return {
        kind: viewport.dataset.aaredBackground,
        source: viewport.dataset.aaredBackgroundSource,
        marker: viewport.querySelector(':scope > [data-aared-element="fill-fallback"]') !== null,
      };
    })()`);
    expect(backgroundMark).toEqual({ kind: "unsupported", source: "gradient-path", marker: true });
  } finally {
    server.stop();
  }
});
