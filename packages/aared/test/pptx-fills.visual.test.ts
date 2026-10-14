import { expect } from "bun:test";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { OpcPackage } from "../src/core/opc.js";
import type { ShapeModel, SlideModel } from "../src/core/model.js";
import { findCoverageSlide } from "./fixture.js";
import { browserModule, visual, withVisualPage } from "./harness.js";
import { withPart } from "./mutate.js";
import { fontAlias } from "../src/core/typography.js";
import { within } from "./harness.js";

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
  await withVisualPage({ server, width: 1280, height: 720 }, async (view) => {

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
  });
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
  await withVisualPage({ server, width: 1280, height: 720 }, async (view) => {
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
  });
});

// Slide 11 (appended through PowerPoint): lines and connectors are `p:cxnSp`.
// Each paints as a stroked path, dashes as authored; a zero-width or
// zero-height box still shows its line.
visual("visual DOM output: connector and line shapes paint with their dash styles", async () => {
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const deck = await indexPptx(pkg);
  const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
  const { slide } = await findCoverageSlide(slides, "lines");
  const module = await browserModule(`${import.meta.dir}/../src/pptx/index.ts`);
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      if (new URL(request.url).pathname === "/pptx.mjs") {
        return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      return new Response(`<!doctype html><style>html,body{margin:0}#host{width:1280px;height:720px}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        window.probe = async (model) => {
          await render({ slide: model, target: document.querySelector("#host") });
          const out = {};
          for (const shape of model.elements.filter((element) => element.kind === "shape" && element.stroke !== undefined)) {
            const host = document.querySelector('[data-aared-shape-id="' + shape.id + '"]');
            const path = host?.querySelector('svg[data-aared-element="shape-geometry"] > path');
            const box = host?.getBoundingClientRect();
            out[shape.name] = {
              stroke: path?.getAttribute("stroke") ?? null,
              fill: path?.getAttribute("fill") ?? null,
              dash: path?.getAttribute("stroke-dasharray") ?? null,
              span: box === undefined ? 0 : Math.max(box.width, box.height),
            };
          }
          // "2.1%" is an en-US run in a right-to-left paragraph: "%" right of the digits.
          const span = [...document.querySelectorAll("span")].find((el) => el.textContent === "2.1%");
          const glyph = (index) => { const range = document.createRange(); range.setStart(span.firstChild, index); range.setEnd(span.firstChild, index + 1); return range.getBoundingClientRect().left; };
          window.ltrRun = { percentRightOfDigits: glyph(3) > glyph(0) };
          // The same box is spaced exactly 30pt and mixes a Calibri run into
          // the Arabic: its one line keeps the pitch (offsetHeight is layout
          // pixels, before the slide's scale).
          const line = span.closest("p");
          const runs = [...line.querySelectorAll("span:not([data-aared-element])")];
          window.exactPitch = {
            lineHeight: getComputedStyle(line).lineHeight,
            height: line.offsetHeight,
            fonts: new Set(runs.map((run) => getComputedStyle(run).fontFamily)).size,
          };
          return out;
        };
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  await withVisualPage({ server, width: 1280, height: 720 }, async (view) => {
    const lines = await view.evaluate(`probe(${JSON.stringify(slide)})`) as Record<string, { stroke: string | null; fill: string | null; dash: string | null; span: number }>;
    expect(await view.evaluate("window.ltrRun") as unknown).toEqual({ percentRightOfDigits: true });
    const pitch = await view.evaluate("window.exactPitch") as { lineHeight: string; height: number; fonts: number };
    expect(pitch.lineHeight).toBe("40px");
    expect(pitch.height).toBe(40);
    expect(pitch.fonts).toBeGreaterThan(1);
    expect(Object.keys(lines).sort()).toEqual(["Diagonal", "Dotted", "Timeline", "Today"]);
    for (const line of Object.values(lines)) {
      expect(line.fill).toBe("none");
      expect(line.span).toBeGreaterThan(100);
    }
    expect(lines.Timeline!.stroke).toBe("#1E3A5A");
    expect(lines.Timeline!.dash).toBeNull();
    expect(lines.Today!.stroke).toBe("#F59B1E");
    expect(lines.Today!.dash).not.toBeNull();
    expect(lines.Dotted!.dash).not.toBeNull();
    const [dotOn, dotOff] = lines.Dotted!.dash!.split(" ").map(Number);
    const [dashOn, dashOff] = lines.Today!.dash!.split(" ").map(Number);
    expect(dotOn).toBe(dotOff);
    expect(dashOn! / dashOff!).toBeCloseTo(4 / 3, 5);
  });
}, 30000);

/**
 * `PptxRenderHandle.warnings` names a family that painted with nothing of
 * its own, the same signal DOCX exposes (Notes §45).
 */
visual("visual DOM output: PptxRenderHandle.warnings names fonts that painted with the generic fallback", async () => {
  const NONEXISTENT_FAMILY = "Definitely Not A Real Font Zzqx";
  const ALIAS = fontAlias(NONEXISTENT_FAMILY);
  const slideWith = (family: string): SlideModel => ({
    part: "ppt/slides/slide1.xml",
    size: { width: 10, height: 7.5 },
    elements: [{
      kind: "shape",
      id: "1",
      name: "Box",
      bounds: { x: 1, y: 1, width: 6, height: 2 },
      text: { paragraphs: [{ runs: [{ text: "Some text", fontFamily: family }] }] },
    } satisfies ShapeModel],
  });
  const renderer = await browserModule(`${import.meta.dir}/../src/pptx/index.ts`);
  const fontBytes = await Bun.file("C:/Windows/Fonts/Cour.ttf").arrayBuffer();
  let releaseFont!: () => void;
  const heldFont = new Promise<void>((resolve) => { releaseFont = resolve; });
  let fontRequests = 0;
  const page = `<!doctype html><style>html,body{margin:0}#host{width:800px;height:600px}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        window.renderSlide = async (model) => {
          window.handle?.destroy();
          window.handle = await render({ slide: model, target: document.querySelector("#host") });
          return window.handle.warnings;
        };
      </script>`;
  const server = Bun.serve({
    port: 0,
    async fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/pptx.mjs") return new Response(renderer, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      if (url.pathname === "/held-font") {
        fontRequests++;
        await heldFont;
        return new Response(fontBytes, { headers: { "content-type": "font/ttf" } });
      }
      return new Response(page, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  await withVisualPage({ server, width: 800, height: 600 }, async (view) => {
    try {
    const realWarnings = (await view.evaluate(`renderSlide(${JSON.stringify(slideWith("Calibri"))})`)) as Array<{ code: string }>;
    expect(realWarnings.filter((w) => w.code === "font-unavailable")).toEqual([]);
    const missingWarnings = (await view.evaluate(`renderSlide(${JSON.stringify(slideWith(NONEXISTENT_FAMILY))})`)) as Array<{ code: string; message: string }>;
    const missing = missingWarnings.find((w) => w.code === "font-unavailable");
    expect(missing).toBeDefined();
    expect(missing!.message).toContain(NONEXISTENT_FAMILY);
    const substituted = await view.evaluate(`(async () => {
      window.substituteFace = new FontFace(${JSON.stringify(ALIAS)}, "local('Courier New')");
      document.fonts.add(window.substituteFace); await window.substituteFace.load();
      return renderSlide(${JSON.stringify(slideWith(NONEXISTENT_FAMILY))});
    })()`) as Array<{ code: string }>;
    expect(substituted.filter((warning) => warning.code === "font-unavailable")).toEqual([]);
    const widths = await view.evaluate(`(() => {
      const span = [...document.querySelectorAll('span')].find(element => element.textContent === 'Some text');
      const style = getComputedStyle(span);
      const measure = family => {
        const control = document.createElement('span');
        control.style.cssText = 'position:absolute;visibility:hidden;white-space:pre;font:' + style.fontSize + ' ' + family;
        control.textContent = span.textContent; document.body.append(control);
        const width = control.getBoundingClientRect().width; control.remove(); return width;
      };
      const scene = document.querySelector('[data-aared-element="slide-scene"]');
      const scale = scene.getBoundingClientRect().width / parseFloat(scene.style.width);
      const range = document.createRange(); range.selectNodeContents(span);
      return { painted: range.getBoundingClientRect().width / scale, courier: measure('"Courier New"'), generic: measure('sans-serif') };
    })()`) as { painted: number; courier: number; generic: number };
    expect(Math.abs(widths.courier - widths.generic)).toBeGreaterThan(5);
    expect(widths.painted).toBeCloseTo(widths.courier, 0);
    // The network response stays withheld until render has returned. This
    // proves the real offline-font bound without sleeping for an assumed load.
    const blocked = await within("PPTX render with a withheld font", view.evaluate(`(async () => {
      document.fonts.delete(window.substituteFace);
      window.heldFace = new FontFace(${JSON.stringify(ALIAS)}, 'url(/held-font)');
      document.fonts.add(window.heldFace);
      const warnings = await renderSlide(${JSON.stringify(slideWith(NONEXISTENT_FAMILY))});
      return { warnings, status: window.heldFace.status };
    })()`), 5000) as { warnings: Array<{ code: string }>; status: string };
    expect(fontRequests).toBe(1);
    expect(blocked.status).toBe("loading");
    expect(blocked.warnings.some((warning) => warning.code === "font-unavailable")).toBe(true);
    releaseFont();
    const recovered = await view.evaluate(`(async () => { await window.heldFace.load(); return renderSlide(${JSON.stringify(slideWith(NONEXISTENT_FAMILY))}); })()`) as Array<{ code: string }>;
    expect(recovered.filter((warning) => warning.code === "font-unavailable")).toEqual([]);
    } finally {
      releaseFont();
    }
  });
}, 30000);

/**
 * Slide 13 of the coverage deck (merged table cells, appended through PowerPoint): the merged cells paint
 * as grid cells spanning two columns and two rows, and the cells they cover add no boxes of their own.
 */
visual("visual DOM output: merged table cells span the grid cells they cover", async () => {
  const source = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const deck = await indexPptx(source);
  const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(source, deck, index)));
  const { slide } = await findCoverageSlide(slides, "merged-cells");
  const renderer = await browserModule(`${import.meta.dir}/../src/pptx/index.ts`);
  const page = `<!doctype html><style>html,body{margin:0}#host{width:1000px;height:600px}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        window.renderSlide = async (model) => {
          window.handle = await render({ slide: model, target: document.querySelector("#host") });
          const grid = [...document.querySelectorAll('[data-aared-element="table"]')].at(0);
          const cells = [...(grid?.children ?? [])].filter((cell) => cell.dataset.aaredElement !== "table-cell-edges" && cell.style.gridColumn !== "");
          return cells.map((cell) => ({ text: (cell.textContent ?? "").trim(), column: cell.style.gridColumn, row: cell.style.gridRow, width: Math.round(cell.getBoundingClientRect().width), height: Math.round(cell.getBoundingClientRect().height) }));
        };
      </script>`;
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      if (new URL(request.url).pathname === "/pptx.mjs") return new Response(renderer, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      return new Response(page, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  await withVisualPage({ server, width: 1000, height: 700 }, async (view) => {
    type Box = { text: string; column: string; row: string; width: number; height: number };
    const cells = (await view.evaluate(`renderSlide(${JSON.stringify(slide)})`)) as Box[];
    // 4 + 4 + 3 + 4 cells in the rows: the covered ones are not painted (14 grid cells for a 4 x 4 table).
    expect(cells).toHaveLength(14);
    const wide = cells.find((cell) => cell.text === "العنوان المدمج")!;
    const tall = cells.find((cell) => cell.text === "جارٍ")!;
    const plain = cells.find((cell) => cell.text === "الحالة")!;
    expect(wide.column).toContain("span 2");
    expect(tall.row).toContain("span 2");
    // The merged cell is as wide as two columns, the tall one as high as two rows.
    expect(wide.width).toBeGreaterThan(plain.width * 1.8);
    expect(tall.height).toBeGreaterThan(plain.height * 1.8);
  });
}, 30000);

/**
 * A slide whose file gave some elements no id renders in full, and its handle
 * says so once per slide (`pptx-missing-id`); a clean slide has no such warning.
 */
visual("visual DOM output: PptxRenderHandle.warnings reports elements the file gave no id, once per slide", async () => {
  const renderer = await browserModule(`${import.meta.dir}/../src/pptx/index.ts`);
  const page = `<!doctype html><style>html,body{margin:0}#host{width:800px;height:600px}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        window.renderSlide = async (model) => {
          window.handle?.destroy();
          window.handle = await render({ slide: model, target: document.querySelector("#host") });
          return { warnings: window.handle.warnings, shapes: document.querySelectorAll("[data-aared-shape-id]").length };
        };
      </script>`;
  const source = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const deck = await indexPptx(source);
  const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(source, deck, index)));
  const { index, slide: clean } = await findCoverageSlide(slides, "shapes");
  // Two elements on the slide with their ids stripped: a group (id 4) and a shape (id 29, inside another group).
  const pkg = await withPart("pptx-coverage.pptx", deck.slideParts[index]!, (xml) =>
    xml.replace('<p:cNvPr id="4" name="Group 3"', '<p:cNvPr name="Group 3"').replace('<p:cNvPr id="29" name="Rectangle 28"', '<p:cNvPr name="Rectangle 28"'));
  const damaged = await parsePptxSlide(pkg, await indexPptx(pkg), index);
  expect(damaged.missingIds).toBe(2);
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/pptx.mjs") return new Response(renderer, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      return new Response(page, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  await withVisualPage({ server, width: 800, height: 600 }, async (view) => {
    type Rendered = { warnings: Array<{ code: string; message: string }>; shapes: number };
    const whole = (await view.evaluate(`renderSlide(${JSON.stringify(clean)})`)) as Rendered;
    expect(whole.warnings.filter((w) => w.code === "pptx-missing-id")).toEqual([]);
    const repaired = (await view.evaluate(`renderSlide(${JSON.stringify(damaged)})`)) as Rendered;
    const warnings = repaired.warnings.filter((w) => w.code === "pptx-missing-id");
    expect(warnings).toHaveLength(1);
    expect(warnings[0]!.message).toContain(`2 elements on slide ${index + 1} have no id`);
    expect(warnings[0]!.message).toMatch(/^[\x20-\x7e]+$/);
    // Drawn in full: the same shapes are on the page as for the clean slide.
    expect(repaired.shapes).toBe(whole.shapes);
  });
}, 30000);
