import { expect, test } from "bun:test";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { OpcPackage } from "../src/core/opc.js";
import { browserModule } from "./harness.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const COVERAGE_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const PRODUCTION_PATH = `${import.meta.dir}/../../../fixtures/post-production-review.pptx`;

visual("visual DOM output: pictures, autofit, numbered bullets, resolved tables, and charts", async () => {
  const coveragePkg = await OpcPackage.open(await Bun.file(COVERAGE_PATH).arrayBuffer());
  const coverageDeck = await indexPptx(coveragePkg);
  const coverageSlides = await Promise.all(coverageDeck.slideParts.map((_, index) => parsePptxSlide(coveragePkg, coverageDeck, index)));
  const productionPkg = await OpcPackage.open(await Bun.file(PRODUCTION_PATH).arrayBuffer());
  const productionDeck = await indexPptx(productionPkg);
  const productionSlides = await Promise.all(productionDeck.slideParts.map((_, index) => parsePptxSlide(productionPkg, productionDeck, index)));
  const productionSlide1 = productionSlides[0]!;
  const productionSlide4 = productionSlides[3]!;
  const productionSlide5 = productionSlides[4]!;
  const picture = productionSlide1.elements.find((element) => element.kind === "picture");
  if (picture?.kind !== "picture") throw new Error("fixture structure changed: expected a production picture");
  const pictureBytes = await productionPkg.readPart(picture.mediaId);
  const pictureBase64 = Buffer.from(pictureBytes).toString("base64");

  const module = await browserModule(`${import.meta.dir}/../src/pptx/index.ts`);
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/pptx.mjs") {
        return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      return new Response(`<!doctype html><style>html,body{margin:0;font-family:Arial,sans-serif}#host{width:1280px}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        window.renderedFixture = null;
        const decode = (base64) => Uint8Array.from(atob(base64), (c) => c.charCodeAt(0)).buffer;
        window.renderFixture = async (model, mediaBase64, mediaId) => {
          window.renderedFixture?.destroy();
          window.renderedFixture = await render({
            slide: model,
            target: document.querySelector("#host"),
            mediaResolver: mediaBase64 === null ? undefined : () => Promise.resolve(decode(mediaBase64)),
          });
        };
        window.renderFixtureRejecting = async (model) => {
          window.renderedFixture?.destroy();
          window.renderedFixture = await render({
            slide: model,
            target: document.querySelector("#host"),
            mediaResolver: () => Promise.reject(new Error("missing")),
          });
        };
        window.destroyFixture = () => window.renderedFixture?.destroy();
        window.waitForPictures = () => new Promise((resolve) => {
          const tick = () => {
            const images = [...document.querySelectorAll('[data-aared-element="picture-image"]')];
            const fallbacks = document.querySelectorAll('[data-aared-element="picture-fallback"]').length;
            if (images.every((img) => img.src.startsWith("blob:") || img.complete) || fallbacks > 0) resolve();
            else requestAnimationFrame(tick);
          };
          requestAnimationFrame(() => requestAnimationFrame(tick));
        });
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1280, height: 800 });
    await view.navigate(server.url.href);

    // Pictures end-to-end: placement, rotation/flips, crop, async load, fallback.
    await view.evaluate(`renderFixture(${JSON.stringify({ size: productionSlide1.size, elements: [picture] })}, "${pictureBase64}")`);
    await view.evaluate(`waitForPictures()`);
    const pictureOutput = await view.evaluate(`(() => {
      const wrapper = document.querySelector('[data-aared-element="picture"]');
      const image = document.querySelector('[data-aared-element="picture-image"]');
      const rect = wrapper?.getBoundingClientRect();
      return {
        mediaId: wrapper?.dataset.aaredMediaId,
        contentType: wrapper?.dataset.aaredContentType,
        left: wrapper?.style.left,
        top: wrapper?.style.top,
        width: wrapper?.style.width,
        height: wrapper?.style.height,
        src: image?.src.slice(0, 5),
        loaded: image?.dataset.aaredMediaLoaded,
        canvasCount: document.querySelectorAll("canvas").length,
        rect: rect ? { width: rect.width, height: rect.height } : null,
      };
    })()`) as {
      mediaId?: string; contentType?: string; left?: string; top?: string; width?: string; height?: string;
      src?: string; loaded?: string; canvasCount: number; rect: { width: number; height: number } | null;
    };
    expect(pictureOutput.mediaId).toBe(picture.mediaId);
    expect(pictureOutput.contentType).toBe("image/png");
    expect(pictureOutput.left).toBe(`${0.75 * 96}px`);
    expect(pictureOutput.width).toBe(`${0.9 * 96}px`);
    expect(pictureOutput.src).toBe("blob:");
    expect(pictureOutput.loaded).toBe("true");
    expect(pictureOutput.canvasCount).toBe(0);
    expect(pictureOutput.rect!.width).toBeGreaterThan(0);

    // Synthetic crop + rotation resolve to CSS without canvas.
    const cropped = {
      kind: "picture", id: "crop", name: "Crop", mediaId: picture.mediaId, contentType: "image/png",
      bounds: { x: 1, y: 1, width: 2, height: 2 },
      crop: { left: 0.1, top: 0.2, right: 0.05, bottom: 0 },
      transform: { rotation: 90, flipHorizontal: true },
    };
    await view.evaluate(`renderFixture(${JSON.stringify({ size: productionSlide1.size, elements: [cropped] })}, "${pictureBase64}")`);
    await view.evaluate(`waitForPictures()`);
    const cropOutput = await view.evaluate(`(() => {
      const wrapper = document.querySelector('[data-aared-element="picture"]');
      const image = document.querySelector('[data-aared-element="picture-image"]');
      return {
        crop: [wrapper?.dataset.aaredCropLeft, wrapper?.dataset.aaredCropTop, wrapper?.dataset.aaredCropRight, wrapper?.dataset.aaredCropBottom],
        imageWidth: image?.style.width,
        imageLeft: image?.style.left,
        transform: wrapper?.style.transform,
      };
    })()`) as { crop: string[]; imageWidth?: string; imageLeft?: string; transform?: string };
    expect(cropOutput.crop).toEqual(["0.1", "0.2", "0.05", "0"]);
    // Visible 85% width -> 117.65% image width; left offset -11.76%.
    expect(Number.parseFloat(cropOutput.imageWidth!)).toBeCloseTo(100 / 0.85, 2);
    expect(Number.parseFloat(cropOutput.imageLeft!)).toBeCloseTo((-0.1 / 0.85) * 100, 2);
    expect(cropOutput.transform).toContain("rotate(90deg)");
    expect(cropOutput.transform).toContain("scale(-1, 1)");

    // Missing resolver and load errors render marked fallbacks, never canvas.
    await view.evaluate(`renderFixture(${JSON.stringify({ size: productionSlide1.size, elements: [cropped] })}, null)`);
    const missingOutput = await view.evaluate(`({
      fallback: document.querySelector('[data-aared-element="picture"]')?.dataset.aaredPictureFallback,
      canvasCount: document.querySelectorAll("canvas").length,
    })`) as { fallback?: string; canvasCount: number };
    expect(missingOutput.fallback).toBe("missing-resolver");
    await view.evaluate(`renderFixtureRejecting(${JSON.stringify({ size: productionSlide1.size, elements: [cropped] })})`);
    await view.evaluate(`new Promise((resolve) => setTimeout(resolve, 50))`);
    const errorOutput = await view.evaluate(`({
      fallback: document.querySelector('[data-aared-element="picture"]')?.dataset.aaredPictureFallback,
      canvasCount: document.querySelectorAll("canvas").length,
    })`) as { fallback?: string; canvasCount: number };
    expect(errorOutput.fallback).toBe("load-error");
    expect(errorOutput.canvasCount).toBe(0);

    // Autofit: shrink and resize contain text; none clips.
    const autofitModel = {
      size: { width: 10, height: 5.625 },
      elements: [
        { kind: "shape", id: "shrink", name: "shrink", bounds: { x: 0.5, y: 0.5, width: 2, height: 0.6 }, text: { paragraphs: [{ runs: [{ text: "Long ".repeat(40), fontSizePt: 18 }] }], autofit: "shrinkText" }, geometry: { path: "M 0 0 H 100000 V 100000 H 0 Z", viewBox: { width: 100000, height: 100000 } } },
        { kind: "shape", id: "resize", name: "resize", bounds: { x: 3, y: 0.5, width: 2, height: 0.6 }, text: { paragraphs: [{ runs: [{ text: "Long ".repeat(40), fontSizePt: 18 }] }], autofit: "resizeShape" }, geometry: { path: "M 0 0 H 100000 V 100000 H 0 Z", viewBox: { width: 100000, height: 100000 } } },
        { kind: "shape", id: "clip", name: "clip", bounds: { x: 5.5, y: 0.5, width: 2, height: 0.6 }, text: { paragraphs: [{ runs: [{ text: "Long ".repeat(40), fontSizePt: 18 }] }], autofit: "none" }, geometry: { path: "M 0 0 H 100000 V 100000 H 0 Z", viewBox: { width: 100000, height: 100000 } } },
      ],
    };
    await view.evaluate(`renderFixture(${JSON.stringify(autofitModel)}, null)`);
    await view.evaluate(`new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))`);
    const autofitOutput = await view.evaluate(`(() => {
      const result = {};
      for (const id of ["shrink", "resize", "clip"]) {
        const shape = document.querySelector('[data-aared-shape-id="' + id + '"]');
        const text = shape?.querySelector('[data-aared-element="shape-text"]');
        result[id] = text ? {
          autofit: text.dataset.aaredAutofit,
          overflow: text.style.overflow,
          contained: text.scrollHeight <= text.clientHeight + 2,
          scroll: text.scrollHeight,
          client: text.clientHeight,
        } : null;
      }
      return result;
    })()`) as Record<string, { autofit?: string; overflow: string; contained: boolean } | null>;
    expect(autofitOutput.shrink?.autofit).toBe("shrinkText");
    expect(autofitOutput.shrink?.contained).toBe(true);
    expect(autofitOutput.resize?.autofit).toBe("resizeShape");
    expect(autofitOutput.resize?.overflow).toBe("visible");
    expect(autofitOutput.resize?.contained).toBe(true);
    expect(autofitOutput.clip?.autofit).toBe("none");
    expect(autofitOutput.clip?.overflow).toBe("hidden");

    // Numbered bullets follow direction for their marker side.
    const numbered = {
      size: { width: 10, height: 5.625 },
      elements: [{
        kind: "shape", id: "num", name: "num", bounds: { x: 0.5, y: 0.5, width: 9, height: 4 },
        text: { paragraphs: [{ direction: "rtl", bullet: { char: "2." }, runs: [{ text: "Second" }] }] },
        geometry: { path: "M 0 0 H 100000 V 100000 H 0 Z", viewBox: { width: 100000, height: 100000 } },
      }],
    };
    await view.evaluate(`renderFixture(${JSON.stringify(numbered)}, null)`);
    const bulletOutput = await view.evaluate(`(() => {
      const line = document.querySelector("p");
      const marker = line?.querySelector('[data-aared-element="bullet-marker"]');
      const text = line?.querySelector('[data-aared-element="paragraph-text"]');
      if (!line || !marker || !text) return null;
      const markerRect = marker.getBoundingClientRect();
      const textRect = text.getBoundingClientRect();
      return { char: marker.textContent, side: marker.dataset.aaredBulletSide, markerLeft: markerRect.left, textRight: textRect.right };
    })()`) as { char: string; side?: string; markerLeft: number; textRight: number } | null;
    expect(bulletOutput?.char).toBe("2.");
    expect(bulletOutput?.side).toBe("right");
    expect(bulletOutput!.markerLeft).toBeGreaterThanOrEqual(bulletOutput!.textRight);

    // Resolved tables keep logical order, fills, and per-edge strokes.
    await view.evaluate(`renderFixture(${JSON.stringify(productionSlide5)}, null)`);
    const tableOutput = await view.evaluate(`(() => {
      const grid = document.querySelector('[data-aared-element="table"]');
      const cells = [...document.querySelectorAll('[data-aared-element="table-cell"]')];
      return {
        direction: grid?.dataset.aaredTableDirection,
        firstRow: cells.slice(0, 4).map((cell) => cell.textContent),
        headerFill: cells[0] ? getComputedStyle(cells[0]).backgroundColor : null,
        margins: cells[0]?.dataset.aaredCellMargins,
        borders: cells[0]?.dataset.aaredCellBorders,
        topEdge: (() => {
          const edge = document.querySelector('[data-aared-element="table-cell-edges"] [data-aared-edge="top"]');
          return edge === null ? null : getComputedStyle(edge).borderTopStyle + " " + getComputedStyle(edge).borderTopColor;
        })(),
        paddingTop: cells[0] ? getComputedStyle(cells[0]).paddingTop : null,
        lefts: cells.slice(0, 2).map((cell) => cell.getBoundingClientRect().left),
        gridHeight: grid.getBoundingClientRect().height,
        scale: document.querySelector('[data-aared-element="slide-scene"]').getBoundingClientRect().width / 1280,
        textClipped: cells.some((cell) => cell.scrollHeight > cell.clientHeight + 1),
      };
    })()`) as {
      direction?: string; firstRow: string[]; headerFill: string | null; margins?: string;
      borders?: string; topEdge: string | null; paddingTop: string | null; lefts: number[];
      gridHeight: number; scale: number; textClipped: boolean;
    };
    expect(tableOutput.direction).toBe("ltr");
    expect(tableOutput.firstRow[0]).toContain("Parameter");
    expect(tableOutput.headerFill).toBe("rgb(18, 23, 31)");
    expect(tableOutput.margins).toBe("resolved");
    expect(tableOutput.borders).toBe("resolved");
    expect(tableOutput.topEdge).toBe("solid rgb(216, 211, 201)");
    // 0.05in top inset at 96 DPI, before the slide's own scale.
    expect(Number.parseFloat(tableOutput.paddingTop!)).toBeCloseTo(4.8, 1);
    // The frame is authored 1in tall but holds six 0.52in rows: PowerPoint
    // grows the table to its rows, so the painted grid is the rows' total,
    // and no cell's text is clipped.
    const authoredRows = productionSlide5.elements.find((element) => element.kind === "table");
    if (authoredRows?.kind !== "table") throw new Error("fixture structure changed: expected slide 5's table");
    const rowTotal = authoredRows.rows.reduce((sum, row) => sum + (row.height ?? 0), 0) * 96 * tableOutput.scale;
    expect(tableOutput.gridHeight).toBeGreaterThanOrEqual(rowTotal - 1);
    expect(tableOutput.textClipped).toBe(false);

    // Charts paint typed primitives to SVG, never canvas.
    await view.evaluate(`renderFixture(${JSON.stringify(productionSlide4)}, null)`);
    const chartOutput = await view.evaluate(`(() => {
      const wrapper = document.querySelector('[data-aared-element="chart"]');
      const svg = document.querySelector('[data-aared-element="chart-geometry"]');
      return {
        primitives: wrapper?.dataset.aaredChartPrimitives,
        rects: svg?.querySelectorAll("rect").length,
        texts: svg?.querySelectorAll("text").length,
        hasLogged: svg?.textContent.includes("Logged"),
        canvasCount: document.querySelectorAll("canvas").length,
      };
    })()`) as { primitives?: string; rects: number; texts: number; hasLogged: boolean; canvasCount: number };
    expect(Number(chartOutput.primitives)).toBeGreaterThan(10);
    expect(chartOutput.rects).toBeGreaterThanOrEqual(6);
    expect(chartOutput.texts).toBeGreaterThanOrEqual(10);
    expect(chartOutput.hasLogged).toBe(true);
    expect(chartOutput.canvasCount).toBe(0);

    // Blood: destroy stops async picture work and revokes object URLs.
    await view.evaluate(`renderFixture(${JSON.stringify({ size: productionSlide1.size, elements: [picture] })}, "${pictureBase64}")`);
    const destroyedOutput = await view.evaluate(`(() => new Promise((resolve) => {
      const before = document.querySelector('[data-aared-element="slide-scene"]')?.style.transform;
      destroyFixture();
      document.querySelector("#host").style.width = "640px";
      requestAnimationFrame(() => requestAnimationFrame(() => resolve({
        before,
        after: document.querySelector('[data-aared-element="slide-scene"]')?.style.transform,
      })));
    }))()`) as { before?: string; after?: string };
    expect(destroyedOutput.after).toBe(destroyedOutput.before);
  } finally {
    server.stop();
  }
});
