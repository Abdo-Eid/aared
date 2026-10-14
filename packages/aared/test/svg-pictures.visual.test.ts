/**
 * Office SVG pictures paint the vector part first and fall back to Office's
 * own raster when the SVG cannot be read or decoded (probe `svg-pictures`).
 * Re-reads actual rendered `<img>` output: blob sources, decode success, and
 * the fallback swap — never an in-memory intermediate.
 */
import { expect } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { readWorkbook } from "../src/core/xlsx/workbook.js";
import { parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { findCoverageSlide } from "./fixture.js";
import { findCoverageSheet } from "./xlsx-fixture.js";
import { browserModule, visual, withVisualPage } from "./harness.js";

const PPTX_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const XLSX_PATH = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;

const base64 = (bytes: Uint8Array): string => Buffer.from(bytes).toString("base64");

visual("visual DOM output: SVG pictures paint the vector part, falling back to the raster", async () => {
  const pptxPkg = await OpcPackage.open(await Bun.file(PPTX_PATH).arrayBuffer());
  const deck = await indexPptx(pptxPkg);
  const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pptxPkg, deck, index)));
  const { slide } = await findCoverageSlide(slides, "svg-picture");
  const picture = slide.elements.find((element) => element.kind === "picture");
  if (picture?.kind !== "picture") throw new Error("fixture structure changed: expected an SVG picture slide");
  const svgBytes = base64(await pptxPkg.readPart(picture.mediaId));

  const xlsxPkg = await OpcPackage.open(await Bun.file(XLSX_PATH).arrayBuffer());
  const workbook = await readWorkbook(xlsxPkg);
  const sheets = [];
  for (let index = 0; index < workbook.sheets.length; index++) sheets.push(await parseXlsxSheet(xlsxPkg, workbook, index));
  const notes = (await findCoverageSheet(sheets, "notes")).sheet;
  const drawing = notes.drawings.find((entry) => entry.kind === "picture");
  if (drawing?.kind !== "picture") throw new Error("fixture structure changed: expected a Notes SVG picture");
  const drawingSvg = base64(await xlsxPkg.readPart(drawing.mediaId));
  const drawingPng = base64(await xlsxPkg.readPart(drawing.fallbackMediaId!));

  const [pptxModule, xlsxModule] = await Promise.all([
    browserModule(`${import.meta.dir}/../src/pptx/renderer.ts`),
    browserModule(`${import.meta.dir}/../src/xlsx/renderer.ts`),
  ]);
  await withVisualPage({
    modules: { "/pptx.mjs": pptxModule, "/xlsx.mjs": xlsxModule },
    page: `<!doctype html><style>html,body{margin:0;font-family:Arial,sans-serif}#host{width:1280px;height:600px}</style><main id="host"></main><script type="module">
        import { render as renderPptx } from "/pptx.mjs";
        import { render as renderXlsx } from "/xlsx.mjs";
        const decode = (b64) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)).buffer;
        window.handles = [];
        // parts: mediaId -> base64 ("CORRUPT" decodes to bytes no image decoder accepts).
        window.renderSlide = async (model, parts) => {
          const handle = await renderPptx({ slide: model, target: document.querySelector("#host"),
            mediaResolver: (id) => {
              const found = parts[id];
              if (found === undefined) return Promise.reject(new Error("missing " + id));
              const bytes = found === "CORRUPT" ? new Uint8Array([1, 2, 3, 4]).buffer : decode(found);
              return Promise.resolve(bytes);
            } });
          window.handles.push(handle);
        };
        window.renderNotes = async (model, parts) => {
          const handle = await renderXlsx({ sheet: model, target: document.querySelector("#host"),
            mediaResolver: (id) => {
              const found = parts[id];
              if (found === undefined) return Promise.reject(new Error("missing " + id));
              const bytes = found === "CORRUPT" ? new Uint8Array([1, 2, 3, 4]).buffer : decode(found);
              return Promise.resolve(bytes);
            } });
          window.handles.push(handle);
        };
        window.imageState = (selector) => {
          const img = document.querySelector(selector);
          return img === null ? null : {
            src: img.src.slice(0, 5),
            loaded: img.dataset.aaredMediaLoaded ?? null,
            width: img.naturalWidth,
            fallbackLeft: img.dataset.aaredImageFallback !== undefined,
          };
        };
        window.waitFor = (selector) => new Promise((resolve) => {
          const tick = () => {
            const img = document.querySelector(selector);
            if (img !== null && (img.dataset.aaredMediaLoaded === "true" || img.complete)) resolve(true);
            else requestAnimationFrame(tick);
          };
          requestAnimationFrame(() => requestAnimationFrame(tick));
        });
        window.waitForElement = (selector) => new Promise((resolve) => {
          const tick = () => {
            if (document.querySelector(selector) !== null) resolve(true);
            else requestAnimationFrame(tick);
          };
          requestAnimationFrame(() => requestAnimationFrame(tick));
        });
      </script>`,
    width: 1280,
    height: 800,
  }, async (view) => {
    // The Office-inserted SVG paints as a decoded vector image.
    const slideModel = { size: slide.size, elements: [picture] };
    await view.evaluate(`renderSlide(${JSON.stringify(slideModel)}, ${JSON.stringify({ [picture.mediaId]: svgBytes })})`);
    await view.evaluate(`waitFor('[data-aared-element="picture-image"]')`);
    const painted = await view.evaluate(`imageState('[data-aared-element="picture-image"]')`) as {
      src?: string; loaded?: string; width?: number; fallbackLeft?: boolean;
    } | null;
    expect(painted?.src).toBe("blob:");
    expect(painted?.loaded).toBe("true");
    expect(painted?.width).toBeGreaterThan(0);
    expect(painted?.fallbackLeft).toBe(false);

    // A corrupt SVG with no raster beside it ends in the marked fallback, never a broken icon.
    await view.evaluate(`renderSlide(${JSON.stringify(slideModel)}, ${JSON.stringify({ [picture.mediaId]: "CORRUPT" })})`);
    await view.evaluate(`waitForElement('[data-aared-element="picture-fallback"]')`);
    const missing = await view.evaluate(`document.querySelector('[data-aared-element="picture"]')?.dataset.aaredPictureFallback ?? null`);
    expect(missing).toBe("load-error");

    // The spreadsheet drawing keeps Office's PNG beside its SVG: a corrupt
    // vector swaps to the raster and still decodes.
    await view.evaluate(`renderNotes(${JSON.stringify(notes)}, ${JSON.stringify({ [drawing.mediaId]: drawingSvg, [drawing.fallbackMediaId!]: drawingPng })})`);
    await view.evaluate(`waitFor('[data-aared-element="picture"] img')`);
    const vector = await view.evaluate(`imageState('[data-aared-element="picture"] img')`) as {
      src?: string; loaded?: string; width?: number; fallbackLeft?: boolean;
    } | null;
    expect(vector?.src).toBe("blob:");
    expect(vector?.width).toBeGreaterThan(0);
    await view.evaluate(`renderNotes(${JSON.stringify(notes)}, ${JSON.stringify({ [drawing.mediaId]: "CORRUPT", [drawing.fallbackMediaId!]: drawingPng })})`);
    await view.evaluate(`waitFor('[data-aared-element="picture"] img')`);
    const raster = await view.evaluate(`((img) => img === null ? null : {
      src: img.src.slice(0, 5), width: img.naturalWidth, fallbackLeft: img.dataset.aaredImageFallback !== undefined,
    })(document.querySelector('[data-aared-element="picture"] img'))`) as {
      src?: string; width?: number; fallbackLeft?: boolean;
    } | null;
    expect(raster?.src).toBe("blob:");
    expect(raster?.width).toBeGreaterThan(0);
    expect(raster?.fallbackLeft).toBe(false);
    // Excel's natively-authored shadow paints through the shared effect
    // painter, and the 12-degree rotation on the box itself.
    const effect = await view.evaluate(`(() => {
      const box = document.querySelector('[data-aared-element="picture"]');
      return {
        filter: box?.style.filter ?? null,
        transform: box?.style.transform ?? null,
        fallback: box?.dataset.aaredEffectFallback ?? null,
        boxes: document.querySelectorAll('[data-aared-element="picture"]').length,
      };
    })()`) as { filter?: string | null; transform?: string | null; fallback?: string | null; boxes?: number };
    expect(effect.filter).toContain("drop-shadow");
    expect(effect.transform).toContain("rotate(12deg)");
    await view.evaluate(`window.handles.forEach((handle) => handle.destroy())`);
  });
}, 60000);
