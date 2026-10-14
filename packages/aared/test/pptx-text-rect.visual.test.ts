/**
 * A preset geometry's own text rectangle reaches the painted text layer, and a
 * preset aared cannot draw still paints its fill so its text stays readable.
 *
 * PowerPoint lays text in a rounded rectangle's narrower column and an
 * ellipse's much narrower one, not across the full box. The offsets are the
 * ECMA `a:rect` guides and were measured against PowerPoint's own laid-out
 * text edges (probe `pptx-text-rectangle`, worst disagreement 0.000pt across
 * nine presets at two aspect ratios), so the numbers here are Office's.
 * `drawingml-catalog.test.ts` pins the formulas themselves; these tests read
 * what the renderer actually painted.
 */
import { expect, test } from "bun:test";
import { geometryForPreset, unresolvedGeometry } from "../src/core/drawingml-shape.js";
import type { ShapeModel, SlideModel } from "../src/core/model.js";
import { browserModule, visual, withVisualPage } from "./harness.js";

/** A shape big enough that the two insets are distinguishable, in inches. */
const WIDTH = 4;
const HEIGHT = 1.5;

/** Office's authored defaults are 0.1in / 0.05in. */
const INSETS = { left: 0.1, right: 0.1, top: 0.05, bottom: 0.05 };

const FILL = "#156082";
const LABEL = "Shaping text inside the preset rectangle needs a column narrow enough to wrap, and a left edge that is the column's own";

function shapeWith(preset: string): ShapeModel {
  const geometry = geometryForPreset(preset, { width: WIDTH * 914400, height: HEIGHT * 914400 });
  if (geometry === undefined) throw new Error(`no geometry for ${preset}`);
  return {
    kind: "shape",
    id: "1",
    name: preset,
    bounds: { x: 0, y: 0, width: WIDTH, height: HEIGHT },
    geometry,
    fill: { kind: "solid", color: FILL },
    text: {
      paragraphs: [{ runs: [{ text: LABEL, color: "#FFFFFF" }], align: "left" }],
      insets: INSETS,
    },
  };
}

/** The runner options for a page that paints a slide and reads it back. */
async function page(readSource: string, width: number, height: number) {
  const module = await browserModule(`${import.meta.dir}/../src/pptx/index.ts`);
  return {
    modules: { "/pptx.mjs": module },
    page: `<!doctype html><style>html,body,#host{margin:0;width:100%;height:100%;overflow:hidden;font-family:Arial,sans-serif}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        window.handle = null;
        window.paint = async (model) => {
          window.handle = await render({ slide: model, target: document.querySelector("#host") });
          await document.fonts.ready;
          await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
          return true;
        };
        ${readSource}
      </script>`,
    width,
    height,
  };
}

visual("visual DOM output: a preset geometry's text rectangle insets the painted text layer", async () => {
  const presets = ["rect", "roundRect", "ellipse", "diamond", "octagon", "hexagon", "parallelogram"];
  await withVisualPage(await page(`window.read = (ids) => ids.map((id) => {
    const host = document.querySelector('[data-aared-shape-id="' + id + '"]');
    const layer = host && host.querySelector('[data-aared-element="shape-text"]');
    if (!layer) return { id, missing: true };
    const style = getComputedStyle(layer);
    return {
      id, missing: false,
      padding: [style.paddingLeft, style.paddingRight, style.paddingTop, style.paddingBottom].map(parseFloat),
      columnWidth: layer.getBoundingClientRect().width - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight),
      canvas: document.querySelectorAll('canvas').length,
    };
  });`, 1400, 900), async (view) => {
    const slide = {
      part: "ppt/slides/slide1.xml",
      size: { width: 10, height: 6 },
      elements: presets.map((preset, index) => ({
        ...shapeWith(preset),
        id: String(index + 1),
        bounds: { x: 0.5, y: 0.4 + index * (HEIGHT + 0.35), width: WIDTH, height: HEIGHT },
      })),
    } as unknown as SlideModel;
    await view.evaluate(`paint(${JSON.stringify(slide)})`);
    type Painted = { id: string; missing: boolean; padding: number[]; columnWidth: number; canvas: number };
    const painted = await view.evaluate(`read(${JSON.stringify(presets.map((_, index) => String(index + 1)))})`) as Painted[];
    const byPreset = new Map(presets.map((preset, index) => [preset, painted[index]!]));
    for (const preset of presets) {
      const entry = byPreset.get(preset)!;
      expect(entry.missing).toBe(false);
      expect(entry.canvas).toBe(0);
      const geometry = geometryForPreset(preset, { width: WIDTH * 914400, height: HEIGHT * 914400 })!;
      const rect = geometry.textRect!;
      // The geometry's share plus the authored inset, which is what PowerPoint's
      // measured column edge is.
      const expectedLeft = INSETS.left * 96 + (rect.left / geometry.viewBox!.width) * WIDTH * 96;
      const expectedTop = INSETS.top * 96 + (rect.top / geometry.viewBox!.height) * HEIGHT * 96;
      expect(entry.padding[0]).toBeCloseTo(expectedLeft, 0);
      expect(entry.padding[2]).toBeCloseTo(expectedTop, 0);
      if (preset === "rect") {
        // A rect adds no geometry inset: the column is the box less its insets.
        expect(entry.padding[0]).toBeCloseTo(INSETS.left * 96, 1);
        expect(entry.padding[2]).toBeCloseTo(INSETS.top * 96, 1);
      } else {
        expect(entry.padding[0]).toBeGreaterThan(INSETS.left * 96);
      }
    }
    // The visible point of the row: the ellipse's column is far narrower than
    // the rectangle's, by the amount PowerPoint measured.
    expect(byPreset.get("rect")!.columnWidth - byPreset.get("ellipse")!.columnWidth).toBeGreaterThan(90);
    await view.evaluate("window.handle?.destroy()");
  });
}, 30000);

/**
 * A preset aared cannot draw still paints the shape's own fill. PowerPoint
 * fills the pentagon and sets white text; painting nothing left black text on
 * a white slide inside a dashed outline, which reads as a missing shape rather
 * than a missing outline (probe `pptx-text-rectangle`, the pentagon case).
 */
visual("visual DOM output: an unrenderable preset still paints its fill so its text stays legible", async () => {
  await withVisualPage(await page(`window.read = () => {
    const host = document.querySelector('[data-aared-shape-id="1"]');
    const path = host.querySelector('[data-aared-element="shape-geometry"] path');
    const run = host.querySelector('[data-aared-element="paragraph-text"] span');
    return {
      fallback: host.dataset.aaredGeometryFallback ?? null,
      pathFill: path.getAttribute("fill"),
      dashed: path.getAttribute("stroke-dasharray"),
      text: host.textContent,
      runColor: run ? getComputedStyle(run).color : null,
      canvas: document.querySelectorAll('canvas').length,
    };
  };`, 1200, 900), async (view) => {
    // A pentagon: no fast-path outline, so it resolves as a marked fallback.
    const shape = shapeWith("rect");
    const slide = {
      part: "ppt/slides/slide1.xml",
      size: { width: 10, height: 6 },
      elements: [{
        ...shape,
        geometry: unresolvedGeometry({ width: WIDTH * 914400, height: HEIGHT * 914400 }, "pentagon", {}, undefined),
      }],
    } as unknown as SlideModel;
    await view.evaluate(`paint(${JSON.stringify(slide)})`);
    const out = await view.evaluate("read()") as {
      fallback: string | null; pathFill: string | null; dashed: string | null; text: string; runColor: string | null; canvas: number;
    };
    expect(out.fallback).toBe("unsupported");
    // The fill is what keeps the text readable; the dashed outline marks the
    // outline as a stand-in without hiding the shape.
    expect(out.pathFill?.toUpperCase()).toBe(FILL);
    expect(out.dashed).not.toBeNull();
    expect(out.text).toContain("Shaping text");
    expect(out.runColor).toBe("rgb(255, 255, 255)");
    expect(out.canvas).toBe(0);
    await view.evaluate("window.handle?.destroy()");
  });
}, 30000);