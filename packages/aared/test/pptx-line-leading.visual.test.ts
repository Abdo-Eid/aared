import { expect } from "bun:test";
import { browserModule, visual, withVisualPage } from "./harness.js";

// Ground truth, PowerPoint's own PDF (probe first-line-leading): Arial 40pt in a
// top-anchored box puts its first line, against the same box at 100% spacing,
// 4.80pt higher at 90%, 9.60pt higher at 80%, 33.14pt lower at 200% and 69.14pt
// lower at 300%. CSS splits the leading evenly, which puts the text 2.40pt
// higher at 90% and 24pt lower at 200%: the paragraph paints the difference as
// a shift of the text, never as a taller or shorter box.
const SPACINGS = [0.9, 1, 2, 3];
const POWERPOINT_SHIFT_PT: Record<string, number> = { "0.9": -4.8, "1": 0, "2": 33.14, "3": 69.14 };

visual("visual DOM output: PPTX places a line inside its box as PowerPoint does when the spacing is not 100%", async () => {
  const module = await browserModule(`${import.meta.dir}/../src/pptx/index.ts`);
  await withVisualPage({
    modules: { "/pptx.mjs": module },
    page: `<!doctype html><style>html,body{margin:0}#host{width:1280px}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        window.rendered = null;
        window.renderFixture = async (model) => {
          window.rendered?.destroy();
          window.rendered = await render({ slide: model, target: document.querySelector("#host"), zoom: 1 });
          await document.fonts.ready;
        };
        // Per shape: where its glyph starts and how tall its paragraph is.
        window.lines = () => Object.fromEntries([...document.querySelectorAll("[data-aared-shape-id]")].map((shape) => {
          const paragraph = shape.querySelector("p");
          const text = [...paragraph.querySelectorAll("span")].find((span) => span.firstChild?.nodeType === 3);
          const range = document.createRange();
          range.selectNodeContents(text.firstChild);
          return [shape.dataset.aaredShapeId, {
            glyphTop: range.getClientRects()[0].top,
            paragraphHeight: paragraph.getBoundingClientRect().height,
            position: getComputedStyle(paragraph).position,
          }];
        }));
      </script>`,
    width: 1280,
    height: 720,
  }, async (view) => {
    const model = {
      size: { width: 13.333, height: 7.5 },
      elements: SPACINGS.map((spacing, index) => ({
        kind: "shape",
        id: String(spacing),
        name: `spacing ${spacing}`,
        bounds: { x: 0.2 + index * 3, y: 0.5, width: 2.8, height: 5 },
        text: {
          paragraphs: [{ lineSpacing: spacing, runs: [{ text: "H", fontSizePt: 40, fontFamily: "Arial" }] }],
          verticalAlign: "top",
        },
        geometry: { path: "M 0 0 H 100000 V 100000 H 0 Z", viewBox: { width: 100000, height: 100000 } },
      })),
    };
    await view.evaluate(`renderFixture(${JSON.stringify(model)})`);
    const lines = await view.evaluate("lines()") as Record<string, { glyphTop: number; paragraphHeight: number; position: string }>;
    for (const spacing of SPACINGS) {
      const line = lines[String(spacing)]!;
      // Painted in points (96 dpi: 0.75pt per px), against the 100% box.
      const shift = (line.glyphTop - lines["1"]!.glyphTop) * 0.75;
      expect(Math.abs(shift - POWERPOINT_SHIFT_PT[String(spacing)]!)).toBeLessThan(1.5);
      // The box keeps its authored height (or the floor under it): only the text moves.
      expect(line.paragraphHeight / lines["1"]!.paragraphHeight).toBeCloseTo(Math.max(spacing * 1.2, 1.1) / 1.2, 1);
      expect(line.position).toBe(spacing === 1 ? "static" : "relative");
    }
  });
});
