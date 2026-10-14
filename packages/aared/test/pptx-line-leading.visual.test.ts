import { expect } from "bun:test";
import { readFileSync } from "node:fs";
import { browserModule, visual, withVisualPage } from "./harness.js";

// Ground truth, PowerPoint's own PDF (probe first-line-leading): one slide per
// font and size (Arial, Calibri, Times New Roman at 18pt and 40pt), one fixed
// top-anchored box per line spacing (0.7-3.0 lines, set through COM as the UI
// writes it), each holding "H<spacing>" then "L<spacing>". The baselines of
// every line are in pptx-first-line-oracle.json (page = slide, in points). How
// far a box's first line sits from the same box at 100% is PowerPoint's rule
// for where leading goes; CSS splits leading evenly, so the paragraph paints
// the difference as a shift of the text, never as a taller or shorter box.
// Arial 40pt, for example: 4.80pt higher at 90%, 33.14pt lower at 200%.
const SPACINGS = [0.7, 0.8, 0.9, 1, 1.2, 1.5, 2, 3];
const CASES = [
  { font: "Arial", size: 18 }, { font: "Arial", size: 40 },
  { font: "Calibri", size: 18 }, { font: "Calibri", size: 40 },
  { font: "Times New Roman", size: 18 }, { font: "Times New Roman", size: 40 },
];
/** The worst gap the rule left against PowerPoint when it was fitted was 1.15pt. */
const TOLERANCE_PT = 1.5;

const baselines = JSON.parse(readFileSync(new URL("./pptx-first-line-oracle.json", import.meta.url), "utf8")) as Array<{ page: number; text: string; baseline: number }>;

/** PowerPoint's first-line shift for one case, in points, against its 100% box. */
function powerpointShift(page: number, spacing: number): number {
  const at = (s: number): number => baselines.find((row) => row.page === page && row.text === `H${s}`)!.baseline;
  return at(spacing) - at(1);
}

visual("visual DOM output: PPTX places a line inside its box as PowerPoint does when the spacing is not 100%", async () => {
  const module = await browserModule(`${import.meta.dir}/../src/pptx/renderer.ts`);
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
    for (const [page, { font, size }] of CASES.entries()) {
      const model = {
        size: { width: 13.333, height: 7.5 },
        elements: SPACINGS.map((spacing, index) => ({
          kind: "shape",
          id: String(spacing),
          name: `spacing ${spacing}`,
          bounds: { x: 0.14 + index * 1.64, y: 0.56, width: 1.53, height: 5.56 },
          text: {
            paragraphs: [{ lineSpacing: spacing, runs: [{ text: "H", fontSizePt: size, fontFamily: font }] }],
            verticalAlign: "top",
          },
          geometry: { path: "M 0 0 H 100000 V 100000 H 0 Z", viewBox: { width: 100000, height: 100000 } },
        })),
      };
      await view.evaluate(`renderFixture(${JSON.stringify(model)})`);
      const lines = await view.evaluate("lines()") as Record<string, { glyphTop: number; paragraphHeight: number; position: string }>;
      for (const spacing of SPACINGS) {
        const line = lines[String(spacing)]!;
        const label = `${font} ${size}pt at ${spacing}`;
        // Painted in points (96 dpi: 0.75pt per px), against the 100% box.
        const shift = (line.glyphTop - lines["1"]!.glyphTop) * 0.75;
        expect(Math.abs(shift - powerpointShift(page, spacing)), label).toBeLessThan(TOLERANCE_PT);
        // The box keeps its authored height (or the floor under it): only the text moves.
        expect(line.paragraphHeight / lines["1"]!.paragraphHeight, label).toBeCloseTo(Math.max(spacing * 1.2, 1.1) / 1.2, 1);
        expect(line.position, label).toBe(spacing === 1 ? "static" : "relative");
      }
    }
  });
});
