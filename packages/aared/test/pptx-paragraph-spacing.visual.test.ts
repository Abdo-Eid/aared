import { expect } from "bun:test";
import { browserModule, visual, withVisualPage } from "./harness.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { OpcPackage } from "../src/core/opc.js";
import { findCoverageSlide } from "./fixture.js";

const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

interface ParagraphBox { top: number; bottom: number; fontSize: number; marginTop: number; marginBottom: number }

// Ground truth, PowerPoint's own export (probe line-spacing-units): a line of
// paragraph spacing is 1.2 x the paragraph's largest run size; space before a
// text body's first paragraph and after its last is dropped, in points or in
// lines; one paragraph's space after adds to the next one's space before.
visual("visual DOM output: PPTX paragraph spacing in lines, text-body edges and additive gaps", async () => {
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const deck = await indexPptx(pkg);
  const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
  const [{ slide: spacingSlide }, { slide: bulletSlide }] = await Promise.all([
    findCoverageSlide(slides, "paragraph-spacing"),
    findCoverageSlide(slides, "bullets"),
  ]);
  const module = await browserModule(`${import.meta.dir}/../src/pptx/renderer.ts`);
  await withVisualPage({
    modules: { "/pptx.mjs": module },
    page: `<!doctype html><style>html,body{margin:0}#host{width:1280px}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        window.rendered = null;
        window.renderFixture = async (model) => {
          window.rendered?.destroy();
          window.rendered = await render({ slide: model, target: document.querySelector("#host") });
        };
        // Every paragraph of the text body holding the given text, top to bottom.
        window.paragraphs = (text, selector) => {
          const host = [...document.querySelectorAll(selector)].find((element) => element.textContent.includes(text));
          return [...host.querySelectorAll("p")].map((p) => {
            const style = getComputedStyle(p);
            const { top, bottom } = p.getBoundingClientRect();
            return { top, bottom, fontSize: parseFloat(style.fontSize), marginTop: parseFloat(style.marginTop), marginBottom: parseFloat(style.marginBottom) };
          });
        };
      </script>`,
    width: 1280,
    height: 720,
  }, async (view) => {
    await view.evaluate(`renderFixture(${JSON.stringify(spacingSlide)})`);
    const box = await view.evaluate(`paragraphs("نصف سطر قبل هذه الفقرة", '[data-aared-element="shape"]')`) as ParagraphBox[];
    const cell = await view.evaluate(`paragraphs("نصف سطر بعد هذه", '[data-aared-element="table-cell"]')`) as ParagraphBox[];
    const gap = (list: ParagraphBox[], index: number): number => list[index]!.top - list[index - 1]!.bottom;
    const line = (paragraph: ParagraphBox): number => 1.2 * paragraph.fontSize;

    expect(box).toHaveLength(6);
    // The first paragraph's authored line before and the last one's line after are dropped.
    expect(box[0]!.marginTop).toBe(0);
    expect(box[5]!.marginBottom).toBe(0);
    // Half a line before, in the paragraph's own 18pt.
    expect(box[1]!.fontSize).toBeCloseTo(24, 1);
    expect(gap(box, 1)).toBeCloseTo(0.5 * line(box[1]!), 0);
    // A line before a paragraph holding a 28pt word counts 28pt.
    expect(box[2]!.fontSize).toBeCloseTo(28 * 4 / 3, 1);
    expect(gap(box, 2)).toBeCloseTo(line(box[2]!), 0);
    // A line after plus a line before: both, not the larger.
    expect(gap(box, 4)).toBeCloseTo(line(box[3]!) + line(box[4]!), 0);

    // The same in a table cell, whose paragraphs must not collapse their margins either.
    expect(cell).toHaveLength(3);
    expect(cell[0]!.marginTop).toBe(0);
    expect(gap(cell, 1)).toBeCloseTo(0, 0);
    expect(gap(cell, 2)).toBeCloseTo(0.5 * line(cell[1]!) + 0.5 * line(cell[2]!), 0);

    // Space in points follows the same edge rule: slide 5's bullets carry 10pt
    // before every paragraph, and PowerPoint shows none above the first.
    await view.evaluate(`renderFixture(${JSON.stringify(bulletSlide)})`);
    const bullets = await view.evaluate(`paragraphs("نقطة افتراضية", '[data-aared-element="shape"]')`) as ParagraphBox[];
    expect(bullets[0]!.marginTop).toBe(0);
    expect(bullets[1]!.marginTop).toBeCloseTo(10 * 4 / 3, 1);
  });
}, 30_000);
