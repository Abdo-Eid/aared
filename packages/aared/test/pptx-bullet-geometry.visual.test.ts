import { expect } from "bun:test";
import { browserModule, visual, withVisualPage } from "./harness.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { OpcPackage } from "../src/core/opc.js";
import { findCoverageSlide } from "./fixture.js";

const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const PX_PER_INCH = 96;

interface Line { text: string; p: Box; marker: Box | null; body: Box | null; markerFont: string | null }
interface Box { left: number; right: number; top: number; bottom: number }

// Ground truth, PowerPoint's own PDF of the coverage deck (probe
// line-spacing-units): the marker starts the first line at the margin plus the
// (negative) first-line indent; the text starts at the margin, or right where a
// wider marker ends; a marker larger than its text does not make the line
// taller; an empty paragraph is still a line; `marL` is counted from the right
// in a right-to-left paragraph; a `+mj-lt` bullet font is the theme's heading font.
visual("visual DOM output: PPTX bullets keep PowerPoint's hanging indent in both directions", async () => {
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const deck = await indexPptx(pkg);
  const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
  const [{ slide: bulletSlide }, { slide: rtlSlide }] = await Promise.all([
    findCoverageSlide(slides, "bullets"),
    findCoverageSlide(slides, "rtl-text"),
  ]);
  const module = await browserModule(`${import.meta.dir}/../src/pptx/index.ts`);
  await withVisualPage({
    modules: { "/pptx.mjs": module },
    page: `<!doctype html><style>html,body{margin:0}#host{width:1280px}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        window.rendered = null;
        window.renderFixture = async (model) => {
          window.rendered?.destroy();
          window.rendered = await render({ slide: model, target: document.querySelector("#host") });
        };
        // Every paragraph of the shape holding the given text, with its marker and text boxes.
        window.lines = (text) => {
          const shape = [...document.querySelectorAll('[data-aared-element="shape"]')].find((s) => s.textContent.includes(text));
          const box = (e) => { if (!e) return null; const { left, right, top, bottom } = e.getBoundingClientRect(); return { left, right, top, bottom }; };
          return [...shape.querySelectorAll("p")].map((p) => {
            const marker = p.querySelector('[data-aared-element="bullet-marker"]');
            const body = p.querySelector('[data-aared-element="paragraph-text"]');
            return { text: body?.textContent ?? "", p: box(p), marker: box(marker), body: body?.textContent ? box(body) : null, markerFont: marker ? getComputedStyle(marker).fontFamily : null };
          });
        };
      </script>`,
    width: 1280,
    height: 720,
  }, async (view) => {
    await view.evaluate(`renderFixture(${JSON.stringify(bulletSlide)})`);
    const list = await view.evaluate(`lines("نقطة افتراضية")`) as Line[];
    const [plain, red, square, check, blank, one] = list;
    const hanging = 0.25 * PX_PER_INCH;
    // Every marker starts at the same place: the margin plus the hanging indent.
    for (const line of [plain, red, square, check, one]) expect(line!.marker!.left).toBeCloseTo(plain!.marker!.left, 0);
    // A marker narrower than the hanging gap: the text starts at the margin.
    expect(plain!.body!.left - plain!.marker!.left).toBeCloseTo(hanging, 0);
    expect(red!.body!.left - red!.marker!.left).toBeCloseTo(hanging, 0);
    // A wider one (the 42pt Wingdings square): the text starts where it ends.
    expect(square!.marker!.right - square!.marker!.left).toBeGreaterThan(hanging + 1);
    expect(square!.body!.left).toBeCloseTo(square!.marker!.right, 0);
    // The 150% markers do not make their 28pt lines taller.
    const height = (line: Line): number => line.p.bottom - line.p.top;
    for (const line of [red, square, check]) expect(height(line!)).toBeCloseTo(height(plain!), 0);
    // An empty paragraph is a 28pt line too, not nothing.
    expect(blank!.text).toBe("");
    expect(height(blank!)).toBeCloseTo(height(plain!), 0);
    // Numbers: `+mj-lt` is the theme's heading font; the text starts at the 0.5625in margin.
    expect(one!.markerFont).toStartWith(`"Aptos Display"`);
    expect(one!.body!.left - one!.marker!.left).toBeCloseTo(0.5625 * PX_PER_INCH, 0);

    // Right to left: the marker starts at the right, the text ends 0.375in in from there.
    await view.evaluate(`renderFixture(${JSON.stringify(rtlSlide)})`);
    const rtl = (await view.evaluate(`lines("Next review")`) as Line[]).filter((line) => line.marker !== null);
    const rightAligned = rtl.filter((line) => !line.text.includes("Next review"));
    expect(rightAligned.length).toBeGreaterThan(0);
    for (const line of rightAligned) {
      expect(line.marker!.right).toBeCloseTo(rightAligned[0]!.marker!.right, 0);
      expect(line.marker!.left - line.body!.right).toBeCloseTo(0, 0);
      expect(line.marker!.right - line.marker!.left).toBeCloseTo(0.375 * PX_PER_INCH, 0);
    }
    // A left-aligned right-to-left line keeps its marker just past its text's right end.
    const left = rtl.find((line) => line.text.includes("Next review"))!;
    expect(left.marker!.left).toBeCloseTo(left.body!.right, 0);
    expect(left.marker!.right - left.marker!.left).toBeCloseTo(0.375 * PX_PER_INCH, 0);
  });
}, 30_000);
