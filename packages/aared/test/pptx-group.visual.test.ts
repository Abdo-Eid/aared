import { expect } from "bun:test";
import { browserModule, visual, withVisualPage } from "./harness.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { OpcPackage } from "../src/core/opc.js";
import { findCoverageSlide } from "./fixture.js";

const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

// Identity child space from the fixture: off == chOff, ext == chExt, so the
// scale is 1 and the mapping only translates the child origin out.
const EMU_X = 4070021 / 914400;
const EMU_Y = 1035031 / 914400;
const SCALE = 1162310 / 1162310;

visual("visual DOM output: PPTX nests grouped shapes in a mapped group wrapper", async () => {
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const deck = await indexPptx(pkg);
  const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
  const { slide: shapesSlide } = await findCoverageSlide(slides, "shapes");
  const module = await browserModule(`${import.meta.dir}/../src/pptx/index.ts`);
  await withVisualPage({
    modules: { "/pptx.mjs": module },
    page: `<!doctype html><style>html,body{margin:0}#host{width:1280px}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        window.renderedFixture = null;
        window.renderFixture = async (model) => {
          window.renderedFixture?.destroy();
          window.renderedFixture = await render({ slide: model, target: document.querySelector("#host") });
        };
      </script>`,
    width: 1280,
    height: 720,
  }, async (view) => {
    await view.evaluate(`renderFixture(${JSON.stringify(shapesSlide)})`);
    const output = await view.evaluate(`(() => {
      const readGroup = (id) => {
        const group = document.querySelector('[data-aared-element="group"][data-aared-group-id="' + id + '"]');
        const inner = group?.querySelector('[data-aared-element="group-children"]');
        return {
          found: group !== null && group !== undefined,
          transform: group ? group.style.transform : null,
          nestedShapes: [...(group?.querySelectorAll('[data-aared-element="shape"]') ?? [])]
            .map((shape) => shape.dataset.aaredShapeId),
          mapping: inner ? inner.style.transform : null,
        };
      };
      const pictures = [...document.querySelectorAll('[data-aared-element="picture"]')].map((picture) => ({
        id: picture.dataset.aaredPictureId,
        inGroup: picture.closest('[data-aared-element="group"]')?.dataset.aaredGroupId ?? null,
      }));
      return { identity: readGroup("3"), scaled: readGroup("4"), pictures };
    })()`) as {
      identity: { found: boolean; transform: string | null; nestedShapes: (string | undefined)[]; mapping: string | null };
      scaled: { found: boolean; transform: string | null; nestedShapes: (string | undefined)[]; mapping: string | null };
      pictures: { id?: string; inGroup: string | null }[];
    };
    expect(output.identity.found).toBe(true);
    expect(output.identity.nestedShapes).toEqual(["29", "31"]);
    // The browser quantizes transform components on serialize, so compare
    // numerically rather than as an exact string.
    const components = output.identity.mapping === null ? null : output.identity.mapping.match(/translate\((-?\d+\.?\d*)px, (-?\d+\.?\d*)px\) scale\((-?\d+\.?\d*), (-?\d+\.?\d*)\)/);
    expect(components).not.toBeNull();
    expect(Math.abs(Number(components![1]) - -(EMU_X * SCALE * 96))).toBeLessThan(0.001);
    expect(Math.abs(Number(components![2]) - -(EMU_Y * SCALE * 96))).toBeLessThan(0.001);
    expect(Number(components![3])).toBe(1);
    expect(Number(components![4])).toBe(1);
    // The rotated, scaled group carries the group rotation on its wrapper and
    // a ~1.4x child-space mapping on the inner div.
    expect(output.scaled.found).toBe(true);
    expect(output.scaled.transform).toContain("rotate(20deg)");
    expect(output.scaled.nestedShapes).toEqual(["43", "45"]);
    const scaledComponents = output.scaled.mapping === null ? null : output.scaled.mapping.match(/translate\((-?\d+\.?\d*)px, (-?\d+\.?\d*)px\) scale\((-?\d+\.?\d*), (-?\d+\.?\d*)\)/);
    expect(scaledComponents).not.toBeNull();
    expect(Math.abs(Number(scaledComponents![1]) - -((4692619 / 914400) * (1627133 / 1162238) * 96))).toBeLessThan(0.01);
    expect(Math.abs(Number(scaledComponents![2]) - -((1892333 / 914400) * (352296 / 251640) * 96))).toBeLessThan(0.01);
    expect(Number(scaledComponents![3])).toBeCloseTo(1627133 / 1162238, 3);
    expect(Number(scaledComponents![4])).toBeCloseTo(352296 / 251640, 3);
    // Z-order guard: grouped pictures stay inside their group wrapper instead
    // of hoisted to the scene; only the two direct pictures sit outside.
    for (const picture of output.pictures) {
      if (picture.id === "60" || picture.id === "68") expect(picture.inGroup).toBeNull();
      else expect(picture.inGroup).not.toBeNull();
    }
  });
});
