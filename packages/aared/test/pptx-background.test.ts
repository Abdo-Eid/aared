import { expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";

const COVERAGE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const PRODUCTION = `${import.meta.dir}/../../../fixtures/post-production-review.pptx`;

async function slideBackgrounds(path: string): Promise<Array<Awaited<ReturnType<typeof parsePptxSlide>>["background"]>> {
  const pkg = await OpcPackage.open(await Bun.file(path).arrayBuffer());
  const deck = await indexPptx(pkg);
  return Promise.all(deck.slideParts.map(async (_, index) => (await parsePptxSlide(pkg, deck, index)).background));
}

test("a slide's own solid background wins (dark title and closing slides)", async () => {
  const backgrounds = await slideBackgrounds(PRODUCTION);
  expect(backgrounds).toHaveLength(7);
  expect(backgrounds[0]).toEqual({ kind: "solid", color: "#12171F" });
  expect(backgrounds[6]).toEqual({ kind: "solid", color: "#12171F" });
  for (const index of [1, 2, 3, 4, 5]) expect(backgrounds[index]).toEqual({ kind: "solid", color: "#FFFFFF" });
});

test("a slide with no background inherits the master's plain background ref", async () => {
  const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
  const deck = await indexPptx(pkg);
  // The master carries a plain `bgRef` on `bg1`, which resolves white.
  for (const index of [0, 1, 2, 3, 4, 5]) {
    const { background } = await parsePptxSlide(pkg, deck, index);
    expect(background).toEqual({ kind: "solid", color: "#FFFFFF" });
  }
});
