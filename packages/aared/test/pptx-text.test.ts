import { expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide, readPptxSlideText } from "../src/core/pptx.js";
import { chartText } from "../src/core/chart/text.js";
import { runsText } from "../src/core/search.js";
import type { SlideModel, TextBodyModel } from "../src/core/model.js";

// Search reads a slide's words without building its model (`readPptxSlideText`).
// The reference is the full model a render paints: each paragraph of its shapes
// and table cells, and each chart's words, in element order, groups included. A
// slide where the two disagree is a hit on words the reader can't see, or words
// they can see that search misses.

const ROOT = `${import.meta.dir}/../../..`;

/** The words of a parsed slide, and which kinds of element gave some. */
function modelText(model: SlideModel): { words: string[]; kinds: Set<string> } {
  const words: string[] = [];
  const kinds = new Set<string>();
  const body = (text: TextBodyModel, kind: string): void => {
    for (const paragraph of text.paragraphs) {
      const line = runsText(paragraph.runs).trim();
      if (line === "") continue;
      words.push(line);
      kinds.add(kind);
    }
  };
  const visit = (element: SlideModel["elements"][number], inGroup: boolean): void => {
    if (element.kind === "group") for (const child of element.children) visit(child, true);
    else if (element.kind === "shape" && element.text !== undefined) body(element.text, inGroup ? "group" : "shape");
    else if (element.kind === "table") for (const row of element.rows) for (const cell of row.cells) body(cell.text, "table");
    else if (element.kind === "chart" && element.spec !== undefined) {
      const chart = chartText(element.spec);
      words.push(...chart);
      if (chart.length > 0) kinds.add("chart");
    }
  };
  for (const element of model.elements) visit(element, false);
  return { words, kinds };
}

async function compareDeck(path: string): Promise<Set<string>> {
  const pkg = await OpcPackage.open(await Bun.file(path).arrayBuffer());
  const deck = await indexPptx(pkg);
  const kinds = new Set<string>();
  for (let index = 0; index < deck.slideParts.length; index++) {
    const reference = modelText(await parsePptxSlide(pkg, deck, index));
    expect(await readPptxSlideText(pkg, deck, index), `slide ${index + 1}`).toEqual(reference.words);
    for (const kind of reference.kinds) kinds.add(kind);
  }
  return kinds;
}

test("every deck's words read without the model equal the model's, on every slide", async () => {
  // Shapes, merged tables, a chartEx waterfall, and the appearance slide's
  // nested bilingual groups all contribute words; reflections add no words.
  expect([...(await compareDeck(`${ROOT}/fixtures/pptx-coverage.pptx`))].sort()).toEqual(["chart", "group", "shape", "table"]);
  expect(await compareDeck(`${ROOT}/fixtures/post-production-review.pptx`)).toContain("chart");
  expect(await compareDeck(`${ROOT}/tools/chart-oracle/chart-oracle.pptx`)).toContain("chart");
});

const LARGE = `${ROOT}/tools/out/search-fixtures/large-300-slides.pptx`;
test.skipIf(!(await Bun.file(LARGE).exists()))("the generated 300-slide deck's words match too (when make-fixtures.py has run)", async () => {
  expect((await compareDeck(LARGE)).size).toBeGreaterThan(0);
});

test("a text read refuses what a slide read refuses, with the same typed error", async () => {
  const pkg = await OpcPackage.open(await Bun.file(`${ROOT}/fixtures/pptx-coverage.pptx`).arrayBuffer());
  const deck = await indexPptx(pkg);
  const outside = deck.slideParts.length;
  const [text, model] = await Promise.all([
    readPptxSlideText(pkg, deck, outside).catch((error: unknown) => error),
    parsePptxSlide(pkg, deck, outside).catch((error: unknown) => error),
  ]);
  expect((text as { code?: string }).code).toBe("not-found");
  expect((text as Error).message).toBe((model as Error).message);
});
