import { expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide, parseTextBody } from "../src/core/pptx.js";
import { parseXml } from "../src/core/xml.js";

const PRODUCTION = `${import.meta.dir}/../../../fixtures/post-production-review.pptx`;
const COVERAGE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

test("a:fld runs keep document order and mark slide-number and date fields", async () => {
  const pkg = await OpcPackage.open(await Bun.file(PRODUCTION).arrayBuffer());
  const { theme } = await indexPptx(pkg);
  const body = parseTextBody(
    parseXml(
      `<a:txBody><a:bodyPr/><a:lstStyle/><a:p>` +
        `<a:r><a:rPr/><a:t>Slide </a:t></a:r>` +
        `<a:fld id="{A}" type="slidenum"><a:rPr/><a:t>‹#›</a:t></a:fld>` +
        `<a:r><a:rPr/><a:t> · </a:t></a:r>` +
        `<a:fld id="{B}" type="datetime1"><a:rPr/><a:t>12 May 2026</a:t></a:fld>` +
        `<a:fld id="{C}" type="datetimeFigureOut"><a:rPr/><a:t>5/12/2026</a:t></a:fld>` +
        `<a:fld id="{D}" type="slidecount"><a:rPr/><a:t>7</a:t></a:fld>` +
        `</a:p></a:txBody>`,
    ),
    theme,
  );
  const runs = body.paragraphs[0]!.runs;
  expect(runs.map((run) => run.text)).toEqual(["Slide ", "‹#›", " · ", "12 May 2026", "5/12/2026", "7"]);
  expect(runs.map((run) => run.field)).toEqual([undefined, "slidenum", undefined, "datetime", "datetime", undefined]);
});

test("real PowerPoint slide-number placeholders parse as slidenum fields matching the slide's number", async () => {
  const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
  const deck = await indexPptx(pkg);
  for (const [index] of deck.slideParts.entries()) {
    const slide = await parsePptxSlide(pkg, deck, index);
    // The blank-layout group-appearance slide deliberately has no number
    // placeholder. Every earlier coverage slide retains its native field.
    if (slide.elements.some((element) => element.name === "Group fill inheritance")) {
      expect(slide.elements.some((element) => element.kind === "shape" && element.placeholder?.type === "sldNum")).toBe(false);
      continue;
    }
    const fields = slide.elements.flatMap((element) =>
      element.kind === "shape"
        ? (element.text?.paragraphs ?? []).flatMap((paragraph) => paragraph.runs.filter((run) => run.field === "slidenum"))
        : [],
    );
    expect(fields).toHaveLength(1);
    // The file caches the number it was saved with; the renderer shows `slide.number`.
    expect(fields[0]!.text).toBe(String(index + 1));
    expect(slide.number).toBe(index + 1);
  }
});

test("a slide-number placeholder takes the master's position even though master and slide number placeholders differently", async () => {
  // The master numbers date/footer/slide-number 2/3/4; layouts and slides use
  // 10/11/12. Matching by index never found the master, so the number was
  // drawn at the slide's top-left corner instead of bottom-right (PowerPoint).
  const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
  const deck = await indexPptx(pkg);
  const slide = await parsePptxSlide(pkg, deck, 0);
  const number = slide.elements.find((element) => element.kind === "shape" && element.placeholder?.type === "sldNum");
  expect(number).toBeDefined();
  const { bounds, text } = number as { bounds?: { x: number; y: number; width: number; height: number }; text?: { paragraphs: Array<{ align?: string }> } };
  expect(bounds).toBeDefined();
  // Bottom-right of the slide, right-aligned.
  expect(bounds!.x + bounds!.width).toBeGreaterThan(deck.size.width - 1.5);
  expect(bounds!.y).toBeGreaterThan(deck.size.height - 1);
  expect(text?.paragraphs[0]?.align).toBe("right");
});
