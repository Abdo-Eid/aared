import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePicture, parsePptxSlide } from "../src/core/pptx.js";
import { parseXml } from "../src/core/xml.js";
import type { PictureModel } from "../src/core/model.js";
import { flattenElements } from "./fixture.js";

const COVERAGE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const PRODUCTION = `${import.meta.dir}/../../../fixtures/post-production-review.pptx`;

async function openDeck(path: string) {
  const pkg = await OpcPackage.open(await Bun.file(path).arrayBuffer());
  const deck = await indexPptx(pkg);
  return { pkg, deck };
}

describe("PPTX pictures", () => {
  test("resolves production pictures through slide rels with MIME types and bounds", async () => {
    const { pkg, deck } = await openDeck(PRODUCTION);
    const slide1 = await parsePptxSlide(pkg, deck, 0);
    const pictures = slide1.elements.filter((element): element is PictureModel => element.kind === "picture");
    expect(pictures).toHaveLength(2);
    expect(pictures[0]).toMatchObject({
      id: "2",
      name: "Image 0",
      mediaId: "ppt/media/image-1-1.png",
      contentType: "image/png",
      bounds: { x: 0.75, y: 0.7, width: 0.9, height: 0.9 },
    });
    expect(pictures[1]).toMatchObject({
      id: "8",
      mediaId: "ppt/media/image-1-2.jpeg",
      contentType: "image/jpeg",
    });
    expect(pictures[0]?.crop).toBeUndefined();
    expect(pictures[0]?.transform).toBeUndefined();
    // Media bytes flow through the existing package contract.
    const bytes = await pkg.readPart(pictures[0]!.mediaId);
    expect(bytes[0]).toBe(0x89);
    expect(bytes[1]).toBe(0x50);
  });

  test("keeps picture z-order and exposes the chart slide separately", async () => {
    const { pkg, deck } = await openDeck(PRODUCTION);
    const slide1 = await parsePptxSlide(pkg, deck, 0);
    expect(slide1.elements.map((element) => element.kind)).toEqual([
      "picture",
      "shape",
      "shape",
      "shape",
      "shape",
      "shape",
      "picture",
    ]);
    const slide4 = await parsePptxSlide(pkg, deck, 3);
    expect(slide4.elements.some((element) => element.kind === "chart")).toBe(true);
    expect(slide4.elements.some((element) => element.kind === "picture")).toBe(false);
  });

  test("renders ink fallbacks as pictures instead of silently dropping them", async () => {
    const { pkg, deck } = await openDeck(COVERAGE);
    const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
    // Ink fallbacks live inside groups; flatten so nested pictures still count.
    const pictures = slides.flatMap((slide) => flattenElements(slide.elements)).filter((element) => element.kind === "picture");
    // 8 ink fallbacks (grouped) plus the direct future-coverage image.
    expect(pictures).toHaveLength(9);
    const direct = flattenElements(slides[2]!.elements).find((element) => element.kind === "picture" && element.id === "68");
    expect(direct).toMatchObject({ mediaId: "ppt/media/image9.png", contentType: "image/png" });
  });

  test("parses srcRect crops as edge fractions", async () => {
    const { pkg, deck } = await openDeck(PRODUCTION);
    const slidePart = deck.slideParts[0]!;
    const relationships = await pkg.relationships(slidePart);
    const rel = relationships.find((candidate) => candidate.type.endsWith("/image"))!;
    const pic = parseXml(
      `<p:pic><p:nvPicPr><p:cNvPr id="99" name="Crop"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="${rel.id}"><a:srcRect l="10000" t="20000" r="5000" b="0"/></a:blip>` +
        `<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
    );
    const picture = parsePicture(pic, relationships, pkg);
    expect(picture).toMatchObject({
      id: "99",
      mediaId: rel.target,
      crop: { left: 0.1, top: 0.2, right: 0.05, bottom: 0 },
      bounds: { x: 1, y: 1, width: 2, height: 2 },
    });
  });

  test("parses picture rotation and flips from its xfrm", async () => {
    const { pkg, deck } = await openDeck(PRODUCTION);
    const slidePart = deck.slideParts[0]!;
    const relationships = await pkg.relationships(slidePart);
    const rel = relationships.find((candidate) => candidate.type.endsWith("/image"))!;
    const pic = parseXml(
      `<p:pic><p:nvPicPr><p:cNvPr id="100" name="Rotated"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="${rel.id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr><a:xfrm rot="5400000" flipH="1"><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
    );
    expect(parsePicture(pic, relationships, pkg)).toMatchObject({
      transform: { rotation: 90, flipHorizontal: true },
    });
  });

  test("serves picture bytes through the worker media route", async () => {
    const { open } = await import("../src/pptx/index.js");
    const deck = await open(await Bun.file(PRODUCTION).arrayBuffer());
    try {
      const slide = await deck.slides.get(0).read();
      const picture = slide.elements.find((element) => element.kind === "picture");
      if (picture?.kind !== "picture") throw new Error("expected a picture element");
      const bytes = await deck.media.get(picture.mediaId);
      expect(bytes.byteLength).toBeGreaterThan(100);
      expect(new Uint8Array(bytes)[0]).toBe(0x89);
    } finally {
      deck.close();
    }
  });
});

describe("PPTX notes", () => {
  test("extracts production notes as text, including Arabic", async () => {
    const { pkg, deck } = await openDeck(PRODUCTION);
    const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
    expect(slides[0]?.notes).toContain("الافتتاحية تُعرض قبل بدء النقاش");
    expect(slides[1]?.notes).toBe("Four items, twenty minutes. Do not open the grade discussion here.");
    expect(slides[3]?.notes).toContain("Episode 06 is not late");
    expect(slides[6]?.notes).toContain("القرارات مطلوبة قبل الأحد");
  });

  test("leaves notes undefined when a deck carries no notes slides", async () => {
    const { pkg, deck } = await openDeck(COVERAGE);
    const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
    expect(slides.every((slide) => slide.notes === undefined)).toBe(true);
  });
});
