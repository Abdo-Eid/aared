import { expect, test } from "bun:test";
import { AaredError, isAaredError } from "../src/core/index.js";
import { open } from "../src/pptx/index.js";
import { rejection } from "./harness.js";

const ARABIC_FIXTURE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;
const PRODUCTION = `${import.meta.dir}/../../../fixtures/post-production-review.pptx`;

test("opens a deck with an addressable slide collection", async () => {
  const deck = await open(await Bun.file(ARABIC_FIXTURE).arrayBuffer());
  try {
    expect(deck.info.format).toBe("pptx");
    expect(deck.info.count).toBe(10);
    const slides = deck.slides.list();
    expect(slides).toHaveLength(10);
    expect(slides.map((slide) => slide.info.index)).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
    for (const slide of slides) {
      expect(slide.info.title).toContain("Slide");
      expect(slide.info.size.width).toBeGreaterThan(0);
    }
    expect(deck.slides.get(0)).toBe(slides[0]!);
  } finally {
    deck.close();
  }
});

test("reads slides lazily through handles", async () => {
  const deck = await open(await Bun.file(ARABIC_FIXTURE).arrayBuffer());
  try {
    const slide = await deck.slides.get(0).read();
    expect(slide.elements.length).toBeGreaterThan(0);
    expect(slide.size.width).toBe(deck.slides.get(0).info.size.width);
  } finally {
    deck.close();
  }
});

test("reports unknown slides with a not-found error", async () => {
  const deck = await open(await Bun.file(ARABIC_FIXTURE).arrayBuffer());
  try {
    const error = await rejection(deck.slides.get(2).read().then(() => deck.slides.get(19).read()));
    expect(error).toBeInstanceOf(AaredError);
    expect((error as AaredError).code).toBe("not-found");
  } finally {
    deck.close();
  }
  const reopened = await open(await Bun.file(ARABIC_FIXTURE).arrayBuffer());
  try {
    expect(() => reopened.slides.get(19)).toThrow("outside this 10-slide deck");
  } finally {
    reopened.close();
  }
});

test("rejects non-PPTX files with a typed format error", async () => {
  const error = await rejection(open(await Bun.file(XLSX_FIXTURE).arrayBuffer()));
  expect(error).toBeInstanceOf(AaredError);
  expect(isAaredError(error) && error.code).toBe("unsupported-format");
  expect(error.message).toContain("aared/xlsx");
});

test("maps corrupt archives to invalid-document errors", async () => {
  const error = await rejection(open(new ArrayBuffer(8)));
  expect(error).toBeInstanceOf(AaredError);
  expect((error as AaredError).code).toBe("invalid-document");
  expect(error.message).toContain("not a ZIP archive");
});

test("detects password-protected files as encrypted documents", async () => {
  const ole = new ArrayBuffer(8);
  new Uint8Array(ole).set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  const error = await rejection(open(ole));
  expect(error).toBeInstanceOf(AaredError);
  expect((error as AaredError).code).toBe("encrypted-document");
});

test("searches slide text", async () => {
  const deck = await open(await Bun.file(ARABIC_FIXTURE).arrayBuffer());
  try {
    expect(await deck.search("")).toEqual([]);
    const matches = await deck.search("الإنتاج");
    expect(matches.length).toBeGreaterThan(0);
    for (const match of matches) {
      expect(match.slide.index).toBeGreaterThanOrEqual(0);
      expect(match.slide.title).toBe(`Slide ${match.slide.index + 1}`);
      expect(match.text.toLowerCase()).toContain("الإنتاج".toLowerCase());
    }
  } finally {
    deck.close();
  }
});

test("serves picture media through the media collection", async () => {
  const deck = await open(await Bun.file(PRODUCTION).arrayBuffer());
  try {
    const mediaIds: string[] = [];
    for (const handle of deck.slides.list()) {
      const slide = await handle.read();
      for (const element of slide.elements) {
        if (element.kind === "picture") mediaIds.push(element.mediaId);
      }
    }
    expect(mediaIds.length).toBeGreaterThan(0);
    const bytes = await deck.media.get(mediaIds[0]!);
    expect(bytes.byteLength).toBeGreaterThan(0);
    const missing = await rejection(deck.media.get("ppt/media/not-present.png"));
    expect((missing as AaredError).code).toBe("media-error");
  } finally {
    deck.close();
  }
});

test("closes cleanly and rejects later reads", async () => {
  const deck = await open(await Bun.file(ARABIC_FIXTURE).arrayBuffer());
  deck.close();
  const error = await rejection(deck.slides.get(0).read());
  expect(error.message).toContain("document is closed");
});
