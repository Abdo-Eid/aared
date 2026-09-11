import { expect, test } from "bun:test";
import { open } from "../src/pptx/index.js";

const ARABIC_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-pptx-rtl-coverage.pptx`;

test("search truncates long matches to 120 characters", async () => {
  const deck = await open(await Bun.file(ARABIC_FIXTURE).arrayBuffer());
  try {
    const matches = await deck.search("slide");
    for (const match of matches) {
      expect(match.text.length).toBeLessThanOrEqual(120);
    }
  } finally {
    deck.close();
  }
});

test("search returns at most one match per slide", async () => {
  const deck = await open(await Bun.file(ARABIC_FIXTURE).arrayBuffer());
  try {
    const matches = await deck.search("slide");
    const slidesSeen = new Set<number>();
    for (const match of matches) {
      expect(slidesSeen.has(match.slide)).toBe(false);
      slidesSeen.add(match.slide);
    }
  } finally {
    deck.close();
  }
});

test("search with abort signal cancels between slides", async () => {
  const deck = await open(await Bun.file(ARABIC_FIXTURE).arrayBuffer());
  try {
    const controller = new AbortController();
    const promise = deck.search("slide", { signal: controller.signal });
    controller.abort();
    const error = await promise.catch((e) => e);
    expect(error.name).toBe("AbortError");
  } finally {
    deck.close();
  }
});

test("search for empty string returns no results", async () => {
  const deck = await open(await Bun.file(ARABIC_FIXTURE).arrayBuffer());
  try {
    expect(await deck.search("")).toEqual([]);
  } finally {
    deck.close();
  }
});
