import { describe, expect, test } from "bun:test";
import { MIN_SLOT_PX, fitUnits } from "../src/docx/split-fit.js";

const SPACE = 4;

describe("splitting a line around a picture", () => {
  test("fills each slot in reading order with as many whole words as fit", () => {
    // Words of width 30, slots of 100: 30+4+30+4+30 = 98 fits three, the next slot takes the following three.
    expect(fitUnits([30, 30, 30, 30, 30, 30, 30], SPACE, [{ widthPx: 100 }, { widthPx: 100 }])).toEqual([3, 3]);
  });

  test("counts the spaces between words in a slot, not before the first", () => {
    expect(fitUnits([48, 48], SPACE, [{ widthPx: 100 }])).toEqual([2]); // 48 + 4 + 48 = 100
    expect(fitUnits([49, 48], SPACE, [{ widthPx: 100 }])).toEqual([1]); // 101 does not fit
  });

  test("a slot too narrow for the next word stays empty and the word goes to the next slot", () => {
    // 80 does not fit in 60; in the next slot it takes 80 + 4 + 10 = 94.
    expect(fitUnits([80, 10], SPACE, [{ widthPx: 60 }, { widthPx: 100 }])).toEqual([0, 2]);
  });

  test("a slot below the minimum takes nothing, even for a short word", () => {
    expect(fitUnits([10], SPACE, [{ widthPx: MIN_SLOT_PX - 1 }, { widthPx: 100 }])).toEqual([0, 1]);
  });

  test("only a forced slot takes a word wider than itself, so a full-width line always advances", () => {
    expect(fitUnits([150], SPACE, [{ widthPx: 100 }])).toEqual([0]);
    expect(fitUnits([150, 20], SPACE, [{ widthPx: 100, force: true }])).toEqual([1]);
  });

  test("stops when the words run out and resumes from `start`", () => {
    expect(fitUnits([30, 30], SPACE, [{ widthPx: 100 }, { widthPx: 100 }])).toEqual([2, 0]);
    expect(fitUnits([30, 30, 30], SPACE, [{ widthPx: 100 }], 2)).toEqual([1]);
  });
});
