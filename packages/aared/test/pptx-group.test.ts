import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { groupChildMapping } from "../src/pptx/group-mapping.js";
import type { GroupModel } from "../src/core/model.js";
import { findCoverageSlide } from "./fixture.js";

const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

async function shapesSlide() {
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const deck = await indexPptx(pkg);
  const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
  return (await findCoverageSlide(slides, "shapes")).slide;
}

describe("PPTX grouped shapes", () => {
  test("parses the PowerPoint-grouped pair as a group with nested shapes", async () => {
    const slide = await shapesSlide();
    const group = slide.elements.find((element): element is GroupModel => element.kind === "group" && element.id === "3");
    if (group === undefined) throw new Error("fixture structure changed: expected Group 2 (id 3) on the shapes slide");
    expect(group.name).toBe("Group 2");
    expect(group.transform).toBeUndefined();
    // Identity child space: PowerPoint wrote off == chOff, ext == chExt.
    expect(group.bounds).toEqual({ x: 4070021 / 914400, y: 1035031 / 914400, width: 1162310 / 914400, height: 251640 / 914400 });
    expect(group.childSpace).toEqual(group.bounds);
    expect(group.children.map((child) => child.kind)).toEqual(["shape", "shape"]);
    expect(group.children).toMatchObject([
      { id: "29", name: "Rectangle 28", geometry: { preset: "rect" } },
      { id: "31", name: "Rectangle: Rounded Corners 30", geometry: { preset: "roundRect" } },
    ]);
  });

  test("parses the rotated, scaled group with a non-identity child space", async () => {
    const slide = await shapesSlide();
    const group = slide.elements.find((element): element is GroupModel => element.kind === "group" && element.id === "4");
    if (group === undefined) throw new Error("fixture structure changed: expected Group 3 (id 4) on the shapes slide");
    expect(group.name).toBe("Group 3");
    expect(group.transform).toEqual({ rotation: 20 });
    expect(group.bounds).toEqual({ x: 4077823 / 914400, y: 2043714 / 914400, width: 1627133 / 914400, height: 352296 / 914400 });
    expect(group.childSpace).toEqual({ x: 4692619 / 914400, y: 1892333 / 914400, width: 1162238 / 914400, height: 251640 / 914400 });
    expect(group.children).toMatchObject([
      { id: "43", name: "Hexagon 42", geometry: { preset: "hexagon" } },
      { id: "45", name: "Octagon 44", geometry: { preset: "octagon" } },
    ]);
  });

  test("keeps ink-fallback pictures nested in their groups, in z-order", async () => {
    const slide = await shapesSlide();
    const groups = slide.elements.filter((element): element is GroupModel => element.kind === "group");
    expect(groups.map((group) => group.id)).toEqual(["3", "4", "56", "59", "63"]);
    for (const id of ["56", "59", "63"]) {
      const group = groups.find((candidate) => candidate.id === id);
      if (group === undefined) throw new Error(`fixture structure changed: expected ink group ${id}`);
      expect(group.children.length).toBeGreaterThan(0);
      expect(group.children.every((child) => child.kind === "picture")).toBe(true);
    }
    // No hoisting: the only top-level pictures are the direct ink fallback
    // (id 60) and the direct future-coverage image (id 68).
    expect(slide.elements.filter((element) => element.kind === "picture").map((element) => element.id)).toEqual(["60", "68"]);
  });
});

describe("groupChildMapping", () => {
  test("maps a scaled child space onto group bounds", () => {
    expect(
      groupChildMapping({ x: 0, y: 0, width: 2, height: 1 }, { x: 0.5, y: 0.25, width: 1, height: 1 }),
    ).toBe("translate(-96px, -24px) scale(2, 1)");
  });

  test("maps an identity child space to a unit scale with origin translation", () => {
    expect(groupChildMapping({ x: 1, y: 2, width: 3, height: 4 }, { x: 1, y: 2, width: 3, height: 4 })).toBe(
      "translate(-96px, -192px) scale(1, 1)",
    );
  });

  test("falls back to identity for missing or degenerate spaces", () => {
    expect(groupChildMapping(undefined, { x: 0, y: 0, width: 1, height: 1 })).toBe("");
    expect(groupChildMapping({ x: 0, y: 0, width: 1, height: 1 }, undefined)).toBe("");
    expect(groupChildMapping({ x: 0, y: 0, width: 0, height: 1 }, { x: 0, y: 0, width: 1, height: 1 })).toBe("");
    expect(groupChildMapping({ x: 0, y: 0, width: 1, height: 1 }, { x: 0, y: 0, width: 1, height: 0 })).toBe("");
  });
});
