import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { groupChildMapping } from "../src/pptx/group-mapping.js";
import type { GroupModel, SlideElementModel } from "../src/core/model.js";
import { findCoverageSlide } from "./fixture.js";
import { withPart } from "./mutate.js";

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

describe("shape-tree elements that state no id", () => {
  // PowerPoint draws a group, a shape and a picture that have no `p:cNvPr/@id` (probe `group-id`; the
  // shape and picture after a repair prompt); aared keeps such an element and gives it an id of its own
  // (`auto-N`, from its position), never a real numeric one.
  async function withoutId(strip: (xml: string) => string) {
    const source = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
    const deck = await indexPptx(source);
    const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(source, deck, index)));
    const { index } = await findCoverageSlide(slides, "shapes");
    const pkg = await withPart("pptx-coverage.pptx", deck.slideParts[index]!, strip);
    const edited = await indexPptx(pkg);
    const slide = await parsePptxSlide(pkg, edited, index);
    return Object.assign(slide.elements, { missingIds: slide.missingIds });
  }
  const idsOf = (elements: readonly SlideElementModel[]): string[] =>
    elements.flatMap((element) => [element.id, ...(element.kind === "group" ? idsOf(element.children) : [])]);

  test("a group with no id opens, with its children, and every id stays unique", async () => {
    const elements = await withoutId((xml) => {
      expect(xml).toContain('<p:cNvPr id="4" name="Group 3"');
      return xml.replace('<p:cNvPr id="4" name="Group 3"', '<p:cNvPr name="Group 3"');
    });
    const group = elements.find((element): element is GroupModel => element.kind === "group" && element.name === "Group 3");
    if (group === undefined) throw new Error("the group with no id was dropped");
    expect(group.id).toMatch(/^auto-\d+$/);
    expect(group.transform).toEqual({ rotation: 20 });
    expect(group.children).toMatchObject([{ name: "Hexagon 42" }, { name: "Octagon 44" }]);
    const ids = idsOf(elements);
    expect(new Set(ids).size).toBe(ids.length);
    // The slide says how many elements it had to name (the renderer's `pptx-missing-id` warning reads it).
    expect(elements.missingIds).toBe(1);
  });

  test("a slide whose every id is stated carries no count", async () => {
    expect((await withoutId((xml) => xml)).missingIds).toBeUndefined();
  });

  test("a shape or a picture with no id opens too, and a group with no p:cNvPr at all has an empty name", async () => {
    const elements = await withoutId((xml) => {
      expect(xml).toContain('<p:cNvPr id="29" name="Rectangle 28"');
      return xml.replace('<p:cNvPr id="29" name="Rectangle 28"', '<p:cNvPr name="Rectangle 28"');
    });
    const group = elements.find((element): element is GroupModel => element.kind === "group" && element.id === "3");
    const rectangle = group?.children.find((child) => child.name === "Rectangle 28");
    if (rectangle === undefined) throw new Error("the shape with no id was dropped");
    expect(rectangle.id).toMatch(/^auto-\d+$/);
    expect(elements.missingIds).toBe(1);
    expect(group!.children.map((child) => child.name)).toEqual(["Rectangle 28", "Rectangle: Rounded Corners 30"]);
    const ids = idsOf(elements);
    expect(new Set(ids).size).toBe(ids.length);

    // A picture too: the slide's own `Picture 67` (a top-level pic), id stripped.
    const picture = await withoutId((xml) => {
      expect(xml).toContain('<p:cNvPr id="68" name="Picture 67">');
      return xml.replace('<p:cNvPr id="68" name="Picture 67">', '<p:cNvPr name="Picture 67">');
    });
    const kept = picture.find((element) => element.kind === "picture" && element.name === "Picture 67");
    if (kept === undefined) throw new Error("the picture with no id was dropped");
    expect(kept.id).toMatch(/^auto-\d+$/);
    expect(picture.missingIds).toBe(1);
    const pictureIds = idsOf(picture);
    expect(new Set(pictureIds).size).toBe(pictureIds.length);

    // A group whose non-visual properties hold no p:cNvPr at all.
    const bare = await withoutId((xml) => xml.replace(/<p:cNvPr id="4" name="Group 3">.*?<\/p:cNvPr>/s, ""));
    const unnamed = bare.find((element): element is GroupModel => element.kind === "group" && element.name === "");
    expect(unnamed?.id).toMatch(/^auto-\d+$/);
    expect(bare.missingIds).toBe(1);
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
