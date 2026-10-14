import { expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import type { GroupModel, ShapeModel, SlideModel } from "../src/core/model.js";
import type { DrawingTheme } from "../src/core/theme.js";
import { parseXml } from "../src/core/xml.js";
import { resolveDirectFill } from "../src/core/fill.js";
import { readDrawingEffects } from "../src/core/drawingml-effects.js";

export async function groupAppearanceSlide(): Promise<SlideModel> {
  const pkg = await OpcPackage.open(await Bun.file(`${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`).arrayBuffer());
  const deck = await indexPptx(pkg);
  for (let index = 0; index < deck.slideParts.length; index++) {
    const slide = await parsePptxSlide(pkg, deck, index);
    if (slide.elements.some((e) => e.name === "Group fill inheritance")) return slide;
  }
  throw new Error("native group appearance coverage is missing");
}

export function namedGroup(slide: SlideModel, name: string): GroupModel {
  const group = slide.elements.find((e): e is GroupModel => e.kind === "group" && e.name === name);
  if (group === undefined) throw new Error(`missing group ${name}`);
  return group;
}

test("native grpFill resolves parent paint while a child's direct fill wins", async () => {
  const group = namedGroup(await groupAppearanceSlide(), "Group fill inheritance");
  const shapes = group.children as ShapeModel[];
  expect(shapes[0]?.fill).toEqual({ kind: "solid", color: "#156082" });
  expect(shapes[0]?.fillBounds).toEqual(group.childSpace);
  expect(shapes[1]?.fill).toEqual({ kind: "solid", color: "#BE2832" });
  expect(shapes[1]?.fillBounds).toBeUndefined();
  // Native PNG has no ink; PDF trace opacity is zero (extracted RGB is black).
  expect(shapes[1]?.text?.paragraphs[0]?.runs[0]?.noFill).toBe(true);
  expect(shapes[1]?.text?.paragraphs[0]?.runs[0]?.groupFill).toBe(true);
});

test("native nested groups carry grpFill through the ancestry without changing child geometry", async () => {
  const group = namedGroup(await groupAppearanceSlide(), "Group nested outer");
  const nested = group.children.find((e): e is GroupModel => e.kind === "group");
  expect(nested).toBeDefined();
  for (const child of nested!.children) {
    expect(child.kind).toBe("shape");
    expect((child as ShapeModel).fill?.kind).toBe("linear-gradient");
    expect((child as ShapeModel).fillBounds).toEqual(group.childSpace);
    expect(child.bounds?.width).toBeGreaterThan(0);
  }
  const transparent = group.children.find((e) => e.name === "Nested fill transparent") as ShapeModel;
  expect(transparent.fill).toBeUndefined();
});

test("group effects remain on the composite and native child effects remain on their own shapes", async () => {
  const slide = await groupAppearanceSlide();
  const shadow = namedGroup(slide, "Group outer shadow");
  expect(shadow.effects?.outerShadow?.color).toBe("rgb(0 0 0 / 0.6)");
  expect(shadow.effects?.outerShadow?.blurPt).toBe(4);
  expect(shadow.effects?.outerShadow?.dxPt).toBeCloseTo(8, 4);
  expect(shadow.effects?.outerShadow?.dyPt).toBeCloseTo(6, 4);
  for (const child of shadow.children) expect((child as ShapeModel).effects).toBeUndefined();
  expect(namedGroup(slide, "Group glow").effects?.glow).toEqual({ color: "rgb(220 120 10 / 0.8)", radiusPt: 8 });
  expect(namedGroup(slide, "Group reflection").effects?.reflection).toMatchObject({ startOpacity: .5, endPosition: .55 });
  expect(namedGroup(slide, "Group soft edge").effects?.softEdgePt).toBe(5);
  const inner = namedGroup(slide, "Group inner shadow");
  expect(inner.effects).toBeUndefined();
  for (const child of inner.children) expect((child as ShapeModel).effects?.innerShadow?.blurPt).toBe(3);
});

test("group fill distinguishes missing context, explicit no paint, and direct overrides", () => {
  const theme: DrawingTheme = { colors: {}, fonts: {
    major: {latin:"",complexScript:"",scripts:{}}, minor: {latin:"",complexScript:"",scripts:{}},
  } };
  const owner = { relationships: [], mediaContentType: () => "image/png", inherited: false };
  const inherited = parseXml('<spPr><grpFill/></spPr>');
  expect(resolveDirectFill(inherited, theme, owner)).toEqual({kind:"unsupported",source:"group"});
  expect(resolveDirectFill(inherited, theme, {...owner,groupFill:null})).toBeNull();
  expect(resolveDirectFill(parseXml('<spPr><noFill/></spPr>'), theme,
    {...owner,groupFill:{kind:"solid",color:"#156082"}})).toBeNull();
  expect(resolveDirectFill(parseXml('<spPr/>'), theme,
    {...owner,groupFill:{kind:"solid",color:"#156082"}})).toBeUndefined();
  expect(readDrawingEffects(parseXml('<grpSpPr><effectLst><outerShdw sx="120000"><srgbClr val="000000"/></outerShdw></effectLst></grpSpPr>'), theme))
    .toEqual({unsupported:["shadow-transform"]});
  expect(readDrawingEffects(parseXml('<grpSpPr><effectDag><cont/></effectDag></grpSpPr>'), theme))
    .toEqual({unsupported:["effectDag"]});
});
