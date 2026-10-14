/**
 * The DOCX text-box shape's own fill and outline, read from a real Word-authored
 * fixture and from mutations of it.
 *
 * The rule these pin: a shape's fill is the fill its `wps:spPr` authors itself,
 * and nothing else. Word gives every inserted shape a black outline, so
 * `a:ln/a:solidFill` sits under almost every text box, and hiding that outline
 * adds `a:extLst/a:ext/a14:hiddenLine` carrying the same black. A subtree search
 * finds either one and paints the box solid black, or in its outline colour -
 * which is how a box with no fill, a gradient or a pattern fill came out black.
 * Probe `autofit` in the gitignored `tools/probe/` measured it against Word 16.0:
 * Word paints those boxes unfilled, and aared painted four of them black.
 */

import { describe, expect, test } from "bun:test";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import type { DocxTextBoxNode } from "../src/core/model.js";
import { OpcPackage } from "../src/core/opc.js";
import { FIXTURES, withPart } from "./mutate.js";

const FIXTURE = "docx-coverage.docx";

async function textBoxes(pkg: OpcPackage): Promise<DocxTextBoxNode[]> {
  const sections = await parseDocxSections(pkg, await indexDocx(pkg));
  const root = sections.flatMap((section) => section.root);
  const anchored = root.flatMap((node) =>
    node.type === "paragraph" || node.type === "heading" ? node.floats ?? [] : [node]);
  return anchored.filter((node): node is DocxTextBoxNode => node.type === "textBox");
}

/** The fixture's filled box, the one carrying a direct `a:solidFill`. */
const FILLED = "English note: confirm the final delivery.";
/** The Word-authored box with the fill and the outline both turned off. */
const BARE = "No fill, no outline";

function boxText(box: DocxTextBoxNode): string {
  return box.content
    .flatMap((node) => node.type === "paragraph" ? node.runs.map((run) => run.text) : [])
    .join(" ");
}

function findBox(boxes: DocxTextBoxNode[], needle: string): DocxTextBoxNode {
  const found = boxes.find((box) => boxText(box).includes(needle));
  if (found === undefined) {
    throw new Error(`fixture structure changed: no text box containing "${needle}"\n${boxes.map(boxText).join("\n")}`);
  }
  return found;
}

/** The `wps:spPr` of the fixture's filled box, the one a mutation replaces. */
async function fixtureShapeProperties(): Promise<string> {
  const pkg = await OpcPackage.open(await Bun.file(`${FIXTURES}/${FIXTURE}`).arrayBuffer());
  const xml = new TextDecoder().decode(await pkg.readPart("word/document.xml"));
  const match = new RegExp(`<wps:spPr\\b[^>]*>[\\s\\S]*?</wps:spPr>(?=[\\s\\S]*${FILLED.slice(0, 20)})`).exec(xml);
  if (match === null) throw new Error("fixture structure changed: no wps:spPr for the filled text box");
  return match[0];
}

async function filledBoxWith(spPr: string): Promise<DocxTextBoxNode> {
  const original = await fixtureShapeProperties();
  const pkg = await withPart(FIXTURE, "word/document.xml", (xml) => xml.replace(original, spPr));
  return findBox(await textBoxes(pkg), FILLED);
}

describe("a DOCX text box's own fill and outline", () => {
  test("the Word-authored filled box keeps its direct solid fill", async () => {
    const pkg = await OpcPackage.open(await Bun.file(`${FIXTURES}/${FIXTURE}`).arrayBuffer());
    const boxes = await textBoxes(pkg);
    expect(boxes).toHaveLength(25);
    expect(findBox(boxes, FILLED)).toMatchObject({ fill: { kind: "solid", color: "#E8F2E8" } });
    expect(findBox(boxes, FILLED).border).toMatchObject({ color: "#3A6128", widthPt: 0.5 });
  });

  test("the Word-authored box with no fill and no outline paints neither", async () => {
    // `Fill.Visible = 0; Line.Visible = 0` is what Word writes for a plain
    // inserted text box: `a:noFill` where the fill goes, `a:noFill` in the
    // outline, and `a:extLst/a:ext/a14:hiddenLine` keeping the hidden line's
    // black for when the outline comes back. Reading any of those three as the
    // shape's fill painted this box solid black.
    const pkg = await OpcPackage.open(await Bun.file(`${FIXTURES}/${FIXTURE}`).arrayBuffer());
    const boxes = await textBoxes(pkg);
    const bare = findBox(boxes, BARE);
    expect(bare.fill).toBeUndefined();
    expect(bare.border).toBeUndefined();
    expect(boxText(bare)).toContain("صندوق بلا تعبئة ولا حد");
  });

  test("an explicit noFill paints no background even with a black outline beside it", async () => {
    const box = await filledBoxWith(
      '<wps:spPr><a:noFill/><a:ln w="6350"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></wps:spPr>');
    expect(box.fill).toBeUndefined();
    expect(box.border).toMatchObject({ color: "#000000" });
  });

  test("a hidden outline beside a direct fill does not become the fill", async () => {
    const box = await filledBoxWith(
      '<wps:spPr><a:solidFill><a:srgbClr val="E8F2E8"/></a:solidFill>'
      + '<a:ln w="6350"><a:noFill/></a:ln>'
      + '<a:extLst><a:ext uri="{91240B29-F687-4F45-9708-019B960494DF}">'
      + '<a14:hiddenLine xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" w="6350">'
      + '<a:solidFill><a:prstClr val="black"/></a:solidFill></a14:hiddenLine>'
      + "</a:ext></a:extLst></wps:spPr>");
    expect(box.fill).toEqual({ kind: "solid", color: "#E8F2E8" });
    expect(box.border).toBeUndefined();
  });

  test("a gradient fill is not read as its first stop or as the outline colour", async () => {
    const box = await filledBoxWith(
      '<wps:spPr><a:gradFill><a:gsLst>'
      + '<a:gs pos="0"><a:srgbClr val="1F4E79"/></a:gs>'
      + '<a:gs pos="100000"><a:srgbClr val="9CC2E5"/></a:gs>'
      + '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
      + '<a:ln w="6350"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></wps:spPr>');
    expect(box.fill).toEqual({
      kind: "linear-gradient", angle: 90, scaled: false, rotatesWithShape: true,
      stops: [{ position: 0, color: "#1F4E79" }, { position: 1, color: "#9CC2E5" }],
    });
    expect(box.border).toMatchObject({ color: "#000000" });
  });
});

describe("native Word text-box appearance and sizing", () => {
  async function fixtureBoxes(): Promise<{ pkg: OpcPackage; boxes: DocxTextBoxNode[] }> {
    const pkg = await OpcPackage.open(await Bun.file(`${FIXTURES}/${FIXTURE}`).arrayBuffer());
    return { pkg, boxes: await textBoxes(pkg) };
  }

  test("all seventeen added semantic cases parse once beside the original two boxes", async () => {
    const { boxes } = await fixtureBoxes();
    const tokens = ["TBGradient", "TBPattern", "TBPicture", "TBCombined", "TBRoundWide", "TBRoundTall",
      "TBOvalWide", "TBOvalTall", "TBRotated", "TBMiddle", "TBBottom", "TBDash", "TBNoWrap",
      "TBFixedOverflow", "TBAutoLong", "TBAutoShort", "TBAutoArabic"];
    expect(boxes).toHaveLength(25);
    for (const token of tokens) expect(boxes.filter((box) => boxText(box).includes(token))).toHaveLength(1);
    expect(boxText(findBox(boxes, "TBRotated"))).toContain("التسليم قريب");
  });

  test("native gradient, pattern and picture fills resolve with usable media", async () => {
    const { pkg, boxes } = await fixtureBoxes();
    const gradient = findBox(boxes, "TBGradient").fill;
    expect(gradient).toMatchObject({ kind: "linear-gradient" });
    if (gradient?.kind !== "linear-gradient") throw new Error("native gradient did not resolve");
    expect(gradient.stops.length).toBeGreaterThanOrEqual(2);
    expect(gradient.stops[0]!.color).toBe("#CBDCEA");
    expect(gradient.stops.at(-1)!.color).toBe("#C0D8B4");
    expect(findBox(boxes, "TBPattern").fill).toMatchObject({
      kind: "pattern", foreground: "#80B0D0", background: "#FFFFFF",
    });
    const picture = findBox(boxes, "TBPicture").fill;
    expect(picture).toMatchObject({ kind: "picture", contentType: "image/png" });
    if (picture?.kind !== "picture") throw new Error("native picture fill did not resolve");
    const bytes = await pkg.readPart(picture.mediaId);
    expect(Array.from(bytes.subarray(0, 8))).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
  });

  test("ellipse and adjusted roundRect carry geometry-local text rectangles at both aspect ratios", async () => {
    const { boxes } = await fixtureBoxes();
    for (const token of ["TBOvalWide", "TBOvalTall"]) {
      const geometry = findBox(boxes, token).geometry;
      expect(geometry?.preset).toBe("ellipse");
      expect(geometry?.fallback).toBeUndefined();
      if (geometry?.textRect === undefined) throw new Error(`${token}: missing text rectangle`);
      expect(geometry.textRect.left / geometry.viewBox.width).toBeCloseTo((1 - Math.SQRT1_2) / 2, 5);
      expect(geometry.textRect.top / geometry.viewBox.height).toBeCloseTo((1 - Math.SQRT1_2) / 2, 5);
      expect(geometry.textRect.right + geometry.textRect.left).toBeCloseTo(geometry.viewBox.width, 0);
      expect(geometry.textRect.bottom + geometry.textRect.top).toBeCloseTo(geometry.viewBox.height, 0);
    }
    for (const [token, adjustment] of [["TBRoundWide", 0.1], ["TBRoundTall", 0.45], ["TBCombined", 0.35]] as const) {
      const geometry = findBox(boxes, token).geometry;
      expect(geometry?.preset).toBe("roundRect");
      expect(geometry?.fallback).toBeUndefined();
      if (geometry?.textRect === undefined) throw new Error(`${token}: missing text rectangle`);
      const inset = Math.min(geometry.viewBox.width, geometry.viewBox.height) * adjustment * (1 - Math.SQRT1_2);
      expect(geometry.textRect.left).toBeCloseTo(inset, -1);
      expect(geometry.textRect.top).toBeCloseTo(inset, -1);
    }
  });

  test("native rotation, middle/bottom alignment, dash and combined styling stay independent", async () => {
    const { boxes } = await fixtureBoxes();
    expect(findBox(boxes, "TBRotated").transform).toMatchObject({ rotation: 25 });
    expect(findBox(boxes, "TBMiddle").verticalAlign).toBe("middle");
    expect(findBox(boxes, "TBBottom").verticalAlign).toBe("bottom");
    expect(findBox(boxes, "TBDash").border).toMatchObject({ style: "dashed", widthPt: 1.5 });
    expect(findBox(boxes, "TBCombined")).toMatchObject({
      geometry: { preset: "roundRect" }, fill: { kind: "linear-gradient" }, verticalAlign: "middle",
    });
    expect(findBox(boxes, "TBRotated").anchor).toMatchObject({
      horizontal: { from: "page", offset: 48 / 72 }, vertical: { from: "page", offset: 310 / 72 },
    });
  });

  test("native resize modes retain authored extents and twelve-point text, including Arabic", async () => {
    const { boxes } = await fixtureBoxes();
    // Native TextFrame.AutoSize = -1 writes spAutoFit, but saves the original
    // extent: Word recomputes 40 -> 132.9/91.5pt and 140 -> 36.3pt on reopen.
    for (const [token, heightPt] of [["TBAutoLong", 40], ["TBAutoShort", 140], ["TBAutoArabic", 40]] as const) {
      const box = findBox(boxes, token);
      expect(box.autofit).toBe("resizeShape");
      expect(box.height).toBeCloseTo(heightPt / 72, 5);
      expect(box.wrapping).toBe("square");
      const runs = box.content.flatMap((node) => node.type === "paragraph" ? node.runs : []);
      expect(runs.length).toBeGreaterThan(0);
      for (const run of runs) expect(run.fontSizePt).toBe(12);
    }
    expect(boxText(findBox(boxes, "TBAutoArabic"))).toContain("هذا نص عربي طويل");
    const fixed = findBox(boxes, "TBFixedOverflow");
    expect(fixed.autofit).toBe("none");
    expect(fixed.height).toBeCloseTo(40 / 72, 5);
    const nowrap = findBox(boxes, "TBNoWrap");
    expect(nowrap).toMatchObject({ wrapping: "none", autofit: "none" });
    // Word widens this to 174.7pt even with AutoSize off, while saving 150pt.
    expect(nowrap.width).toBeCloseTo(150 / 72, 5);
  });

  test("a carried normAutofit scale keeps Word text unscaled and the frame fixed", async () => {
    // Supplemental mutation of the native resize case, not a permanent
    // fixture: the earlier native Word PDF probe painted normAutofit at the
    // original size even with fontScale=62500 and lnSpcReduction=20000.
    let replaced = false;
    const pkg = await withPart(FIXTURE, "word/document.xml", (xml) => xml.replace(
      /<wps:wsp\b[^>]*>[\s\S]*?<\/wps:wsp>/g,
      (shape) => {
        if (!shape.includes("TBAutoLong")) return shape;
        replaced = true;
        return shape.replace("<a:spAutoFit/>", '<a:normAutofit fontScale="62500" lnSpcReduction="20000"/>');
      },
    ));
    expect(replaced).toBe(true);
    const box = findBox(await textBoxes(pkg), "TBAutoLong");
    expect(box.autofit).toBe("none");
    expect(box.height).toBeCloseTo(40 / 72, 5);
    const runs = box.content.flatMap((node) => node.type === "paragraph" ? node.runs : []);
    expect(runs.length).toBeGreaterThan(0);
    for (const run of runs) expect(run.fontSizePt).toBe(12);
  });
});
