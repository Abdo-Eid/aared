import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide, parseTextBody } from "../src/core/pptx.js";
import { attr, childElements, firstChild, parseXml } from "../src/core/xml.js";
import type { XmlElement } from "../src/core/xml.js";
import { findCoverageSlide, findShapesWithPreset, findShapeWithText, findTableWithText, joined } from "./fixture.js";

const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/arabic-pptx-rtl-coverage.pptx`;

async function fixture() {
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const deck = await indexPptx(pkg);
  return { deck, slides: await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index))) };
}

function placeholder(root: XmlElement, type: string, index: string): XmlElement | undefined {
  const slide = firstChild(root, "cSld");
  const tree = slide === undefined ? undefined : firstChild(slide, "spTree");
  if (tree === undefined) return undefined;
  return childElements(tree, "sp").find((shape) => {
    const nvPr = firstChild(firstChild(shape, "nvSpPr") ?? shape, "nvPr");
    const ph = nvPr === undefined ? undefined : firstChild(nvPr, "ph");
    return ph !== undefined && (attr(ph, "type") ?? "body") === type && (attr(ph, "idx") ?? "0") === index;
  });
}

describe("PPTX DrawingML parser", () => {
  test("re-reads slides and RTL text from the real Arabic fixture", async () => {
    const doc = await fixture();
    expect(doc.deck.size).toEqual({ width: 12192000 / 914400, height: 6858000 / 914400 });
    expect(doc.deck.slideParts).toHaveLength(4);
    expect(doc.deck.theme.colors.dk1).toBe("#000000");
    expect(doc.deck.theme.fonts.minor.scripts.Arab).toBe("Arial");
    expect(doc.deck.tableStyles.get("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")).toMatchObject({ whole: "#E7EAED", firstRow: "#156082", band1H: "#CCD2D8" });

    const { slide } = await findCoverageSlide(doc.slides, "rtl-text");
    const title = findShapeWithText(slide, "تقرير الإنتاج الأسبوعي");
    const bullets = findShapeWithText(slide, "Nile Frame Studio");
    if (title.text === undefined || bullets.text === undefined) throw new Error("fixture structure changed: expected RTL text shapes");
    expect(joined(title.text)).toBe("تقرير الإنتاج الأسبوعي");
    expect(title.text.paragraphs[0]).toMatchObject({ direction: "rtl", align: "right", fontSizePt: 60, lineSpacing: 0.9 });
    expect(title.text).toMatchObject({ verticalAlign: "bottom" });
    expect(title.text.insets).toEqual({ left: 0.1, right: 0.1, top: 0.05, bottom: 0.05 });
    expect(title.text.autofit).toBeUndefined();
    expect(title.text.paragraphs[0]?.runs[0]).toMatchObject({
      bold: true,
      fontSizePt: 60,
      latinTypeface: "+mj-lt",
      complexScriptTypeface: "+mj-cs",
      fontFamily: "Times New Roman",
    });
    expect(title.text.paragraphs[0]?.runs.every((run) => run.fontFamily === "Times New Roman")).toBe(true);
    expect(title.bounds).toEqual({ x: 2390273 / 914400, y: 493880 / 914400, width: 9144000 / 914400, height: 1027279 / 914400 });
    expect(title.fill).toBeUndefined();
    expect(title.geometry?.fallback).toBe("missing");
    expect(bullets.text.paragraphs).toHaveLength(5);
    expect(bullets.text.paragraphs[0]).toMatchObject({
      direction: "rtl",
      align: "right",
      leftIndent: 0.375,
      firstLineIndent: -0.375,
      bullet: { char: "•", typeface: "Arial" },
    });
    expect(bullets.text.paragraphs[0]?.runs[0]).toMatchObject({
      fontSizePt: 24,
      latinTypeface: "+mn-lt",
      complexScriptTypeface: "+mn-cs",
      fontFamily: "Arial",
    });
    expect(bullets.text.paragraphs[0]?.runs.every((run) => run.fontFamily === "Arial")).toBe(true);
    expect(bullets.text.paragraphs[2]?.runs.at(-1)).toMatchObject({
      text: " — Nile Frame Studio — Shot 4 locked",
      fontFamily: "Aptos",
    });
    // Direction and alignment remain independent: the English control line
    // is left-aligned while its paragraph remains explicitly RTL.
    expect(bullets.text.paragraphs[3]).toMatchObject({ direction: "rtl", align: "left" });
    expect(bullets.text.paragraphs[3]?.runs[0]).toMatchObject({ text: "Next review: Thursday 12:00", fontFamily: "Aptos" });
  });

  test("matches the master title placeholder for a centered-title layout placeholder", async () => {
    const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
    const deck = await indexPptx(pkg);
    const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
    const { index, slide } = await findCoverageSlide(slides, "rtl-text");
    const slidePart = deck.slideParts[index]!;
    const layoutPart = (await pkg.relationships(slidePart)).find((relationship) => relationship.type.endsWith("/slideLayout"))?.target;
    if (layoutPart === undefined) throw new Error("fixture structure changed: slide 1 has no layout");
    const masterPart = (await pkg.relationships(layoutPart)).find((relationship) => relationship.type.endsWith("/slideMaster"))?.target;
    if (masterPart === undefined) throw new Error("fixture structure changed: layout has no master");
    const layoutTitle = placeholder(parseXml(await pkg.readText(layoutPart)), "ctrTitle", "0");
    const masterTitle = placeholder(parseXml(await pkg.readText(masterPart)), "title", "0");
    const layoutBodyPr = layoutTitle === undefined ? undefined : firstChild(firstChild(layoutTitle, "txBody") ?? layoutTitle, "bodyPr");
    const masterBodyPr = masterTitle === undefined ? undefined : firstChild(firstChild(masterTitle, "txBody") ?? masterTitle, "bodyPr");
    const title = findShapeWithText(slide, "تقرير الإنتاج الأسبوعي");
    if (title.text === undefined) throw new Error("fixture structure changed: expected title shape");

    expect(attr(layoutBodyPr ?? layoutTitle!, "lIns")).toBeUndefined();
    expect(attr(masterBodyPr ?? masterTitle!, "lIns")).toBe("91440");
    expect(title.text.verticalAlign).toBe("bottom");
    expect(title.text.paragraphs[0]?.runs[0]).toMatchObject({
      latinTypeface: "+mj-lt",
      complexScriptTypeface: "+mj-cs",
      fontFamily: "Times New Roman",
    });
    // A layout can override the anchor while each omitted inset falls through
    // to the matching master attribute. Ground truth: tools/probe/title-insets.
    expect(title.text.insets).toEqual({ left: 0.1, right: 0.1, top: 0.05, bottom: 0.05 });
    expect(title.text.autofit).toBeUndefined();
  });

  test("inherits each missing text-frame inset without overriding an explicit zero", async () => {
    const doc = await fixture();
    const text = parseTextBody(
      parseXml(`<a:txBody><a:bodyPr bIns="0"/><a:lstStyle/><a:p><a:r><a:t>title</a:t></a:r></a:p></a:txBody>`),
      doc.deck.theme,
      {
        bodies: [
          parseXml(`<a:txBody><a:bodyPr lIns="91440" tIns="45720" bIns="137160"/></a:txBody>`),
          parseXml(`<a:txBody><a:bodyPr rIns="182880" bIns="182880"/></a:txBody>`),
        ],
      },
    );

    expect(text.insets).toEqual({ left: 0.1, right: 0.2, top: 0.05, bottom: 0 });
  });

  test("preserves table-level RTL separately from explicit LTR cells", async () => {
    const doc = await fixture();
    const { slide } = await findCoverageSlide(doc.slides, "rtl-table");
    const title = findShapeWithText(slide, "جدول المتابعة");
    const table = findTableWithText(slide, "الاسم");
    if (title.text === undefined) throw new Error("fixture structure changed: expected table title");
    expect(joined(title.text)).toBe("جدول المتابعة");
    expect(title.bounds).toEqual({ x: 838200 / 914400, y: 365125 / 914400, width: 10515600 / 914400, height: 1325563 / 914400 });
    expect(table).toMatchObject({
      direction: "rtl",
      columns: [2709333 / 914400, 2709333 / 914400, 2709333 / 914400],
      bounds: { x: 2032000 / 914400, y: 2687320 / 914400, width: 8127999 / 914400, height: 2560320 / 914400 },
    });
    expect(table.rows).toHaveLength(4);
    expect(table.cellFills[0]).toEqual(["#156082", "#156082", "#156082"]);
    expect(table.cellFills).toEqual([
      ["#156082", "#156082", "#156082"],
      ["#CCD2D8", "#CCD2D8", "#CCD2D8"],
      ["#E7EAED", "#E7EAED", "#E7EAED"],
      ["#CCD2D8", "#CCD2D8", "#CCD2D8"],
    ]);
    expect(table.rows.map((row) => row.cells)).toSatisfy((rows) => rows.every((cells) => cells.length === 3));
    expect(table.rows[0]!.cells.map((cell) => joined(cell.text))).toEqual(["الاسم", "الدور", "الحالة"]);
    expect(table.rows[0]!.cells.every((cell) => cell.text.paragraphs[0]?.direction === "ltr")).toBe(true);
    expect(table.rows[0]!.cells.every((cell) => cell.verticalAlign === "middle")).toBe(true);
    expect(table.rows[0]!.cells[0]!.text.paragraphs[0]!.runs[0]).toMatchObject({ bold: true, fontSizePt: 32, color: "#FF0000" });
    expect(table.rows[1]!.cells[0]!.text.paragraphs[0]!.runs[0]).toMatchObject({ bold: false, fontSizePt: 28, color: "#000000", fontFamily: "Arial" });
  });

  test("returns a structured-cloneable canonical slide model", async () => {
    const doc = await fixture();
    const { slide } = await findCoverageSlide(doc.slides, "rtl-table");
    const copy = structuredClone(slide);

    expect(copy).toEqual(slide);
    expect(copy.elements.map((element) => element.kind)).toEqual(["shape", "table"]);
    expect(copy.size).toEqual({ width: 12192000 / 914400, height: 6858000 / 914400 });
  });

  test("rejects an already-aborted parser request with AbortError", async () => {
    const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
    const controller = new AbortController();
    controller.abort();

    await expect(indexPptx(pkg, controller.signal)).rejects.toMatchObject({
      name: "AbortError",
      message: "aared/core: request was cancelled",
    });
  });

  test("resolves positioned, styled common shapes from the real LibreOffice fixture", async () => {
    const doc = await fixture();
    const { slide } = await findCoverageSlide(doc.slides, "shapes");
    const presets = ["rect", "roundRect", "ellipse", "line", "triangle", "rtTriangle", "diamond", "parallelogram", "hexagon", "octagon"];
    const authored = presets.flatMap((preset) => findShapesWithPreset(slide, preset));

    expect(authored.every((shape) => shape.geometry?.fallback === undefined && shape.geometry?.path !== "")).toBe(true);
    expect(new Set(authored.map((shape) => shape.geometry?.preset))).toEqual(new Set(presets));
    const rectangle = findShapesWithPreset(slide, "rect").find((shape) => shape.fill?.kind === "solid" && shape.fill.color === "#73A0B4" && shape.stroke?.color === "#156082");
    if (rectangle === undefined) throw new Error("fixture structure changed: expected a blue outlined rectangle");
    expect(rectangle).toMatchObject({
      fill: { kind: "solid", color: "#73A0B4" },
      stroke: { color: "#156082", width: 28800 / 914400 },
      geometry: { preset: "rect" },
    });
    const firstBounds = rectangle.bounds;
    if (firstBounds === undefined) throw new Error("fixture structure changed: expected positioned rectangle");
    expect(firstBounds.width).toBeGreaterThan(0);
    expect(firstBounds.height).toBeGreaterThan(0);
    expect(firstBounds.x).toBeGreaterThan(0);
    expect(firstBounds.y).toBeGreaterThan(0);
    const rotated = authored.find((shape) => shape.transform?.rotation === 330);
    expect(rotated).toMatchObject({
      fill: { kind: "solid", color: "#44809B" },
      stroke: { color: "#156082", width: 43200 / 914400 },
      transform: { rotation: 330 },
    });
    expect(authored.find((shape) => shape.stroke?.color === "#E8A33D")).toMatchObject({ stroke: { color: "#E8A33D", width: 36000 / 914400 } });
    const fallback = slide.elements.find((element) => element.kind === "shape" && element.geometry?.fallback === "unsupported");
    expect(fallback).toMatchObject({ fill: { kind: "solid", color: "#73A0B4" }, stroke: { color: "#156082", width: 28800 / 914400 }, geometry: { preset: "star5", fallback: "unsupported" } });

    const title = findShapeWithText(slide, "أشكال ورسومات");
    expect(joined(title.text!)).toBe("أشكال ورسومات");
    expect(title.text?.paragraphs[0]?.runs[0]).toMatchObject({ fontFamily: "Arial" });
  });

  test("resolves LibreOffice linear gradients and explicitly degrades path gradients", async () => {
    const doc = await fixture();
    const { slide } = await findCoverageSlide(doc.slides, "gradients");
    const linear = slide.elements.find((element) => element.kind === "shape" && element.fill?.kind === "linear-gradient");
    const path = slide.elements.find((element) => element.kind === "shape" && element.fill?.kind === "unsupported");
    if (linear?.kind !== "shape" || linear.fill?.kind !== "linear-gradient" || path?.kind !== "shape" || path.fill?.kind !== "unsupported"
    ) {
      throw new Error("fixture structure changed: expected linear and path gradients");
    }

    expect(linear.fill).toEqual({
      kind: "linear-gradient",
      angle: 45,
      scaled: true,
      rotatesWithShape: false,
      stops: [
        { position: 0, color: "#156082" },
        { position: 1, color: "#E8A33D" },
      ],
    });
    expect(path.fill).toEqual({ kind: "unsupported", source: "gradient-path" });
    const title = findShapeWithText(slide, "تدرجات لونية");
    expect(joined(title.text!)).toBe("تدرجات لونية");
    expect(title.text?.paragraphs[0]?.runs[0]).toMatchObject({ fontFamily: "Arial" });
  });
});
