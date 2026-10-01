import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide, parseTextBody, resolveStyleRef } from "../src/core/pptx.js";
import { attr, childElements, firstChild, parseXml } from "../src/core/xml.js";
import type { XmlElement } from "../src/core/xml.js";
import { findCoverageSlide, flattenElements, findShapesWithPreset, findShapeWithText, findTableWithText, joined } from "./fixture.js";

const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

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
    expect(doc.deck.slideParts).toHaveLength(12);
    expect(doc.deck.theme.colors.dk1).toBe("#000000");
    expect(doc.deck.theme.fonts.minor.scripts.Arab).toBe("Arial");
    // Medium Style 2 Accent 1, every region read: fills, text, fonts and all six edges.
    expect(doc.deck.tableStyles.get("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")).toMatchObject({
      wholeTbl: { fill: "#E7EAED", text: "#000000", font: "+mn-lt", borders: { insideH: { color: "#FFFFFF" }, insideV: { color: "#FFFFFF" } } },
      firstRow: { fill: "#156082", bold: true },
      band1H: { fill: "#CCD2D8" },
    });

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

  // A subtitle placeholder inherits the master's body style, not `otherStyle`:
  // PowerPoint paints its lines at 90% with 10 pt before each paragraph, while
  // date/footer/slide-number keep the plain other style.
  test("gives a subtitle placeholder the master body spacing and the slide number none", async () => {
    const doc = await fixture();
    const { slide } = await findCoverageSlide(doc.slides, "rtl-text");
    const subtitle = findShapeWithText(slide, "Nile Frame Studio");
    const number = flattenElements(slide.elements).find((element) => element.kind === "shape" && element.placeholder?.type === "sldNum");
    if (subtitle.text === undefined || number?.kind !== "shape" || number.text === undefined) throw new Error("fixture structure changed: expected subtitle and slide-number shapes");

    expect(subtitle.placeholder?.type).toBe("subTitle");
    for (const paragraph of subtitle.text.paragraphs) expect(paragraph).toMatchObject({ lineSpacing: 0.9, spaceBeforePt: 10 });
    expect(number.text.paragraphs[0]?.lineSpacing).toBeUndefined();
    expect(number.text.paragraphs[0]?.spaceBeforePt).toBeUndefined();
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

  // Ground truth: tools/probe/line-spacing-units (PowerPoint's own PDF of nine
  // stored-scale boxes). PowerPoint never re-measures on open: only the stored
  // `fontScale` / `lnSpcReduction` change what it paints.
  test("applies the autofit scale and line-spacing cut PowerPoint stored, and nothing else", async () => {
    const doc = await fixture();
    const body = (autofit: string, paragraphs: string) => parseTextBody(
      parseXml(`<a:txBody><a:bodyPr>${autofit}</a:bodyPr><a:lstStyle/>${paragraphs}</a:txBody>`),
      doc.deck.theme,
    );
    const run = (size: number) => `<a:r><a:rPr sz="${size}"/><a:t>text</a:t></a:r>`;

    // Each size is scaled and rounded half up to a whole point; the scale is on the model too.
    const scaled = body(`<a:normAutofit fontScale="92500"/>`, `<a:p>${run(1800)}</a:p><a:p>${run(2800)}</a:p>`);
    expect(scaled.autofit).toBe("shrinkText");
    expect(scaled.fontScale).toBe(0.925);
    expect(scaled.paragraphs.map((paragraph) => paragraph.runs[0]!.fontSizePt)).toEqual([17, 26]);
    expect(body(`<a:normAutofit fontScale="62500"/>`, `<a:p>${run(2800)}</a:p>`).paragraphs[0]!.runs[0]!.fontSizePt).toBe(18);
    expect(body(`<a:normAutofit fontScale="47500"/>`, `<a:p>${run(2400)}</a:p>`).paragraphs[0]!.runs[0]!.fontSizePt).toBe(11);

    // The line-spacing cut comes off each paragraph's own spacing (100% when unauthored);
    // spacing in points is not touched.
    const cut = body(
      `<a:normAutofit fontScale="62500" lnSpcReduction="20000"/>`,
      `<a:p>${run(2800)}</a:p>`
        + `<a:p><a:pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcAft><a:spcPts val="600"/></a:spcAft></a:pPr>${run(2800)}</a:p>`
        + `<a:p><a:pPr><a:lnSpc><a:spcPts val="3000"/></a:lnSpc></a:pPr>${run(2800)}</a:p>`,
    );
    expect(cut.lineSpacingReduction).toBe(0.2);
    expect(cut.paragraphs[0]!.lineSpacing).toBeCloseTo(0.8, 10);
    expect(cut.paragraphs[1]!.lineSpacing).toBeCloseTo(0.7, 10);
    expect(cut.paragraphs[1]!.spaceAfterPt).toBe(6);
    expect(cut.paragraphs[2]).toMatchObject({ lineSpacingPt: 30 });
    expect(cut.paragraphs[2]!.lineSpacing).toBeUndefined();
    // A paragraph with no run size scales the size it inherits, so the empty line follows the text.
    const inherited = body(`<a:normAutofit fontScale="50000"/>`, `<a:p><a:pPr><a:defRPr sz="2000"/></a:pPr><a:endParaRPr sz="2000"/></a:p>`);
    expect(inherited.paragraphs[0]!.fontSizePt).toBe(10);

    // Bare, absent, and unscaled autofit paint at 100%: an overflowing box is not shrunk on open.
    for (const autofit of [`<a:normAutofit/>`, `<a:normAutofit fontScale="100000"/>`, `<a:noAutofit/>`, ``]) {
      const plain = body(autofit, `<a:p>${run(1800)}</a:p>`);
      expect(plain.paragraphs[0]!.runs[0]!.fontSizePt).toBe(18);
      expect(plain.fontScale).toBeUndefined();
      expect(plain.lineSpacingReduction).toBeUndefined();
      expect(plain.paragraphs[0]!.lineSpacing).toBeUndefined();
    }

    // A layout's or master's `normAutofit` never carries a slide's scale.
    const inheritedAutofit = parseTextBody(
      parseXml(`<a:txBody><a:bodyPr/><a:lstStyle/><a:p>${run(1800)}</a:p></a:txBody>`),
      doc.deck.theme,
      { bodies: [parseXml(`<a:txBody><a:bodyPr><a:normAutofit fontScale="50000"/></a:bodyPr></a:txBody>`)] },
    );
    expect(inheritedAutofit.autofit).toBe("shrinkText");
    expect(inheritedAutofit.fontScale).toBeUndefined();
    expect(inheritedAutofit.paragraphs[0]!.runs[0]!.fontSizePt).toBe(18);
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
    const band = (color: string) => [1, 2, 3].map(() => ({ kind: "solid" as const, color }));
    expect(table.cellFills).toEqual([band("#156082"), band("#CCD2D8"), band("#E7EAED"), band("#CCD2D8")]);
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
    // Title, the table, then the slide-number placeholder the fixture carries.
    expect(copy.elements.map((element) => element.kind)).toEqual(["shape", "table", "shape"]);
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
    const presets = ["rect", "roundRect", "ellipse", "line", "triangle", "rtTriangle", "diamond", "parallelogram", "hexagon", "octagon", "star5"];
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
    // star5 resolves its ECMA path; the first remaining unsupported preset is
    // the up arrow, whose fill and outline come from its shape style.
    expect(fallback).toMatchObject({ fill: { kind: "solid", color: "#156082" }, stroke: { color: "#030E14", width: 12700 / 914400 }, geometry: { preset: "upArrow", fallback: "unsupported" } });

    const title = findShapeWithText(slide, "أشكال ورسومات");
    expect(joined(title.text!)).toBe("أشكال ورسومات");
    expect(title.text?.paragraphs[0]?.runs[0]).toMatchObject({ fontFamily: "Arial" });
  });

  test("style references with idx 0 resolve to nothing; direct fills win", async () => {
    const doc = await fixture();
    const { slide } = await findCoverageSlide(doc.slides, "shapes");
    // The unfilled right brace carries fillRef idx="0" with a vestigial
    // accent1 child: PowerPoint paints no fill, so neither do we — while its
    // lnRef idx="2" sibling still resolves the outline.
    const bare = slide.elements.find((element) => element.kind === "shape" && element.id === "66");
    if (bare?.kind !== "shape") throw new Error("fixture structure changed: expected Right Brace 65 (id 66)");
    expect(bare.fill).toBeUndefined();
    expect(bare.stroke).toMatchObject({ color: "#156082", width: 12700 / 914400 });
    // The filled left brace next to it keeps its direct black fill and red
    // outline: direct formatting always wins over style references.
    const filled = slide.elements.find((element) => element.kind === "shape" && element.id === "5");
    if (filled?.kind !== "shape") throw new Error("fixture structure changed: expected Left Brace 4 (id 5)");
    expect(filled).toMatchObject({
      fill: { kind: "solid", color: "#000000" },
      stroke: { color: "#FF0000", width: 76200 / 914400 },
      geometry: { preset: "leftBrace" },
    });
  });

  test("resolveStyleRef honors the idx-0-means-none convention", async () => {
    const doc = await fixture();
    const theme = doc.deck.theme;
    const ref = (xml: string): Parameters<typeof resolveStyleRef>[0] => firstChild(parseXml(`<root>${xml}</root>`), "fillRef");
    expect(resolveStyleRef(ref(`<fillRef idx="0"><schemeClr val="accent1"/></fillRef>`), theme)).toBeUndefined();
    expect(resolveStyleRef(ref(`<fillRef idx="1"><schemeClr val="accent1"/></fillRef>`), theme)).toBe(theme.colors.accent1);
    expect(resolveStyleRef(ref(`<fillRef idx="2"><srgbClr val="FF0000"/></fillRef>`), theme)).toBe("#FF0000");
    expect(resolveStyleRef(ref(`<fillRef><schemeClr val="accent1"/></fillRef>`), theme)).toBeUndefined();
    expect(resolveStyleRef(undefined, theme)).toBeUndefined();
  });

  test("resolves linear gradients and every path-gradient type", async () => {
    const doc = await fixture();
    const { slide } = await findCoverageSlide(doc.slides, "gradients");
    const fillOf = (name: string) => {
      const shape = slide.elements.find((element) => element.kind === "shape" && element.name === name);
      if (shape?.kind !== "shape") throw new Error(`fixture structure changed: expected shape "${name}"`);
      return shape.fill;
    };
    const linear = { fill: fillOf("Rectangle 4") };
    const radial = { fill: fillOf("Oval 6") };

    expect(linear.fill).toEqual({
      kind: "linear-gradient",
      angle: 45,
      // `<a:lin ang="2700000"/>` states no `scaled`: false, as PowerPoint draws it.
      scaled: false,
      rotatesWithShape: false,
      stops: [
        { position: 0, color: "#156082" },
        { position: 1, color: "#E8A33D" },
      ],
    });
    const centre = { left: 0.5, top: 0.5, right: 0.5, bottom: 0.5 };
    const box = { left: 0, top: 0, right: 0, bottom: 0 };
    // PowerPoint's centred radial preset: `path="circle"`, focus 50/50/50/50.
    expect(radial.fill).toEqual({
      kind: "path-gradient",
      path: "circle",
      focus: centre,
      extent: box,
      stops: [
        { position: 0, color: "#156082" },
        { position: 1, color: "#73A0B4" },
      ],
    });
    // Hand-made in PowerPoint: Radial "From Bottom Right Corner" (the extent
    // doubled to the right and down, the focus at its far corner) with a
    // tinted middle stop, and a Path gradient on a star.
    expect(fillOf("Rectangle 3")).toEqual({
      kind: "path-gradient",
      path: "circle",
      focus: { left: 1, top: 1, right: 0, bottom: 0 },
      extent: { left: 0, top: 0, right: -1, bottom: -1 },
      stops: [
        { position: 0, color: "#156082" },
        { position: 0.2, color: "#74C4E9" },
        { position: 1, color: "#E8A33D" },
      ],
    });
    expect(fillOf("Star: 5 Points 8")).toEqual({
      kind: "path-gradient",
      path: "shape",
      focus: centre,
      extent: box,
      stops: [
        { position: 0, color: "#156082" },
        { position: 1, color: "#E8A33D" },
      ],
    });
    const title = findShapeWithText(slide, "تدرجات لونية");
    expect(joined(title.text!)).toBe("تدرجات لونية");
    expect(title.text?.paragraphs[0]?.runs[0]).toMatchObject({ fontFamily: "Arial" });
  });

  test("resolves pattern, picture and rect-path fills on shapes", async () => {
    const doc = await fixture();
    const { slide } = await findCoverageSlide(doc.slides, "shape-fills");
    const fillOf = (name: string) => {
      const shape = slide.elements.find((element) => element.kind === "shape" && element.name === name);
      if (shape?.kind !== "shape") throw new Error(`fixture structure changed: expected shape "${name}"`);
      return shape.fill;
    };
    expect(fillOf("Pattern DkDnDiag")).toEqual({ kind: "pattern", preset: "dkDnDiag", foreground: "#156082", background: "#FFFFFF" });
    expect(fillOf("Pattern SmGrid")).toEqual({ kind: "pattern", preset: "smGrid", foreground: "#E8A33D", background: "#FFF4E0" });
    // A plain `a:stretch/a:fillRect` carries no insets, so no `stretch` field.
    expect(fillOf("Picture Stretch")).toMatchObject({ kind: "picture", contentType: "image/jpeg" });
    expect(fillOf("Picture Stretch")).not.toHaveProperty("stretch");
    expect(fillOf("Picture Tile")).toMatchObject({
      kind: "picture",
      contentType: "image/png",
      tile: { scaleX: 0.25, scaleY: 0.25, offsetX: 0, offsetY: 0, align: "tl", flip: "none" },
    });
    const stops = [{ position: 0, color: "#156082" }, { position: 1, color: "#E8A33D" }];
    // `tileRect l=t=-100%` + `fillToRect r=b=100%`: the focus is the box's
    // top-left corner, as PowerPoint paints it.
    expect(fillOf("Gradient Rect Corner")).toEqual({
      kind: "path-gradient",
      path: "rect",
      focus: { left: 0, top: 0, right: 1, bottom: 1 },
      extent: { left: -1, top: -1, right: 0, bottom: 0 },
      stops,
    });
    expect(fillOf("Gradient Rect Center")).toEqual({
      kind: "path-gradient",
      path: "rect",
      focus: { left: 0.5, top: 0.5, right: 0.5, bottom: 0.5 },
      extent: { left: 0, top: 0, right: 0, bottom: 0 },
      stops,
    });
  });

  test("resolves pattern and tiled-picture slide backgrounds", async () => {
    const doc = await fixture();
    const pattern = await findCoverageSlide(doc.slides, "pattern-background");
    expect(pattern.slide.background).toEqual({ kind: "pattern", preset: "dkDnDiag", foreground: "#D6E4F0", background: "#FFFFFF" });
    const tiled = await findCoverageSlide(doc.slides, "tiled-background");
    // Slide-owned, so not `shared`; the tile carries PowerPoint's 30 % scale.
    expect(tiled.slide.background).toMatchObject({
      kind: "picture",
      contentType: "image/png",
      tile: { scaleX: 0.3, scaleY: 0.3, align: "tl", flip: "none" },
    });
    expect(tiled.slide.background).not.toHaveProperty("shared");
  });

  test("resolves every table-cell fill kind, direct fills over style bands", async () => {
    const doc = await fixture();
    const { slide } = await findCoverageSlide(doc.slides, "table-fills");
    const table = findTableWithText(slide, "البند");
    const header = { kind: "linear-gradient" as const, angle: 90, scaled: true, rotatesWithShape: true, stops: [{ position: 0, color: "#156082" }, { position: 1, color: "#73A0B4" }] };
    expect(table.cellFills[0]).toEqual([header, header, header]);
    // Unfilled cells keep the Medium Style 2 Accent 1 band colours.
    expect(table.cellFills[1]!.slice(0, 2)).toEqual([{ kind: "solid", color: "#CCD2D8" }, { kind: "solid", color: "#CCD2D8" }]);
    expect(table.cellFills[1]![2]).toEqual({ kind: "pattern", preset: "smGrid", foreground: "#E8A33D", background: "#FFF4E0" });
    expect(table.cellFills[2]![2]).toMatchObject({ kind: "path-gradient", path: "rect", focus: { left: 0.5, top: 0.5, right: 0.5, bottom: 0.5 } });
  });

  // PowerPoint writes every line and connector as `p:cxnSp` (slide 11,
  // appended through PowerPoint's object model): they parse as shapes.
  test("parses connector and line shapes with their dash style", async () => {
    const doc = await fixture();
    const { slide } = await findCoverageSlide(doc.slides, "lines");
    const byName = new Map(flattenElements(slide.elements).filter((element) => element.kind === "shape").map((shape) => [shape.name, shape]));
    const timeline = byName.get("Timeline");
    const today = byName.get("Today");
    const dotted = byName.get("Dotted");
    const diagonal = byName.get("Diagonal");
    expect(timeline?.stroke).toMatchObject({ color: "#1E3A5A" });
    expect(timeline?.stroke?.dash).toBeUndefined();
    expect(timeline?.bounds?.height).toBe(0);
    expect(today?.stroke).toMatchObject({ color: "#F59B1E", dash: "dash" });
    expect(today?.bounds?.width).toBe(0);
    expect(dotted?.stroke).toMatchObject({ dash: "dot" });
    expect(diagonal?.transform?.flipVertical).toBe(true);
    for (const line of [timeline, today, dotted, diagonal]) expect(line?.fill).toBeUndefined();
  });

  // PowerPoint keeps an en-US run's number together left to right inside a
  // right-to-left paragraph ("2.1%", never "%2.1"; probe run-direction).
  test("marks a left-to-right-language run inside a right-to-left paragraph", async () => {
    const doc = await fixture();
    const { slide } = await findCoverageSlide(doc.slides, "lines");
    const paragraph = findShapeWithText(slide, "2.1%").text!.paragraphs[0]!;
    expect(paragraph.direction).toBe("rtl");
    expect(paragraph.runs.map((run) => [run.text, run.ltrRun ?? false])).toEqual([
      ["الإيرادات دون الموازنة بنسبة ", false],
      ["2.1%", true],
    ]);
  });

  // Space before/after in lines (`a:spcPct`, slide 12, set through
  // PowerPoint's object model) stays in lines; the renderer turns it into
  // 1.2 x the paragraph's largest run size (probe line-spacing-units).
  test("keeps paragraph spacing in lines, in text boxes and table cells", async () => {
    const doc = await fixture();
    const { slide } = await findCoverageSlide(doc.slides, "paragraph-spacing");
    const box = findShapeWithText(slide, "نصف سطر قبل هذه الفقرة").text!.paragraphs;
    expect(box.map((paragraph) => [paragraph.spaceBeforeLines, paragraph.spaceAfterLines])).toEqual([
      [1, 0], [0.5, 0], [1, 0], [0, 1], [1, 0], [0, 1],
    ]);
    for (const paragraph of box) {
      expect(paragraph.spaceBeforePt).toBeUndefined();
      expect(paragraph.spaceAfterPt).toBeUndefined();
    }
    expect(box[2]!.runs.map((run) => run.fontSizePt)).toContain(28);
    const cell = findTableWithText(slide, "نصف سطر بعد هذه").rows[0]!.cells[0]!.text.paragraphs;
    expect(cell.map((paragraph) => [paragraph.spaceBeforeLines, paragraph.spaceAfterLines])).toEqual([[1, 0], [0, 0.5], [0.5, 0]]);
  });
});
