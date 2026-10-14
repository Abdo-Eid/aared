import { expect, spyOn, test } from "bun:test";
import { open as openXlsx, render as renderXlsx } from "../src/xlsx/index.js";
import { open as openPptx, render as renderPptx } from "../src/pptx/index.js";
import { open as openDocx, render as renderDocx } from "../src/docx/index.js";
import { readSheetModel, type SheetHandle } from "../src/xlsx/renderer.js";
import { sectionsOf, type DocxDocument } from "../src/docx/renderer.js";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide, readPptxSlideText } from "../src/core/pptx.js";
import { WorkerSession } from "../src/core/session.js";
import { cellData } from "../src/xlsx/data.js";
import type { CellModel } from "../src/core/model.js";

const FIXTURES = `${import.meta.dir}/../../../fixtures`;

test("format entries export operations, not implementation constructors or painter primitives", async () => {
  for (const api of [await import("../src/pptx/index.js"), await import("../src/xlsx/index.js"), await import("../src/docx/index.js")]) {
    expect(Object.keys(api).sort()).toEqual(["open", "preview", "render"]);
  }
});

test("spreadsheet data projects cached values and clones every nested container", async () => {
  const requestedFonts: string[] = [];
  const workbook = await openXlsx(await Bun.file(`${FIXTURES}/xlsx-coverage.xlsx`).arrayBuffer(), {
    fontResolver: family => { requestedFonts.push(family); return undefined; },
  });
  let formula = false;
  let hyperlink = false;
  const types = new Set<string>();
  try {
    for (const handle of workbook.sheets.list()) {
      const data = await handle.read();
      expect(requestedFonts).toEqual([]);
      const model = await readSheetModel(handle as SheetHandle, undefined, false);
      expect(Object.keys(data).sort()).toEqual(["extent", "frozen", "merges", "name", "range", "rightToLeft", "rows"]);
      expect(data.name).toBe(model.name);
      expect(data.extent).toEqual(model.extent);
      expect(data.merges).toEqual(model.merges);
      expect(data.frozen).toEqual(model.frozen);
      for (const [i, row] of data.rows.entries()) {
        expect(Object.keys(row).sort()).toEqual(["cells", "index"]);
        for (const [j, entry] of row.cells.entries()) {
          const reference = model.rows[i]!.cells[j]!.cell;
          expect(entry.cell.value).toBe(reference.value);
          expect(entry.cell.formatted).toBe(reference.formatted);
          expect(entry.cell.formula).toBe(reference.formula);
          expect(entry.cell.hyperlink).toEqual(reference.hyperlink);
          expect(entry.cell.type).toBe(reference.type === "shared" || reference.type === "inline" || reference.type === "formula-string" ? "string" : reference.type);
          expect(Object.keys(entry.cell).every(key => ["value", "formatted", "type", "formula", "hyperlink"].includes(key))).toBe(true);
          formula ||= entry.cell.formula !== undefined;
          hyperlink ||= entry.cell.hyperlink !== undefined;
          types.add(entry.cell.type);
        }
      }
      const reference = await handle.read();
      data.range.start.row = 999;
      data.extent.rows = 999;
      data.frozen.rows = 999;
      if (data.merges[0]) data.merges[0].end.col = 999;
      for (const row of data.rows) for (const { cell } of row.cells) {
        cell.formatted = "changed";
        if (cell.hyperlink) cell.hyperlink.target = "changed";
      }
      data.rows.push({ index: 999, cells: [] });
      expect(await handle.read()).toEqual(reference);
      const window = await handle.read({ range: "A1:C3" });
      expect(window.range).toEqual({ start: { row: 0, col: 0 }, end: { row: 2, col: 2 } });
      const windowModel = await readSheetModel(handle as SheetHandle, { range: "A1:C3" }, false);
      expect(window.merges).toEqual(windowModel.merges); // Preserve intersected, clipped merges.
      expect(window.rows.every(row => row.index <= 2 && row.cells.every(cell => cell.col <= 2))).toBe(true);
    }
    expect(formula).toBe(true);
    expect(hyperlink).toBe(true);
    expect(types.has("string")).toBe(true);
    expect(types.has("number")).toBe(true);
  } finally { workbook.close(); }
});

test("cell storage kinds project to semantic types without exposing paint fields", () => {
  for (const type of ["shared", "inline", "formula-string", "number", "boolean", "error", "empty"] as const) {
    const input: CellModel = { value: null, formatted: "", type, style: { fill: { kind: "solid", color: "#FF0000" } }, conditional: { fill: "#000000" } };
    expect(cellData(input)).toEqual({ value: null, formatted: "", type: type === "shared" || type === "inline" || type === "formula-string" ? "string" : type });
  }
});

test("public slide reads return fixture paragraphs and notes without models or font resolution", async () => {
  let notesFound = false;
  for (const name of ["pptx-coverage.pptx", "post-production-review.pptx"]) {
    const bytes = await Bun.file(`${FIXTURES}/${name}`).arrayBuffer();
    const pkg = await OpcPackage.open(bytes);
    const parsed = await indexPptx(pkg);
    const fonts: string[] = [];
    const deck = await openPptx(bytes, { fontResolver: family => { fonts.push(family); return undefined; } });
    const unit = spyOn(WorkerSession.prototype, "unit");
    const text = spyOn(WorkerSession.prototype, "text");
    try {
      for (const slide of deck.slides.list()) {
        const data = await slide.read();
        const model = await parsePptxSlide(pkg, parsed, slide.info.index);
        expect(data.paragraphs).toEqual(await readPptxSlideText(pkg, parsed, slide.info.index));
        expect(data.notes).toBe(model.notes);
        notesFound ||= data.notes !== undefined;
        expect(Object.keys(data).every(key => key === "paragraphs" || key === "notes")).toBe(true);
        data.paragraphs.push("changed");
        expect((await slide.read()).paragraphs).not.toContain("changed");
      }
      expect(unit).not.toHaveBeenCalled();
      expect(fonts).toEqual([]);
      expect(text.mock.calls.every(call => call[1]?.includeNotes === true)).toBe(true);
      text.mockClear();
      await deck.search("budget");
      expect(text.mock.calls.length).toBeGreaterThan(0);
      expect(text.mock.calls.every(call => call[1]?.includeNotes !== true)).toBe(true);
      expect(unit).not.toHaveBeenCalled();
    } finally { unit.mockRestore(); text.mockRestore(); deck.close(); }
  }
  expect(notesFound).toBe(true);
});

test("section metadata converts fixture twips to inches and isolates nested page data", async () => {
  const document = await openDocx(await Bun.file(`${FIXTURES}/docx-coverage.docx`).arrayBuffer());
  try {
    const reference = sectionsOf(document as DocxDocument);
    const data = document.sections.list();
    for (const [index, info] of data.entries()) {
      const section = reference[index]!;
      expect(Object.keys(info).sort()).toEqual(["breakType", "index", "page"]);
      expect(info.index).toBe(index);
      expect(info.breakType).toBe(section.properties.breakType);
      expect(info.page.width).toBe(section.properties.page.widthTwips / 1440);
      expect(info.page.height).toBe(section.properties.page.heightTwips / 1440);
      for (const key of ["top", "right", "bottom", "left", "header", "footer", "gutter"] as const) {
        expect(info.page.margins[key]).toBe(section.properties.page.margins[key] / 1440);
      }
      const expected = document.sections.get(index);
      info.page.width = 999;
      info.page.margins.top = 999;
      expect(document.sections.get(index)).toEqual(expected);
    }
    expect(await document.search("Objectives")).not.toEqual([]);
  } finally { document.close(); }
});

test("public renderers reject models, primitives and forged handles before touching the target", async () => {
  const target = {} as HTMLElement;
  for (const input of [null, undefined, 1, "model", [], { read: () => undefined }, { elements: [], size: { width: 1, height: 1 } }, { rows: [], extent: { rows: 1, cols: 1 } }]) {
    await expect(renderPptx(input as never)).rejects.toMatchObject({ code: "invalid-argument" });
    await expect(renderXlsx(input as never)).rejects.toMatchObject({ code: "invalid-argument" });
    await expect(renderDocx(input as never)).rejects.toMatchObject({ code: "invalid-argument" });
    await expect(renderPptx({ slide: input as never, target })).rejects.toMatchObject({ code: "invalid-argument" });
    await expect(renderXlsx({ sheet: input as never, target })).rejects.toMatchObject({ code: "invalid-argument" });
    await expect(renderDocx({ document: input as never, target })).rejects.toMatchObject({ code: "invalid-argument" });
  }
});
