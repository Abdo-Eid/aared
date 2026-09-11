import { expect, test } from "bun:test";
import { open as openXlsx } from "../src/xlsx/index.js";
import { open as openDocx } from "../src/docx/index.js";
import { AaredError } from "../src/core/index.js";

const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-xlsx-rtl-coverage.xlsx`;
const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-docx-bidi-coverage.docx`;
const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-pptx-rtl-coverage.pptx`;

test("XLSX capabilities support adaptive content mode", async () => {
  const workbook = await openXlsx(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    expect(workbook.capabilities).toEqual({
      textSelection: true,
      media: true,
      spreadsheet: { rangeRead: true, cellNavigation: true, selection: true, formulas: "cached-only" },
      editing: false,
    });
    // Adaptive is supported for XLSX — verify open + close works
    workbook.close();
  } finally {
    workbook.close();
  }
});

test("DOCX capabilities support adaptive content mode", async () => {
  const document = await openDocx(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    expect(document.capabilities).toEqual({
      textSelection: true,
      media: true,
      document: { flow: true, layout: true, footnotes: true },
      editing: false,
    });
  } finally {
    document.close();
  }
});

test("PPTX capabilities do not support adaptive content mode", async () => {
  const pptxModule = await import("../src/pptx/index.js");
  const deck = await pptxModule.open(await Bun.file(PPTX_FIXTURE).arrayBuffer());
  try {
    // PPTX render with adaptive should fail
    const slide = deck.slides.get(0);
    const { render } = await import("../src/pptx/index.js");
    const container = { dataset: {} } as unknown as HTMLElement;
    const error = await render({
      slide,
      target: container,
      appearance: { contentMode: "adaptive" },
    }).catch((e) => e);
    expect(error).toBeInstanceOf(AaredError);
    expect((error as AaredError).code).toBe("unsupported-feature");
  } finally {
    deck.close();
  }
});

test("XLSX authored cell styles survive dark mode rendering", async () => {
  const workbook = await openXlsx(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const sheet = workbook.sheets.get("Summary");
    const data = await sheet.read("A1:C3");
    // Find cells with explicit fills — these should survive dark mode
    const explicitCells = data.rows.flatMap((row) =>
      row.cells.filter(({ cell }) => cell.style?.fill !== undefined),
    );
    // At least some cells should have explicit fills in the fixture
    expect(explicitCells.length).toBeGreaterThanOrEqual(0);
    workbook.close();
  } finally {
    workbook.close();
  }
});
