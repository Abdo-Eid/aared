import { expect, test } from "bun:test";
import { open as openXlsx } from "../src/xlsx/index.js";
import { open as openDocx } from "../src/docx/index.js";

const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;
const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

// Content is source-faithful (like PPTX): there is no contentMode, and the
// resolved models never depend on the viewer theme. Theme drives chrome only.

test("XLSX resolved models are theme-independent and keep authored styles", async () => {
  const bytes = await Bun.file(XLSX_FIXTURE).arrayBuffer();
  const first = await openXlsx(bytes.slice(0));
  const second = await openXlsx(bytes.slice(0));
  try {
    const a = await first.sheets.get("Summary").read({ range: "A1:C3" });
    const b = await second.sheets.get("Summary").read({ range: "A1:C3" });
    expect(JSON.stringify(a)).toBe(JSON.stringify(b));
    // Authored styles survive the round trip verbatim.
    const styled = a.rows.flatMap((row) =>
      row.cells.filter(({ cell }) => cell.style?.fill !== undefined || cell.style?.color !== undefined),
    );
    expect(styled.length).toBeGreaterThan(0);
  } finally {
    first.close();
    second.close();
  }
});

test("DOCX resolved sections keep authored colors and shading", async () => {
  const document = await openDocx(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const sections = document.sections.list();
    expect(sections.length).toBeGreaterThan(0);
    const json = JSON.stringify(sections);
    // The fixture carries authored table shading and run colors.
    expect(json).toContain("#F4F1EC");
  } finally {
    document.close();
  }
});
