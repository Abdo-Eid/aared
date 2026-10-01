import { expect, test } from "bun:test";
import { AaredError } from "../src/core/index.js";
import { open } from "../src/xlsx/index.js";
import { rejection } from "./harness.js";
import { storedZipFromTuples } from "./zip.js";

const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;
const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

test("opens a workbook with an addressable sheet collection", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    expect(workbook.info.format).toBe("xlsx");
    const sheets = workbook.sheets.list();
    expect(sheets.map((sheet) => sheet.info.title)).toEqual(["الميزانية", "Summary", "Shot Log", "Notes"]);
    for (const sheet of sheets) expect(sheet.info.state).toBe("visible");
    expect(workbook.sheets.get(1).info.title).toBe("Summary");
  } finally {
    workbook.close();
  }
});

test("looks sheets up by title — exact, then ignoring case, never a substring", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    expect(workbook.sheets.get("Summary").info.index).toBe(1);
    expect(workbook.sheets.get("summary").info.index).toBe(1);
    expect(() => workbook.sheets.get("Revenue")).toThrow('no sheet titled "Revenue"');
    expect(() => workbook.sheets.get("Summ")).toThrow("no sheet titled");
  } finally {
    workbook.close();
  }
});

test("reads a whole sheet as a full-extent range model", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const sheet = workbook.sheets.get("Summary");
    const model = await sheet.read();
    expect(model.name).toBe("Summary");
    expect(model.range).toEqual({
      start: { row: 0, col: 0 },
      end: { row: model.extent.rows - 1, col: model.extent.cols - 1 },
    });
    expect(model.rows.length).toBeGreaterThan(0);
  } finally {
    workbook.close();
  }
});

test("reads A1 windows with clipped merges and an echoed range", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const sheet = workbook.sheets.get("Summary");
    const window = await sheet.read({ range: "A1:C3" });
    expect(window.range).toEqual({ start: { row: 0, col: 0 }, end: { row: 2, col: 2 } });
    for (const row of window.rows) {
      expect(row.index).toBeLessThan(3);
      for (const { col } of row.cells) expect(col).toBeLessThan(3);
    }
    for (const merge of window.merges) {
      expect(merge.start.row).toBeGreaterThanOrEqual(0);
      expect(merge.end.row).toBeLessThanOrEqual(2);
      expect(merge.start.col).toBeGreaterThanOrEqual(0);
      expect(merge.end.col).toBeLessThanOrEqual(2);
    }
    // A repeated windowed read serves the cached sheet without re-parsing.
    const again = await sheet.read({ range: "A1:C3" });
    expect(again).toEqual(window);
  } finally {
    workbook.close();
  }
});

test("reads single cells and structured ranges", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const sheet = workbook.sheets.get("Summary");
    const single = await sheet.read({ range: "B2" });
    expect(single.range).toEqual({ start: { row: 1, col: 1 }, end: { row: 1, col: 1 } });
    const structured = await sheet.read({ range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } } });
    expect(structured.range).toEqual({ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } });
    const reversed = await sheet.read({ range: "C3:A1" });
    expect(reversed.range).toEqual({ start: { row: 0, col: 0 }, end: { row: 2, col: 2 } });
  } finally {
    workbook.close();
  }
});

test("rejects malformed ranges with invalid-range errors", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const sheet = workbook.sheets.get("Summary");
    for (const bad of ["A:A", "XFE1", "A1:M200:Z9", ""]) {
      const error = await rejection(sheet.read({ range: bad }));
      expect((error as AaredError).code).toBe("invalid-range");
    }
    const numeric = await rejection(
      sheet.read({ range: { start: { row: 5, col: 5 }, end: { row: 1, col: 1 } } }),
    );
    expect((numeric as AaredError).code).toBe("invalid-range");
  } finally {
    workbook.close();
  }
});

test("searches formatted cell text across sheets", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    expect(await workbook.search("")).toEqual([]);
    const matches = await workbook.search("budget");
    expect(matches.length).toBeGreaterThan(0);
    expect(matches.some((match) => match.sheet.title === "Summary")).toBe(true);
    for (const match of matches) {
      // A match names its sheet as { index, title } and its cell as { address, row, col }.
      expect(workbook.sheets.get(match.sheet.index).info.title).toBe(match.sheet.title);
      expect(match.address).toMatch(/^[A-Z]+[0-9]+$/);
      expect(Number.isInteger(match.row) && Number.isInteger(match.col)).toBe(true);
      expect(match.text.length).toBeGreaterThan(0);
    }
  } finally {
    workbook.close();
  }
});

test("rejects non-XLSX files", async () => {
  const error = await rejection(open(await Bun.file(PPTX_FIXTURE).arrayBuffer()));
  expect((error as AaredError).code).toBe("unsupported-format");
  expect(error.message).toContain("aared/pptx");
});

/** A one-sheet workbook around `sheet` (worksheet XML); a damaged sheet is something Excel never writes. */
function workbookWithSheet(sheet: string): ArrayBuffer {
  return storedZipFromTuples([
    [
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
    ],
    [
      "_rels/.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    ],
    [
      "xl/workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Damaged" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
    ],
    [
      "xl/_rels/workbook.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
    ],
    ["xl/worksheets/sheet1.xml", sheet],
  ]);
}

test("a damaged sheet keeps what it can read and says what it skipped", async () => {
  // Two cells and a merge with unreadable addresses, a hyperlink and a
  // conditional-format range likewise: each used to vanish without a trace.
  const workbook = await open(workbookWithSheet(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>kept</t></is></c><c r="??" t="inlineStr"><is><t>lost</t></is></c><c r="1A" t="inlineStr"><is><t>lost</t></is></c></row>
</sheetData>
<mergeCells count="2"><mergeCell ref="A2:B2"/><mergeCell ref="B?:C3"/></mergeCells>
<conditionalFormatting sqref="??"><cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>1</formula></cfRule></conditionalFormatting>
<hyperlinks><hyperlink ref="!!" location="Damaged!A1"/></hyperlinks>
</worksheet>`));
  try {
    const sheet = await workbook.sheets.get(0).read();
    expect(sheet.rows[0]?.cells.map((entry) => entry.cell.formatted)).toEqual(["kept"]);
    expect(sheet.merges).toHaveLength(1);
    const damaged = sheet.unsupported.filter((entry) => entry.kind === "damaged").map((entry) => entry.detail);
    expect(damaged).toEqual([
      "2 cells could not be read and were skipped; the file may be damaged",
      "1 hyperlink could not be read and was skipped; the file may be damaged",
      "1 merged range could not be read and was skipped; the file may be damaged",
      "1 conditional-format range could not be read and was skipped; the file may be damaged",
    ]);
  } finally {
    workbook.close();
  }
});
