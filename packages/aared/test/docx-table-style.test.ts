/**
 * A table Word draws from a `w:tblStyle` paints that style: banded rows, a
 * shaded header, style borders, and direct formatting on top.
 *
 * The expectations come from Word 16, measured against the file itself
 * (probe `docx-table-style`). Three findings drive the code:
 *
 * - Word authors **no** per-cell shading for a style-driven table. All of the
 *   look lives in `styles.xml`, so a parser that reads only direct formatting
 *   paints such a table plain.
 * - Band counting **excludes the header row**: with `w:tblLook/@firstRow="1"`
 *   the painted band rows were 2, 4, 6, and with `firstRow="0"` they were
 *   1, 3, 5, 7. A vertical merge does not shift the count.
 * - A style region layers *below* the paragraph style and below direct
 *   formatting: a `Heading 2` paragraph in a header row kept its own colour,
 *   and a cell with its own `w:shd` kept that fill.
 */
import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import type { DocxTableNode } from "../src/core/model.js";
import { tableLookFromVal } from "../src/core/table-style.js";
import { FIXTURES, withPart } from "./mutate.js";
import { inheritedTableSections } from "./docx-table-style-harness.js";

/** Every table a parsed fixture holds, in document order. */
async function tablesOf(pkg: OpcPackage): Promise<DocxTableNode[]> {
  const doc = await indexDocx(pkg);
  const sections = await parseDocxSections(pkg, doc);
  const tables: DocxTableNode[] = [];
  const walk = (node: unknown): void => {
    if (Array.isArray(node)) {
      for (const child of node) walk(child);
      return;
    }
    if (node === null || typeof node !== "object") return;
    if ((node as { type?: string }).type === "table") tables.push(node as DocxTableNode);
    // A cell's `nodes` are its content, which never holds a table.
    for (const [key, value] of Object.entries(node as Record<string, unknown>)) {
      if (key !== "nodes") walk(value);
    }
  };
  for (const section of sections) walk(section);
  return tables;
}

/** The fixture's last table: the one Word styles (the rest carry direct formatting). */
async function styledTable(): Promise<DocxTableNode> {
  const pkg = await OpcPackage.open(await Bun.file(`${FIXTURES}/docx-coverage.docx`).arrayBuffer());
  const table = (await tablesOf(pkg)).at(-1);
  if (table === undefined) throw new Error("the fixture has no table");
  return table;
}

/** The shading each row's cells resolved to, `-` for none. */
function shades(table: DocxTableNode): string[][] {
  return table.rows.map((row) => row.cells.map((cell) => cell.shading ?? "-"));
}

describe("a DOCX table style", () => {
  test("inherits every table edge and padding side through styles with no table properties", async () => {
    const sections = await inheritedTableSections();
    const table = sections.flatMap(section => section.root).filter(node => node.type === "table").at(-1)!;
    for (const side of ["top", "right", "bottom", "left", "insideH", "insideV"] as const) {
      expect(table.borders?.[side]).toMatchObject({ color: "#123456", widthPt: 1, style: "solid" });
    }
    expect(table.cellMargins).toEqual({ top: .1, right: .2, bottom: .3, left: .4 });
  });

  test("merges partial style overrides base-first and keeps direct formatting last, including none and zero", async () => {
    const middle = '<w:tblPr><w:tblBorders><w:right w:val="double" w:sz="16" w:color="654321"/></w:tblBorders>' +
      '<w:tblCellMar><w:right w:w="720" w:type="dxa"/></w:tblCellMar></w:tblPr>';
    const child = '<w:tblPr><w:tblBorders><w:top w:val="nil"/></w:tblBorders>' +
      '<w:tblCellMar><w:left w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>';
    const direct = '<w:tblBorders><w:right w:val="none"/></w:tblBorders>' +
      '<w:tblCellMar><w:right w:w="0" w:type="dxa"/></w:tblCellMar>';
    const inherited = await inheritedTableSections(child, "", middle);
    const table = inherited.flatMap(section => section.root).filter(node => node.type === "table").at(-1)!;
    expect(table.borders?.top).toBe("none");
    expect(table.borders?.right).toMatchObject({ color: "#654321", widthPt: 2, style: "double" });
    expect(table.borders?.insideH).toMatchObject({ color: "#123456", widthPt: 1 });
    expect(table.cellMargins).toEqual({ top: .1, right: .5, bottom: .3, left: 0 });
    const sections = await inheritedTableSections(child, direct, middle);
    const overridden = sections.flatMap(section => section.root).filter(node => node.type === "table").at(-1)!;
    expect(overridden.borders).toEqual({ ...table.borders, right: "none" });
    expect(overridden.cellMargins).toEqual({ ...table.cellMargins, right: 0 });
  });

  test("bands the body rows and skips the header row in the count", async () => {
    const table = await styledTable();
    // `GridTable4-Accent1`: firstRow fills `156082`, band1Horz fills `C1E4F5`,
    // and there is no band2Horz, so band 2 paints nothing.
    expect(shades(table)).toEqual([
      ["#156082", "#156082", "#156082"],
      // The middle cell of this band row carries a direct fill of its own.
      ["#C1E4F5", "#00FF00", "#C1E4F5"],
      ["-", "-", "-"],
      ["#C1E4F5", "#C1E4F5", "#C1E4F5"],
      // Row 5 keeps the band's count (it is band 2) and loses column one to
      // the merge above, so it holds two cells.
      ["-", "-"],
      ["#C1E4F5", "#C1E4F5", "#C1E4F5"],
    ]);
    // A vertical merge does not shift the band count: the covered row keeps
    // its band but loses the cell the merge takes.
    expect(table.rows[3]!.cells[0]!.rowSpan).toBe(2);
  });

  test("paints the style's own borders when the table states none", async () => {
    const table = await styledTable();
    // Word paints `Table Grid`'s edges from `styles.xml`; a table with no
    // `w:tblBorders` of its own is not unbordered.
    expect(table.borders?.top).toMatchObject({ style: "solid", widthPt: 0.5 });
    expect(table.borders?.insideH).toMatchObject({ style: "solid", widthPt: 0.5 });
    expect(table.borders?.insideV).toMatchObject({ style: "solid", widthPt: 0.5 });
  });

  test("a paragraph style's colour beats the region's", async () => {
    // The header row's region authors white ink, and the fixture's header cells
    // carry no `w:pStyle`, so they take it. Give the cell holding "Q1" the
    // `Heading2` paragraph style (which authors `12171F`) and that colour must
    // win instead: a style region layers *below* the paragraph style (probe
    // `docx-table-style`, T5).
    const mutated = await withPart("docx-coverage.docx", "word/document.xml", (xml) => {
      const text = xml.indexOf("<w:t>Q1</w:t>");
      if (text < 0) throw new Error("the header cell holding Q1 is gone");
      // The paragraph that owns that run, then its `w:pPr` (every fixture
      // paragraph has one), then the insertion point inside it.
      const paragraph = xml.lastIndexOf("<w:p ", text);
      const pPr = xml.indexOf("<w:pPr>", paragraph);
      if (paragraph < 0 || pPr < 0 || pPr > text) throw new Error("that paragraph has no w:pPr");
      const at = pPr + "<w:pPr>".length;
      return `${xml.slice(0, at)}<w:pStyle w:val="Heading2"/>${xml.slice(at)}`;
    });
    const cell = (await tablesOf(mutated)).at(-1)!.rows[0]!.cells[1]!.nodes[0];
    const run = cell && "runs" in cell ? cell.runs[0] : undefined;
    expect(run?.color).toBe("#12171F");
  });

  test("a region's colour stands when the paragraph style states none", async () => {
    // The mirror of the case above: no `w:pStyle`, so nothing outranks the
    // region and the header paints the region's white ink.
    const cell = (await styledTable()).rows[0]!.cells[1]!.nodes[0];
    const run = cell && "runs" in cell ? cell.runs[0] : undefined;
    expect(run?.color).toBe("#FFFFFF");
  });

  test("a table with no style keeps painting only what it authors directly", async () => {
    const table = await styledTable();
    // The fixture's first table states its own formatting and names no style;
    // this is the shape every other fixture table has, and it must be unchanged.
    const pkg = await OpcPackage.open(await Bun.file(`${FIXTURES}/docx-coverage.docx`).arrayBuffer());
    const tables = await tablesOf(pkg);
    // A direct fill still paints; the style contributes nothing to it.
    const styled = new Set(["#156082", "#C1E4F5"]);
    for (const row of tables[0]!.rows) {
      for (const cell of row.cells) {
        if (cell.shading !== undefined) expect(styled.has(cell.shading.toUpperCase())).toBe(false);
      }
    }
    expect(table.rows.length).toBe(6);
  });
});

describe("w:tblLook", () => {
  test("reads the explicit attributes", () => {
    expect(tableLookFromVal("0420")).toEqual({
      firstRow: true, lastRow: false, firstColumn: false, lastColumn: false, bandRow: true, bandColumn: false,
    });
  });

  test("treats the no-band bits as the absence of banding", () => {
    // `0600` is what Word wrote for a table with both bandings off.
    expect(tableLookFromVal("0600")).toEqual({
      firstRow: false, lastRow: false, firstColumn: false, lastColumn: false, bandRow: false, bandColumn: false,
    });
  });

  test("rejects anything that is not a four-digit bitmask", () => {
    expect(tableLookFromVal(undefined)).toBeUndefined();
    expect(tableLookFromVal("")).toBeUndefined();
    expect(tableLookFromVal("42")).toBeUndefined();
    expect(tableLookFromVal("zzzz")).toBeUndefined();
  });
});
