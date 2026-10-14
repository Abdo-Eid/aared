import { afterEach, describe, expect, test } from "bun:test";
import { createTableCell, createTableGrid } from "../src/table-dom.js";

// The shells need a DOM; the unit suite is headless, so a stand-in records what
// is set on each element. The real-DOM output of PPTX and DOCX tables is covered
// by their visual suites.
interface FakeElement {
  attributes: Record<string, string>;
  dataset: Record<string, string>;
  style: Record<string, string>;
  setAttribute(name: string, value: string): void;
}

const original = (globalThis as { document?: unknown }).document;
afterEach(() => {
  (globalThis as { document?: unknown }).document = original;
});

function withFakeDocument<T>(run: () => T): T {
  (globalThis as { document?: unknown }).document = {
    createElement: (): FakeElement => ({
      attributes: {},
      dataset: {},
      style: {},
      setAttribute(name, value) {
        this.attributes[name] = value;
      },
    }),
  };
  return run();
}

const grid = (direction: "ltr" | "rtl", columns: number[]): FakeElement =>
  withFakeDocument(() => createTableGrid(direction, columns)) as unknown as FakeElement;
const cell = (options: Parameters<typeof createTableCell>[0]): FakeElement =>
  withFakeDocument(() => createTableCell(options)) as unknown as FakeElement;

describe("table grid", () => {
  test("is an ARIA table laid out as a CSS grid in the table's own direction", () => {
    const element = grid("rtl", [1, 1]);
    expect(element.attributes.role).toBe("table");
    expect(element.dataset).toMatchObject({ aaredElement: "table", aaredTableDirection: "rtl" });
    expect(element.style).toMatchObject({ display: "grid", direction: "rtl" });
  });

  test("proportions its columns to their widths, or equalises them when none is authored", () => {
    expect(grid("ltr", [1, 3]).style.gridTemplateColumns).toBe("25% 75%");
    expect(grid("ltr", [2, 2, 4]).style.gridTemplateColumns).toBe("25% 25% 50%");
    expect(grid("ltr", [0, 0]).style.gridTemplateColumns).toBe("1fr 1fr");
  });

  test("a table with no column widths at all sets no template", () => {
    expect(grid("ltr", []).style.gridTemplateColumns).toBeUndefined();
  });
});

describe("table cell", () => {
  test("is an ARIA cell that may shrink below its content", () => {
    const element = cell({});
    expect(element.attributes.role).toBe("cell");
    expect(element.dataset.aaredElement).toBe("table-cell");
    expect(element.style).toMatchObject({ minWidth: "0", boxSizing: "border-box" });
    expect(element.style.gridColumn).toBeUndefined();
    expect(element.style.display).toBeUndefined();
  });

  test("spans columns and rows, and marks them for scripts", () => {
    const element = cell({ columnSpan: 2, rowSpan: 3 });
    expect(element.style).toMatchObject({ gridColumn: "span 2", gridRow: "span 3" });
    expect(element.dataset).toMatchObject({ aaredColumnSpan: "2", aaredRowSpan: "3" });
  });

  test("aligns its content vertically as a flex column", () => {
    expect(cell({ verticalAlign: "top" }).style.justifyContent).toBe("flex-start");
    expect(cell({ verticalAlign: "middle" }).style).toMatchObject({ display: "flex", flexDirection: "column", justifyContent: "center" });
    expect(cell({ verticalAlign: "bottom" }).style.justifyContent).toBe("flex-end");
  });
});
