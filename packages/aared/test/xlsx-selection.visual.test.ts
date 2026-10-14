import { expect } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { readWorkbook } from "../src/core/xlsx/workbook.js";
import { parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { browserModule, visual, withVisualPage } from "./harness.js";

interface CellRef {
  address: string;
  row: number;
  col: number;
}

const TRACKER_PATH = `${import.meta.dir}/../../../fixtures/post-production-tracker.xlsx`;
const COVERAGE_PATH = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;

/**
 * The empty area past the data selects like Excel's: a click, a drag and the
 * arrow keys reach its cells, numbered past the sheet's extent, and the
 * selection is drawn there too. Cells on it show Excel's plus pointer; the
 * headings do not. The wide host has empty columns past the Dashboard's
 * eight; the tall host has empty rows under the budget sheet's data.
 */
visual("visual DOM output: XLSX selection reaches the empty area past the data", async () => {
  const trackerPkg = await OpcPackage.open(await Bun.file(TRACKER_PATH).arrayBuffer());
  const dashboard = await parseXlsxSheet(trackerPkg, await readWorkbook(trackerPkg), 0);
  const coveragePkg = await OpcPackage.open(await Bun.file(COVERAGE_PATH).arrayBuffer());
  const budget = await parseXlsxSheet(coveragePkg, await readWorkbook(coveragePkg), 0);
  const module = await browserModule(`${import.meta.dir}/../src/xlsx/renderer.ts`);
  await withVisualPage({
    modules: { "/xlsx.mjs": module },
    page: `<!doctype html><style>html,body{margin:0}#wide{width:1280px;height:500px}#tall{width:1280px;height:700px}</style><main id="wide"></main><main id="tall"></main><script type="module">
        import { render } from "/xlsx.mjs";
        const appearance = { spreadsheet: { formulaBar: true } };
        window.wide = await render({ sheet: ${JSON.stringify(dashboard)}, target: document.querySelector("#wide"), appearance });
        window.tall = await render({ sheet: ${JSON.stringify(budget)}, target: document.querySelector("#tall"), appearance });
        window.trackCenter = (host, row, track) => {
          const td = document.querySelector('#' + host + ' tr[data-aared-row="' + row + '"] td[data-aared-fill="tracks"]');
          const rect = td.getBoundingClientRect();
          return { x: rect.left + (rect.width / Number(td.colSpan)) * (track + 0.5), y: rect.top + rect.height / 2 };
        };
        window.press = (x, y) => {
          const element = document.elementFromPoint(x, y);
          for (const type of ["mousedown", "mouseup", "click"]) {
            element.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, button: 0, clientX: x, clientY: y }));
          }
          return element;
        };
        window.dragTo = (from, to) => {
          const start = document.elementFromPoint(from.x, from.y);
          start.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true, button: 0, clientX: from.x, clientY: from.y }));
          const end = document.elementFromPoint(to.x, to.y);
          end.dispatchEvent(new MouseEvent("mousemove", { bubbles: true, cancelable: true, button: 0, clientX: to.x, clientY: to.y }));
          document.querySelector("#wide").dispatchEvent(new MouseEvent("mouseup", { bubbles: true, cancelable: true, button: 0 }));
          end.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, button: 0, clientX: to.x, clientY: to.y }));
        };
        window.nameBox = (host) => document.querySelector('#' + host + ' [data-aared-element="sheet-name-box"]').textContent;
        window.pressKey = (key, shift) => document.activeElement.dispatchEvent(new KeyboardEvent("keydown", { key, shiftKey: !!shift, bubbles: true, cancelable: true }));
      </script>`,
    width: 1280,
    height: 900,
  }, async (view) => {
    await view.evaluate(`new Promise((resolve) => { const poll = () => (window.wide && window.tall ? resolve(true) : setTimeout(poll, 20)); poll(); })`);

    // The plus pointer shows over data and empty cells, never over a heading.
    expect(await view.evaluate<string>(`getComputedStyle(document.querySelector('#wide td[data-aared-fill="tracks"]')).cursor`)).toBe("cell");
    expect(await view.evaluate<string>(`getComputedStyle(document.querySelector('#wide th[data-aared-element="column-header"]')).cursor`)).toBe("auto");

    // A click on the first empty column names it, numbered past the Dashboard's eight columns.
    await view.evaluate(`(() => { const p = trackCenter("wide", 0, 0); press(p.x, p.y); })()`);
    expect(await view.evaluate<CellRef | null>(`window.wide.selection`)).toEqual({ address: "I1", row: 0, col: 8 });
    expect(await view.evaluate<string>(`nameBox("wide")`)).toBe("I1");
    // Its frame is drawn over the empty area itself.
    expect(await view.evaluate<boolean>(`document.querySelector('#wide [data-aared-element="sheet-selection-cursor"]') !== null`)).toBe(true);

    // The next empty column is J: each track is its own column.
    await view.evaluate(`(() => { const p = trackCenter("wide", 0, 1); press(p.x, p.y); })()`);
    expect(await view.evaluate<string>(`nameBox("wide")`)).toBe("J1");

    // A drag from a data cell over the empty columns selects the whole run.
    await view.evaluate(`(() => {
      const from = document.querySelector('#wide td[data-aared-row="10"][data-aared-col="1"]').getBoundingClientRect();
      const to = trackCenter("wide", 12, 0);
      dragTo({ x: from.left + from.width / 2, y: from.top + from.height / 2 }, to);
    })()`);
    expect(await view.evaluate<{ start: CellRef; end: CellRef } | null>(`window.wide.range`)).toEqual({
      start: { address: "B11", row: 10, col: 1 },
      end: { address: "I13", row: 12, col: 8 },
    });
    expect(await view.evaluate<string>(`nameBox("wide")`)).toBe("B11:I13");
    // The wash sits over a header's own fill: its dark fill and white text stay readable.
    expect(await view.evaluate<{ fill: string; ink: string; wash: boolean }>(`(() => {
      const td = document.querySelector('#wide td[data-aared-row="10"][data-aared-col="1"]');
      const style = getComputedStyle(td);
      return { fill: style.backgroundColor, ink: style.color, wash: style.boxShadow.includes("inset") };
    })()`)).toEqual({ fill: "rgb(18, 23, 31)", ink: "rgb(255, 255, 255)", wash: true });
    expect(await view.evaluate<number>(`document.querySelectorAll('#wide [data-aared-element="sheet-selection-range"]').length`)).toBeGreaterThan(0);
    expect(await view.evaluate<number>(`document.querySelectorAll('#wide [data-aared-element="sheet-selection-boundary"]').length`)).toBe(0);
    expect(await view.evaluate<string>(`getComputedStyle(document.querySelector('#wide [data-aared-element="sheet-selection-cursor"]')).outlineStyle`)).toBe("solid");

    // Restore the original selection appearance: a cell tint and a solid
    // outline only on the active cell.
    await view.evaluate(`window.wide.focus("B11")`);
    await view.evaluate(`pressKey("ArrowDown", true)`);
    expect(await view.evaluate<number>(`document.querySelectorAll('#wide [data-aared-element="sheet-selection-boundary"]').length`)).toBe(0);
    expect(await view.evaluate<string[]>(`[...document.querySelectorAll('#wide td[data-aared-in-range="true"]')].map((td) => getComputedStyle(td).outlineStyle)`))
      .toEqual(["none", "solid"]);

    // The arrow keys walk into the empty columns: from H1 one step right is I1.
    await view.evaluate(`window.wide.focus("H1")`);
    await view.evaluate(`pressKey("ArrowRight")`);
    expect(await view.evaluate<CellRef | null>(`window.wide.selection`)).toEqual({ address: "I1", row: 0, col: 8 });

    // Empty rows under the budget sheet's data: a click names a row past the extent, and Home stays on the data.
    const rows = budget.extent.rows;
    await view.evaluate(`(() => {
      const td = document.querySelector('#tall tr[data-aared-fill-row="0"] td[data-aared-fill="row"]');
      td.scrollIntoView({ block: "center" });
      const rect = td.getBoundingClientRect();
      press(rect.right - 40, rect.top + rect.height / 2);
    })()`);
    expect(await view.evaluate<CellRef | null>(`window.tall.selection`)).toEqual({ address: `A${rows + 1}`, row: rows, col: 0 });
    expect(await view.evaluate<string>(`nameBox("tall")`)).toBe(`A${rows + 1}`);
    // ArrowDown from the last data row walks into the empty row below it.
    await view.evaluate(`window.tall.focus("A${rows}")`);
    await view.evaluate(`pressKey("ArrowDown")`);
    expect(await view.evaluate<CellRef | null>(`window.tall.selection`)).toEqual({ address: `A${rows + 1}`, row: rows, col: 0 });
  });
});
