import { expect } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { readWorkbook } from "../src/core/xlsx/workbook.js";
import { parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { browserModule, visual, withVisualPage } from "./harness.js";

const TRACKER_PATH = `${import.meta.dir}/../../../fixtures/post-production-tracker.xlsx`;

interface ChromeText {
  name: string;
  formula: string;
  aggregate: string | null;
}

interface CellRef {
  address: string;
  row: number;
  col: number;
}

interface CellRange {
  start: CellRef;
  end: CellRef;
}

/**
 * The Excel-like viewer chrome re-reads rendered DOM: the name box and
 * formula text follow the cursor, Shift+arrows and drags extend a painted
 * range, and the formula-bar dropdown aggregates cached values. Nothing here
 * evaluates a formula.
 */
visual("visual DOM output: XLSX formula bar with cached range aggregates", async () => {
  const pkg = await OpcPackage.open(await Bun.file(TRACKER_PATH).arrayBuffer());
  const workbook = await readWorkbook(pkg);
  const sheet = await parseXlsxSheet(pkg, workbook, 0);
  const coveragePackage = await OpcPackage.open(await Bun.file(`${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`).arrayBuffer());
  const coverageWorkbook = await readWorkbook(coveragePackage);
  const budgetReferenceSheet = await parseXlsxSheet(coveragePackage, coverageWorkbook, 0);
  const referenceSheet = await parseXlsxSheet(coveragePackage, coverageWorkbook, 1);

  const module = await browserModule(`${import.meta.dir}/../src/xlsx/index.ts`);
  await withVisualPage({
    modules: { "/xlsx.mjs": module },
    page: `<!doctype html><style>html,body{margin:0}#host{width:1280px;height:500px}#plain{width:1280px;height:500px}#refs{width:1280px;height:500px}#budget{width:1280px;height:500px}</style><main id="host"></main><main id="plain"></main><main id="refs"></main><main id="budget"></main><script type="module">
        import { render } from "/xlsx.mjs";
        const chrome = { spreadsheet: { formulaBar: true } };
        window.viewer = await render({ sheet: ${JSON.stringify(sheet)}, target: document.querySelector("#host"), appearance: chrome });
        window.plain = await render({ sheet: ${JSON.stringify(sheet)}, target: document.querySelector("#plain") });
        window.refsViewer = await render({ sheet: ${JSON.stringify(referenceSheet)}, target: document.querySelector("#refs"), appearance: chrome });
        window.budgetRefsViewer = await render({ sheet: ${JSON.stringify(budgetReferenceSheet)}, target: document.querySelector("#budget"), appearance: chrome });
        window.events = [];
        window.viewer.addEventListener("selectionchange", (event) => window.events.push(event.detail));
        window.cell = (row, col) => document.querySelector('#host td[data-aared-row="' + row + '"][data-aared-col="' + col + '"]');
        window.refCell = (row, col) => document.querySelector('#refs td[data-aared-row="' + row + '"][data-aared-col="' + col + '"]');
        window.budgetRefCell = (row, col) => document.querySelector('#budget td[data-aared-row="' + row + '"][data-aared-col="' + col + '"]');
        window.clickCell = (row, col) => {
          const td = window.cell(row, col);
          td.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
          td.focus();
        };
        window.pressKey = (key, shift) => {
          document.activeElement.dispatchEvent(new KeyboardEvent("keydown", { key, shiftKey: !!shift, bubbles: true, cancelable: true }));
        };
        window.dragRange = (fromRow, fromCol, toRow, toCol) => {
          const from = window.cell(fromRow, fromCol);
          from.dispatchEvent(new MouseEvent("mousedown", { button: 0, bubbles: true, cancelable: true }));
          const to = window.cell(toRow, toCol);
          to.dispatchEvent(new MouseEvent("mousemove", { button: 0, bubbles: true, cancelable: true }));
          document.querySelector("#host").dispatchEvent(new MouseEvent("mouseup", { bubbles: true, cancelable: true }));
          // The release click must not collapse the range the drag made.
          to.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
        };
        window.chromeText = () => ({
          name: document.querySelector('#host [data-aared-element="sheet-name-box"]')?.textContent ?? null,
          formula: document.querySelector('#host [data-aared-element="sheet-formula-text"]')?.textContent ?? null,
          aggregate: document.querySelector('#host [data-aared-element="sheet-aggregate-button"]')?.textContent ?? null,
        });
        window.inRange = () => [...document.querySelectorAll('#host td[data-aared-in-range="true"]')].map((td) =>
          [Number(td.dataset.aaredRow), Number(td.dataset.aaredCol)]);
        window.references = (root) => [...document.querySelectorAll(root + ' td[data-aared-formula-reference="true"]')]
          .map((td) => [Number(td.dataset.aaredRow), Number(td.dataset.aaredCol)])
          .sort((a, b) => a[0] - b[0] || a[1] - b[1]);
        window.changeColors = () => [4, 5, 6, 7, 8].map((row) => getComputedStyle(window.refCell(row, 4)).color);
      </script>`,
    width: 1280,
    height: 900,
  }, async (view) => {
      await view.evaluate(`new Promise((resolve) => { const poll = () => (window.viewer && window.refsViewer && window.budgetRefsViewer ? resolve(true) : setTimeout(poll, 20)); poll(); })`);

    // The chrome stays out of a default render.
    expect(await view.evaluate(`document.querySelector('#plain [data-aared-element="sheet-formula-bar"]')`)).toBeNull();

    // A formula cell: the name box names it, the bar shows its source.
    await view.evaluate(`clickCell(11, 1)`);
    expect(await view.evaluate<ChromeText>(`window.chromeText()`)).toEqual({
      name: "B12",
      formula: "SUMIFS('Shot Log'!$F$5:$F$16,'Shot Log'!$B$5:$B$16,4)",
      aggregate: "",
    });
    expect(await view.evaluate<number[][]>(`window.references('#host')`)).toEqual([]);

    // A same-sheet formula lights up the cells it reads; its authored
    // formula result remains cached and no other-sheet cells are guessed.
    // Summary's table runs A4:G9 under its title band: E5 = D5/C5-1.
    expect(await view.evaluate<boolean>(`window.refCell(4, 4) !== null`)).toBe(true);
    await view.evaluate(`window.refCell(4, 4).click()`);
    expect(await view.evaluate<number[][]>(`window.references('#refs')`)).toEqual([[4, 2], [4, 3]]);
    expect(await view.evaluate<string[]>(`[...document.querySelectorAll('#refs [data-aared-element="sheet-reference-boundary"]')].map((frame) => getComputedStyle(frame).borderTopColor)`))
      .toEqual(["rgb(47, 111, 219)", "rgb(232, 119, 46)"]);
    await view.evaluate(`window.refCell(4, 5).click()`);
    expect(await view.evaluate<number[][]>(`window.references('#refs')`)).toEqual([[4, 4]]);
    await view.evaluate(`window.refCell(8, 2).click()`);
    expect(await view.evaluate<number[][]>(`window.references('#refs')`)).toEqual([[4, 2], [5, 2], [6, 2], [7, 2]]);
    // The one range this formula reads is the first colour; its cells carry
    // the tint, and one dashed boundary encloses all of them.
    expect(await view.evaluate<string[]>(`[...document.querySelectorAll('#refs td[data-aared-formula-reference="true"]')].map((td) => td.dataset.aaredReferenceIndex)`)).toEqual(["0", "0", "0", "0"]);
    expect(await view.evaluate<boolean>(`getComputedStyle(window.refCell(4, 2)).boxShadow.includes("47, 111, 219")`)).toBe(true);
    expect(await view.evaluate<{ count: number; style: string; color: string }>(`(() => {
      const frames = document.querySelectorAll('#refs [data-aared-element="sheet-reference-boundary"]');
      const frame = getComputedStyle(frames[0]);
      return { count: frames.length, style: frame.borderTopStyle, color: frame.borderTopColor };
    })()`)).toEqual({ count: 1, style: "dashed", color: "rgb(47, 111, 219)" });
    const referenceBounds = await view.evaluate<number[]>(`(() => {
      const frame = document.querySelector('#refs [data-aared-element="sheet-reference-boundary"]').getBoundingClientRect();
      const cells = [...document.querySelectorAll('#refs td[data-aared-formula-reference="true"]')].map((cell) => cell.getBoundingClientRect());
      return [frame.left, frame.right, frame.top, frame.bottom,
        Math.min(...cells.map((rect) => rect.left)), Math.max(...cells.map((rect) => rect.right)),
        Math.min(...cells.map((rect) => rect.top)), Math.max(...cells.map((rect) => rect.bottom))];
    })()`);
    expect(referenceBounds.slice(0, 4).every((value, index) => Math.abs(value - referenceBounds[index + 4]!) < 1)).toBe(true);

    // The coverage workbook's budget sheet has the exact C9 = SUM(C4:C8)
    // case: one dashed rectangle encloses its five precedent cells.
    await view.evaluate(`window.budgetRefCell(8, 2).click()`);
    expect(await view.evaluate<string>(`document.querySelector('#budget [data-aared-element="sheet-formula-text"]').textContent`)).toBe("SUM(C4:C8)");
    expect(await view.evaluate<number[][]>(`[...document.querySelectorAll('#budget td[data-aared-formula-reference="true"]')].map((td) => [Number(td.dataset.aaredRow), Number(td.dataset.aaredCol)])`))
      .toEqual([[3, 2], [4, 2], [5, 2], [7, 2]]);
    expect(await view.evaluate<{ count: number; style: string; color: string }>(`(() => {
      const frames = document.querySelectorAll('#budget [data-aared-element="sheet-reference-boundary"]');
      const frame = getComputedStyle(frames[0]);
      return { count: frames.length, style: frame.borderTopStyle, color: frame.borderTopColor };
    })()`)).toEqual({ count: 1, style: "dashed", color: "rgb(47, 111, 219)" });
    const budgetReferenceBounds = await view.evaluate<number[]>(`(() => {
      const frame = document.querySelector('#budget [data-aared-element="sheet-reference-boundary"]').getBoundingClientRect();
      const cells = [...document.querySelectorAll('#budget td[data-aared-formula-reference="true"]')].map((cell) => cell.getBoundingClientRect());
      return [frame.left, frame.right, frame.top, frame.bottom,
        Math.min(...cells.map((rect) => rect.left)), Math.max(...cells.map((rect) => rect.right)),
        Math.min(...cells.map((rect) => rect.top)), Math.max(...cells.map((rect) => rect.bottom))];
    })()`);
    expect(budgetReferenceBounds.slice(0, 4).every((value, index) => Math.abs(value - budgetReferenceBounds[index + 4]!) < 1)).toBe(true);

    await view.evaluate(`window.refCell(5, 0).click()`);
    expect(await view.evaluate<number[][]>(`window.references('#refs')`)).toEqual([]);
    // A range is its own highlight: the formula cell's references are not outlined under it.
    await view.evaluate(`window.refCell(4, 4).click()`);
    expect(await view.evaluate<number[][]>(`window.references('#refs')`)).toEqual([[4, 2], [4, 3]]);
    await view.evaluate(`window.refCell(4, 4).dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", shiftKey: true, bubbles: true, cancelable: true }))`);
    expect(await view.evaluate<number[][]>(`window.references('#refs')`)).toEqual([]);
    expect(await view.evaluate<string[]>(`window.changeColors()`)).toEqual([
      "rgb(46, 122, 59)", "rgb(46, 122, 59)", "rgb(176, 63, 46)", "rgb(46, 122, 59)", "rgb(46, 122, 59)",
    ]);

    // Shift+Down extends from the anchor; the bar dropdown aggregates the
    // cached numeric values, and selecting an item changes its summary.
    await view.evaluate(`clickCell(11, 2)`);
    await view.evaluate(`pressKey("ArrowDown", true)`);
    await view.evaluate(`pressKey("ArrowDown", true)`);
    expect(await view.evaluate<CellRef>(`window.viewer.selection`)).toEqual({ address: "C14", row: 13, col: 2 });
    expect(await view.evaluate<CellRange>(`window.viewer.range`))
      .toEqual({ start: { address: "C12", row: 11, col: 2 }, end: { address: "C14", row: 13, col: 2 } });
    expect(await view.evaluate<ChromeText>(`window.chromeText()`)).toEqual({
      name: "C12:C14",
      formula: "SUMIFS('Shot Log'!$H$5:$H$16,'Shot Log'!$B$5:$B$16,6)",
      aggregate: "Sum: 74",
    });
    expect(await view.evaluate(`document.querySelector('#host [data-aared-element="sheet-status-bar"]')`)).toBeNull();
    await view.evaluate(`document.querySelector('#host [data-aared-element="sheet-aggregate-button"]').click()`);
    expect(await view.evaluate<{ insideBar: boolean; menuFits: boolean; itemsFit: boolean; menuOnTop: boolean }>(`(() => {
      const bar = document.querySelector('#host [data-aared-element="sheet-formula-bar"]');
      const menu = document.querySelector('#host [data-aared-element="sheet-aggregate-menu"]');
      bar.style.width = '220px';
      const result = {
        insideBar: menu.getBoundingClientRect().right <= bar.getBoundingClientRect().right,
        menuFits: menu.scrollWidth <= menu.clientWidth,
        itemsFit: [...menu.querySelectorAll('button')].every((item) => item.scrollWidth <= item.clientWidth),
        menuOnTop: (() => {
          const rect = menu.getBoundingClientRect();
          return menu.contains(document.elementFromPoint(rect.left + 24, rect.top + 18));
        })(),
      };
      bar.style.width = '';
      return result;
    })()`)).toEqual({ insideBar: true, menuFits: true, itemsFit: true, menuOnTop: true });
    expect(await view.evaluate(` [...document.querySelectorAll('#host [role="menuitemradio"]')].map((item) => [item.textContent, item.getAttribute('aria-checked')])`) as string[][]).toEqual([
      ["xSum74", "true"], ["Average24.67", "false"], ["Min8", "false"], ["Max39", "false"], ["Count3", "false"],
    ]);
    await view.evaluate(`document.querySelector('#host [data-aggregate="average"]').click()`);
    expect(await view.evaluate<ChromeText>(`window.chromeText()`)).toMatchObject({ aggregate: "Average: 24.67" });
    expect(await view.evaluate(`document.querySelector('#host [data-aared-element="sheet-aggregate-menu"]').hidden`) as boolean).toBe(true);
    expect(await view.evaluate<number[][]>(`window.inRange()`)).toEqual([[11, 2], [12, 2], [13, 2]]);
    const last = await view.evaluate(`window.events.at(-1)`) as { selection: unknown; range: unknown };
    expect(last.selection).toEqual({ address: "C14", row: 13, col: 2 });
    expect(last.range).toEqual({ start: { address: "C12", row: 11, col: 2 }, end: { address: "C14", row: 13, col: 2 } });

    // A plain click collapses back to the cursor.
    await view.evaluate(`clickCell(10, 0)`);
    expect(await view.evaluate<CellRange | null>(`window.viewer.range`)).toBeNull();
    expect(await view.evaluate<number[][]>(`window.inRange()`)).toEqual([]);
    expect(await view.evaluate<ChromeText>(`window.chromeText()`)).toEqual({ name: "A11", formula: "Episode", aggregate: "" });

    // A drag paints the same range, and its release click keeps it.
    await view.evaluate(`dragRange(11, 2, 13, 2)`);
    expect(await view.evaluate<CellRange>(`window.viewer.range`))
      .toEqual({ start: { address: "C12", row: 11, col: 2 }, end: { address: "C14", row: 13, col: 2 } });
    expect(await view.evaluate<number[][]>(`window.inRange()`)).toEqual([[11, 2], [12, 2], [13, 2]]);
    expect(await view.evaluate<ChromeText>(`window.chromeText()`)).toEqual({
      name: "C12:C14",
      formula: "SUMIFS('Shot Log'!$H$5:$H$16,'Shot Log'!$B$5:$B$16,6)",
      aggregate: "Average: 24.67",
    });
  });
});
