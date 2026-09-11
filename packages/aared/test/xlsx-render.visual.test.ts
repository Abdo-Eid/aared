import { expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { readWorkbook } from "../src/core/xlsx/workbook.js";
import { parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { findCoverageSheet } from "./xlsx-fixture.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const COVERAGE_PATH = `${import.meta.dir}/../../../fixtures/arabic-xlsx-rtl-coverage.xlsx`;
const TRACKER_PATH = `${import.meta.dir}/../../../fixtures/post-production-tracker.xlsx`;

async function sheetModels() {
  const coveragePkg = await OpcPackage.open(await Bun.file(COVERAGE_PATH).arrayBuffer());
  const coverageWorkbook = await readWorkbook(coveragePkg);
  const coverage: Awaited<ReturnType<typeof parseXlsxSheet>>[] = [];
  for (let index = 0; index < coverageWorkbook.sheets.length; index++) {
    coverage.push(await parseXlsxSheet(coveragePkg, coverageWorkbook, index));
  }
  const trackerPkg = await OpcPackage.open(await Bun.file(TRACKER_PATH).arrayBuffer());
  const trackerWorkbook = await readWorkbook(trackerPkg);
  return {
    rtl: (await findCoverageSheet(coverage, "rtl-grid")).sheet,
    summary: (await findCoverageSheet(coverage, "summary")).sheet,
    dashboard: await parseXlsxSheet(trackerPkg, trackerWorkbook, 0),
    log: await parseXlsxSheet(trackerPkg, trackerWorkbook, 1),
  };
}

visual("visual DOM output: XLSX grid preserves RTL order, merges, frozen stickiness, and formats", async () => {
  const models = await sheetModels();
  const build = await Bun.build({
    entrypoints: [`${import.meta.dir}/../src/xlsx/index.ts`],
    target: "browser",
    format: "esm",
  });
  if (!build.success) throw new Error(build.logs.map((log) => log.message).join("\n"));
  const module = await build.outputs[0]!.text();
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/xlsx.mjs") {
        return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      return new Response(`<!doctype html><style>html,body{margin:0}#host{width:1280px;height:400px}</style><main id="host"></main><script type="module">
        import { render } from "/xlsx.mjs";
        window.renderedSheet = null;
        window.renderSheet = async (model, options) => {
          window.renderedSheet?.destroy();
          const spreadsheet = {};
          if (options?.headers !== undefined) spreadsheet.headers = options.headers;
          if (options?.gridlines !== undefined) spreadsheet.gridlines = options.gridlines;
          const input = { sheet: model, target: document.querySelector("#host") };
          if (Object.keys(spreadsheet).length > 0) input.appearance = { spreadsheet };
          window.renderedSheet = await render(input);
        };
        window.destroySheet = () => window.renderedSheet?.destroy();
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1280, height: 800 });
    await view.navigate(server.url.href);

    await view.evaluate(`renderSheet(${JSON.stringify(models.rtl)})`);
    const rtlOutput = await view.evaluate(`(() => {
      const grid = document.querySelector('[data-aared-element="grid"]');
      const firstRow = grid?.querySelector('tr[data-aared-row="0"]');
      const titleCell = firstRow?.querySelector('td');
      const headerCell = grid?.querySelector('tr[data-aared-row="2"] td');
      const frozenColumnCell = grid?.querySelector('tr[data-aared-row="3"] td[data-aared-col="0"]');
      const currencyCell = grid?.querySelector('tr[data-aared-row="3"] td[data-aared-col="1"]');
      const dateCell = grid?.querySelector('tr[data-aared-row="9"] td[data-aared-col="1"]');
      const texts = [...grid?.querySelectorAll('[data-aared-element="cell-text"]') ?? []].map((cell) => cell.textContent);
      return {
        direction: grid?.dataset.aaredSheetDirection,
        dir: grid?.dir,
        title: titleCell?.textContent,
        titleSpan: titleCell?.colSpan,
        firstCellCol: titleCell?.dataset.aaredCol,
        titleDir: titleCell?.dir,
        titleFill: titleCell ? getComputedStyle(titleCell).backgroundColor : null,
        headerFill: headerCell ? getComputedStyle(headerCell).backgroundColor : null,
        currencyDir: currencyCell?.dir,
        currencyComputed: currencyCell ? getComputedStyle(currencyCell).direction : null,
        dateDir: dateCell?.dir,
        dateComputed: dateCell ? getComputedStyle(dateCell).direction : null,
        arabicDir: frozenColumnCell?.dir,
        frozenTop: firstRow?.querySelector('td')?.style.top,
        frozenFlag: firstRow?.querySelector('td')?.dataset.aaredFrozen,
        frozenColumnSide: frozenColumnCell?.dataset.aaredFrozen,
        frozenColumnOffset: frozenColumnCell?.style.right,
        frozenColumnLeft: frozenColumnCell?.style.left,
        hiddenColumns: grid?.dataset.aaredHiddenColumns,
        hiddenRows: grid?.dataset.aaredHiddenRows,
        hiddenColumnRendered: grid?.querySelector('td[data-aared-col="3"]') !== null,
        hiddenRowRendered: grid?.querySelector('tr[data-aared-row="6"]') !== null,
        hasEgp: texts.includes("180,000 EGP"),
        hasDate: texts.includes("13 Sep 2026"),
        notices: [...document.querySelectorAll('[data-aared-element="sheet-notices"] li')].map((item) => item.dataset.aaredNotice),
        firstColumnWidth: document.querySelector('colgroup col')?.style.width,
        canvasCount: document.querySelectorAll("canvas").length,
      };
    })()`) as Record<string, unknown>;
    expect(rtlOutput.direction).toBe("rtl");
    expect(rtlOutput.dir).toBe("rtl");
    // DOM order stays logical: the Arabic title is the first cell, and CSS
    // direction mirrors display without reordering columns.
    expect(rtlOutput.title).toBe("ميزانية الإنتاج الشهرية");
    expect(rtlOutput.firstCellCol).toBe("0");
    // The A1:E1 merge spans four *visible* columns: hidden D collapses.
    expect(rtlOutput.titleSpan).toBe(4);
    expect(rtlOutput.titleDir).toBe("rtl");
    // The title keeps its authored dark fill: guards the reserved gray125
    // fill slot (Excel repairs workbooks that put a custom fill there).
    expect(rtlOutput.titleFill).toBe("rgb(18, 23, 31)");
    expect(rtlOutput.headerFill).toBe("rgb(21, 96, 130)");
    // Sheet RTL flips column order, but context-ordered cells keep their own
    // base direction so "180,000 EGP" and "13 Sep 2026" do not visually swap.
    expect(rtlOutput.currencyDir).toBe("auto");
    expect(rtlOutput.currencyComputed).toBe("ltr");
    expect(rtlOutput.dateDir).toBe("auto");
    expect(rtlOutput.dateComputed).toBe("ltr");
    expect(rtlOutput.arabicDir).toBe("rtl");
    // Headers are on by default: frozen cells sit past the 22px header row
    // and the 40px header column.
    expect(rtlOutput.frozenTop).toBe("22px");
    expect(rtlOutput.frozenFlag).toBe("corner");
    expect(rtlOutput.frozenColumnSide).toBe("column");
    // Frozen columns stick to the inline-start edge: physical right in RTL.
    expect(rtlOutput.frozenColumnOffset).toBe("40px");
    expect(rtlOutput.frozenColumnLeft).toBe("");
    expect(rtlOutput.hiddenColumns).toBe("3");
    expect(rtlOutput.hiddenRows).toBe("6");
    expect(rtlOutput.hiddenColumnRendered).toBe(false);
    expect(rtlOutput.hiddenRowRendered).toBe(false);
    expect(rtlOutput.hasEgp).toBe(true);
    expect(rtlOutput.hasDate).toBe(true);
    expect(rtlOutput.notices).toContain("formulas");
    // The first colgroup column is the 40px header gutter (headers default on).
    expect(rtlOutput.firstColumnWidth).toBe("40px");
    expect(rtlOutput.canvasCount).toBe(0);

    await view.evaluate(`renderSheet(${JSON.stringify(models.rtl)}, { headers: true })`);
    const headingsOutput = await view.evaluate(`(() => {
      const grid = document.querySelector('[data-aared-element="grid"]');
      const letters = [...grid?.querySelectorAll('[data-aared-element="column-header"]') ?? []].map((cell) => cell.textContent);
      const numbers = [...grid?.querySelectorAll('[data-aared-element="row-header"]') ?? []].map((cell) => cell.textContent);
      const corner = grid?.querySelector('[data-aared-element="corner"]');
      const letterTop = grid?.querySelector('[data-aared-element="column-header"]')?.style.top;
      const rowHead = grid?.querySelector('[data-aared-element="row-header"]');
      const frozenCorner = grid?.querySelector('tr[data-aared-row="0"] td');
      const frozenColumn = grid?.querySelector('tr[data-aared-row="3"] td[data-aared-col="0"]');
      return {
        flag: grid?.dataset.aaredHeaders,
        letters,
        numbers,
        cornerTop: corner?.style.top,
        cornerSide: corner?.style.right,
        letterTop,
        rowHeadSide: rowHead?.style.right,
        frozenTop: frozenCorner?.style.top,
        frozenColumnSide: frozenColumn?.style.right,
        frameBorder: grid?.style.border,
        canvasCount: document.querySelectorAll("canvas").length,
      };
    })()`) as Record<string, unknown>;
    expect(headingsOutput.flag).toBe("true");
    // Headings skip hidden rows/columns exactly like Excel: C→E, …6, 8….
    expect(headingsOutput.letters).toEqual(["A", "B", "C", "E"]);
    expect(headingsOutput.numbers).toEqual(["1", "2", "3", "4", "5", "6", "8", "9", "10", "11"]);
    // Headings stick to the top and inline-start edges (physical right in RTL).
    expect(headingsOutput.cornerTop).toBe("0px");
    expect(headingsOutput.cornerSide).toBe("0px");
    expect(headingsOutput.letterTop).toBe("0px");
    expect(headingsOutput.rowHeadSide).toBe("0px");
    // Frozen cells sit past the heading chrome: 22px header row, 40px header column.
    expect(headingsOutput.frozenTop).toBe("22px");
    expect(headingsOutput.frozenColumnSide).toBe("40px");
    // The print frame is opt-in: no border without it.
    expect(headingsOutput.frameBorder).toBe("");
    expect(headingsOutput.canvasCount).toBe(0);

    await view.evaluate(`renderSheet(${JSON.stringify(models.rtl)}, { headers: true })`);
    const chromeOutput = await view.evaluate(`(() => {
      const grid = document.querySelector('[data-aared-element="grid"]');
      const link = grid?.querySelector('tr[data-aared-row="10"] a');
      return {
        gridlines: grid?.dataset.aaredGridlines,
        linkText: link?.textContent,
        linkDecoration: link?.style.textDecoration,
        canvasCount: document.querySelectorAll("canvas").length,
      };
    })()`) as Record<string, unknown>;
    expect(chromeOutput.gridlines).toBe("true");
    // Native print shows no underline: the link keeps cell styling only.
    expect(chromeOutput.linkText).toBe("دليل الإنتاج");
    expect(chromeOutput.linkDecoration).toBe("none");
    expect(chromeOutput.canvasCount).toBe(0);

    await view.evaluate(`renderSheet(${JSON.stringify(models.rtl)}, { headers: false, gridlines: "hide" })`);
    const bareOutput = await view.evaluate(`(() => {
      const grid = document.querySelector('[data-aared-element="grid"]');
      return {
        flag: grid?.dataset.aaredHeaders,
        gridlines: grid?.dataset.aaredGridlines,
        letters: [...grid?.querySelectorAll('[data-aared-element="column-header"]') ?? []].length,
        numbers: [...grid?.querySelectorAll('[data-aared-element="row-header"]') ?? []].length,
      };
    })()`) as Record<string, unknown>;
    expect(bareOutput.flag).toBe("false");
    expect(bareOutput.gridlines).toBe("false");
    expect(bareOutput.letters).toBe(0);
    expect(bareOutput.numbers).toBe(0);

    await view.evaluate(`renderSheet(${JSON.stringify(models.summary)})`);
    const summaryOutput = await view.evaluate(`(() => {
      const grid = document.querySelector('[data-aared-element="grid"]');
      const frozen = [...grid?.querySelectorAll('[data-aared-frozen]') ?? []].map((cell) => cell.dataset.aaredFrozen);
      return {
        direction: grid?.dataset.aaredSheetDirection,
        dir: grid?.dir,
        titleSpan: grid?.querySelector('tr[data-aared-row="0"] td')?.colSpan,
        frozen,
      };
    })()`) as { direction?: string; dir?: string; titleSpan?: number; frozen: string[] };
    expect(summaryOutput.direction).toBe("ltr");
    expect(summaryOutput.dir).toBe("ltr");
    expect(summaryOutput.titleSpan).toBe(3);
    expect(summaryOutput.frozen.every((flag) => flag === "row")).toBe(true);
    expect(summaryOutput.frozen.length).toBeGreaterThan(0);

    await view.evaluate(`renderSheet(${JSON.stringify(models.dashboard)})`);
    const drawingOutput = await view.evaluate(`(() => ({
      chartCount: document.querySelectorAll('[data-aared-element="chart-graphic"]').length,
      chartRects: document.querySelectorAll('[data-aared-element="chart-graphic"] rect').length,
      canvasCount: document.querySelectorAll("canvas").length,
    }))()`) as { chartCount: number; chartRects: number; canvasCount: number };
    expect(drawingOutput.chartCount).toBe(1);
    expect(drawingOutput.chartRects).toBeGreaterThan(0);
    expect(drawingOutput.canvasCount).toBe(0);

    await view.evaluate(`renderSheet(${JSON.stringify(models.log)})`);
    const barOutput = await view.evaluate(`(() => {
      const bar = document.querySelector('[data-aared-element="data-bar"]');
      const locked = [...document.querySelectorAll('td')].find((cell) => cell.textContent === "Locked");
      return {
        bar: bar ? { width: bar.style.width, color: bar.style.backgroundColor } : null,
        lockedFill: locked ? getComputedStyle(locked).backgroundColor : null,
      };
    })()`) as { bar: { width: string; color: string } | null; lockedFill: string | null };
    expect(barOutput.bar).not.toBeNull();
    expect(barOutput.bar!.width).toBe("32.5%");
    // Cached conditional-formatting results paint from the model snapshot.
    expect(barOutput.lockedFill).toBe("rgb(221, 235, 232)");

    // Virtualization: inside a short viewport, scrolled rows outside the
    // window collapse into spacers. (Small sheets fit inside overscan and
    // render whole, so this uses the 34-row dashboard.)
    await view.evaluate(`renderSheet(${JSON.stringify(models.dashboard)})`);
    await view.evaluate(`(() => new Promise((resolve) => {
      document.querySelector("#host").style.height = "140px";
      document.querySelector('[data-aared-element="sheet-scroll"]').scrollTop = 100000;
      requestAnimationFrame(() => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    }))()`);
    const scrolledOutput = await view.evaluate(`(() => ({
      lastRowKept: document.querySelector('tr[data-aared-row="33"]') !== null,
      spacers: document.querySelectorAll('tr[data-aared-element="spacer"]').length,
      middleDropped: document.querySelector('tr[data-aared-row="4"]') === null,
    }))()`) as { lastRowKept: boolean; spacers: number; middleDropped: boolean };
    expect(scrolledOutput.lastRowKept).toBe(true);
    expect(scrolledOutput.spacers).toBeGreaterThan(0);
    expect(scrolledOutput.middleDropped).toBe(true);

    const destroyedOutput = await view.evaluate(`(() => {
      destroySheet();
      document.querySelector('[data-aared-element="sheet-scroll"]').scrollTop = 0;
      return { rows: document.querySelectorAll('tr[data-aared-row]').length };
    })()`) as { rows: number };
    expect(destroyedOutput.rows).toBeGreaterThan(0);
  } finally {
    server.stop();
  }
});
