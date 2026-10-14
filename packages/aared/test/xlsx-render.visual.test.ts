import { expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { readWorkbook } from "../src/core/xlsx/workbook.js";
import { parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { findCoverageSheet } from "./xlsx-fixture.js";
import { browserModule, visual, withVisualPage } from "./harness.js";
import { withPart } from "./mutate.js";
import { conditionalFormulaFixture } from "./xlsx-conditional-fixture.js";
import { fontAlias } from "../src/core/typography.js";

const COVERAGE_PATH = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;
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
    deep: (await findCoverageSheet(coverage, "deep-log")).sheet,
    dashboard: await parseXlsxSheet(trackerPkg, trackerWorkbook, 0),
    log: await parseXlsxSheet(trackerPkg, trackerWorkbook, 1),
  };
}

visual("visual DOM output: XLSX grid preserves RTL order, merges, frozen stickiness, and formats", async () => {
  const models = await sheetModels();
  const module = await browserModule(`${import.meta.dir}/../src/xlsx/index.ts`);
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
        // A bounded card wrapping a content-sized host, the shape a bench or
        // embedder card gives the sheet: no height of its own, clipped by an
        // ancestor with overflow.
        window.renderInCard = async (model) => {
          window.renderedCard?.destroy();
          let card = document.querySelector("#card");
          if (card === null) {
            card = document.createElement("div");
            card.id = "card";
            card.style.cssText = "width:900px;height:300px;overflow:auto;border:1px solid #999";
            const inner = document.createElement("div");
            inner.id = "card-host";
            card.append(inner);
            document.body.append(card);
          }
          window.renderedCard = await render({ sheet: model, target: document.querySelector("#card-host") });
        };
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  await withVisualPage({ server, width: 1280, height: 800 }, async (view) => {

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
      const letter = grid?.querySelector('[data-aared-element="column-header"]');
      const rowHead = grid?.querySelector('[data-aared-element="row-header"]');
      const frozenCorner = grid?.querySelector('tr[data-aared-row="0"] td');
      const frozenColumn = grid?.querySelector('tr[data-aared-row="3"] td[data-aared-col="0"]');
      return {
        flag: grid?.dataset.aaredHeaders,
        letters,
        numbers,
        // Headings take their look from the grid's class rules, not inline style.
        cornerInline: corner?.getAttribute("style"),
        cornerTop: corner && getComputedStyle(corner).top,
        cornerSide: corner && getComputedStyle(corner).right,
        letterTop: letter && getComputedStyle(letter).top,
        rowHeadSide: rowHead && getComputedStyle(rowHead).right,
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
    expect(headingsOutput.cornerInline).toBeNull();
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
        linkDecoration: link && getComputedStyle(link).textDecorationLine,
        linkColorInline: link?.style.color,
        canvasCount: document.querySelectorAll("canvas").length,
      };
    })()`) as Record<string, unknown>;
    expect(chromeOutput.gridlines).toBe("true");
    // Native print shows no underline: the link keeps cell styling only.
    expect(chromeOutput.linkText).toBe("دليل الإنتاج");
    expect(chromeOutput.linkDecoration).toBe("none");
    // The colour varies per cell, so it stays inline.
    expect(chromeOutput.linkColorInline).not.toBe("");
    expect(chromeOutput.canvasCount).toBe(0);

    // Gridlines beside text that may run into empty cells are settled from
    // the painted text: B11's "دليل الإنتاج" fits its cell, so the lines of
    // the empty C11 and E11 stay, as Excel draws them; the Dashboard's A16
    // runs into B16 only, so A16's edge goes and B16's stays.
    await view.evaluate(`renderSheet(${JSON.stringify(models.rtl)})`);
    const fitLines = await view.evaluate(`[1, 2, 4].map((col) => getComputedStyle(document.querySelector('td[data-aared-row="10"][data-aared-col="' + col + '"]')).borderRightColor)`) as string[];
    expect(fitLines.every((color) => color !== "rgba(0, 0, 0, 0)")).toBe(true);
    // Text sits 1.5px inside the 1px border, Excel's 2.4-2.5px from the cell box.
    expect(await view.evaluate<string>(`getComputedStyle(document.querySelector('td[data-aared-row="10"][data-aared-col="1"]')).paddingLeft`)).toBe("1.5px");
    await view.evaluate(`renderSheet(${JSON.stringify(models.dashboard)})`);
    const runLines = await view.evaluate(`[0, 1].map((col) => {
      const cell = document.querySelector('td[data-aared-row="15"][data-aared-col="' + col + '"]');
      return { text: cell.textContent, line: getComputedStyle(cell).borderRightColor };
    })`) as Array<{ text: string; line: string }>;
    expect(runLines[0]!.text).toStartWith("Busiest editor");
    expect(runLines[0]!.line).toBe("rgba(0, 0, 0, 0)");
    expect(runLines[1]!.line).not.toBe("rgba(0, 0, 0, 0)");

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
      // The total row's status (F9) is centred beside an empty G9 and a filled
      // E9: its text must sit centred on F9 itself, as Excel draws it.
      const status = grid?.querySelector('td[data-aared-row="8"][data-aared-col="5"]');
      const statusText = status?.querySelector('[data-aared-element="cell-text"]');
      const middle = (rect) => rect.left + rect.width / 2;
      return {
        direction: grid?.dataset.aaredSheetDirection,
        dir: grid?.dir,
        gridlines: grid?.dataset.aaredGridlines,
        titleSpan: grid?.querySelector('tr[data-aared-row="0"] td')?.colSpan,
        frozen,
        statusText: statusText?.textContent,
        statusOffset: status && statusText ? middle(statusText.getBoundingClientRect()) - middle(status.getBoundingClientRect()) : null,
      };
    })()`) as { direction?: string; dir?: string; gridlines?: string; titleSpan?: number; frozen: string[]; statusText?: string; statusOffset: number | null };
    expect(summaryOutput.direction).toBe("rtl");
    expect(summaryOutput.dir).toBe("rtl");
    // The sheet turns its gridlines off in Excel, and aared follows the file.
    expect(summaryOutput.gridlines).toBe("false");
    // The title band is merged across A1:G1.
    expect(summaryOutput.titleSpan).toBe(7);
    expect(summaryOutput.frozen.every((flag) => flag === "row")).toBe(true);
    expect(summaryOutput.frozen.length).toBeGreaterThan(0);
    expect(summaryOutput.statusText).toBe("زيادة طفيفة");
    expect(Math.abs(summaryOutput.statusOffset!)).toBeLessThan(1);

    await view.evaluate(`renderSheet(${JSON.stringify(models.dashboard)})`);
    const drawingOutput = await view.evaluate(`(() => ({
      chartCount: document.querySelectorAll('[data-aared-element="chart-graphic"]').length,
      chartPaths: document.querySelectorAll('[data-aared-element="chart-graphic"] path').length,
      chartTexts: [...document.querySelectorAll('[data-aared-element="chart-graphic"] text')].map((text) => text.textContent),
      chartA11y: [...document.querySelectorAll('[data-aared-element="chart-graphic"]')].map((svg) => ({
        role: svg.getAttribute("role"),
        label: svg.getAttribute("aria-label"),
        desc: svg.querySelector("desc")?.textContent,
        described: svg.getAttribute("aria-describedby") === svg.querySelector("desc")?.id,
      })),
      canvasCount: document.querySelectorAll("canvas").length,
    }))()`) as { chartCount: number; chartPaths: number; chartTexts: string[]; chartA11y: Array<{ role: string; label: string; desc: string; described: boolean }>; canvasCount: number };
    expect(drawingOutput.chartCount).toBe(1);
    expect(drawingOutput.chartPaths).toBeGreaterThan(0);
    // Laid out as Excel draws it: an axis of 0.0-45.0 by 5.
    expect(drawingOutput.chartTexts).toEqual(expect.arrayContaining(["Hours logged by episode", "0.0", "45.0"]));
    // An image to a screen reader, named by the chart's title (the file states no alt text) and described by its plot.
    expect(drawingOutput.chartA11y).toEqual([{
      role: "img",
      label: "Hours logged by episode",
      desc: "Column chart. 1 series: Hours logged. 3 categories: Episode 4, Episode 5, Episode 6.",
      described: true,
    }]);
    expect(drawingOutput.canvasCount).toBe(0);

    await view.evaluate(`renderSheet(${JSON.stringify(models.log)})`);
    const barOutput = await view.evaluate(`(() => {
      const bar = document.querySelector('[data-aared-element="data-bar"]');
      const locked = [...document.querySelectorAll('td')].find((cell) => cell.textContent === "Locked");
      return {
        bar: bar ? {
          width: bar.style.width,
          gradient: bar.style.backgroundImage,
          opacity: getComputedStyle(bar).opacity,
        } : null,
        lockedFill: locked ? getComputedStyle(locked).backgroundColor : null,
      };
    })()`) as { bar: { width: string; gradient: string; opacity: string } | null; lockedFill: string | null };
    expect(barOutput.bar).not.toBeNull();
    expect(barOutput.bar!.width).toBe("36%");
    expect(barOutput.bar!.gradient).toContain("linear-gradient(to right");
    expect(barOutput.bar!.gradient).toContain("rgb(46, 125, 116)");
    expect(barOutput.bar!.gradient).toContain("transparent");
    expect(barOutput.bar!.opacity).toBe("1");
    // Cached conditional-formatting results paint from the model snapshot. Excel's x14 default uses a gradient, not 45% opacity.
    expect(barOutput.lockedFill).toBe("rgb(221, 235, 232)");

    const iconPkg = await withPart("xlsx-coverage.xlsx", "xl/worksheets/sheet3.xml", (xml) =>
      xml.replace(/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/, "").replace("</sheetData>",
        `</sheetData><conditionalFormatting sqref="F5:F16"><cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>`));
    const iconModel = await parseXlsxSheet(iconPkg, await readWorkbook(iconPkg), 2);
    await view.evaluate(`renderSheet(${JSON.stringify(iconModel)})`);
    const iconOutput = await view.evaluate(`(() => {
      const icon = document.querySelector('[data-aared-element="cell-icon"]');
      const svg = icon?.querySelector("svg");
      const text = icon?.parentElement?.querySelector('[data-aared-element="cell-text"]');
      return {
        present: icon !== null,
        svg: svg !== null && svg.getAttribute("viewBox") === "0 0 16 16",
        shapes: svg?.querySelectorAll("polygon,circle,rect,path,polyline").length ?? 0,
        textPadded: text ? getComputedStyle(text).paddingInlineStart !== "0px" : false,
      };
    })()`) as { present: boolean; svg: boolean; shapes: number; textPadded: boolean };
    expect(iconOutput.present).toBe(true);
    expect(iconOutput.svg).toBe(true);
    expect(iconOutput.shapes).toBeGreaterThan(0);
    expect(iconOutput.textPadded).toBe(true);

    // Re-read visible paint for the repaired IFS error path and positive
    // missing-blank materialization; the rules run over real fixture cells.
    const paintFormula = async (formula: string, blankCells: string[] = []) => {
      const pkg = await conditionalFormulaFixture({ formula, sqref: "F5:F7", blankCells });
      const sheet = await parseXlsxSheet(pkg, await readWorkbook(pkg), 2);
      await view.evaluate(`renderSheet(${JSON.stringify(sheet)})`);
      return await view.evaluate(`(() => [4, 5, 6].map(row => {
        const cell = document.querySelector('td[data-aared-row="' + row + '"][data-aared-col="5"]');
        return { fill: getComputedStyle(cell).backgroundColor, text: cell.querySelector('[data-aared-element="cell-text"]')?.textContent ?? '' };
      }))()`) as Array<{ fill: string; text: string }>;
    };
    const blankPaint = await paintFormula("ISBLANK(F5)", ["F5", "F7"]);
    expect(blankPaint[0]).toEqual({ fill: "rgb(255, 220, 180)", text: "" });
    expect(blankPaint[2]).toEqual({ fill: "rgb(255, 220, 180)", text: "" });
    expect(blankPaint[1]!.fill).not.toBe("rgb(255, 220, 180)");
    expect(blankPaint[1]!.text.length).toBeGreaterThan(0);
    for (const name of ["SUMIFS", "AVERAGEIFS"]) {
      const formula = `${name}($F$5:$F$7,$E$5:$E$6,"Locked")`;
      const unpainted = await paintFormula(`${formula}>0`);
      expect(unpainted.every((cell) => cell.fill !== "rgb(255, 220, 180)")).toBe(true);
      const errors = await paintFormula(`ISERROR(${formula})`);
      expect(errors.map((cell) => cell.fill)).toEqual(Array(3).fill("rgb(255, 220, 180)"));
      expect(errors.map((cell) => cell.text)).toEqual(unpainted.map((cell) => cell.text));
    }

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

    // A content-sized sheet inside a clipping card (a bench card, a viewer
    // frame) scrolls itself, as Excel's grid does: the card never scrolls,
    // and its own scroll box keeps frozen rows, their row numbers, and the
    // column letters pinned in the visible room instead of letting them
    // scroll away with the card.
    await view.evaluate(`renderInCard(${JSON.stringify(models.deep)})`);
    await view.evaluate(`(() => new Promise((resolve) => {
      const scroller = document.querySelector('#card-host [data-aared-element="sheet-scroll"]');
      scroller.scrollTop = 100000;
      requestAnimationFrame(() => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    }))()`);
    const cardOutput = await view.evaluate(`(() => {
      const card = document.querySelector('#card');
      const scroller = document.querySelector('#card-host [data-aared-element="sheet-scroll"]');
      const frozenRowHeader = document.querySelector('#card-host th[data-aared-element="row-header"][data-aared-row="0"]');
      const colHeader = document.querySelector('#card-host [data-aared-element="column-header"]');
      // The sheet's own visible area starts at the scroller's box, below any
      // notice bar; that is what "pinned" means, not the card's own edge.
      const scrollerTop = scroller.getBoundingClientRect().top;
      return {
        cardScrollTop: card.scrollTop,
        scrollerScrolled: scroller.scrollTop > 0,
        scrollerMaxHeight: scroller.style.maxHeight,
        frozenRowHeaderOffset: frozenRowHeader ? frozenRowHeader.getBoundingClientRect().top - scrollerTop : null,
        colHeaderOffset: colHeader ? colHeader.getBoundingClientRect().top - scrollerTop : null,
      };
    })()`) as { cardScrollTop: number; scrollerScrolled: boolean; scrollerMaxHeight: string; frozenRowHeaderOffset: number | null; colHeaderOffset: number | null };
    // The card holds still; the sheet's own box does the scrolling.
    expect(cardOutput.cardScrollTop).toBe(0);
    expect(cardOutput.scrollerScrolled).toBe(true);
    expect(cardOutput.scrollerMaxHeight).toMatch(/^\d+px$/);
    // Column letters sit right at the sheet's own visible top; the frozen
    // row's number sits just below them, both still in view after scrolling.
    expect(cardOutput.colHeaderOffset).not.toBeNull();
    expect(Math.abs(cardOutput.colHeaderOffset!)).toBeLessThan(2);
    expect(cardOutput.frozenRowHeaderOffset).not.toBeNull();
    expect(cardOutput.frozenRowHeaderOffset!).toBeGreaterThanOrEqual(0);
    expect(cardOutput.frozenRowHeaderOffset!).toBeLessThan(60);
  });
});

visual("visual DOM output: XLSX viewport fill: fills the rendered space in bounded and content-sized scrollers without feedback growth", async () => {
  const models = await sheetModels();
  const module = await browserModule(`${import.meta.dir}/../src/xlsx/index.ts`);
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/xlsx.mjs") {
        return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      // Two hosts: #unbounded sizes itself to its content (no height), the
      // way a plain embedder div lays out; #bounded gives a fixed 1000px.
      return new Response(`<!doctype html><style>
        html,body{margin:0}
        #unbounded{width:900px;border:1px solid #999}
        #bounded{width:900px;height:1000px;border:1px solid #999}
        /* Simulate font metrics that lay rows out taller than their modeled
           heights (Arabic text does this on real machines): with a
           content-sized scroller this used to feed the fill measurement
           back into itself, growing the sheet on every re-render. */
        td[data-aared-element="cell"], td[data-aared-element="fill-cell"],
        tr[data-aared-element="fill-row"] th {
          font-size: 13pt !important; line-height: 2.4 !important;
        }
      </style>
      <div id="unbounded"></div><div id="bounded"></div>
      <script type="module">
        import { render } from "/xlsx.mjs";
        window.renderBoth = async (model) => {
          const handles = await Promise.all([
            render({ sheet: model, target: document.querySelector("#unbounded") }),
            render({ sheet: model, target: document.querySelector("#bounded") }),
          ]);
          return handles.length;
        };
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  await withVisualPage({ server, width: 1280, height: 1400 }, async (view) => {
    await view.evaluate(`renderBoth(${JSON.stringify(models.dashboard)})`);
    // The first windowed paint re-runs inside a rAF once real layout
    // metrics exist; sample only after that settles.
    await view.evaluate(`(() => new Promise((resolve) => {
      requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(resolve, 80)));
    }))()`);

    const sample = `(() => {
      const fillCount = (root) => root.querySelectorAll('[data-aared-element="fill-row"]').length;
      const lastFillNumber = (root) => {
        const headers = [...root.querySelectorAll('[data-aared-element="fill-row-header"]')].map((h) => Number(h.textContent));
        return headers.length > 0 ? headers[headers.length - 1] : 0;
      };
      const unbounded = document.querySelector("#unbounded [data-aared-element='sheet-scroll']");
      const bounded = document.querySelector("#bounded [data-aared-element='sheet-scroll']");
      return {
        unbounded: { fill: fillCount(document.querySelector("#unbounded")), last: lastFillNumber(document.querySelector("#unbounded")), clientH: unbounded?.clientHeight ?? -1, scrollH: unbounded?.scrollHeight ?? -1 },
        bounded: { fill: fillCount(document.querySelector("#bounded")), last: lastFillNumber(document.querySelector("#bounded")), clientH: bounded?.clientHeight ?? -1, scrollH: bounded?.scrollHeight ?? -1 },
      };
    })()`;
    const initial = await view.evaluate(sample) as {
      unbounded: { fill: number; last: number; clientH: number; scrollH: number };
      bounded: { fill: number; last: number; clientH: number; scrollH: number };
    };

    // Bounded: the sheet is shorter than the host, so the viewport fill
    // paints numbered empty rows down to the host bottom. (Rows may lay out
    // a few pixels taller than their tracks under the simulated metrics, so
    // exact top-out is not asserted — stability is.)
    expect(initial.bounded.fill).toBeGreaterThan(0);
    expect(initial.bounded.last).toBeGreaterThan(models.dashboard.extent.rows);
    expect(initial.bounded.scrollH).toBeGreaterThanOrEqual(initial.bounded.clientH);
    // Content-sized: no scroller of its own, so the grid adopts the
    // document viewport's visible room (the 1400px test window) as its
    // height and fills it — numbered rows down toward the window bottom,
    // without anywhere near the old feedback cap.
    expect(initial.unbounded.fill).toBeGreaterThan(0);
    expect(initial.unbounded.last).toBeGreaterThan(models.dashboard.extent.rows);
    expect(initial.unbounded.last).toBeLessThan(models.dashboard.extent.rows + 200);
    expect(initial.unbounded.scrollH).toBeGreaterThanOrEqual(1300);

    // Re-renders (scroll, resize) must not compound: cycle scroll positions
    // and return to the start — the same scroll state must paint the same
    // sheet. Under the feedback bug each pass grew the fill (to its
    // 500-track cap), so identical scroll state ⇒ identical height is the
    // regression pin. Drift between *different* scroll positions is expected
    // here (windowed spacers are model-sized while the simulated rows are
    // not) and is not a defect.
    for (let step = 0; step < 4; step++) {
      await view.evaluate(`(() => {
        for (const scroller of document.querySelectorAll('[data-aared-element="sheet-scroll"]')) {
          scroller.scrollTop = Math.min(scroller.scrollTop + 200, scroller.scrollHeight);
          scroller.dispatchEvent(new Event("scroll"));
        }
        window.dispatchEvent(new Event("resize"));
        return true;
      })()`);
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
    await view.evaluate(`(() => new Promise((resolve) => {
      for (const scroller of document.querySelectorAll('[data-aared-element="sheet-scroll"]')) {
        scroller.scrollTop = 0;
        scroller.dispatchEvent(new Event("scroll"));
      }
      requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(resolve, 80)));
    }))()`);
    const after = await view.evaluate(sample) as typeof initial;
    expect(after.unbounded.fill).toBe(initial.unbounded.fill);
    expect(after.unbounded.last).toBe(initial.unbounded.last);
    expect(after.unbounded.scrollH).toBe(initial.unbounded.scrollH);
    expect(after.bounded.fill).toBe(initial.bounded.fill);
    expect(after.bounded.last).toBe(initial.bounded.last);
    expect(after.bounded.scrollH).toBe(initial.bounded.scrollH);
  });
});

visual("visual DOM output: XLSX zoom: a sheet opens at 100%, and fit-width / fit-page really fit its used area", async () => {
  const models = await sheetModels();
  const module = await browserModule(`${import.meta.dir}/../src/xlsx/index.ts`);
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/xlsx.mjs") {
        return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      return new Response(`<!doctype html><style>
        html,body{margin:0}
        #bounded{width:900px;height:500px}
        #short{width:900px;height:450px}
        #tiny{width:900px;height:200px}
        #unbounded{width:900px}
      </style>
      <div id="bounded"></div><div id="short"></div><div id="tiny"></div><div id="unbounded"></div>
      <script type="module">
        import { render } from "/xlsx.mjs";
        const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(resolve, 30))));
        const handles = {};
        window.show = async (id, model, zoom, appearance) => {
          handles[id]?.destroy();
          const input = { sheet: model, target: document.getElementById(id) };
          if (zoom !== undefined) input.zoom = zoom;
          if (appearance !== undefined) input.appearance = appearance;
          handles[id] = await render(input);
          await frame();
          return window.measure(id);
        };
        window.act = async (id, action) => {
          const handle = handles[id];
          if (action.zoom !== undefined) handle.setZoom(action.zoom);
          if (action.appearance !== undefined) handle.setAppearance(action.appearance);
          if (action.width !== undefined) document.getElementById(id).style.width = action.width;
          await frame();
          return window.measure(id);
        };
        window.measure = (id) => {
          const root = document.getElementById(id);
          const scroller = root.querySelector('[data-aared-element="sheet-scroll"]');
          return {
            zoom: handles[id].zoom,
            scale: handles[id].scale,
            hostZoom: root.querySelector('[data-aared-element="sheet"]').style.zoom,
            overflowX: scroller.scrollWidth - scroller.clientWidth,
            overflowY: scroller.scrollHeight - scroller.clientHeight,
            fillColumns: root.querySelectorAll('[data-aared-element="fill-column-header"]').length,
            rows: root.querySelectorAll('tr[data-aared-row]').length,
            headers: root.querySelector('[data-aared-element="corner"]') !== null,
          };
        };
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  interface Measure {
    zoom: unknown; scale: number; hostZoom: string; overflowX: number; overflowY: number;
    fillColumns: number; rows: number; headers: boolean;
  }
  await withVisualPage({ server, width: 1280, height: 1200 }, async (view) => {
    const log = JSON.stringify(models.log);
    const dashboard = JSON.stringify(models.dashboard);

    // 100% by default, like Excel: no zoom on the host, and a sheet wider
    // than its host scrolls sideways.
    const opened = await view.evaluate(`show("bounded", ${log})`) as Measure;
    expect(opened.zoom).toBe(1);
    expect(opened.scale).toBe(1);
    expect(opened.hostZoom).toBe("");
    expect(opened.overflowX).toBeGreaterThan(0);

    // fit-width: the used columns fill the width exactly — no sideways
    // scroll, and no viewport-fill columns (the fit leaves no room for them).
    const fitted = await view.evaluate(`act("bounded", { zoom: "fit-width" })`) as Measure;
    expect(fitted.zoom).toBe("fit-width");
    expect(fitted.scale).toBeLessThan(1);
    expect(fitted.overflowX).toBeLessThanOrEqual(1);
    expect(fitted.fillColumns).toBe(0);

    // The fit follows the room: a narrower host re-fits smaller.
    const narrower = await view.evaluate(`act("bounded", { width: "600px" })`) as Measure;
    expect(narrower.scale).toBeLessThan(fitted.scale);
    expect(narrower.overflowX).toBeLessThanOrEqual(1);

    // A headers toggle changes the used width, so the fit is recomputed.
    const bare = await view.evaluate(`act("bounded", { appearance: { spreadsheet: { headers: false } } })`) as Measure;
    expect(bare.headers).toBe(false);
    expect(bare.scale).toBeGreaterThan(narrower.scale);
    expect(bare.overflowX).toBeLessThanOrEqual(1);
    // setAppearance changes only what it sets: a theme switch keeps the
    // headers hidden, and passing the default brings them back.
    const themed = await view.evaluate(`act("bounded", { appearance: { theme: "light" } })`) as Measure;
    expect(themed.headers).toBe(false);
    const restored = await view.evaluate(`act("bounded", { appearance: { spreadsheet: { headers: true } } })`) as Measure;
    expect(restored.headers).toBe(true);

    // fit-page keeps the whole used area in view in a short host — the
    // painted rows included, which outgrow the model below 100% as borders
    // snap to device pixels.
    const page = await view.evaluate(`show("short", ${dashboard}, "fit-page")`) as Measure;
    expect(page.scale).toBeGreaterThan(0.25);
    expect(page.overflowX).toBeLessThanOrEqual(1);
    expect(page.overflowY).toBeLessThanOrEqual(1);
    const pageWidth = await view.evaluate(`act("short", { zoom: "fit-width" })`) as Measure;
    expect(page.scale).toBeLessThan(pageWidth.scale);

    // A host with no height of its own offers none: fit-page = fit-width.
    const unboundedPage = await view.evaluate(`show("unbounded", ${dashboard}, "fit-page")`) as Measure;
    const unboundedWidth = await view.evaluate(`act("unbounded", { zoom: "fit-width" })`) as Measure;
    expect(unboundedPage.scale).toBeCloseTo(unboundedWidth.scale, 3);

    // A deep sheet's fit-page stops at 25%: it scrolls rather than shrinking
    // to nothing, and the row window stays bounded.
    const floored = await view.evaluate(`show("tiny", ${JSON.stringify(models.deep)}, "fit-page")`) as Measure;
    expect(floored.scale).toBe(0.25);
    expect(floored.rows).toBeLessThan(models.deep.extent.rows);

    // RTL sheets fit the same way.
    const rtl = await view.evaluate(`show("bounded", ${JSON.stringify(models.rtl)}, "fit-width")`) as Measure;
    expect(rtl.overflowX).toBeLessThanOrEqual(1);
  });
}, 30000);

/**
 * XLSX's render handle names fonts that painted with the browser's generic
 * fallback, as DOCX and PPTX do: its `warnings` were always empty.
 */
visual("visual DOM output: XlsxRenderHandle.warnings names fonts that painted with the generic fallback", async () => {
  const { rtl } = await sheetModels();
  const NONEXISTENT_FAMILY = "Definitely Not A Real Font Zzqx";
  const renamed = structuredClone(rtl);
  for (const row of renamed.rows) {
    for (const entry of row.cells) {
      if (entry.cell.style !== undefined) entry.cell.style = { ...entry.cell.style, fontFamily: NONEXISTENT_FAMILY };
    }
  }
  const module = await browserModule(`${import.meta.dir}/../src/xlsx/index.ts`);
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      if (new URL(request.url).pathname === "/xlsx.mjs") {
        return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      return new Response(`<!doctype html><main id="host" style="width:900px;height:400px"></main><script type="module">
        import { render } from "/xlsx.mjs";
        window.warningsFor = async (model) => {
          const handle = await render({ sheet: model, target: document.querySelector("#host") });
          const warnings = handle.warnings.map((warning) => ({ code: warning.code, message: warning.message }));
          handle.destroy();
          return warnings;
        };
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  await withVisualPage({ server, width: 1000, height: 600 }, async (view) => {
    const real = (await view.evaluate(`warningsFor(${JSON.stringify(rtl)})`)) as Array<{ code: string }>;
    expect(real.filter((warning) => warning.code === "font-unavailable")).toEqual([]);
    const missing = (await view.evaluate(`warningsFor(${JSON.stringify(renamed)})`)) as Array<{ code: string; message: string }>;
    const unavailable = missing.filter((warning) => warning.code === "font-unavailable");
    expect(unavailable).toHaveLength(1);
    expect(unavailable[0]!.message).toContain(NONEXISTENT_FAMILY);
    const substituted = await view.evaluate(`(async () => {
      const face = new FontFace(${JSON.stringify(fontAlias(NONEXISTENT_FAMILY))}, "local('Courier New')");
      document.fonts.add(face); await face.load();
      return warningsFor(${JSON.stringify(renamed)});
    })()`) as Array<{ code: string }>;
    expect(substituted.filter((warning) => warning.code === "font-unavailable")).toEqual([]);
  });
}, 30000);

/**
 * A sheet whose chart the file gave no id renders with its chart and says so once
 * (`xlsx-missing-id`); the same sheet with its id has no such warning.
 */
visual("visual DOM output: XlsxRenderHandle.warnings reports drawings the file gave no id, once", async () => {
  const TRACKER = "post-production-tracker.xlsx";
  const whole = await OpcPackage.open(await Bun.file(`${import.meta.dir}/../../../fixtures/${TRACKER}`).arrayBuffer());
  const clean = await parseXlsxSheet(whole, await readWorkbook(whole), 0);
  const stripped = await withPart(TRACKER, "xl/drawings/drawing1.xml", (xml) => xml.replace(/(<xdr:cNvPr) id="[^"]*"/g, "$1"));
  const damaged = await parseXlsxSheet(stripped, await readWorkbook(stripped), 0);
  const module = await browserModule(`${import.meta.dir}/../src/xlsx/index.ts`);
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      if (new URL(request.url).pathname === "/xlsx.mjs") {
        return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      return new Response(`<!doctype html><main id="host" style="width:900px;height:700px"></main><script type="module">
        import { render } from "/xlsx.mjs";
        window.renderSheet = async (model) => {
          const handle = await render({ sheet: model, target: document.querySelector("#host") });
          const out = {
            warnings: handle.warnings.map((warning) => ({ code: warning.code, message: warning.message })),
            charts: document.querySelector("#host").querySelectorAll('[data-aared-element="chart-geometry"], svg[data-aared-chart-primitives]').length,
          };
          handle.destroy();
          return out;
        };
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  await withVisualPage({ server, width: 1000, height: 800 }, async (view) => {
    type Rendered = { warnings: Array<{ code: string; message: string }>; charts: number };
    const intact = (await view.evaluate(`renderSheet(${JSON.stringify(clean)})`)) as Rendered;
    expect(intact.warnings.filter((warning) => warning.code === "xlsx-missing-id")).toEqual([]);
    const repaired = (await view.evaluate(`renderSheet(${JSON.stringify(damaged)})`)) as Rendered;
    const warnings = repaired.warnings.filter((warning) => warning.code === "xlsx-missing-id");
    expect(warnings).toHaveLength(1);
    expect(warnings[0]!.message).toContain(`1 drawing on sheet "${damaged.name}" has no id`);
    expect(warnings[0]!.message).toMatch(/^[\x20-\x7e]+$/);
    // Drawn in full: the chart is on the page as for the sheet with its id.
    expect(intact.charts).toBeGreaterThan(0);
    expect(repaired.charts).toBe(intact.charts);
  });
}, 30000);
