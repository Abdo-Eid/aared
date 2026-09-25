import { expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { readWorkbook } from "../src/core/xlsx/workbook.js";
import { parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { findCoverageSheet } from "./xlsx-fixture.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
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

visual("viewport fill: fills the rendered space in bounded and content-sized scrollers without feedback growth", async () => {
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
  try {
    await using view = new Bun.WebView({ width: 1280, height: 1400 });
    await view.navigate(server.url.href);
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
  } finally {
    server.stop();
  }
});

visual("zoom: a sheet opens at 100%, and fit-width / fit-page really fit its used area", async () => {
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
  try {
    await using view = new Bun.WebView({ width: 1280, height: 1200 });
    await view.navigate(server.url.href);
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
    // setAppearance replaces, never merges: knobs left out return to their defaults.
    const replaced = await view.evaluate(`act("bounded", { appearance: { theme: "light" } })`) as Measure;
    expect(replaced.headers).toBe(true);

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
  } finally {
    server.stop();
  }
}, 30000);
