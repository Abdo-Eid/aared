import { expect, test } from "bun:test";
import { browserModule } from "./harness.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;
const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

visual("visual preview output: auto-detect renders PPTX, XLSX, and DOCX into one handle", async () => {
  const [previewModule, workerModule, pptx, xlsx, docx] = await Promise.all([
    browserModule(`${import.meta.dir}/../src/preview/index.ts`),
    browserModule(`${import.meta.dir}/../src/core/worker.ts`),
    Bun.file(PPTX_FIXTURE).arrayBuffer(),
    Bun.file(XLSX_FIXTURE).arrayBuffer(),
    Bun.file(DOCX_FIXTURE).arrayBuffer(),
  ]);
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/preview.mjs") {
        return new Response(previewModule, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      if (url.pathname === "/worker.mjs") {
        return new Response(workerModule, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      if (url.pathname === "/arabic.pptx") return new Response(pptx);
      if (url.pathname === "/arabic.xlsx") return new Response(xlsx);
      if (url.pathname === "/arabic.docx") return new Response(docx);
      return new Response(`<!doctype html><body style="margin:0"><main id="host" style="height:600px"></main><script type="module">
        import { preview } from "/preview.mjs";
        window.events = [];
        window.take = () => window.events.splice(0);
        window.frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        window.openPreview = async (source, options) => {
          const host = document.querySelector("#host");
          const handle = await preview({ ...options, source, target: host });
          window.handle = handle;
          window.events = [];
          for (const type of ["indexchange", "slidechange", "zoomchange", "sheetchange"]) {
            handle.addEventListener(type, (event) => window.events.push({ type, detail: event.detail }));
          }
          handle.addEventListener("error", (event) => window.events.push({
            type: "error",
            detail: { message: event.detail.message, code: event.detail.cause.code, same: event.detail.message === event.detail.cause.message },
          }));
          return { format: handle.format, count: handle.info.count, index: handle.index, titles: handle.info.titles, zoom: handle.zoom };
        };
        window.slideScale = () => {
          const host = document.querySelector("#host");
          const slide = host.querySelector('[data-aared-element="slide"]');
          const scene = host.querySelector('[data-aared-element="slide-scene"]');
          return slide.getBoundingClientRect().width / parseFloat(scene.style.width);
        };
        window.destroyPreview = () => {
          window.handle.destroy();
          return { children: document.querySelector("#host").childElementCount };
        };
        window.snapshot = () => {
          const host = document.querySelector("#host");
          return {
            text: host.textContent,
            canvasCount: host.querySelectorAll("canvas").length,
            cells: host.querySelectorAll('td[data-aared-element="cell"]').length,
            pages: host.querySelectorAll('section[data-aared-element="page"]').length,
          };
        };
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1280, height: 720 });
    await view.navigate(server.url.href);

    // PPTX: detected, first slide painted; the starting state is the resolved handle.
    const opened = await view.evaluate(`openPreview("/arabic.pptx")`) as {
      format: string; count: number; index: number; titles: string[]; zoom: unknown;
    };
    expect(opened.format).toBe("pptx");
    expect(opened.count).toBe(18);
    expect(opened.index).toBe(0);
    expect(opened.zoom).toBe("fit-page");
    const first = await view.evaluate(`snapshot()`) as { text?: string; canvasCount: number };
    expect(first.canvasCount).toBe(0);
    expect(first.text).toContain("تقرير الإنتاج الأسبوعي");

    // Navigation announces once, with the unit reference; going nowhere is silent.
    const moved = await view.evaluate(`(async () => {
      await handle.goTo(1);
      const events = take();
      await handle.goTo(1);
      return { events, repeated: take(), text: snapshot().text };
    })()`) as { events: Array<{ type: string; detail: unknown }>; repeated: unknown[]; text: string };
    expect(moved.text).toContain("الاسم");
    expect(moved.events).toEqual([
      { type: "indexchange", detail: { index: 1 } },
      { type: "slidechange", detail: { index: 1, title: "Slide 2" } },
    ]);
    expect(moved.repeated).toEqual([]);

    // Zoom and appearance belong to the facade: they survive navigation, and
    // the zoom is announced once — never again for the unit change.
    const carried = await view.evaluate(`(async () => {
      handle.setZoom(1.5);
      handle.setZoom(1.5);
      await frame();
      const zoomed = { scale: slideScale(), events: take() };
      handle.setAppearance({ theme: "dark" });
      await handle.goTo(2);
      await frame();
      return {
        zoomed,
        after: { zoom: handle.zoom, scale: slideScale(), events: take().map((event) => event.type) },
        theme: document.querySelector('[data-aared-element="slide"]').dataset.aaredTheme,
        controls: { xlsx: handle.xlsx, docx: handle.docx },
      };
    })()`) as {
      zoomed: { scale: number; events: unknown[] };
      after: { zoom: unknown; scale: number; events: string[] };
      theme: string;
      controls: { xlsx: unknown; docx: unknown };
    };
    expect(carried.zoomed.scale).toBeCloseTo(1.5, 2);
    expect(carried.zoomed.events).toEqual([{ type: "zoomchange", detail: { zoom: 1.5 } }]);
    expect(carried.after.zoom).toBe(1.5);
    expect(carried.after.scale).toBeCloseTo(1.5, 2);
    expect(carried.after.events).toEqual(["indexchange", "slidechange"]);
    expect(carried.theme).toBe("dark");
    expect(carried.controls).toEqual({ xlsx: null, docx: null });

    // A unit that cannot be shown rejects and also fires `error` once, with the
    // element's `{ message, cause }` shape; the unit on screen stays.
    const failed = await view.evaluate(`(async () => {
      take();
      const rejected = await handle.goTo(99).then(() => null, (error) => error.code);
      return { rejected, index: handle.index, events: take() };
    })()`) as { rejected: unknown; index: number; events: Array<{ type: string; detail: { message: string; code: string; same: boolean } }> };
    expect(failed.rejected).toBe("not-found");
    expect(failed.index).toBe(2);
    expect(failed.events).toHaveLength(1);
    expect(failed.events[0]).toMatchObject({ type: "error", detail: { code: "not-found", same: true } });
    expect(failed.events[0]!.detail.message).toContain("99");
    // Not an index at all: the call is wrong, not the document.
    const notAnIndex = await view.evaluate(`handle.goTo(1.5).then(() => null, (error) => error.code)`);
    expect(notAnIndex).toBe("invalid-argument");
    await view.evaluate(`take()`);

    // Two goTo() calls in one tick: the second overtakes the first, which then
    // changes nothing; it used to set `index` to its own unit and announce it.
    const raced = await view.evaluate(`(async () => {
      take();
      await Promise.all([handle.goTo(3), handle.goTo(4)]);
      return { index: handle.index, events: take() };
    })()`) as { index: number; events: Array<{ type: string; detail: unknown }> };
    expect(raced.index).toBe(4);
    expect(raced.events).toEqual([
      { type: "indexchange", detail: { index: 4 } },
      { type: "slidechange", detail: { index: 4, title: "Slide 5" } },
    ]);

    // Back to the unit on screen while a move is on its way: the newest call
    // wins, so the move is overtaken and the slide on screen stays (it used
    // to land on the slide the older call asked for).
    const back = await view.evaluate(`(async () => {
      take();
      await frame();
      const before = snapshot().text;
      await Promise.all([handle.goTo(1), handle.goTo(4)]);
      await frame();
      return { index: handle.index, events: take(), same: snapshot().text === before };
    })()`) as { index: number; events: unknown[]; same: boolean };
    expect(back).toEqual({ index: 4, events: [], same: true });

    // The toolbar methods. next()/previous() step on from where the newest
    // move is heading, so quick presses add up; at either end nothing changes.
    // zoomIn()/zoomOut() step from the scale on screen in a fit mode, round to
    // whole percents and stay within 25%–300%; resetZoom() is the format's own.
    const toolbar = await view.evaluate(`(async () => {
      take();
      await Promise.all([handle.next(), handle.next(), handle.next()]);
      const added = { index: handle.index, events: take().map((event) => event.type) };
      await handle.previous();
      const stepBack = handle.index;
      await handle.goTo(0);
      take();
      await handle.previous();
      const atStart = { index: handle.index, events: take() };
      await handle.goTo(17);
      take();
      await handle.next();
      const atEnd = { index: handle.index, events: take() };
      handle.resetZoom();
      await frame();
      const fitted = { zoom: handle.zoom, defaultZoom: handle.defaultZoom, scale: handle.scale };
      handle.zoomIn();
      const stepped = handle.zoom;
      handle.zoomOut();
      const steppedBack = handle.zoom;
      handle.zoomIn(0.5);
      const bigStep = handle.zoom;
      for (let press = 0; press < 30; press++) handle.zoomIn();
      const capped = handle.zoom;
      for (let press = 0; press < 40; press++) handle.zoomOut();
      const floored = handle.zoom;
      let badStep = null;
      try { handle.zoomIn(0); } catch (error) { badStep = error.code; }
      take();
      handle.resetZoom();
      return { added, stepBack, atStart, atEnd, fitted, stepped, steppedBack, bigStep, capped, floored, badStep, reset: { zoom: handle.zoom, events: take() } };
    })()`) as {
      added: { index: number; events: string[] };
      stepBack: number;
      atStart: { index: number; events: unknown[] };
      atEnd: { index: number; events: unknown[] };
      fitted: { zoom: unknown; defaultZoom: unknown; scale: number };
      stepped: number;
      steppedBack: number;
      bigStep: number;
      capped: number;
      floored: number;
      badStep: unknown;
      reset: { zoom: unknown; events: unknown[] };
    };
    const percent = (value: number): number => Math.round(value * 100) / 100;
    // Three quick presses from slide 5: three slides on, announced once.
    expect(toolbar.added).toEqual({ index: 7, events: ["indexchange", "slidechange"] });
    expect(toolbar.stepBack).toBe(6);
    expect(toolbar.atStart).toEqual({ index: 0, events: [] });
    expect(toolbar.atEnd).toEqual({ index: 17, events: [] });
    expect(toolbar.fitted.zoom).toBe("fit-page");
    expect(toolbar.fitted.defaultZoom).toBe("fit-page");
    expect(toolbar.fitted.scale).toBeLessThan(1);
    // The default step adds 10 points to the fitted scale; the same step back takes them away.
    expect(toolbar.stepped).toBe(percent(toolbar.fitted.scale + 0.1));
    expect(toolbar.steppedBack).toBe(percent(toolbar.stepped - 0.1));
    expect(toolbar.bigStep).toBe(percent(toolbar.steppedBack + 0.5));
    expect(toolbar.capped).toBe(3);
    expect(toolbar.floored).toBe(0.25);
    expect(toolbar.badStep).toBe("invalid-argument");
    expect(toolbar.reset).toEqual({ zoom: "fit-page", events: [{ type: "zoomchange", detail: { zoom: "fit-page" } }] });

    // search(): the deck the viewer already has open, one result per occurrence,
    // each a place to go with goTo(); an empty query finds nothing.
    const searched = await view.evaluate(`(async () => {
      take();
      const matches = await handle.search("الاسم");
      const onSlide2 = matches.find((match) => match.slide.index === 1);
      if (onSlide2 !== undefined) await handle.goTo(onSlide2.slide.index);
      return {
        slides: matches.map((match) => match.slide.index),
        index: handle.index,
        text: snapshot().text,
        empty: await handle.search(""),
      };
    })()`) as { slides: number[]; index: number; text: string; empty: unknown[] };
    expect(searched.slides).toContain(1);
    expect(searched.index).toBe(1);
    expect(searched.text).toContain("الاسم");
    expect(searched.empty).toEqual([]);

    const cleared = await view.evaluate(`destroyPreview()`) as { children: number };
    expect(cleared.children).toBe(0);
    // A destroyed viewer says so instead of resolving as if it moved.
    const afterDestroy = await view.evaluate(`handle.goTo(1).then(() => "resolved", (error) => ({ code: error.code, message: error.message }))`);
    expect(afterDestroy).toEqual({ code: "closed", message: "aared/preview: goTo() was called after destroy()" });
    const nextAfterDestroy = await view.evaluate(`handle.next().then(() => "resolved", (error) => ({ code: error.code, message: error.message }))`);
    expect(nextAfterDestroy).toEqual({ code: "closed", message: "aared/preview: next() was called after destroy()" });
    const searchAfterDestroy = await view.evaluate(`handle.search("x").then(() => "resolved", (error) => ({ code: error.code, message: error.message }))`);
    expect(searchAfterDestroy).toEqual({ code: "closed", message: "aared/preview: search() was called after destroy()" });

    // XLSX: detected at 100%, sheet tabs shown by default (an embed has no
    // other way to switch sheets), places within the sheet on `xlsx`, tab
    // clicks followed.
    const workbook = await view.evaluate(`openPreview("/arabic.xlsx")`) as { format: string; count: number; titles: string[]; zoom: unknown };
    expect(workbook.format).toBe("xlsx");
    expect(workbook.count).toBe(workbook.titles.length);
    expect(workbook.zoom).toBe(1);
    const sheet = await view.evaluate(`(async () => {
      const tabCount = () => document.querySelectorAll('[data-aared-element="sheet-tab"]').length;
      const tabs = { start: tabCount() };
      // setAppearance() replaces the appearance; what it leaves out takes the
      // embed default, so a theme switch keeps the tabs.
      handle.setAppearance({ theme: "dark" });
      await frame();
      tabs.afterTheme = tabCount();
      handle.setAppearance({});
      const controls = handle.xlsx;
      const stable = controls === handle.xlsx;
      await controls.goToCell("B2");
      const selection = controls.selection;
      const tab = document.querySelectorAll('[data-aared-element="sheet-tab"]')[1];
      tab.click();
      await new Promise((resolve) => handle.addEventListener("indexchange", resolve, { once: true }));
      const events = take();
      // The same object keeps working after goTo() renders another sheet (it used to act on the old one).
      await handle.goTo(0);
      await controls.goToCell("C3");
      take();
      return {
        tabs,
        stable,
        selection,
        index: handle.index,
        events,
        kept: { same: controls === handle.xlsx, selection: controls.selection },
        cells: snapshot().cells,
        canvasCount: snapshot().canvasCount,
      };
    })()`) as {
      tabs: { start: number; afterTheme: number };
      stable: boolean;
      selection: unknown;
      index: number;
      events: Array<{ type: string; detail: unknown }>;
      kept: { same: boolean; selection: unknown };
      cells: number;
      canvasCount: number;
    };
    expect(sheet.tabs.start).toBeGreaterThan(1);
    expect(sheet.tabs.afterTheme).toBe(sheet.tabs.start);
    expect(sheet.stable).toBe(true);
    expect(sheet.selection).toEqual({ address: "B2", row: 1, col: 1 });
    expect(sheet.index).toBe(0);
    expect(sheet.events).toEqual([
      { type: "indexchange", detail: { index: 1 } },
      { type: "sheetchange", detail: { index: 1, title: workbook.titles[1] } },
    ]);
    expect(sheet.kept).toEqual({ same: true, selection: { address: "C3", row: 2, col: 2 } });
    expect(sheet.cells).toBeGreaterThan(0);
    expect(sheet.canvasCount).toBe(0);
    // search() on a workbook searches the sheet on screen, as Excel's Find
    // does: each match is a cell of that sheet, reached with goToCell(address).
    // The query is a painted cell's own text, so it is known to be on it.
    const found = await view.evaluate(`(async () => {
      const query = [...document.querySelectorAll('td[data-aared-element="cell"]')]
        .map((cell) => cell.textContent.trim())
        .find((text) => text.length >= 3);
      const matches = await handle.search(query);
      const [match] = matches;
      if (match === undefined) return { found: false, address: null, selected: null, onScreen: false, elsewhere: null };
      const onScreen = matches.every((m) => m.sheet.index === handle.index);
      await handle.xlsx.goToCell(match.address);
      take();
      const selected = handle.xlsx.selection?.address ?? null;
      // On another sheet the same query is that sheet's own search, not sheet 0's.
      await handle.goTo(1);
      const elsewhere = (await handle.search(query)).every((m) => m.sheet.index === 1);
      await handle.goTo(0);
      take();
      return { found: true, address: match.address, selected, onScreen, elsewhere };
    })()`) as { found: boolean; address: string | null; selected: string | null; onScreen: boolean; elsewhere: boolean | null };
    expect(found.found).toBe(true);
    expect(found.selected).toBe(found.address);
    expect(found.onScreen).toBe(true);
    expect(found.elsewhere).toBe(true);
    // After destroy() a move says so, like goTo(); the rest do nothing.
    const closedMove = await view.evaluate(`(async () => {
      const controls = handle.xlsx;
      destroyPreview();
      controls.focus("A1");
      return { code: await controls.goToCell("A1").then(() => "resolved", (error) => error.code), selection: controls.selection };
    })()`);
    expect(closedMove).toEqual({ code: "closed", selection: null });
    // `sheetTabs: false` turns the embed default off.
    const untabbed = await view.evaluate(`(async () => {
      await openPreview("/arabic.xlsx", { appearance: { spreadsheet: { sheetTabs: false } } });
      await frame();
      const tabs = document.querySelectorAll('[data-aared-element="sheet-tab"]').length;
      destroyPreview();
      return tabs;
    })()`);
    expect(untabbed).toBe(0);

    // DOCX: detected; layout pagination paints pages; places on `docx`.
    const document = await view.evaluate(
      `openPreview("/arabic.docx", { pagination: { mode: "layout" } })`,
    ) as { format: string; count: number; titles: string[] };
    expect(document.format).toBe("docx");
    expect(document.count).toBe(1);
    expect(document.titles).toHaveLength(1);
    const paged = await view.evaluate(`(async () => {
      const controls = handle.docx;
      await controls.goToPage(controls.pageCount - 1);
      await controls.goToHeading(0);
      return {
        stable: controls === handle.docx,
        pageCount: controls.pageCount,
        firstPage: controls.pages[0],
        headings: controls.headings.map((heading) => heading.title),
        ...snapshot(),
      };
    })()`) as {
      stable: boolean;
      pageCount: number;
      firstPage: { index: number };
      headings: string[];
      text?: string;
      pages: number;
      canvasCount: number;
    };
    expect(paged.stable).toBe(true);
    expect(paged.pageCount).toBe(paged.pages);
    expect(paged.firstPage.index).toBe(0);
    expect(paged.headings[0]).toContain("Project Overview");
    expect(paged.text).toContain("1. Project Overview");
    expect(paged.canvasCount).toBe(0);

    // Moving inside the unit fails like goTo(): it rejects and fires `error` too.
    const inside = await view.evaluate(`(async () => {
      take();
      const codes = [];
      for (const move of [() => handle.docx.goToHeading("No such heading"), () => handle.docx.goToPage(1.5)]) {
        codes.push(await move().then(() => null, (error) => error.code));
      }
      return { codes, events: take().map((event) => ({ type: event.type, code: event.detail.code })) };
    })()`) as { codes: unknown[]; events: unknown[] };
    expect(inside.codes).toEqual(["not-found", "invalid-argument"]);
    expect(inside.events).toEqual([
      { type: "error", code: "not-found" },
      { type: "error", code: "invalid-argument" },
    ]);
    // search() on a document finds the heading; a failed search only
    // rejects: it fires no `error`, which reports moves.
    const docSearch = await view.evaluate(`(async () => {
      take();
      const matches = await handle.search("Project Overview");
      const cancelled = await handle.search("Project", { signal: AbortSignal.abort() }).then(() => "resolved", (error) => error.code);
      return {
        heading: matches.some((match) => match.kind === "heading" && match.text.includes("Project Overview")),
        cancelled,
        events: take(),
      };
    })()`) as { heading: boolean; cancelled: unknown; events: unknown[] };
    expect(docSearch).toEqual({ heading: true, cancelled: "aborted", events: [] });
    await view.evaluate(`destroyPreview()`);
  } finally {
    server.stop();
  }
});
