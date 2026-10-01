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
          const handle = await preview(source, host, options);
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
    expect(opened.count).toBe(12);
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

    const cleared = await view.evaluate(`destroyPreview()`) as { children: number };
    expect(cleared.children).toBe(0);
    // A destroyed viewer says so instead of resolving as if it moved.
    const afterDestroy = await view.evaluate(`handle.goTo(1).then(() => "resolved", (error) => ({ code: error.code, message: error.message }))`);
    expect(afterDestroy).toEqual({ code: "closed", message: "aared/preview: goTo() was called after destroy()" });

    // XLSX: detected at 100%, places within the sheet on `xlsx`, tab clicks followed.
    const workbook = await view.evaluate(
      `openPreview("/arabic.xlsx", { appearance: { spreadsheet: { sheetTabs: true } } })`,
    ) as { format: string; count: number; titles: string[]; zoom: unknown };
    expect(workbook.format).toBe("xlsx");
    expect(workbook.count).toBe(workbook.titles.length);
    expect(workbook.zoom).toBe(1);
    const sheet = await view.evaluate(`(async () => {
      const controls = handle.xlsx;
      const stable = controls === handle.xlsx;
      await controls.goToCell("B2");
      const selection = controls.selection;
      const tab = document.querySelectorAll('[data-aared-element="sheet-tab"]')[1];
      tab.click();
      await new Promise((resolve) => handle.addEventListener("indexchange", resolve, { once: true }));
      return { stable, selection, index: handle.index, events: take(), cells: snapshot().cells, canvasCount: snapshot().canvasCount };
    })()`) as {
      stable: boolean;
      selection: unknown;
      index: number;
      events: Array<{ type: string; detail: unknown }>;
      cells: number;
      canvasCount: number;
    };
    expect(sheet.stable).toBe(true);
    expect(sheet.selection).toEqual({ address: "B2", row: 1, col: 1 });
    expect(sheet.index).toBe(1);
    expect(sheet.events).toEqual([
      { type: "indexchange", detail: { index: 1 } },
      { type: "sheetchange", detail: { index: 1, title: workbook.titles[1] } },
    ]);
    expect(sheet.cells).toBeGreaterThan(0);
    expect(sheet.canvasCount).toBe(0);
    await view.evaluate(`destroyPreview()`);

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
    await view.evaluate(`destroyPreview()`);
  } finally {
    server.stop();
  }
});
