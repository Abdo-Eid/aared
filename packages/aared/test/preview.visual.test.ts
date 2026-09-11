import { expect, test } from "bun:test";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-pptx-rtl-coverage.pptx`;
const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-xlsx-rtl-coverage.xlsx`;
const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-docx-bidi-coverage.docx`;

async function browserModule(entrypoint: string): Promise<string> {
  const build = await Bun.build({ entrypoints: [entrypoint], target: "browser", format: "esm" });
  if (!build.success) throw new Error(build.logs.map((log) => log.message).join("\n"));
  return build.outputs[0]!.text();
}

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
      return new Response(`<!doctype html><main id="host"></main><script type="module">
        import { preview } from "/preview.mjs";
        window.events = [];
        window.openPreview = async (source, options) => {
          const host = document.querySelector("#host");
          const handle = await preview(source, host, options);
          window.handle = handle;
          for (const type of ["indexchange", "slidechange", "zoomchange", "sheetchange"]) {
            handle.addEventListener(type, (event) => window.events.push({ type, detail: event.detail }));
          }
          return { format: handle.format, count: handle.info.count, index: handle.index, titles: handle.info.titles };
        };
        window.goTo = async (index) => {
          await window.handle.goTo(index);
          return { index: window.handle.index };
        };
        window.setZoom = (zoom) => {
          window.handle.setZoom(zoom);
          return { zoom: window.handle.zoom };
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
            events: window.events,
          };
        };
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1280, height: 720 });
    await view.navigate(server.url.href);

    // PPTX: detected, first slide painted, navigation + zoom + teardown work.
    const opened = await view.evaluate(`openPreview("/arabic.pptx")`) as {
      format: string; count: number; index: number; titles: string[];
    };
    expect(opened.format).toBe("pptx");
    expect(opened.count).toBe(4);
    expect(opened.index).toBe(0);
    const first = await view.evaluate(`snapshot()`) as {
      text?: string; canvasCount: number; cells: number; pages: number;
      events: Array<{ type: string; detail: unknown }>;
    };
    expect(first.canvasCount).toBe(0);
    expect(first.text).toContain("تقرير الإنتاج الأسبوعي");

    await view.evaluate(`goTo(1)`);
    const second = await view.evaluate(`snapshot()`) as {
      text?: string; canvasCount: number; events: Array<{ type: string; detail: unknown }>;
    };
    expect(second.text).toContain("الاسم");
    expect(second.canvasCount).toBe(0);
    expect(second.events.some((event) => event.type === "indexchange")).toBe(true);
    expect(second.events).toContainEqual({ type: "slidechange", detail: { slide: 1 } });

    const zoomed = await view.evaluate(`setZoom(1.5)`) as { zoom: number };
    expect(zoomed.zoom).toBe(1.5);

    const cleared = await view.evaluate(`destroyPreview()`) as { children: number };
    expect(cleared.children).toBe(0);

    // XLSX: detected, grid painted, sheet navigation works.
    const workbook = await view.evaluate(`openPreview("/arabic.xlsx")`) as {
      format: string; count: number; titles: string[];
    };
    expect(workbook.format).toBe("xlsx");
    expect(workbook.count).toBe(workbook.titles.length);
    const grid = await view.evaluate(`snapshot()`) as { cells: number; canvasCount: number };
    expect(grid.cells).toBeGreaterThan(0);
    expect(grid.canvasCount).toBe(0);
    const summaryIndex = workbook.titles.findIndex((title) => title.toLowerCase().includes("summary"));
    if (summaryIndex > 0) {
      await view.evaluate(`goTo(${summaryIndex})`);
      const summary = await view.evaluate(`snapshot()`) as {
        text?: string;
        events: Array<{ type: string; detail: unknown }>;
      };
      expect(summary.text).toContain("Production budget");
      expect(summary.events.some((event) => event.type === "sheetchange")).toBe(true);
    }
    await view.evaluate(`destroyPreview()`);

    // DOCX: detected, layout pagination paints physical pages.
    const document = await view.evaluate(
      `openPreview("/arabic.docx", { docx: { pagination: { mode: "layout" } } })`,
    ) as { format: string; count: number };
    expect(document.format).toBe("docx");
    const paged = await view.evaluate(`snapshot()`) as { text?: string; pages: number; canvasCount: number };
    expect(paged.pages).toBeGreaterThanOrEqual(1);
    expect(paged.text).toContain("1. Project Overview");
    expect(paged.canvasCount).toBe(0);
    await view.evaluate(`destroyPreview()`);
  } finally {
    server.stop();
  }
});
