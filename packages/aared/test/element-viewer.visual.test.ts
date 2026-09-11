import { expect, test } from "bun:test";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/arabic-pptx-rtl-coverage.pptx`;
const XLSX_FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/arabic-xlsx-rtl-coverage.xlsx`;
const DOCX_FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/arabic-docx-bidi-coverage.docx`;

async function browserModule(entrypoint: string): Promise<string> {
  const build = await Bun.build({ entrypoints: [entrypoint], target: "browser", format: "esm" });
  if (!build.success) throw new Error(build.logs.map((log) => log.message).join("\n"));
  return build.outputs[0]!.text();
}

visual("visual element output: the PPTX viewer renders, navigates, and reports lifecycle events", async () => {
  const [elementModule, workerModule, fixture, xlsxFixture, docxFixture] = await Promise.all([
    browserModule(`${import.meta.dir}/../src/element/index.ts`),
    browserModule(`${import.meta.dir}/../src/core/worker.ts`),
    Bun.file(FIXTURE_PATH).arrayBuffer(),
    Bun.file(XLSX_FIXTURE_PATH).arrayBuffer(),
    Bun.file(DOCX_FIXTURE_PATH).arrayBuffer(),
  ]);
  const server = Bun.serve({
    port: 0,
    async fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/element.mjs") {
        return new Response(elementModule, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      if (url.pathname === "/worker.mjs") {
        return new Response(workerModule, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      if (url.pathname === "/slow.pptx") await Bun.sleep(100);
      if (url.pathname === "/arabic.pptx" || url.pathname === "/slow.pptx") {
        return new Response(fixture, {
          headers: { "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
        });
      }
      if (url.pathname === "/arabic.xlsx") {
        return new Response(xlsxFixture, {
          headers: { "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        });
      }
      if (url.pathname === "/arabic.docx") {
        return new Response(docxFixture, {
          headers: { "content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
        });
      }
      return new Response(`<!doctype html><aared-viewer id="viewer"></aared-viewer><script type="module">
        import "/element.mjs";
        const viewer = document.querySelector("#viewer");
        window.events = [];
        for (const type of ["load", "pagechange", "progress", "error"]) {
          viewer.addEventListener(type, (event) => window.events.push({ type, detail: event.detail }));
        }
        window.openSource = (source) => new Promise((resolve, reject) => {
          const load = (event) => { cleanup(); resolve(event.detail); };
          const error = (event) => { cleanup(); reject(new Error(event.detail.message)); };
          const cleanup = () => { viewer.removeEventListener("load", load); viewer.removeEventListener("error", error); };
          viewer.addEventListener("load", load, { once: true });
          viewer.addEventListener("error", error, { once: true });
          viewer.src = source;
        });
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1280, height: 720 });
    await view.navigate(server.url.href);
    await view.evaluate(`openSource("/arabic.pptx")`);
    const firstSlide = await view.evaluate(`(() => {
      const viewer = document.querySelector("#viewer");
      return {
        title: viewer.shadowRoot.querySelector('[part="page"]')?.textContent,
        stagePart: viewer.shadowRoot.querySelector('[part="stage"]')?.getAttribute("part"),
        pagePart: viewer.shadowRoot.querySelector('[part="page"]')?.getAttribute("part"),
        canvasCount: viewer.shadowRoot.querySelectorAll("canvas").length,
        events: window.events,
      };
    })()`) as { title?: string; stagePart?: string; pagePart?: string; canvasCount: number; events: Array<{ type: string }> };
    expect(firstSlide.title).toContain("تقرير الإنتاج الأسبوعي");
    expect(firstSlide.stagePart).toBe("stage");
    expect(firstSlide.pagePart).toBe("page");
    expect(firstSlide.canvasCount).toBe(0);
    expect(firstSlide.events.some((event: { type: string }) => event.type === "load")).toBe(true);
    expect(firstSlide.events.some((event: { type: string }) => event.type === "progress")).toBe(true);

    await view.evaluate(`(() => new Promise((resolve) => {
      const viewer = document.querySelector("#viewer");
      viewer.addEventListener("pagechange", () => resolve(), { once: true });
      viewer.goTo(1);
    }))()`);
    const secondSlide = await view.evaluate(`(() => {
      const viewer = document.querySelector("#viewer");
      return {
        page: viewer.page,
        table: viewer.shadowRoot.querySelector('[data-aared-element="table"]')?.textContent,
        canvasCount: viewer.shadowRoot.querySelectorAll("canvas").length,
      };
    })()`) as { page: number; table?: string; canvasCount: number };
    expect(secondSlide.page).toBe(1);
    expect(secondSlide.table).toContain("الاسم");
    expect(secondSlide.canvasCount).toBe(0);

    const chrome = await view.evaluate(`(() => {
      const viewer = document.querySelector("#viewer");
      viewer.controls = true;
      viewer.zoom = 1.5;
      return {
        controls: viewer.shadowRoot.querySelector('[part="controls"]')?.hidden,
        controlsPart: viewer.shadowRoot.querySelector('[part="controls"]')?.getAttribute("part"),
        width: viewer.shadowRoot.querySelector('[part="page"]')?.style.width,
      };
    })()`) as { controls: boolean; controlsPart?: string; width?: string };
    expect(chrome.controls).toBe(false);
    expect(chrome.controlsPart).toBe("controls");
    expect(chrome.width).toBe("150%");

    const errors = await view.evaluate(`(() => {
      const viewer = document.querySelector("#viewer");
      let thrown = "";
      try { viewer.goTo(9); } catch (error) { thrown = error.message; }
      viewer.setAttribute("type", "xlsx");
      return new Promise((resolve) => setTimeout(() => resolve({ thrown, errors: window.events.filter((event) => event.type === "error") }), 25));
    })()`) as { thrown: string; errors: Array<{ detail: { message: string } }> };
    expect(errors.thrown).toContain("outside this 4-unit document");
    expect(errors.errors.some((event: { detail: { message: string } }) => event.detail.message.includes("explicit type"))).toBe(true);

    const stale = await view.evaluate(`(() => new Promise((resolve, reject) => {
      const viewer = document.querySelector("#viewer");
      viewer.removeAttribute("type");
      viewer.page = 0;
      const loadsBefore = window.events.filter((event) => event.type === "load").length;
      const load = () => {
        viewer.removeEventListener("error", error);
        resolve({ loads: window.events.filter((event) => event.type === "load").length - loadsBefore, text: viewer.shadowRoot.querySelector('[part="page"]')?.textContent });
      };
      const error = (event) => { viewer.removeEventListener("load", load); reject(new Error(event.detail.message)); };
      viewer.addEventListener("load", load, { once: true });
      viewer.addEventListener("error", error, { once: true });
      viewer.src = "/slow.pptx";
      viewer.src = "/arabic.pptx";
    }))()`) as { loads: number; text?: string };
    expect(stale.loads).toBe(1);
    expect(stale.text).toContain("تقرير الإنتاج الأسبوعي");

    const xlsxLoad = await view.evaluate(`(() => {
      const viewer = document.querySelector("#viewer");
      return openSource("/arabic.xlsx");
    })()`) as { format: string; count: number };
    expect(xlsxLoad.format).toBe("xlsx");
    const xlsxDom = await view.evaluate(`(() => {
      const viewer = document.querySelector("#viewer");
      return {
        hasSheet: Boolean(viewer.shadowRoot.querySelector('[data-aared-element="sheet"]')),
        cellCount: viewer.shadowRoot.querySelectorAll('[data-aared-element="cell"]').length,
        canvasCount: viewer.shadowRoot.querySelectorAll("canvas").length,
      };
    })()`) as { hasSheet: boolean; cellCount: number; canvasCount: number };
    expect(xlsxDom.hasSheet).toBe(true);
    expect(xlsxDom.cellCount).toBeGreaterThan(0);
    expect(xlsxDom.canvasCount).toBe(0);

    const docxLoad = await view.evaluate(`(() => {
      const viewer = document.querySelector("#viewer");
      viewer.setAttribute("type", "docx");
      return openSource("/arabic.docx");
    })()`) as { format: string; count: number };
    expect(docxLoad.format).toBe("docx");
    const docxDom = await view.evaluate(`(() => {
      const viewer = document.querySelector("#viewer");
      return {
        hasDocument: Boolean(viewer.shadowRoot.querySelector('[data-aared-element="document"]')),
        paragraphCount: viewer.shadowRoot.querySelectorAll('[data-aared-element="paragraph"]').length,
        canvasCount: viewer.shadowRoot.querySelectorAll("canvas").length,
      };
    })()`) as { hasDocument: boolean; paragraphCount: number; canvasCount: number };
    expect(docxDom.hasDocument).toBe(true);
    expect(docxDom.paragraphCount).toBeGreaterThan(0);
    expect(docxDom.canvasCount).toBe(0);
  } finally {
    server.stop();
  }
});
