import { expect, test } from "bun:test";
import { browserModule } from "./harness.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const XLSX_FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;
const DOCX_FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

visual("visual element output: renders, navigates, zooms, themes and reports lifecycle events", async () => {
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
      if (url.pathname === "/missing.pptx") return new Response("gone", { status: 404 });
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
      return new Response(`<!doctype html><body style="margin:0"><aared-viewer id="viewer" style="height:600px"></aared-viewer><script type="module">
        import "/element.mjs";
        const viewer = document.querySelector("#viewer");
        window.events = [];
        for (const type of ["load", "indexchange", "slidechange", "sheetchange", "zoomchange", "progress", "error"]) {
          viewer.addEventListener(type, (event) => window.events.push({ type, detail: event.detail }));
        }
        window.take = () => window.events.splice(0).filter((event) => event.type !== "progress");
        window.openSource = (source) => new Promise((resolve, reject) => {
          const load = (event) => { cleanup(); resolve(event.detail); };
          const error = (event) => { cleanup(); reject(new Error(event.detail.message)); };
          const cleanup = () => { viewer.removeEventListener("load", load); viewer.removeEventListener("error", error); };
          viewer.addEventListener("load", load, { once: true });
          viewer.addEventListener("error", error, { once: true });
          viewer.src = source;
        });
        window.frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        window.slideScale = () => {
          const shadow = viewer.shadowRoot;
          const slide = shadow.querySelector('[data-aared-element="slide"]');
          const scene = shadow.querySelector('[data-aared-element="slide-scene"]');
          return slide.getBoundingClientRect().width / parseFloat(scene.style.width);
        };
        window.label = () => viewer.shadowRoot.querySelector('[data-aared-controls="zoom"]').textContent;
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1280, height: 720 });
    await view.navigate(server.url.href);
    // Script-created elements (unlike parsed markup) throw if the
    // constructor adds attributes — the playground/home demos rely on this.
    expect(await view.evaluate(`document.createElement("aared-viewer").tagName`) as string).toBe("AARED-VIEWER");

    // --- Load: the starting state is `load`, never an indexchange -----------
    const pptxLoad = await view.evaluate(`openSource("/arabic.pptx")`) as { format: string; count: number };
    expect(pptxLoad.format).toBe("pptx");
    expect(pptxLoad.count).toBe(10);
    const firstSlide = await view.evaluate(`(async () => {
      await frame();
      const viewer = document.querySelector("#viewer");
      return {
        title: viewer.shadowRoot.querySelector('[part="page"]')?.textContent,
        format: viewer.format,
        count: viewer.info?.count,
        canvasCount: viewer.shadowRoot.querySelectorAll("canvas").length,
        events: take().map((event) => event.type),
      };
    })()`) as { title?: string; format: string; count: number; canvasCount: number; events: string[] };
    expect(firstSlide.title).toContain("تقرير الإنتاج الأسبوعي");
    expect(firstSlide.format).toBe("pptx");
    expect(firstSlide.count).toBe(10);
    expect(firstSlide.canvasCount).toBe(0);
    expect(firstSlide.events).toEqual(["load"]);

    // --- goTo(): a promise, one announcement, and going nowhere is silent ---
    const navigation = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      await viewer.goTo(1);
      const moved = take();
      await viewer.goTo(1);
      viewer.index = 1;
      await frame();
      return {
        index: viewer.index,
        table: viewer.shadowRoot.querySelector('[data-aared-element="table"]')?.textContent,
        moved,
        repeated: take(),
      };
    })()`) as { index: number; table?: string; moved: Array<{ type: string; detail: unknown }>; repeated: unknown[] };
    expect(navigation.index).toBe(1);
    expect(navigation.table).toContain("الاسم");
    expect(navigation.moved).toEqual([
      { type: "indexchange", detail: { index: 1 } },
      { type: "slidechange", detail: { index: 1, title: "Slide 2" } },
    ]);
    expect(navigation.repeated).toEqual([]);

    // --- Zoom drives the render handle: real scale, one event, kept on navigation ---
    const zoom = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      viewer.controls = true;
      const fitLabel = label();
      viewer.zoom = 1.5;
      await frame();
      const zoomed = { scale: slideScale(), label: label(), events: take() };
      viewer.zoom = 1.5;
      await viewer.goTo(2);
      await frame();
      const kept = { scale: slideScale(), events: take().map((event) => event.type) };
      viewer.removeAttribute("zoom");
      await frame();
      const fitScaleNow = slideScale();
      take();
      viewer.shadowRoot.querySelector('[aria-label="Zoom in"]').click();
      await frame();
      const stepped = Number(viewer.getAttribute("zoom"));
      // An explicit zoom attribute persists across src changes; clear it.
      viewer.removeAttribute("zoom");
      return {
        fitLabel,
        zoomed,
        kept,
        fitScale: fitScaleNow,
        stepped,
        controlsHidden: viewer.shadowRoot.querySelector('[part~="controls"]').hidden,
        controlsParts: [...viewer.shadowRoot.querySelector('[part~="controls"]').part],
      };
    })()`) as {
      fitLabel: string;
      zoomed: { scale: number; label: string; events: Array<{ type: string; detail: unknown }> };
      kept: { scale: number; events: string[] };
      fitScale: number;
      stepped: number;
      controlsHidden: boolean;
      controlsParts: string[];
    };
    expect(zoom.fitLabel).toBe("Fit page");
    expect(zoom.zoomed.scale).toBeCloseTo(1.5, 2);
    expect(zoom.zoomed.label).toBe("150%");
    expect(zoom.zoomed.events).toEqual([{ type: "zoomchange", detail: { zoom: 1.5 } }]);
    // Same value again, then navigation: zoom survives, and nothing is announced but the unit.
    expect(zoom.kept.scale).toBeCloseTo(1.5, 2);
    expect(zoom.kept.events).toEqual(["indexchange", "slidechange"]);
    // "+" from a fit mode steps from the scale it actually painted, not from 100%.
    expect(zoom.fitScale).toBeLessThan(1);
    expect(zoom.stepped).toBeCloseTo(Math.round(zoom.fitScale * 1.25 * 100) / 100, 2);
    expect(zoom.controlsHidden).toBe(false);
    // `toolbar` is the documented alias of `controls`, so the element carries both parts.
    expect(zoom.controlsParts).toEqual(["controls", "toolbar"]);

    // --- Errors: AaredError, same codes as preview(), reported once, rethrown ---
    const errors = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      take();
      let thrown = null;
      try { await viewer.goTo(19); } catch (error) { thrown = error; }
      const reported = take().filter((event) => event.type === "error");
      viewer.setAttribute("type", "xlsx");
      await frame();
      const mismatch = take().filter((event) => event.type === "error");
      return {
        thrown: { name: thrown?.name, code: thrown?.code, message: thrown?.message },
        reportedCount: reported.length,
        sameInstance: reported[0]?.detail.cause === thrown,
        mismatch: mismatch.map((event) => event.detail.cause?.code),
      };
    })()`) as { thrown: { name: string; code: string; message: string }; reportedCount: number; sameInstance: boolean; mismatch: string[] };
    expect(errors.thrown).toMatchObject({ name: "AaredError", code: "not-found" });
    expect(errors.thrown.message).toContain(`outside this ${pptxLoad.count}-unit document`);
    expect(errors.reportedCount).toBe(1);
    expect(errors.sameInstance).toBe(true);
    expect(errors.mismatch).toEqual(["unsupported-format"]);

    const loadErrors = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      viewer.removeAttribute("type");
      const failOf = (attempt) => attempt.then(() => null, (error) => error.message);
      const codeOf = () => take().filter((event) => event.type === "error").map((event) => event.detail.cause?.code);
      await failOf(openSource("/missing.pptx"));
      const missing = codeOf();
      viewer.setAttribute("type", "pdf");
      await frame();
      const badType = codeOf();
      viewer.removeAttribute("type");
      return { missing, badType };
    })()`) as { missing: string[]; badType: string[] };
    expect(loadErrors.missing).toEqual(["network-error"]);
    expect(loadErrors.badType).toEqual(["unsupported-format"]);

    // --- A newer src wins; the slow one never paints -------------------------
    const stale = await view.evaluate(`(() => new Promise((resolve, reject) => {
      const viewer = document.querySelector("#viewer");
      viewer.index = 0;
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

    // --- Theme: set before load, it reaches the slide's marker -----------------
    const darkSlide = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      viewer.theme = "dark";
      viewer.removeAttribute("src");
      await openSource("/arabic.pptx");
      return viewer.shadowRoot.querySelector('[data-aared-element="slide"]')?.dataset.aaredTheme;
    })()`) as string;
    expect(darkSlide).toBe("dark");

    // --- A consumer's CSS variables win in the dark theme too ---------------
    const stageColors = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      const stage = viewer.shadowRoot.querySelector('[part="stage"]');
      const own = getComputedStyle(stage).backgroundColor;
      viewer.style.setProperty("--aared-background", "rgb(1, 2, 3)");
      const custom = getComputedStyle(stage).backgroundColor;
      viewer.style.removeProperty("--aared-background");
      viewer.theme = "light";
      return { own, custom };
    })()`) as { own: string; custom: string };
    expect(stageColors.own).toBe("rgb(28, 28, 31)");
    expect(stageColors.custom).toBe("rgb(1, 2, 3)");

    // --- XLSX: 100% by default, real fit, tabs survive a theme toggle --------
    const xlsxLoad = await view.evaluate(`openSource("/arabic.xlsx")`) as { format: string; count: number };
    expect(xlsxLoad.format).toBe("xlsx");
    const xlsx = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      await frame();
      const shadow = viewer.shadowRoot;
      const tabs = () => shadow.querySelectorAll('[data-aared-element="sheet-tab"]').length;
      const cell = shadow.querySelector('[data-aared-element="cell"]');
      const heading = shadow.querySelector('[data-aared-element="row-header"]');
      const initial = {
        label: label(),
        tabs: tabs(),
        cells: shadow.querySelectorAll('[data-aared-element="cell"]').length,
        // The grid's class rules must reach inside the shadow root, not just the document.
        classSheets: shadow.adoptedStyleSheets.length,
        cellLine: cell && getComputedStyle(cell).borderRightColor,
        headingFill: heading && getComputedStyle(heading).backgroundColor,
      };
      viewer.theme = "dark";
      await frame();
      const bar = shadow.querySelector('[data-aared-element="sheet-tabs"]');
      const dark = { tabs: tabs(), surface: bar?.style.getPropertyValue("--aared-surface") };
      viewer.theme = "light";
      await frame();
      const light = { tabs: tabs() };
      viewer.zoom = "fit-width";
      await frame();
      const scroller = shadow.querySelector('[data-aared-element="sheet-scroll"]');
      const fitted = {
        label: label(),
        overflow: scroller.scrollWidth - scroller.clientWidth,
        fillColumns: shadow.querySelectorAll('[data-aared-element="fill-column-header"]').length,
      };
      viewer.removeAttribute("zoom");
      return { initial, dark, light, fitted, canvasCount: shadow.querySelectorAll("canvas").length };
    })()`) as {
      initial: { label: string; tabs: number; cells: number; classSheets: number; cellLine: string; headingFill: string };
      dark: { tabs: number; surface?: string };
      light: { tabs: number };
      fitted: { label: string; overflow: number; fillColumns: number };
      canvasCount: number;
    };
    expect(xlsx.initial.label).toBe("100%");
    expect(xlsx.initial.tabs).toBeGreaterThan(0);
    expect(xlsx.initial.cells).toBeGreaterThan(0);
    expect(xlsx.initial.classSheets).toBe(1);
    expect(xlsx.initial.cellLine).toBe("rgb(231, 233, 238)");
    expect(xlsx.initial.headingFill).toBe("rgb(244, 244, 247)");
    expect(xlsx.dark.tabs).toBe(xlsx.initial.tabs);
    expect(xlsx.dark.surface?.toLowerCase()).toBe("#1a1a1a");
    expect(xlsx.light.tabs).toBe(xlsx.initial.tabs);
    expect(xlsx.fitted.label).toBe("Fit width");
    expect(xlsx.fitted.overflow).toBeLessThanOrEqual(1);
    expect(xlsx.fitted.fillColumns).toBe(0);
    expect(xlsx.canvasCount).toBe(0);

    // --- DOCX through the type property; dark theme reaches its marker -------
    const docx = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      viewer.type = "docx";
      viewer.theme = "dark";
      const info = await openSource("/arabic.docx");
      const root = viewer.shadowRoot.querySelector('[data-aared-element="document"]');
      const result = {
        format: info.format,
        typeAttribute: viewer.getAttribute("type"),
        theme: root?.dataset.aaredTheme,
        paragraphs: viewer.shadowRoot.querySelectorAll('[data-aared-element="paragraph"]').length,
        canvasCount: viewer.shadowRoot.querySelectorAll("canvas").length,
      };
      viewer.type = null;
      viewer.theme = "light";
      return { ...result, typeCleared: !viewer.hasAttribute("type") };
    })()`) as { format: string; typeAttribute: string; theme?: string; paragraphs: number; canvasCount: number; typeCleared: boolean };
    expect(docx.format).toBe("docx");
    expect(docx.typeAttribute).toBe("docx");
    expect(docx.theme).toBe("dark");
    expect(docx.paragraphs).toBeGreaterThan(0);
    expect(docx.canvasCount).toBe(0);
    expect(docx.typeCleared).toBe(true);
  } finally {
    server.stop();
  }
});
