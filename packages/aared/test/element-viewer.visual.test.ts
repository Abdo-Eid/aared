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
    expect(pptxLoad.count).toBe(18);
    const firstSlide = await view.evaluate(`(async () => {
      await frame();
      const viewer = document.querySelector("#viewer");
      return {
        title: viewer.shadowRoot.querySelector('[part="page"]')?.textContent,
        format: viewer.viewer?.format,
        count: viewer.viewer?.info.count,
        canvasCount: viewer.shadowRoot.querySelectorAll("canvas").length,
        progress: window.events.filter((event) => event.type === "progress").length,
        events: take().map((event) => event.type),
      };
    })()`) as { title?: string; format: string; count: number; canvasCount: number; progress: number; events: string[] };
    expect(firstSlide.title).toContain("تقرير الإنتاج الأسبوعي");
    expect(firstSlide.format).toBe("pptx");
    expect(firstSlide.count).toBe(18);
    expect(firstSlide.canvasCount).toBe(0);
    expect(firstSlide.progress).toBeGreaterThan(0);
    expect(firstSlide.events).toEqual(["load"]);

    // --- viewer.goTo(): a promise, one announcement, the attribute follows, going nowhere is silent ---
    const navigation = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      await viewer.viewer.goTo(1);
      const attribute = viewer.getAttribute("index");
      const moved = take();
      await viewer.viewer.goTo(1);
      viewer.index = 1;
      await frame();
      return {
        index: viewer.index,
        attribute,
        table: viewer.shadowRoot.querySelector('[data-aared-element="table"]')?.textContent,
        moved,
        repeated: take(),
      };
    })()`) as { index: number; attribute: string; table?: string; moved: Array<{ type: string; detail: unknown }>; repeated: unknown[] };
    expect(navigation.index).toBe(1);
    expect(navigation.attribute).toBe("1");
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
      await viewer.viewer.goTo(2);
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
    // "+" from a fit mode steps from the scale it actually painted, not from 100%: 10 points on.
    expect(zoom.fitScale).toBeLessThan(1);
    expect(zoom.stepped).toBeCloseTo(Math.round((zoom.fitScale + 0.1) * 100) / 100, 2);
    expect(zoom.controlsHidden).toBe(false);
    // `toolbar` is the documented alias of `controls`, so the element carries both parts.
    expect(zoom.controlsParts).toEqual(["controls", "toolbar"]);

    // --- The attributes and the viewer are one state ---
    const sync = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      take();
      viewer.viewer.setZoom(2);
      await frame();
      const zoomAttribute = viewer.getAttribute("zoom");
      // Removing the attribute goes back to the default, and it stays absent.
      viewer.removeAttribute("zoom");
      await frame();
      const cleared = { attribute: viewer.getAttribute("zoom"), zoom: viewer.viewer.zoom };
      await viewer.viewer.goTo(3);
      const indexAttribute = viewer.getAttribute("index");
      take();
      // A unit that cannot be shown: reported once, and the attribute goes back to the unit on screen.
      viewer.index = 99;
      await frame();
      const restored = { attribute: viewer.getAttribute("index"), shown: viewer.viewer.index };
      const errors = take().filter((event) => event.type === "error").map((event) => event.detail.cause.code);
      // The bar's buttons are the viewer's toolbar methods; the attributes follow.
      const button = (text) => [...viewer.shadowRoot.querySelectorAll('[part~="button"]')].find((element) => element.textContent === text);
      button("Next").click();
      await new Promise((resolve) => viewer.addEventListener("indexchange", resolve, { once: true }));
      const nextAttribute = viewer.getAttribute("index");
      viewer.viewer.zoomIn(0.5);
      await frame();
      const zoomedIn = viewer.getAttribute("zoom");
      viewer.viewer.resetZoom();
      await frame();
      const reset = viewer.getAttribute("zoom");
      viewer.removeAttribute("zoom");
      // The element's viewer searches the deck it has open.
      const searched = (await viewer.viewer.search("الاسم")).map((match) => match.slide.index);
      take();
      return { zoomAttribute, cleared, indexAttribute, restored, errors, nextAttribute, zoomedIn, reset, searched };
    })()`) as {
      zoomAttribute: string;
      cleared: { attribute: string | null; zoom: unknown };
      indexAttribute: string;
      restored: { attribute: string; shown: number };
      errors: string[];
      nextAttribute: string;
      zoomedIn: string;
      reset: string;
      searched: number[];
    };
    expect(sync.zoomAttribute).toBe("2");
    expect(sync.cleared).toEqual({ attribute: null, zoom: "fit-page" });
    expect(sync.indexAttribute).toBe("3");
    expect(sync.restored).toEqual({ attribute: "3", shown: 3 });
    expect(sync.errors).toEqual(["not-found"]);
    expect(sync.nextAttribute).toBe("4");
    expect(Number(sync.zoomedIn)).toBeGreaterThan(0);
    // Back at the format's own default: an attribute that was set says so.
    expect(sync.reset).toBe("fit-page");
    expect(sync.searched).toContain(1);

    // --- Errors: AaredError, same codes as preview(), reported once, rethrown ---
    const errors = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      take();
      let thrown = null;
      try { await viewer.viewer.goTo(19); } catch (error) { thrown = error; }
      const reported = take().filter((event) => event.type === "error");
      // Same split as preview(): a whole number outside is not-found, below 0 too.
      let negative = null;
      try { await viewer.viewer.goTo(-1); } catch (error) { negative = error.code; }
      let fractional = null;
      try { await viewer.viewer.goTo(1.5); } catch (error) { fractional = error.code; }
      take();
      viewer.setAttribute("type", "xlsx");
      await frame();
      const mismatch = take().filter((event) => event.type === "error");
      return {
        thrown: { name: thrown?.name, code: thrown?.code, message: thrown?.message },
        reportedCount: reported.length,
        sameInstance: reported[0]?.detail.cause === thrown,
        negative,
        fractional,
        mismatch: mismatch.map((event) => event.detail.cause?.code),
      };
    })()`) as { thrown: { name: string; code: string; message: string }; reportedCount: number; sameInstance: boolean; negative: string; fractional: string; mismatch: string[] };
    expect(errors.thrown).toMatchObject({ name: "AaredError", code: "not-found" });
    expect(errors.thrown.message).toContain(`outside this ${pptxLoad.count}-unit document`);
    expect(errors.reportedCount).toBe(1);
    expect(errors.sameInstance).toBe(true);
    expect(errors.negative).toBe("not-found");
    expect(errors.fractional).toBe("invalid-argument");
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

    // The open() a newer one replaced says so: it rejects `aborted` (no `error`
    // event) instead of resolving as if it had finished.
    const superseded = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      take();
      viewer.src = "/slow.pptx";
      const first = viewer.open().then(() => "resolved", (error) => error.code);
      viewer.src = "/arabic.pptx";
      const outcome = await first;
      await new Promise((resolve) => viewer.addEventListener("load", resolve, { once: true }));
      return { outcome, errors: take().filter((event) => event.type === "error").length };
    })()`) as { outcome: unknown; errors: number };
    expect(superseded).toEqual({ outcome: "aborted", errors: 0 });

    // --- index and zoom set while the file loads: the open has already read
    // the old values, so they are applied before `load`, and no change event
    // comes ahead of it.
    const caughtUp = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      take();
      const loaded = new Promise((resolve) => viewer.addEventListener("load", () => resolve(take().map((event) => event.type)), { once: true }));
      viewer.src = "/slow.pptx";
      viewer.index = 2;
      viewer.zoom = 1.5;
      const beforeLoad = await loaded;
      await frame();
      const result = { beforeLoad, shown: viewer.viewer.index, attribute: viewer.getAttribute("index"), scale: slideScale() };
      viewer.removeAttribute("zoom");
      const back = new Promise((resolve) => viewer.addEventListener("indexchange", resolve, { once: true }));
      viewer.index = 0;
      await back;
      take();
      return result;
    })()`) as { beforeLoad: string[]; shown: number; attribute: string; scale: number };
    expect(caughtUp.beforeLoad).toEqual(["load"]);
    expect(caughtUp.shown).toBe(2);
    expect(caughtUp.attribute).toBe("2");
    expect(caughtUp.scale).toBeCloseTo(1.5, 2);

    // --- type on an open document: the same type, or removing it, needs no
    // download and changes nothing (no load, progress or error; same viewer).
    const typed = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      await frame();
      take();
      const before = viewer.viewer;
      viewer.setAttribute("type", "pptx");
      viewer.removeAttribute("type");
      await frame();
      return { events: window.events.map((event) => event.type), same: viewer.viewer === before };
    })()`) as { events: string[]; same: boolean };
    expect(typed).toEqual({ events: [], same: true });

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
      // A sheet-tab click moves the viewer itself; the index attribute follows.
      take();
      shadow.querySelectorAll('[data-aared-element="sheet-tab"]')[1].click();
      await new Promise((resolve) => viewer.addEventListener("indexchange", resolve, { once: true }));
      const tabbed = { index: viewer.getAttribute("index"), shown: viewer.viewer.index, events: take().map((event) => event.type) };
      viewer.zoom = "fit-width";
      await frame();
      const scroller = shadow.querySelector('[data-aared-element="sheet-scroll"]');
      const fitted = {
        label: label(),
        overflow: scroller.scrollWidth - scroller.clientWidth,
        fillColumns: shadow.querySelectorAll('[data-aared-element="fill-column-header"]').length,
      };
      viewer.removeAttribute("zoom");
      return { initial, dark, light, tabbed, fitted, canvasCount: shadow.querySelectorAll("canvas").length };
    })()`) as {
      initial: { label: string; tabs: number; cells: number; classSheets: number; cellLine: string; headingFill: string };
      dark: { tabs: number; surface?: string };
      light: { tabs: number };
      tabbed: { index: string; shown: number; events: string[] };
      fitted: { label: string; overflow: number; fillColumns: number };
      canvasCount: number;
    };
    expect(xlsx.initial.label).toBe("100%");
    expect(xlsx.initial.tabs).toBeGreaterThan(0);
    expect(xlsx.initial.cells).toBeGreaterThan(0);
    // The element's own chrome sheet, plus the grid's class rules beside it.
    expect(xlsx.initial.classSheets).toBe(2);
    expect(xlsx.initial.cellLine).toBe("rgb(231, 233, 238)");
    expect(xlsx.initial.headingFill).toBe("rgb(244, 244, 247)");
    expect(xlsx.dark.tabs).toBe(xlsx.initial.tabs);
    expect(xlsx.dark.surface?.toLowerCase()).toBe("#1a1a1a");
    expect(xlsx.light.tabs).toBe(xlsx.initial.tabs);
    expect(xlsx.tabbed).toEqual({ index: "1", shown: 1, events: ["indexchange", "sheetchange"] });
    expect(xlsx.fitted.label).toBe("Fit width");
    expect(xlsx.fitted.overflow).toBeLessThanOrEqual(1);
    expect(xlsx.fitted.fillColumns).toBe(0);
    expect(xlsx.canvasCount).toBe(0);

    // --- Whichever of file and src was set last is what opens ---------------
    const lastSet = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      viewer.index = 0;
      const loaded = () => new Promise((resolve) => viewer.addEventListener("load", (event) => resolve(event.detail.format), { once: true }));
      const file = new File([await (await fetch("/arabic.pptx")).arrayBuffer()], "deck.pptx");
      let next = loaded();
      viewer.file = file;
      const fileAfterSrc = await next;
      next = loaded();
      viewer.src = "/arabic.xlsx";
      const srcAfterFile = await next;
      next = loaded();
      viewer.file = null;
      await next;
      // Clearing both empties the viewer quietly: nothing on the page, no event.
      take();
      viewer.removeAttribute("src");
      await frame();
      const page = () => viewer.shadowRoot.querySelector('[part="page"]').childElementCount;
      const emptied = { viewer: viewer.viewer, children: page(), events: take().map((event) => event.type) };
      // Cleared while it loads: the open stops, and still nothing fires.
      viewer.src = "/arabic.pptx";
      viewer.src = null;
      await new Promise((resolve) => setTimeout(resolve, 300));
      const cancelled = { viewer: viewer.viewer, children: page(), events: take().map((event) => event.type) };
      // Asking to open with nothing to open is a mistake: it rejects and is reported once.
      const asked = await viewer.open().then(() => "resolved", (error) => error.code);
      const reported = take().map((event) => event.type);
      return { fileAfterSrc, srcAfterFile, emptied, cancelled, asked, reported };
    })()`) as {
      fileAfterSrc: string;
      srcAfterFile: string;
      emptied: { viewer: unknown; children: number; events: string[] };
      cancelled: { viewer: unknown; children: number; events: string[] };
      asked: string;
      reported: string[];
    };
    expect(lastSet).toEqual({
      fileAfterSrc: "pptx",
      srcAfterFile: "xlsx",
      emptied: { viewer: null, children: 0, events: [] },
      cancelled: { viewer: null, children: 0, events: [] },
      asked: "invalid-state",
      reported: ["error"],
    });

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

    // --- DOCX physical pages through the pagination attribute ---------------
    const layout = await view.evaluate(`(async () => {
      const viewer = document.querySelector("#viewer");
      const shadow = viewer.shadowRoot;
      const loaded = () => new Promise((resolve) => viewer.addEventListener("load", resolve, { once: true }));
      const pageCount = () => shadow.querySelectorAll('section[data-aared-element="page"]').length;
      let next = loaded();
      viewer.pagination = "layout";
      await next;
      await frame();
      const paged = {
        pages: pageCount(),
        pageCount: viewer.viewer.docx.pageCount,
        marker: viewer.dataset.aaredPagination,
        // Each page carries the element's page shadow; the box around them none.
        partShadow: getComputedStyle(shadow.querySelector('[part="page"]')).boxShadow,
        pageShadow: getComputedStyle(shadow.querySelector('section[data-aared-element="page"]')).boxShadow,
      };
      // The same mode again re-opens nothing.
      const same = viewer.viewer;
      viewer.pagination = "layout";
      await frame();
      const kept = viewer.viewer === same;
      next = loaded();
      viewer.removeAttribute("pagination");
      await next;
      await frame();
      return { paged, kept, flow: { pages: pageCount(), marker: viewer.dataset.aaredPagination ?? null } };
    })()`) as {
      paged: { pages: number; pageCount: number; marker: string; partShadow: string; pageShadow: string };
      kept: boolean;
      flow: { pages: number; marker: string | null };
    };
    expect(layout.paged.pages).toBeGreaterThan(1);
    expect(layout.paged.pageCount).toBe(layout.paged.pages);
    expect(layout.paged.marker).toBe("layout");
    expect(layout.paged.partShadow).toBe("none");
    expect(layout.paged.pageShadow).toContain("0px 1px 5px");
    expect(layout.kept).toBe(true);
    expect(layout.flow).toEqual({ pages: 0, marker: null });
  } finally {
    server.stop();
  }
});
