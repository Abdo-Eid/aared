import { expect, test } from "bun:test";
import { browserModule } from "./harness.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;

visual("viewer search highlights every format, clears, opts out, and follows navigation and virtualization", async () => {
  const [previewModule, workerModule] = await Promise.all([
    browserModule(`${import.meta.dir}/../src/preview/index.ts`),
    browserModule(`${import.meta.dir}/../src/core/worker.ts`),
  ]);
  const files = new Map<string, ArrayBuffer>();
  for (const format of ["pptx", "xlsx", "docx"]) {
    files.set(`/fixture.${format}`, await Bun.file(`${import.meta.dir}/../../../fixtures/${format}-coverage.${format}`).arrayBuffer());
  }
  const server = Bun.serve({ port: 0, fetch(request) {
    const path = new URL(request.url).pathname;
    if (path === "/preview.mjs" || path === "/worker.mjs") return new Response(path === "/preview.mjs" ? previewModule : workerModule, { headers: { "content-type": "text/javascript" } });
    if (files.has(path)) return new Response(files.get(path));
    return new Response(`<!doctype html><main id="host" style="width:1100px;height:650px"></main><script type="module">
      import { preview } from "/preview.mjs";
      const host = document.querySelector("#host");
      window.probeFind = async (format) => {
        const viewer = await preview({source:'/fixture.' + format, target:host, pagination:{mode:'layout'}});
        try {
          if (format === 'pptx') await viewer.goTo(4);
          const matches = await viewer.search(format === 'pptx' ? 'نقطة' : format === 'xlsx' ? 'الانتاج' : 'nile');
          const state = () => ({index:viewer.searchNavigation.index,count:viewer.searchNavigation.count});
          const outline = () => document.querySelector('[data-aared-search-active="' + host.dataset.aaredSearchHighlight + '"]');
          const first = state();
          const firstBox = outline().firstElementChild?.getBoundingClientRect().toJSON();
          await viewer.searchNavigation.next();
          await new Promise(r=>setTimeout(r,450));
          const second = state();
          const secondBox = outline().firstElementChild?.getBoundingClientRect().toJSON();
          const outlined = outline().childElementCount > 0 && getComputedStyle(outline().firstElementChild).borderStyle === 'solid';
          const moved = firstBox && secondBox && (firstBox.x !== secondBox.x || firstBox.y !== secondBox.y);
          viewer.setZoom(0.8);
          await new Promise(r=>setTimeout(r,100));
          const zoomBox = outline().firstElementChild?.getBoundingClientRect().toJSON();
          const followsZoom = !!zoomBox && Math.abs(zoomBox.width - secondBox.width) > 0.1;
          await viewer.searchNavigation.previous();
          await viewer.searchNavigation.previous();
          const wrapped = state();
          await viewer.searchNavigation.next();
          const back = state();
          await viewer.search('', {highlight:false});
          const cleared = state().count === 0 && state().index === null && outline().childElementCount === 0;
          return {first,second,wrapped,back,outlined,moved:!!moved,followsZoom,cleared,matchCount:matches.length};
        } finally {viewer.destroy();}
      };
      window.probeIsolation = async () => {
        const shadowHost = document.body.appendChild(document.createElement("div"));
        const shadow = shadowHost.attachShadow({mode: "open"});
        const target = shadow.appendChild(document.createElement("div"));
        target.style.cssText = "width:1100px;height:650px";
        const first = await preview({ source: '/fixture.docx', target: host });
        const second = await preview({ source: '/fixture.docx', target });
        try {
          await first.search("nile");
          await second.search("nile");
          const a = host.dataset.aaredSearchHighlight, b = target.dataset.aaredSearchHighlight;
          const separate = a !== b && CSS.highlights.get(a).size > 0 && CSS.highlights.get(b).size > 0;
          const scopedStyle = shadow.querySelector('style[data-aared-search-highlight-style]').parentNode === shadow;
          await first.search("");
          const secondKept = !CSS.highlights.has(a) && CSS.highlights.get(b).size > 0;
          first.destroy();
          const destroyIsolated = CSS.highlights.get(b).size > 0;
          second.destroy();
          const cleaned = !CSS.highlights.has(b) && !shadow.querySelector('style[data-aared-search-highlight-style]');
          return { separate, scopedStyle, secondKept, destroyIsolated, cleaned };
        } finally { first.destroy(); second.destroy(); shadowHost.remove(); }
      };
      window.probe = async (format) => {
        const viewer = await preview({ source: '/fixture.' + format, target: host, pagination: { mode: "layout" } });
        const ranges = () => [...(CSS.highlights.get(host.dataset.aaredSearchHighlight) ?? [])];
        const query = format === "docx" ? "nile" : format === "xlsx" ? "الانتاج" : "الانتاج";
        const plain = await viewer.search(query, { highlight: false });
        const noController = host.dataset.aaredSearchHighlight === undefined;
        const before = host.firstElementChild.getBoundingClientRect().height;
        const matches = await viewer.search(query);
        const highlighted = ranges().map(r => r.toString());
        const sameLayout = before === host.firstElementChild.getBoundingClientRect().height;
        let headerCopies = 0, headerHighlights = 0, footerCopies = 0, footerHighlights = 0, navigated = true, virtualized = true, occurrencesAgree = true;
        if (format === "docx") {
          headerCopies = [...host.querySelectorAll('[data-aared-element="page-header"]')].filter(n => n.textContent.toLowerCase().includes("nile")).length;
          headerHighlights = ranges().filter(r => r.startContainer.parentElement.closest('[data-aared-element="page-header"]')).length;
          await viewer.docx.goToSearchMatch(matches.find(m => m.kind === "header"));
          await viewer.search("confidential");
          footerCopies = host.querySelectorAll('[data-aared-element="page-footer"]').length;
          footerHighlights = ranges().filter(r => r.startContainer.parentElement.closest('[data-aared-element="page-footer"]')).length;
          await viewer.search("استوديو");
          navigated = ranges().filter(r => r.startContainer.parentElement.closest('[data-aared-element="page-header"]')).length === headerCopies;
        } else if (format === "pptx") {
          await viewer.goTo(matches[matches.length - 1].slide.index);
          await new Promise(r => setTimeout(r, 50));
          navigated = ranges().length > 0 && ranges().every(r => r.startContainer.isConnected);
          const bullets = await viewer.search("نقطة");
          await viewer.goTo(bullets[0].slide.index);
          await new Promise(r => setTimeout(r, 50));
          occurrencesAgree = bullets.length === 5 && ranges().length === 5 && ranges().every(r => r.toString() === "نقطة");
        } else {
          await viewer.goTo(1);
          await viewer.search("budget");
          await viewer.goTo(0);
          await viewer.goTo(1);
          await new Promise(r => setTimeout(r, 50));
          navigated = ranges().length > 0 && ranges().every(r => r.startContainer.isConnected);
          await viewer.goTo(2);
          await viewer.search("0");
          const oldNodes = ranges().map(r => r.startContainer);
          await viewer.xlsx.goToCell("A80");
          await new Promise(r => setTimeout(r, 100));
          virtualized = ranges().length > 0 && ranges().every(r => r.startContainer.isConnected) && oldNodes.some(n => !n.isConnected);
        }
        const opted = await viewer.search(query, { highlight: false });
        const noHighlight = ranges().length === 0;
        await viewer.search(query);
        const beforeAbort = ranges().length;
        const aborted = await viewer.search("missing", { signal: AbortSignal.abort() }).then(() => "resolved", e => e.code);
        const keptOnAbort = ranges().length === beforeAbort;
        const cleared = await viewer.search("");
        const clearCount = ranges().length;
        await Promise.all([viewer.search(query), viewer.search("")]);
        const newestWins = ranges().length === 0;
        await viewer.search(query);
        const name = host.dataset.aaredSearchHighlight;
        viewer.destroy();
        return { plainCount: plain.length, matchCount: matches.length, optedCount: opted.length, noController, highlighted, sameLayout, headerCopies, headerHighlights, footerCopies, footerHighlights, navigated, virtualized, occurrencesAgree, noHighlight, aborted, keptOnAbort, cleared, clearCount, newestWins, removed: !CSS.highlights.has(name) };
      };
    </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
  } });
  try {
    await using view = new Bun.WebView({ width: 1200, height: 800 });
    await view.navigate(server.url.href);
    for (const format of ["pptx", "xlsx", "docx"] as const) {
      const result = await view.evaluate(`probe(${JSON.stringify(format)})`) as {
        plainCount: number; matchCount: number; optedCount: number; noController: boolean; highlighted: string[]; sameLayout: boolean;
        headerCopies: number; headerHighlights: number; footerCopies: number; footerHighlights: number; navigated: boolean; virtualized: boolean; occurrencesAgree: boolean;
        noHighlight: boolean; aborted: string; keptOnAbort: boolean; cleared: unknown[]; clearCount: number; newestWins: boolean; removed: boolean;
      };
      expect(result.matchCount).toBeGreaterThan(0);
      expect(result.noController).toBe(true);
      expect(result.matchCount).toBe(result.plainCount);
      expect(result.highlighted.length).toBeGreaterThan(0);
      expect(result.sameLayout).toBe(true);
      expect(result.navigated).toBe(true);
      expect(result.virtualized).toBe(true);
      expect(result.occurrencesAgree).toBe(true);
      expect(result.noHighlight).toBe(true);
      expect(result.aborted).toBe("aborted");
      // The query may not occur on the slide/sheet navigated to; cancellation
      // must preserve that state rather than install the cancelled query.
      expect(result.keptOnAbort).toBe(true);
      expect(result.cleared).toEqual([]);
      expect(result.clearCount).toBe(0);
      expect(result.newestWins).toBe(true);
      expect(result.removed).toBe(true);
      if (format === "docx") {
        expect(result.headerCopies).toBeGreaterThan(2);
        expect(result.headerHighlights).toBe(result.headerCopies);
        expect(result.footerHighlights).toBe(result.footerCopies);
      }
    }
    expect(await view.evaluate("probeIsolation()") as Record<string, boolean>).toEqual({
      separate: true, scopedStyle: true, secondKept: true, destroyIsolated: true, cleaned: true,
    });
    for (const format of ['pptx','xlsx','docx']) {
      const result = await view.evaluate(`probeFind(${JSON.stringify(format)})`) as {
        first:{index:number;count:number};second:{index:number;count:number};wrapped:{index:number;count:number};back:{index:number;count:number};outlined:boolean;moved:boolean;followsZoom:boolean;cleared:boolean;matchCount:number;
      };
      expect(result.first.index).toBe(0);
      expect(result.first.count).toBeGreaterThan(1);
      expect(result.second.index).toBe(1);
      expect(result.wrapped.index).toBe(result.first.count - 1);
      expect(result.back.index).toBe(0);
      expect(result.outlined).toBe(true);
      expect(result.moved).toBe(true);
      expect(result.followsZoom).toBe(true);
      expect(result.cleared).toBe(true);
      if(format==='pptx') expect(result.first.count).toBe(5);
      if(format==='docx') expect(result.first.count).toBeGreaterThan(result.matchCount);
    }
  } finally { server.stop(); }
}, 60000);
