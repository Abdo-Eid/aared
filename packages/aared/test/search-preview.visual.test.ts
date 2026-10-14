import { expect } from "bun:test";
import { browserModule, visual, withVisualPage } from "./harness.js";

visual("viewer search highlights every format, clears, opts out, and follows navigation and virtualization", async () => {
  const [previewModule, workerModule] = await Promise.all([
    browserModule(`${import.meta.dir}/../src/preview/index.ts`),
    browserModule(`${import.meta.dir}/../src/core/worker.ts`),
  ]);
  const files = new Map<string, ArrayBuffer>();
  for (const format of ["pptx", "xlsx", "docx"]) {
    files.set(`/fixture.${format}`, await Bun.file(`${import.meta.dir}/../../../fixtures/${format}-coverage.${format}`).arrayBuffer());
  }
  await withVisualPage({
    modules: { "/preview.mjs": previewModule, "/worker.mjs": workerModule },
    files: Object.fromEntries(files),
    page: `<!doctype html><main id="host" style="width:1100px;height:650px"></main><script type="module">
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
      // A results list with no branch per format: every result goes through goToSearchMatch().
      window.probeResults = async (format, query) => {
        // A host that scrolls its own pages, as an app's viewer pane does.
        host.style.overflow = 'auto';
        const viewer = await preview({ source: '/fixture.' + format, target: host, ...(format === 'docx' ? { pagination: { mode: 'layout' } } : {}) });
        const errors = [];
        viewer.addEventListener('error', (event) => errors.push(event.detail.cause.code));
        try {
          const matches = await viewer.search(query);
          const last = matches[matches.length - 1];
          const where = () => ({ index: viewer.index, active: viewer.searchNavigation.index, cell: viewer.xlsx?.selection?.address ?? null });
          await viewer.goToSearchMatch(last);
          const listed = where();
          let stale = null, docx = null;
          if (format === 'docx') {
            // The block under the active outline: what a reader sees outlined.
            const outlined = async () => {
              await new Promise((resolve) => setTimeout(resolve, 600));
              const box = document.querySelector('[data-aared-search-active="' + host.dataset.aaredSearchHighlight + '"]')?.firstElementChild?.getBoundingClientRect();
              if (!box) return null;
              const under = document.elementFromPoint(box.left + box.width / 2, box.top + box.height / 2);
              return { header: under?.closest('[data-aared-element="page-header"]') !== null, text: under?.closest('p,h1,h2,h3,h4,h5,h6')?.textContent ?? '' };
            };
            const header = matches.find((match) => match.kind === 'header');
            const body = matches.find((match) => match.story === undefined);
            // This header part repeats on pages 1-3 (later pages use other parts).
            host.scrollTop = 0;
            await viewer.goToSearchMatch(header);
            const headerTop = { active: viewer.searchNavigation.index, under: await outlined() };
            await viewer.docx.goToPage(2);
            await new Promise((resolve) => setTimeout(resolve, 300));
            await viewer.goToSearchMatch(header);
            const headerLater = { active: viewer.searchNavigation.index, under: await outlined() };
            // Past every copy of it: the first copy, as there is none below.
            await viewer.docx.goToPage(viewer.docx.pageCount - 1);
            await new Promise((resolve) => setTimeout(resolve, 300));
            await viewer.goToSearchMatch(header);
            const headerPast = viewer.searchNavigation.index;
            await viewer.goToSearchMatch(body);
            const bodyHit = { active: viewer.searchNavigation.index, under: await outlined(), text: body.text };
            docx = { headerTop, headerLater, headerPast, bodyHit, count: viewer.searchNavigation.count };
          }
          if (format === 'xlsx') {
            // Another sheet on screen: the result brings its own sheet back and selects its cell.
            await viewer.goTo(1);
            await viewer.goToSearchMatch(matches[0]);
            stale = where();
          }
          const foreign = format === 'pptx' ? { kind: 'paragraph', section: 0, heading: null, text: '' } : { slide: { index: 0, title: '' }, text: '' };
          const wrong = await viewer.goToSearchMatch(foreign).then(() => 'resolved', (error) => error.code);
          return {
            count: matches.length,
            last: { slide: last.slide?.index ?? null, sheet: last.sheet?.index ?? null, address: last.address ?? null },
            first: { sheet: matches[0].sheet?.index ?? null, address: matches[0].address ?? null },
            listed, stale, docx, wrong, errors,
          };
        } finally { viewer.destroy(); host.style.overflow = ''; host.scrollTop = 0; }
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
          const ruleFor = (scope, name) => scope.adoptedStyleSheets.some((sheet) => [...sheet.cssRules].some((rule) => rule.cssText.includes("::highlight(" + name + ")")));
          const scopedStyle = ruleFor(shadow, b) && !ruleFor(document, b) && shadow.querySelector("style") === null;
          await first.search("");
          const secondKept = !CSS.highlights.has(a) && CSS.highlights.get(b).size > 0;
          first.destroy();
          const destroyIsolated = CSS.highlights.get(b).size > 0;
          second.destroy();
          const cleaned = !CSS.highlights.has(b) && !ruleFor(shadow, b);
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
        let headerCopies = 0, headerHighlights = 0, footerCopies = 0, footerHighlights = 0, navigated = true, virtualized = true, occurrencesAgree = true, summaryHits = null;
        if (format === "docx") {
          headerCopies = [...host.querySelectorAll('[data-aared-element="page-header"]')].filter(n => n.textContent.toLowerCase().includes("nile")).length;
          headerHighlights = ranges().filter(r => r.startContainer.parentElement.closest('[data-aared-element="page-header"]')).length;
          await viewer.goToSearchMatch(matches.find(m => m.kind === "header"));
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
          // A query on the Summary sheet ("نمو قوي" twice): its highlights must come back after leaving and returning.
          await viewer.goTo(1);
          summaryHits = (await viewer.search("نمو")).length;
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
        return { plainCount: plain.length, matchCount: matches.length, optedCount: opted.length, noController, highlighted, sameLayout, headerCopies, headerHighlights, footerCopies, footerHighlights, navigated, virtualized, occurrencesAgree, summaryHits, noHighlight, aborted, keptOnAbort, cleared, clearCount, newestWins, removed: !CSS.highlights.has(name) };
      };
    </script>`,
    width: 1200,
    height: 800,
  }, async (view) => {
    for (const [format, query] of [["pptx", "الانتاج"], ["xlsx", "الانتاج"], ["docx", "nile"]] as const) {
      const result = await view.evaluate(`probeResults(${JSON.stringify(format)}, ${JSON.stringify(query)})`) as {
        count: number;
        last: { slide: number | null; sheet: number | null; address: string | null };
        first: { sheet: number | null; address: string | null };
        listed: { index: number; active: number | null; cell: string | null };
        stale: { index: number; active: number | null; cell: string | null } | null;
        docx: {
          headerTop: { active: number | null; under: { header: boolean; text: string } | null };
          headerLater: { active: number | null; under: { header: boolean; text: string } | null };
          headerPast: number | null;
          bodyHit: { active: number | null; under: { header: boolean; text: string } | null; text: string };
          count: number;
        } | null;
        wrong: string;
        errors: string[];
      };
      expect(result.count).toBeGreaterThan(0);
      // A public call is refused through its promise, never a duplicate event.
      expect(result.wrong).toBe("invalid-argument");
      expect(result.errors).toEqual([]);
      if (format === "pptx") {
        expect(result.listed.index).toBe(result.last.slide!);
        expect(result.listed.active).toBe(result.count - 1);
      } else if (format === "xlsx") {
        expect(result.listed).toEqual({ index: result.last.sheet!, active: result.count - 1, cell: result.last.address });
        // The sheet switch cleared the sheet's search, so the result only moves.
        expect(result.stale).toEqual({ index: result.first.sheet!, active: null, cell: result.first.address });
      } else {
        expect(result.listed.index).toBe(0);
        const docx = result.docx!;
        // A header result becomes one painted copy: page 1's at the top of the
        // document, page 3's once the reader is there, and page 1's again once
        // the reader is past every page carrying that header.
        expect(docx.headerTop.active).not.toBeNull();
        expect(docx.headerTop.under?.header).toBe(true);
        expect(docx.headerLater.active!).toBeGreaterThan(docx.headerTop.active!);
        expect(docx.headerLater.under?.header).toBe(true);
        expect(docx.headerPast).toBe(docx.headerTop.active);
        // A body result outlines an occurrence in its own paragraph.
        expect(docx.bodyHit.active).not.toBeNull();
        expect(docx.bodyHit.under?.header).toBe(false);
        expect(docx.bodyHit.under?.text.toLowerCase()).toContain(docx.bodyHit.text.slice(0, 30).toLowerCase());
        expect(docx.count).toBeGreaterThan(2);
      }
    }
    for (const format of ["pptx", "xlsx", "docx"] as const) {
      const result = await view.evaluate(`probe(${JSON.stringify(format)})`) as {
        plainCount: number; matchCount: number; optedCount: number; noController: boolean; highlighted: string[]; sameLayout: boolean;
        headerCopies: number; headerHighlights: number; footerCopies: number; footerHighlights: number; navigated: boolean; virtualized: boolean; occurrencesAgree: boolean;
        summaryHits: number | null; noHighlight: boolean; aborted: string; keptOnAbort: boolean; cleared: unknown[]; clearCount: number; newestWins: boolean; removed: boolean;
      };
      expect(result.matchCount).toBeGreaterThan(0);
      expect(result.noController).toBe(true);
      expect(result.matchCount).toBe(result.plainCount);
      expect(result.highlighted.length).toBeGreaterThan(0);
      expect(result.sameLayout).toBe(true);
      // The return-to-sheet check needs a query that is on the sheet; zero hits would prove nothing.
      if (format === "xlsx") expect(result.summaryHits).toBeGreaterThan(0);
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
  });
}, 60000);
