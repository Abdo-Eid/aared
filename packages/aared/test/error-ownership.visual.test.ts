import { expect } from "bun:test";
import { browserModule, visual, withVisualPage } from "./harness.js";

visual("public operations reject only; built-in tabs and element controls report once without losing content", async () => {
  const [preview, element, xlsx, worker] = await Promise.all([
    browserModule(`${import.meta.dir}/../src/preview/index.ts`),
    browserModule(`${import.meta.dir}/../src/element/index.ts`),
    browserModule(`${import.meta.dir}/../src/xlsx/renderer.ts`),
    browserModule(`${import.meta.dir}/../src/core/worker.ts`),
  ]);
  const formats = ["pptx", "xlsx", "docx"] as const;
  const files = new Map(await Promise.all(formats.map(async (format) => [format,
    await Bun.file(`${import.meta.dir}/../../../fixtures/${format}-coverage.${format}`).arrayBuffer()] as const)));
  const scripts: Record<string, string> = { "/preview.mjs": preview, "/element.mjs": element, "/xlsx.mjs": xlsx, "/worker.mjs": worker };
  const served: Record<string, ArrayBuffer> = {};
  for (const format of formats) served[`/${format}`] = files.get(format)!;
  await withVisualPage({
    modules: scripts,
    files: served,
    page: `<!doctype html><main id="host" style="width:1000px;height:600px"></main>
      <aared-viewer id="el" controls style="height:600px"></aared-viewer><script type="module">
      import { preview } from '/preview.mjs';
      import { open, render } from '/xlsx.mjs';
      import '/element.mjs';
      window.unhandled = [];
      window.addEventListener('unhandledrejection', (event) => window.unhandled.push(String(event.reason)));
      const host = document.querySelector('#host');
      const workers = [];
      const worker = { create: () => { const worker = new Worker('/worker.mjs', {type:'module'}); workers.push(worker); return worker; } };
      const failWorker = () => {
        const worker = workers.at(-1);
        worker.terminate();
        worker.dispatchEvent(new ErrorEvent('error', {message:'test Worker failure'}));
      };
      const errors = (target) => { const out = []; target.addEventListener('error', e => out.push(e.detail.cause.code)); return out; };
      const code = (promise) => promise.then(() => 'resolved', e => e.code);
      const settle = () => new Promise(resolve => setTimeout(resolve, 150));
      window.run = async () => {
        const calls = [];
        for (const format of ['pptx','xlsx','docx']) {
          const viewer = await preview({source:'/' + format, target:host, worker});
          const reported = errors(viewer);
          const before = host.textContent;
          const badIndex = await code(viewer.goTo(999));
          const badSearch = await code(viewer.search('nile', {signal:AbortSignal.abort()}));
          calls.push({format, badIndex, badSearch, reported, kept:host.textContent === before});
          viewer.destroy();
        }
        const tabs = [];
        for (const lowLevel of [false,true]) {
          const doc = lowLevel ? await open('/xlsx') : null;
          const handle = lowLevel ? await render({sheet:doc.sheets.get(0),target:host,appearance:{spreadsheet:{sheetTabs:true}}})
            : await preview({source:'/xlsx',target:host,worker});
          const reported = errors(handle);
          const before = host.textContent;
          if (doc !== null) doc.close(); else failWorker();
          const direct = await code(lowLevel ? handle.goToSheet(1) : handle.goTo(1));
          const afterDirect = reported.slice();
          host.querySelectorAll('[data-aared-element="sheet-tab"]')[1].click();
          await settle();
          tabs.push({direct, afterDirect, reported, kept:host.textContent === before});
          handle.destroy();
        }
        const el = document.querySelector('#el');
        el.worker = worker;
        const loaded = new Promise(resolve => el.addEventListener('load', resolve, {once:true}));
        el.src = '/pptx'; await loaded;
        const reported = errors(el);
        failWorker();
        const before = el.shadowRoot.querySelector('[part="page"]').textContent;
        el.shadowRoot.querySelectorAll('[part~="button"]')[1].click();
        await settle();
        const controls = {reported:reported.slice(), kept:el.shadowRoot.querySelector('[part="page"]').textContent === before};
        const sheetLoaded = new Promise(resolve => el.addEventListener('load', resolve, {once:true}));
        el.src = '/xlsx'; await sheetLoaded;
        // The element's responsive grid fills its viewport after load. Let
        // that layout settle before comparing content across a failed action.
        await settle();
        failWorker();
        const sheetBefore = el.shadowRoot.querySelector('[part="page"]').textContent;
        el.shadowRoot.querySelectorAll('[data-aared-element="sheet-tab"]')[1].click();
        await settle();
        const elementTabs = {reported:reported.slice(), kept:el.shadowRoot.querySelector('[part="page"]').textContent === sheetBefore};
        el.removeAttribute('src');
        const explicitOpen = await code(el.open());
        await settle();
        return {calls,tabs,controls,elementTabs,explicitOpen,reported,unhandled:window.unhandled};
      };
      </script>`,
    width: 1200,
    height: 800,
  }, async (view) => {
    const result = await view.evaluate("run()") as {
      calls: Array<{ format: string; badIndex: string; badSearch: string; reported: string[]; kept: boolean }>;
      tabs: Array<{ direct: string; afterDirect: string[]; reported: string[]; kept: boolean }>;
      controls: { reported: string[]; kept: boolean }; explicitOpen: string; reported: string[]; unhandled: string[];
      elementTabs: { reported: string[]; kept: boolean };
    };
    for (const call of result.calls) expect(call).toEqual({ format: call.format, badIndex: "not-found", badSearch: "aborted", reported: [], kept: true });
    expect(result.tabs).toEqual([
      { direct: "worker-error", afterDirect: [], reported: ["worker-error"], kept: true },
      { direct: "closed", afterDirect: [], reported: ["closed"], kept: true },
    ]);
    expect(result.controls).toEqual({ reported: ["worker-error"], kept: true });
    expect(result.elementTabs).toEqual({ reported: ["worker-error", "worker-error"], kept: true });
    expect(result.explicitOpen).toBe("invalid-state");
    expect(result.reported).toEqual(["worker-error", "worker-error"]);
    expect(result.unhandled).toEqual([]);
  });
}, 60000);
