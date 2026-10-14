import { expect } from "bun:test";
import { browserModule, visual, within, withVisualPage } from "./harness.js";

visual("manual-rendering site demos show paging failures, recover and remove their controls on cleanup", async () => {
  const modules = {
    "/pptx-demo.mjs": await browserModule(`${import.meta.dir}/../../../site/src/demos/headless-pptx.ts`),
    "/xlsx-demo.mjs": await browserModule(`${import.meta.dir}/../../../site/src/demos/headless-xlsx.ts`),
    "/worker.mjs": await browserModule(`${import.meta.dir}/../src/core/worker.ts`),
  };
  const files: Record<string, ArrayBuffer> = {};
  for (const format of ["pptx", "xlsx"]) {
    files[`/fixture.${format}`] = await Bun.file(`${import.meta.dir}/../../../fixtures/${format}-coverage.${format}`).arrayBuffer();
  }
  await withVisualPage({ modules, files, page: `<!doctype html>
    <main id="host" style="width:900px;height:600px"></main>
    <script type="module">
      import {runHeadlessPptx} from '/pptx-demo.mjs';
      import {runHeadlessXlsx} from '/xlsx-demo.mjs';
      const host = document.querySelector('#host');
      const post = Worker.prototype.postMessage;
      // Fault one real unit request, leaving the existing Office fixture and
      // normal parser/render path intact. No fake rendering model is supplied.
      Worker.prototype.postMessage = function(message, ...args) {
        if (window.failNext && message.type === 'unit' && message.index === 1) {
          window.failNext = false;
          queueMicrotask(() => this.dispatchEvent(new MessageEvent('message', {data:{
            type:'response',sessionId:message.sessionId,requestId:message.requestId,
            ok:false,error:{code:'invalid-document',message:'The next unit could not be read.'}
          }})));
          return;
        }
        return post.call(this, message, ...args);
      };
      window.mount = async format => {
        window.demo?.dispose(); host.replaceChildren();
        window.demo = await (format === 'pptx' ? runHeadlessPptx : runHeadlessXlsx)(host, '/fixture.' + format);
      };
      window.cleanup = () => {window.demo?.dispose();Worker.prototype.postMessage=post;};
    </script>` }, async view => {
    const wait = (condition: string) => within("site demo state", view.evaluate(`new Promise(resolve => {
      const check = () => {if (${condition}) resolve(true); else requestAnimationFrame(check)}; check();
    })`), 10000);
    try {
      for (const format of ["pptx", "xlsx"]) {
        await view.evaluate(`mount(${JSON.stringify(format)})`);
        const original = await view.evaluate<string>(`document.querySelector('#host').textContent`);
        expect(original.length).toBeGreaterThan(0);
        await view.evaluate(`(() => {window.failNext=true;[...document.querySelectorAll('.demo-controls button')].find(b=>b.textContent==='Next').click()})()`);
        await wait(`document.querySelector('.demo-controls [role="status"]').textContent.length > 0`);
        expect(await view.evaluate<string>(`document.querySelector('#host').textContent`)).toBe(original);
        expect(await view.evaluate<string>(`document.querySelector('.demo-controls [role="status"]').textContent`)).not.toBe("");
        await view.evaluate(`[...document.querySelectorAll('.demo-controls button')].find(b=>b.textContent==='Next').click()`);
        await wait(`document.querySelector('.demo-controls [role="status"]').textContent === '' && document.querySelector('.demo-controls-label').textContent.startsWith('2 /')`);
        expect(await view.evaluate<string>(`document.querySelector('#host').textContent`)).not.toBe(original);
        await view.evaluate(`window.demo.dispose()`);
        expect(await view.evaluate<number>(`document.querySelectorAll('.demo-controls').length`)).toBe(0);
      }
    } finally { await view.evaluate(`cleanup()`); }
  });
}, 60000);
