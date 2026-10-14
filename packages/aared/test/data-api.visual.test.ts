import { expect } from "bun:test";
import { browserModule, visual, withVisualPage } from "./harness.js";

visual("public handle rendering preserves painter output and ignores mutations to semantic data", async () => {
  const formats = ["pptx", "xlsx", "docx"] as const;
  const modules = {
    "/api.mjs": await browserModule(`${import.meta.dir}/data-api-harness.ts`),
    "/worker.mjs": await browserModule(`${import.meta.dir}/../src/core/worker.ts`),
  };
  const files: Record<string, ArrayBuffer> = {};
  for (const format of formats) files[`/fixture.${format}`] = await Bun.file(`${import.meta.dir}/../../../fixtures/${format}-coverage.${format}`).arrayBuffer();
  await withVisualPage({ modules, files, width: 1000, height: 800, page: `<!doctype html>
    <main id="host" style="width:900px;height:650px"></main>
    <script type="module">
    import * as api from '/api.mjs';
    const host = document.querySelector('#host');
    const summarize = () => {
      // Each paint intentionally owns distinct SVG ids. Compare the reference
      // relationships in encounter order, not those per-render identifiers.
      const resources = new Map();
      return [...host.querySelectorAll('[data-aared-element]')].map(el => ({
      kind:el.dataset.aaredElement, text:el.children.length ? '' : el.textContent,
      style:el.getAttribute('style')?.replace(/#aared-[a-zA-Z0-9_-]+/g, id => {
        if (!resources.has(id)) resources.set(id, '#resource-' + resources.size);
        return resources.get(id);
      }) ?? null, direction:getComputedStyle(el).direction,
      width:el.getBoundingClientRect().width, height:el.getBoundingClientRect().height,
    })); };
    const imagesReady = () => Promise.all([...host.querySelectorAll('img')].map(img => img.complete && img.src
      ? undefined : new Promise(resolve => {img.addEventListener('load',resolve,{once:true});img.addEventListener('error',resolve,{once:true})})));
    window.run = async format => {
      const suffix = format[0].toUpperCase() + format.slice(1);
      const doc = await api['open' + suffix]('/fixture.' + format, {worker:{url:'/worker.mjs'}});
      let view;
      try {
        const unit = format === 'pptx' ? doc.slides.get(0) : format === 'xlsx' ? doc.sheets.get(0) : doc;
        const field = format === 'pptx' ? 'slide' : format === 'xlsx' ? 'sheet' : 'document';
        const model = format === 'pptx' ? await api.readSlideModel(unit) : format === 'xlsx' ? await api.readSheetModel(unit) : api.sectionsOf(unit);
        const data = format === 'docx' ? doc.sections.list() : await unit.read();
        const invalid = [];
        for (const candidate of [model, data, {...unit}, Object.create(unit)]) {
          invalid.push(await api['render' + suffix]({[field]:candidate,target:host}).then(()=>'resolved',error=>error.code));
        }
        const mediaResolver = id => doc.media.get(id);
        view = await api['paint' + suffix]({[field]:model,target:host,zoom:1,mediaResolver});
        await imagesReady();
        const original = summarize();
        view.destroy(); host.replaceChildren();
        if (format === 'pptx') data.paragraphs.splice(0, data.paragraphs.length, 'changed');
        else if (format === 'xlsx') { for (const row of data.rows) for (const entry of row.cells) entry.cell.formatted = 'changed'; data.merges.length = 0; }
        else data[0].page.width = 999;
        view = await api['render' + suffix]({[field]:unit,target:host,zoom:1,mediaResolver});
        await imagesReady();
        const actual = summarize();
        const result = {original,actual,invalid,text:host.textContent};
        if (format === 'xlsx') {
          let event;
          view.addEventListener('cellactivate', e => {event=e.detail});
          const td = host.querySelector('td[data-aared-element="cell"]');
          td.dispatchEvent(new MouseEvent('dblclick',{bubbles:true}));
          result.activation=event;
          event.cell.formatted='changed-event';
          result.isolated=(await unit.read()).rows[0].cells[0].cell.formatted !== 'changed-event';
        }
        return result;
      } finally {view?.destroy();host.replaceChildren();doc.close()}
    };
    </script>` }, async view => {
    for (const format of formats) {
      const output = await view.evaluate<Record<string, unknown>>(`run(${JSON.stringify(format)})`);
      expect(output.invalid).toEqual(["invalid-argument", "invalid-argument", "invalid-argument", "invalid-argument"]);
      expect(output.actual).toEqual(output.original);
      expect(String(output.text).length).toBeGreaterThan(0);
      expect(String(output.text)).not.toContain("changed");
      if (format === "xlsx") {
        expect(output.activation).toHaveProperty("cell.formatted");
        expect(output.activation).not.toHaveProperty("model");
        expect(output.activation).not.toHaveProperty("cell.style");
        expect(output.isolated).toBe(true);
      }
    }
  });
}, 90000);
