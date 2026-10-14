import { expect } from "bun:test";
import { browserModule, visual, withVisualPage } from "./harness.js";

visual("a viewer lends restricted content: reads, media and secondary renders without lifecycle access or a second parse", async () => {
  const formats = ["pptx", "xlsx", "docx"] as const;
  const entry = await browserModule(`${import.meta.dir}/preview-content-harness.ts`);
  const fixtures = new Map(await Promise.all(formats.map(async (format) => [format,
    await Bun.file(`${import.meta.dir}/../../../fixtures/${format}-coverage.${format}`).arrayBuffer()] as const)));
  const worker = await browserModule(`${import.meta.dir}/../src/core/worker.ts`);
  const modules: Record<string, string> = { "/worker.mjs": worker, "/content.mjs": entry };
  const files: Record<string, ArrayBuffer> = {};
  for (const format of formats) files[`/fixture.${format}`] = fixtures.get(format)!;
  await withVisualPage({
    modules,
    files,
    page: `<!doctype html>
      <main id="main" style="width:900px;height:500px"></main>
      <div id="strip" style="display:flex;gap:8px"></div>
      <script type="module">
      import '/content.mjs'; // registers the element in the same consumer module graph
      const main = document.querySelector('#main');
      const strip = document.querySelector('#strip');
      window.runElement = async (format, replacementFormat = format) => {
        const api = await import('/content.mjs');
        const render = api[format + 'Render'];
        const el = window.document.createElement('aared-viewer');
        el.style.height = '500px'; window.document.body.append(el);
        const loaded = () => new Promise(resolve => el.addEventListener('load', resolve, {once:true}));
        let load = loaded(); el.src = '/fixture.' + format; await load;
        const content = el.viewer.content;
        const saved = format === 'pptx' ? content.slides.get(0) : format === 'xlsx' ? content.sheets.get(0) : content;
        const input = format === 'pptx' ? {slide:saved} : format === 'xlsx' ? {sheet:saved} : {document:saved};
        const read = () => format === 'docx' ? saved.sections.list() : saved.read();
        const target = window.document.body.appendChild(window.document.createElement('div'));
        target.style.cssText = 'width:160px;height:90px';
        const thumb = await render({...input,target});
        const marker = format === 'pptx' ? 'slide' : format === 'xlsx' ? 'grid' : 'document';
        const painted = target.querySelector('[data-aared-element="' + marker + '"]') !== null;
        thumb.destroy(); target.textContent = 'untouched';
        load = loaded(); el.src = '/fixture.' + replacementFormat + '?replacement'; await load;
        const rejected = await Promise.resolve().then(read).then(() => 'resolved', error => error.code);
        const staleRender = await render({...input,target}).then(() => 'resolved', error => error.code);
        const clean = !('document' in el.viewer) && !('close' in content);
        const next = el.viewer.content;
        const removedRead = () => replacementFormat === 'docx' ? next.sections.list() : (replacementFormat === 'pptx' ? next.slides : next.sheets).get(0).read();
        el.remove();
        const removed = await Promise.resolve().then(removedRead).then(() => 'resolved', error => error.code);
        const result = {painted,rejected,staleRender,clean,removed,kept:target.textContent === 'untouched'};
        target.remove(); return result;
      };
      window.run = async (format, auto) => {
        const api = await import('/content.mjs');
        const preview = auto ? api.autoPreview : api[format + 'Preview'];
        const render = api[format + 'Render'];
        const requests = [];
        const viewer = await preview({ source: '/fixture.' + format, target: main, worker: { create: () => {
          const worker = new Worker('/worker.mjs', { type: 'module' });
          const post = worker.postMessage.bind(worker);
          worker.postMessage = (message, transfers) => { requests.push(message); post(message, transfers); };
          return worker;
        } } });
        const content = viewer.content;
        const views = [];
        const code = (action) => Promise.resolve().then(action).then(() => 'resolved', error => error.code);
        const facadeObjects = [content, content.media];
        const result = { same: content === viewer.content, count: content.info.count === viewer.info.count,
          documentAbsent: !('document' in viewer), routeAbsent: !('route' in viewer), searchOnViewer: typeof viewer.search === 'function',
        };
        let savedInput;
        if (format === 'pptx') {
          for (let index = 0; index < 3; index++) {
            const target = strip.appendChild(window.document.createElement('div'));
            target.style.cssText = 'width:160px;height:90px';
            views.push(await render({ slide: content.slides.get(index), target }));
          }
          result.thumbnails = [...strip.children].filter((thumb) => thumb.querySelector('[data-aared-element="slide"]') !== null).length;
          await viewer.goTo(2);
          result.mainMoved = viewer.index === 2;
          savedInput = content.slides.get(0);
          const model = await savedInput.read();
           result.read = model.paragraphs.length > 0 && !('elements' in model);
          result.stable = savedInput === content.slides.list()[0];
          facadeObjects.push(content.slides, ...content.slides.list());
          const pictureTarget = strip.appendChild(window.document.createElement('div'));
          pictureTarget.style.cssText = 'width:160px;height:90px';
          views.push(await render({slide:content.slides.get(content.slides.list().length - 1),target:pictureTarget}));
        } else if (format === 'xlsx') {
          const data = await content.sheets.get(1).read({ range: 'A1:D8' });
          result.rows = data.rows.length;
          let activation;
          viewer.addEventListener('cellactivate', event => {activation = event.detail});
          const activated = main.querySelector('td[data-aared-element="cell"]');
          activated.dispatchEvent(new MouseEvent('dblclick', {bubbles:true}));
          result.activation = structuredClone(activation);
          activation.cell.formatted = 'changed-activation';
          activated.dispatchEvent(new MouseEvent('dblclick', {bubbles:true}));
          result.activationIsolated = activation.cell.formatted !== 'changed-activation';
          const target = strip.appendChild(window.document.createElement('div'));
          target.style.cssText = 'width:400px;height:200px';
          views.push(await render({ sheet: content.sheets.get(0), target, appearance: {spreadsheet: {sheetTabs: true, fill:false}} }));
          result.secondView = target.querySelectorAll('td[data-aared-element="cell"]').length > 0;
          await views[0].goToSheet(1);
          result.secondaryNavigation = views[0].sheet.index === 1 && viewer.index === 0;
          await views[0].goToSheet(content.sheets.get('Notes').info.index); // The fixture's SVG picture.
          savedInput = content.sheets.get(0);
          result.stable = savedInput === content.sheets.get(savedInput.info.title);
          facadeObjects.push(content.sheets, ...content.sheets.list());
        } else {
          result.headings = content.headings.list().length === viewer.docx.headings.length && viewer.docx.headings.length > 0;
          const target = strip.appendChild(window.document.createElement('div'));
          target.style.cssText = 'width:400px;height:200px;overflow:auto';
          views.push(await render({ document: content, target, pagination: {mode:'layout'} }));
          result.secondView = target.textContent.length > 0;
          await views[0].goToHeading(0);
          await views[0].goToPage(views[0].pageCount - 1);
          result.secondaryNavigation = views[0].pageCount > 1;
          savedInput = content;
          facadeObjects.push(content.sections, content.headings);
        }
        // The facades expose closures and data, never a backing owner, session
        // or internal renderer method. There is no hidden symbol carrying one.
        result.restricted = facadeObjects.every(value => Object.isFrozen(value)
          && Object.getOwnPropertySymbols(value).length === 0
          && ['close','destroy','session','owner','sheetDocument','outline','search','sharedMedia'].every(key => !(key in value)));
        // Picture media resolves after render (render waits for fonts/effects,
        // not image downloads). Wait for that source to be attached first.
        for (let attempt = 0; attempt < 200 && !strip.querySelector('img[src]'); attempt++)
          await new Promise(resolve => setTimeout(resolve, 10));
        const images = [...main.querySelectorAll('img'), ...strip.querySelectorAll('img')];
        await Promise.all(images.map(image => image.complete ? undefined : new Promise(resolve => {
          image.addEventListener('load', resolve, {once:true}); image.addEventListener('error', resolve, {once:true});
        })));
        result.images = images.filter(image => image.naturalWidth > 0).length;
        const mediaRequests = requests.filter(message => message.type === 'media');
        // Inspect the bytes via the same media capability, not the hidden session.
        if (mediaRequests.length > 0) result.mediaBytes = (await content.media.get(mediaRequests[0].mediaId)).byteLength;
        result.opens = requests.filter((message) => message.type === 'open').length;
        for (const view of views) view.destroy();
        strip.replaceChildren();
        const pendingTarget = window.document.createElement('div');
        pendingTarget.textContent = 'untouched';
        window.document.body.append(pendingTarget);
        const input = format === 'pptx' ? {slide:savedInput} : format === 'xlsx' ? {sheet:savedInput} : {document:savedInput};
        const pending = render({...input, target:pendingTarget});
        const rejected = pending.then(() => 'resolved', error => error.code);
        viewer.destroy();
        result.pending = await rejected;
        result.pendingKeptTarget = pendingTarget.textContent === 'untouched';
        const read = format === 'docx' ? () => content.sections.list() : () => savedInput.read();
        result.afterDestroy = await code(read);
        const collection = format === 'pptx' ? content.slides : format === 'xlsx' ? content.sheets : content.sections;
        result.listAfterDestroy = await code(() => collection.list());
        result.getAfterDestroy = await code(() => collection.get(0));
        if (format === 'docx') result.headingsAfterDestroy = await code(() => content.headings.list());
        result.mediaAfterDestroy = await code(() => content.media.get('missing'));
        result.renderAfterDestroy = await code(() => render({...input,target:pendingTarget}));
        result.staleKeptTarget = pendingTarget.textContent === 'untouched';
        pendingTarget.remove();
        return result;
      };
    </script>`,
    width: 1200,
    height: 900,
  }, async (view) => {
    for (const auto of [false, true]) for (const format of formats) {
      const result = await view.evaluate(`run(${JSON.stringify(format)}, ${auto})`) as Record<string, unknown>;
      expect(result.same).toBe(true);
      expect(result.count).toBe(true);
      expect(result.documentAbsent).toBe(true);
      expect(result.routeAbsent).toBe(true);
      expect(result.searchOnViewer).toBe(true);
      expect(result.restricted).toBe(true);
      // One open: the thumbnails, reads and second view use the viewer's own parse.
      expect(result.opens).toBe(1);
      expect(result.afterDestroy).toBe("closed");
      expect(result.listAfterDestroy).toBe("closed");
      expect(result.getAfterDestroy).toBe("closed");
      expect(result.mediaAfterDestroy).toBe("closed");
      expect(result.renderAfterDestroy).toBe("closed");
      expect(result.pending).toBe("aborted");
      expect(result.pendingKeptTarget).toBe(true);
      expect(result.staleKeptTarget).toBe(true);
      expect(result.images, `${format} (${auto ? "auto" : "local"})`).toBeGreaterThan(0);
      expect(result.mediaBytes).toBeGreaterThan(0);
      if (format === "pptx") {
        expect(result.thumbnails).toBe(3);
        expect(result.mainMoved).toBe(true);
        expect(result.read).toBe(true);
        expect(result.stable).toBe(true);
      } else if (format === "xlsx") {
        expect(result.rows).toBeGreaterThan(0);
        expect(result.activation).toHaveProperty("cell.formatted");
        expect(result.activation).not.toHaveProperty("model");
        expect(result.activation).not.toHaveProperty("cell.style");
        expect(result.activation).not.toHaveProperty("cell.conditional");
        expect(result.activationIsolated).toBe(true);
        expect(result.secondView).toBe(true);
        expect(result.secondaryNavigation).toBe(true);
        expect(result.stable).toBe(true);
      } else {
        expect(result.headingsAfterDestroy).toBe("closed");
        expect(result.headings).toBe(true);
        expect(result.secondView).toBe(true);
        expect(result.secondaryNavigation).toBe(true);
      }
    }
    for (const format of formats) expect(await view.evaluate(`runElement(${JSON.stringify(format)})`) as Record<string, unknown>).toEqual({
      painted: true, rejected: 'closed', staleRender: 'closed', clean: true, removed: 'closed', kept: true,
    });
    // Keep the original cross-format case as well as same-format replacement.
    expect(await view.evaluate(`runElement("pptx", "xlsx")`) as Record<string, unknown>).toEqual({
      painted: true, rejected: 'closed', staleRender: 'closed', clean: true, removed: 'closed', kept: true,
    });
  });
}, 60000);
