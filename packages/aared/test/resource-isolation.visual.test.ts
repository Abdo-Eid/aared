import { expect } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { findCoverageSlide } from "./fixture.js";
import { browserModule, visual, within, withVisualPage } from "./harness.js";

interface ResourceOutput {
  definitions: number;
  references: number;
  own: boolean;
  unique: boolean;
  paint: string[];
}

visual("paint resources and search names belong to their render, including independent bundle copies", async () => {
  const pkg = await OpcPackage.open(await Bun.file(`${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`).arrayBuffer());
  const deck = await indexPptx(pkg);
  const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
  const models = await Promise.all(["gradients", "wordart", "table-fills"].map(async (key) => (await findCoverageSlide(slides, key)).slide));
  const groups = slides.find((slide) => slide.elements.some((element) => element.name === "Group fill inheritance"));
  expect(groups).toBeDefined();
  models.push(groups!);
  // A real chart exercises shared description/fill resource names as well.
  models.push((await findCoverageSlide(slides, "waterfall")).slide);
  const module = await browserModule(`${import.meta.dir}/resource-isolation-harness.ts`);
  await withVisualPage({
    modules: { "/one.mjs": module, "/two.mjs": module },
    page: `<!doctype html><style>html,body{margin:0}.host{width:960px}</style>
      <main id="one" class="host"></main><main id="two" class="host"></main>
      <script type="module">
      import * as one from '/one.mjs'; import * as two from '/two.mjs';
      const hosts = [document.querySelector('#one'), document.querySelector('#two')];
      const copies = [one, two];
      window.handles = []; window.highlights = [];
      window.paint = async (model, independent) => {
        for (const highlighter of highlights) highlighter.destroy(); highlights = [];
        for (const handle of handles) handle.destroy(); handles = [];
        for (const [index, host] of hosts.entries()) {
          handles.push(await (independent ? copies[index] : one).render({slide:model,target:host,zoom:1}));
        }
      };
      window.inspect = index => {
        const host = hosts[index];
        const ids = [...document.querySelectorAll('svg [id]')].map(node => node.id);
        let references = 0, own = true;
        for (const node of host.querySelectorAll('*')) {
          const css = getComputedStyle(node);
          for (const value of [node.getAttribute('fill'), node.getAttribute('stroke'), css.filter, css.clipPath]) {
            for (const match of (value ?? '').matchAll(/url\\(["']?#([^"')]+)["']?\\)/g)) {
              references++; own &&= host.contains(document.getElementById(match[1]));
            }
          }
          const description = node.getAttribute('aria-describedby');
          if (description) { references++; own &&= host.contains(document.getElementById(description)); }
        }
        return { definitions:host.querySelectorAll('svg [id]').length, references, own,
          unique:new Set(ids).size === ids.length,
          paint:[...host.querySelectorAll('linearGradient,radialGradient,pattern,filter,desc')].map(node=>node.outerHTML) };
      };
      window.search = () => {
        for (const [index, host] of hosts.entries()) {
          const highlighter = copies[index].createSearchHighlighter(host);
          const paragraph = [...host.querySelectorAll('p')].find(node => node.textContent.trim() && !node.closest('[hidden],[aria-hidden="true"]'));
          const text = paragraph.querySelector('[data-aared-element="paragraph-text"]') ?? paragraph;
          highlighter.setQuery(text.textContent.trim().split(/\\s+/)[0]); highlights.push(highlighter);
        }
        return hosts.map((host,index)=>({name:host.dataset.aaredSearchHighlight,count:highlights[index].count}));
      };
      window.ready = true;
      </script>`,
  }, async (view) => {
    for (const model of models) {
      // One package copy is sufficient to reproduce the old path-based fill
      // collision; independent evaluations additionally catch module counters.
      for (const independent of [false, true]) {
        await within("paint sibling resources", view.evaluate(`paint(${JSON.stringify(model)},${independent})`), 15_000);
        const before = await view.evaluate("inspect(1)") as ResourceOutput;
        expect(before.definitions).toBeGreaterThan(0);
        expect(before.references).toBeGreaterThan(0);
        expect(before.own).toBe(true);
        expect(before.unique).toBe(true);
        await view.evaluate("handles[0].destroy()");
        const after = await view.evaluate("inspect(1)") as ResourceOutput;
        expect(after).toEqual(before);
      }
    }
    await view.evaluate(`paint(${JSON.stringify(models[0])},true)`);
    const search = await view.evaluate("search()") as Array<{ name: string; count: number }>;
    expect(search[0]!.name).not.toBe(search[1]!.name);
    expect(search.every((item) => item.count > 0)).toBe(true);
    const lifecycle = await view.evaluate(`(() => {
      const [first, second] = [...document.querySelectorAll('.host')].map(host => host.dataset.aaredSearchHighlight);
      const hasRule = name => document.adoptedStyleSheets.some(sheet => [...sheet.cssRules].some(rule => rule.cssText.includes(name)));
      document.adoptedStyleSheets = document.adoptedStyleSheets.filter(sheet => ![...sheet.cssRules].some(rule => rule.cssText.includes(second)));
      const removed = !hasRule(second);
      highlights[1].activate(0);
      const restored = hasRule(second);
      highlights[0].destroy(); handles[0].destroy();
      const independent = !CSS.highlights.has(first) && CSS.highlights.has(second) && hasRule(second);
      highlights[1].destroy(); handles[1].destroy();
      return {removed,restored,independent,clean:CSS.highlights.size === 0 && document.adoptedStyleSheets.length === 0};
    })()`) as Record<string, boolean>;
    expect(lifecycle).toEqual({ removed: true, restored: true, independent: true, clean: true });
  });
}, 90_000);
