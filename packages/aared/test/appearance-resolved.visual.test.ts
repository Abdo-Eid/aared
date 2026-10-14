import { expect } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import type { SlideModel } from "../src/core/model.js";
import { serveDocx } from "./docx-harness.js";
import { browserModule, visual, withVisualPage } from "./harness.js";

const FIXTURES = `${import.meta.dir}/../../../fixtures`;

visual("visual DOM output: Word shade/tint paints native ink on Latin and Arabic text", async () => {
  const { server, sections } = await serveDocx(`${FIXTURES}/docx-coverage.docx`, `<!doctype html>
    <main id="host" style="width:960px"></main><script type="module">
    import { render } from '/docx.mjs';
    window.paint = model => render({document:model,target:document.querySelector('#host')});
    </script>`);
  await withVisualPage({ server, width: 1000, height: 800 }, async (view) => {
    await view.evaluate(`paint(${JSON.stringify(sections)})`);
    const output = await view.evaluate(`Array.from(document.querySelectorAll('p')).filter(p=>p.textContent.startsWith('Theme ink')).map(p=>({
      text:p.textContent, colors:[...p.querySelectorAll('span')].filter(s=>s.textContent.trim()).map(s=>getComputedStyle(s).color)}))`) as Array<{text:string;colors:string[]}>;
    expect(output).toHaveLength(5);
    const expected = ["rgb(21, 96, 130)", "rgb(15, 71, 97)", "rgb(69, 176, 225)", "rgb(10, 47, 64)", "rgb(193, 228, 245)"];
    for (const [index, p] of output.entries()) {
      expect(p.colors.length).toBeGreaterThanOrEqual(2);
      expect(p.colors.every(c => c === expected[index])).toBe(true);
    }
  });
}, 30000);

visual("visual DOM output: group paint and effects use native DOM/SVG composite filters", async () => {
  const pkg = await OpcPackage.open(await Bun.file(`${FIXTURES}/pptx-coverage.pptx`).arrayBuffer());
  const deck = await indexPptx(pkg);
  let model: SlideModel | undefined;
  for (let i = 0; i < deck.slideParts.length; i++) {
    const slide = await parsePptxSlide(pkg, deck, i);
    if (slide.elements.some(e => e.name === "Group fill inheritance")) { model = slide; break; }
  }
  expect(model).toBeDefined();
  const module = await browserModule(`${import.meta.dir}/../src/pptx/renderer.ts`);
  await withVisualPage({
    modules: { '/pptx.mjs': module },
    page: `<!doctype html><style>html,body{margin:0;overflow:hidden}#host{width:960px}</style><main id="host"></main>
      <script type="module">import {render} from '/pptx.mjs';
      window.paint = async (model,target=document.querySelector('#host')) => {
        window.handle?.destroy(); window.handle=await render({slide:model,target,zoom:'fit-width'});
      };</script>`,
    width: 960,
    height: 540,
  }, async (view) => {
    const before = JSON.stringify(model);
    await view.evaluate(`paint(${before})`);
    const groupNames = model!.elements.filter(e=>e.kind==='group').map(e=>({id:e.id,name:e.name}));
    const result = await view.evaluate(`(${JSON.stringify(groupNames)}).map(spec=>{
      const group=document.querySelector('[data-aared-group-id="'+spec.id+'"]');
      const filter=getComputedStyle(group).filter;
      const filterId=filter.match(/#([^"\)]+)/)?.[1];
      const def=filterId ? document.getElementById(filterId) : null;
      return {name:spec.name, filter, primitives:def?[...def.children].map(e=>e.localName):[],
        reflection:getComputedStyle(group).getPropertyValue('-webkit-box-reflect'),
        ink:[...group.querySelectorAll('[data-aared-element="paragraph-text"] span')].map(r=>getComputedStyle(r).color),
        fills:[...group.querySelectorAll('[data-aared-element="shape-geometry"] path')].map(p=>p.getAttribute('fill')),
        gradients:[...group.querySelectorAll('linearGradient')].map(g=>({units:g.getAttribute('gradientUnits'),end:g.getAttribute('y2'),mapping:g.getAttribute('gradientTransform')})),
        childFilters:[...group.querySelectorAll('[data-aared-element="shape"]')].map(p=>getComputedStyle(p).filter),
        fallback:group.querySelectorAll('[data-aared-fill-fallback]').length};})`) as Array<{name:string;filter:string;primitives:string[];reflection:string;fills:string[];ink:string[];childFilters:string[];fallback:number;gradients:Array<{units:string;end:string;mapping:string|null}>}>;
    const byName = (name:string) => result.find(r=>r.name===name)!;
    expect(byName('Group fill inheritance').fills).toEqual(['#156082','#BE2832']);
    expect(byName('Group fill inheritance').ink).toContain('rgba(0, 0, 0, 0)');
    expect(byName('Group nested outer').fills.filter(f=>f.startsWith('url('))).toHaveLength(2);
    expect(byName('Group nested outer').fills).toContain('none');
    const gradients = byName('Group nested outer').gradients;
    expect(gradients.map(g=>g.units)).toEqual(['userSpaceOnUse','userSpaceOnUse']);
    expect(gradients[0]!.end).toBe(gradients[1]!.end);
    expect(gradients[0]!.mapping).not.toBe(gradients[1]!.mapping);
    expect(byName('Group outer shadow').filter).toContain('drop-shadow');
    expect(byName('Group outer shadow').primitives).toEqual([]);
    expect(byName('Group outer shadow').childFilters.every(f=>f==='none')).toBe(true);
    expect(byName('Group glow').primitives).toContain('feComponentTransfer');
    expect(byName('Group soft edge').primitives).toContain('feComposite');
    expect(byName('Group reflection').reflection).toContain('below');
    expect(byName('Group inner shadow').childFilters.every(f=>f.includes('aared-effect-'))).toBe(true);
    for (const output of result) expect(output.fallback).toBe(0);
    expect(JSON.stringify(model)).toBe(before);
    expect(await view.evaluate("document.querySelectorAll('canvas').length") as number).toBe(0);
    // Review artifact paired against the native export from the ground-truth probe.
    if (process.env.AARED_ARTIFACTS_DIR !== undefined) {
      await Bun.write(`${process.env.AARED_ARTIFACTS_DIR}/group-appearance-aared.png`, await view.screenshot({ encoding: "buffer", format: "png" }));
    }
    const fallback = await view.evaluate(`(async()=>{
      const supports=CSS.supports.bind(CSS);
      window.originalSupports=supports;
      CSS.supports=(...args)=>args[0]==='-webkit-box-reflect'?false:supports(...args);
      await paint(${before});
      await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
      const mirror=document.querySelector('[data-aared-element="reflection"]');
      const ids=[...document.querySelectorAll('[id]')].map(e=>e.id);
      const result={found:!!mirror,hidden:mirror?.getAttribute('aria-hidden'),inert:mirror?.inert,
        text:mirror?.textContent,mask:mirror?.firstElementChild.style.maskImage,unique:new Set(ids).size===ids.length};
      return result;})()`) as {found:boolean;hidden:string;inert:boolean;text:string;mask:string;unique:boolean};
    expect(fallback).toMatchObject({found:true,hidden:"true",inert:true,unique:true});
    expect(fallback.text).toContain("Group");
    expect(fallback.mask).toContain("linear-gradient");
    if (process.env.AARED_ARTIFACTS_DIR !== undefined) {
      await Bun.write(`${process.env.AARED_ARTIFACTS_DIR}/group-reflection-fallback.png`, await view.screenshot({encoding:"buffer",format:"png"}));
    }
    const lifecycle = await view.evaluate(`(async()=>{
      handle.destroy(); const result={removed:!document.querySelector('[data-aared-element="reflection"]')};
      CSS.supports=window.originalSupports;
      const root=document.querySelector('#host').attachShadow({mode:'open'});
      const target=document.createElement('main');target.style.width='960px';root.append(target);
      await paint(${before},target);
      result.shadowFilters=[...root.querySelectorAll('[data-aared-element="group"]')]
        .filter(g=>getComputedStyle(g).filter!=='none').length;
      handle.destroy();return result;})()`) as {removed:boolean;shadowFilters:number};
    expect(lifecycle.removed).toBe(true);
    expect(lifecycle.shadowFilters).toBeGreaterThanOrEqual(3);
  });
}, 30000);
