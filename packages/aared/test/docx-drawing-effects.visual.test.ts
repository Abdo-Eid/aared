import { expect } from "bun:test";
import { visual, withVisualPage } from "./harness.js";
import { serveDocx } from "./docx-harness.js";

const FIXTURE = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

visual("visual DOM output: Word pictures and text-box surfaces share the PPTX effect painter without changing layout", async () => {
  const { server, sections } = await serveDocx(FIXTURE, `<!doctype html><style>html,body{margin:0}*{box-sizing:border-box}img{display:block;max-width:100%}#host{width:816px}</style>
    <main id="host"></main><script type="module">
    import {render} from '/docx.mjs';
    window.paint=async model=>{window.handle?.destroy();window.handle=await render({document:model,target:document.querySelector('#host'),
      pagination:{mode:'layout'},zoom:1,mediaResolver:async id=>(await fetch('/media/'+id.split('/').at(-1))).arrayBuffer()});
      await document.fonts.ready;await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
      const deadline=Date.now()+5000;while([...document.querySelectorAll('img')].some(i=>!i.complete||!i.naturalWidth)){
        if(Date.now()>deadline)throw new Error('media timeout');await new Promise(resolve=>setTimeout(resolve,25));}
      return {pages:handle.pageCount,heights:[...document.querySelectorAll('[data-aared-element="page"]')].map(p=>p.getBoundingClientRect().height)};
    };
    window.read=()=>{
      const boxes=[...document.querySelectorAll('[data-aared-text-box-frame]')].filter(b=>b.textContent.startsWith('Word effect '));
      return {boxes:boxes.map(b=>{const s=b.querySelector('[data-aared-element="text-box-surface"]');return {text:b.textContent,
        filter:getComputedStyle(s).filter,reflection:getComputedStyle(s).getPropertyValue('-webkit-box-reflect'),
        frameFilter:getComputedStyle(b).filter,transform:getComputedStyle(s).transform,width:b.clientWidth,height:b.clientHeight};}),
        pictures:[...document.querySelectorAll('img')].filter(i=>i.alt.startsWith('Word effect ')).map(i=>({alt:i.alt,filter:getComputedStyle(i.parentElement).filter,
          inline:!!i.closest('[data-aared-element="inline-image"]'),decoded:i.complete&&i.naturalWidth>0})),
        canvas:document.querySelectorAll('canvas').length};
    };
    </script>`);
  await withVisualPage({ server, width: 816, height: 1056 }, async (view) => {
    const before = JSON.stringify(sections);
    const layout = await view.evaluate(`paint(${before})`) as {pages:number;heights:number[]};
    const output = await view.evaluate('read()') as {boxes:Array<{text:string;filter:string;reflection:string;frameFilter:string;transform:string;width:number;height:number}>;
      pictures:Array<{alt:string;filter:string;inline:boolean;decoded:boolean}>;canvas:number};
    expect(output.boxes).toHaveLength(6);
    const byName=(name:string)=>output.boxes.find(b=>b.text.startsWith(`Word effect ${name}`))!;
    expect(byName('outer shadow').filter).toContain('drop-shadow');
    expect(byName('inner shadow').filter).toContain('aared-effect-');
    expect(byName('glow').filter).toContain('aared-effect-');
    expect(byName('soft edge').filter).toContain('aared-effect-');
    expect(byName('reflection').reflection).toContain('below');
    expect(byName('rotated shadow').transform).not.toBe('none');
    for(const box of output.boxes){expect(box.frameFilter).toBe('none');expect(box.width).toBeGreaterThan(200);}
    expect(output.pictures).toHaveLength(3);
    expect(output.pictures.every(p=>p.decoded&&p.filter!=='none')).toBe(true);
    expect(output.pictures.find(p=>p.alt==='Word effect inline picture')?.inline).toBe(true);
    expect(output.canvas).toBe(0);
    expect(JSON.stringify(sections)).toBe(before);
    await view.evaluate("(()=>{[...document.querySelectorAll('[data-aared-element=page]')].at(-1).scrollIntoView();return new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)))})()");
    if(process.env.AARED_ARTIFACTS_DIR)await Bun.write(`${process.env.AARED_ARTIFACTS_DIR}/word-effects-aared.png`,await view.screenshot({encoding:'buffer',format:'png'}));
    await view.evaluate("(()=>{const supports=CSS.supports.bind(CSS);CSS.supports=(name,value)=>name==='-webkit-box-reflect'?false:supports(name,value)})()");
    const fallbackLayout = await view.evaluate(`paint(${before})`);
    expect(fallbackLayout).toEqual(layout);
    const mirror = await view.evaluate(`(()=>{const m=document.querySelector('[data-aared-element="reflection"]');
      const s=[...document.querySelectorAll('[data-aared-element="text-box-surface"]')].find(s=>s.textContent.startsWith('Word effect reflection'));
      return {position:getComputedStyle(m).position,inert:m.inert,hidden:m.getAttribute('aria-hidden'),
        width:parseFloat(getComputedStyle(m).width),sourceWidth:parseFloat(getComputedStyle(s).width),
        height:parseFloat(getComputedStyle(m).height),sourceHeight:parseFloat(getComputedStyle(s).height)}})()`) as
      {position:string;inert:boolean;hidden:string;width:number;sourceWidth:number;height:number;sourceHeight:number};
    expect(mirror.position).toBe('absolute');expect(mirror.inert).toBe(true);expect(mirror.hidden).toBe('true');
    expect(Math.abs(mirror.width - mirror.sourceWidth)).toBeLessThanOrEqual(1);
    expect(Math.abs(mirror.height - mirror.sourceHeight)).toBeLessThanOrEqual(1);
    // Remove only effect paint, preserving real authored geometry/flow, to
    // prove filters/reflections do not influence the pagination measurement.
    const plain = structuredClone(sections);
    const strip=(value:unknown):void=>{if(!value||typeof value!=='object')return;if(Array.isArray(value)){value.forEach(strip);return;}
      const obj=value as Record<string,unknown>;delete obj.effects;for(const item of Object.values(obj))strip(item);};
    strip(plain);
    const control = await view.evaluate(`paint(${JSON.stringify(plain)})`) as {pages:number;heights:number[]};
    expect(control).toEqual(layout);
    expect(await view.evaluate('document.querySelectorAll("[data-aared-element=reflection]").length') as number).toBe(0);
    await view.evaluate('handle.destroy()');
    });
},30000);
