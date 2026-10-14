/** Bounded warm-render/zoom comparison using a production-configured internal painter. */
import { join } from "node:path";
import { open, readSlideModel } from "../../packages/aared/src/pptx/renderer.js";
import type { DrawingEffectsModel, SlideModel, SlideElementModel } from "../../packages/aared/src/core/model.js";
import { REPO_ROOT, OUTPUT_ROOT } from "../shared/paths.js";
import { browserModule } from "../../packages/aared/test/harness.js";

const root = REPO_ROOT;
const media = new Map<string, ArrayBuffer>();
async function readDeck(file: string): Promise<SlideModel[]> {
  const deck = await open(await Bun.file(join(root, "fixtures", file)).arrayBuffer());
  try {
    const models = [];
    for (const slide of deck.slides.list()) models.push(await readSlideModel(slide));
    for (const match of JSON.stringify(models).matchAll(/"(?:mediaId|pictureMediaId)":"([^"]+)"/g)) {
      if (!media.has(match[1]!)) media.set(match[1]!, await deck.media.get(match[1]!));
    }
    return models;
  } finally { deck.close(); }
}
const ordinary = await readDeck("post-production-review.pptx");
const appearance = (await readDeck("pptx-coverage.pptx")).find(m=>m.elements.some(e=>e.name==="Group fill inheritance"));
if (!appearance) throw new Error("group appearance fixture is missing");
const heavy = structuredClone(appearance);
// Deliberate stress, not a consumer-frequency claim: eight overlapping copies
// of the real authored groups, at their actual sizes, in one visible slide.
heavy.elements = Array.from({length:8},()=>structuredClone(appearance.elements.filter(e=>e.kind==="group"))).flat();
function withoutEffects(model: SlideModel): SlideModel {
  const copy = structuredClone(model);
  const visit = (elements: SlideElementModel[]): void => {
    for (const e of elements) {
      if (e.kind === "shape" || e.kind === "group" || e.kind === "picture") delete e.effects;
      if (e.kind === "group") visit(e.children);
    }
  };
  visit(copy.elements); if(copy.layout) visit(copy.layout.graphics);
  return copy;
}
function onlyEffect(kind: keyof DrawingEffectsModel): SlideModel {
  const model = structuredClone(heavy);
  const visit = (elements: SlideElementModel[]): void => {
    for (const e of elements) {
      if ((e.kind === "shape" || e.kind === "group" || e.kind === "picture") && e.effects) {
        const effect = e.effects[kind];
        if (effect === undefined) delete e.effects;
        else e.effects = { [kind]: effect } as DrawingEffectsModel;
      }
      if (e.kind === "group") visit(e.children);
    }
  };
  visit(model.elements); if (model.layout) visit(model.layout.graphics);
  return model;
}
// This diagnostic deliberately transforms private models. It must not turn
// detached-model rendering back into a package export. Warm paint is measured,
// not parse or bundle cost; match production minification/diagnostic settings.
const painter = await browserModule(join(root, "packages/aared/src/pptx/renderer.ts"), { production: true });
const server = Bun.serve({ hostname:"127.0.0.1", port:0, async fetch(request) {
  const url = new URL(request.url);
  if (url.pathname === '/painter.mjs') return new Response(painter, {headers:{'content-type':'text/javascript'}});
  if(url.pathname.startsWith('/media/')) return new Response(media.get(decodeURIComponent(url.pathname.slice(7))) ?? null);
  return new Response(`<!doctype html><style>html,body{margin:0;overflow:hidden}#host{width:960px;height:540px}</style><main id="host"></main>
    <script type="module">
    import {render} from '/painter.mjs';
    const target=document.querySelector('#host');
    const frame=()=>new Promise(resolve=>requestAnimationFrame(resolve));
    const percentile=(values,p)=>[...values].sort((a,b)=>a-b)[Math.max(0,Math.ceil(values.length*p)-1)];
    const counters=()=>({nodes:document.querySelectorAll('*').length,
      filters:target.querySelectorAll('filter').length,mirrors:target.querySelectorAll('[data-aared-element="reflection"]').length,
      cssShadows:[...target.querySelectorAll('[data-aared-element="group"],[data-aared-element="shape"],[data-aared-element="picture"]')].filter(e=>getComputedStyle(e).filter.startsWith('drop-shadow')).length});
    const supports=CSS.supports.bind(CSS);
    window.measure=async(models,fallback)=>{
      CSS.supports=(...args)=>args[0]==='-webkit-box-reflect'&&fallback?false:supports(...args);
      let handle;
      const display=[],teardown=[]; let peakNodes=0,filterCount=0,mirrorCount=0,cssShadowCount=0;
      // Warm code, fonts, decoded media and layout; timings exclude download/parse.
      const paint=async model=>{
        handle?.destroy();
        handle=await render({slide:model,target,zoom:'fit-width',mediaResolver:async id=>(await fetch('/media/'+encodeURIComponent(id))).arrayBuffer()});
        await document.fonts.ready;
        const deadline=performance.now()+5000;
        while(target.querySelector('[data-aared-fill-loading]')||[...target.querySelectorAll('img')].some(i=>!i.complete||!i.naturalWidth)){
          if(performance.now()>deadline)throw new Error('media timeout');await frame();
        }
        await frame();await frame();
      };
      for(const model of models)await paint(model);
      // Alternate across real slides instead of repeatedly measuring one layout.
      for(let round=0;round<3;round++)for(const model of models){
        const start=performance.now();await paint(model);display.push(performance.now()-start);
        const c=counters();peakNodes=Math.max(peakNodes,c.nodes);filterCount=Math.max(filterCount,c.filters);mirrorCount=Math.max(mirrorCount,c.mirrors);cssShadowCount=Math.max(cssShadowCount,c.cssShadows);
      }
      const frames=[];
      for(let i=0;i<60;i++){
        const start=performance.now();handle.setZoom(.65+(i%20)/100);await frame();frames.push(performance.now()-start);
      }
      const end=performance.now();handle.destroy();target.replaceChildren();await frame();teardown.push(performance.now()-end);
      const remaining=counters();CSS.supports=supports;
      return {displayMedianMs:percentile(display,.5),displayP95Ms:percentile(display,.95),
        zoomFrameMedianMs:percentile(frames,.5),zoomFrameP95Ms:percentile(frames,.95),
        framesOver25Ms:frames.filter(n=>n>25).length,peakNodes,filterCount,mirrorCount,cssShadowCount,
        teardownMs:teardown[0],remaining,viewport:[innerWidth,innerHeight],userAgent:navigator.userAgent};
    };
    </script>`, {headers:{'content-type':'text/html'}});
} });
try {
  await using view = new Bun.WebView({width:960,height:540});
  await view.navigate(server.url.href);await view.resize(960,540);
  const results: Record<string,unknown> = {};
  const isolate = process.argv.includes('--isolate');
  const cases: Array<readonly [string, readonly SlideModel[]]> = isolate
    ? [['control',[withoutEffects(heavy)]], ...(['outerShadow','innerShadow','glow','softEdgePt','reflection'] as const).map(kind=>[kind,[onlyEffect(kind)]] as const)]
    : [["ordinary",ordinary],["authored-effects",[appearance]],["stress-8x",[heavy]]];
  for(const [label,models] of cases) {
    // Alternate order on the repeat to reduce warmup/order bias.
    for(let repeat=0;repeat<2;repeat++)for(const mode of isolate?["on"]:repeat===0?["off","on","fallback"]:["fallback","on","off"]) {
      const input=mode==='off'?models.map(withoutEffects):models;
      const result=await view.evaluate(`measure(${JSON.stringify(input)},${mode==='fallback'})`);
      results[`${label}/${mode}/${repeat+1}`]=result;
      console.log(`${label}/${mode}/${repeat+1}`,JSON.stringify(result));
    }
  }
  const output=join(OUTPUT_ROOT,`benchmarks/group-effects-performance${isolate?'-isolated':process.argv.includes('--optimized')?'-optimized':''}.json`);
  await Bun.write(output,JSON.stringify({date:new Date().toISOString(),method:"warm production render + two frames; 60 zoom frames; 8x authored group stress; no GPU-memory inference",results},null,2));
  console.log(output);
} finally {server.stop();}
