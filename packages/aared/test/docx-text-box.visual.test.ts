import { expect } from "bun:test";
import { visual, withVisualPage } from "./harness.js";
import type { DocxTextBoxNode, DocumentSection, FlowNode } from "../src/core/model.js";
import { serveDocx } from "./docx-harness.js";

const FIXTURE = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

function textBoxes(sections: readonly DocumentSection[]): DocxTextBoxNode[] {
  const out: DocxTextBoxNode[] = [];
  const visit = (nodes: readonly FlowNode[]): void => {
    for (const node of nodes) {
      if (node.type === "textBox") { out.push(node); visit(node.content); }
      else if (node.type === "paragraph" || node.type === "heading") {
        for (const drawing of node.floats ?? []) if (drawing.type === "textBox") { out.push(drawing); visit(drawing.content); }
      } else if (node.type === "table") {
        for (const row of node.rows) for (const cell of row.cells) visit(cell.nodes);
      }
    }
  };
  for (const section of sections) visit(section.root);
  return out;
}

function boxText(box: DocxTextBoxNode): string {
  return box.content.flatMap((node) => node.type === "paragraph" || node.type === "heading" ? node.runs.map((run) => run.text) : []).join("").replace(/\s+/g, " ").trim();
}

/** Compare actual frame/text/paint output, including a consumer reset and zoom. */
visual("visual DOM output: Word-authored text boxes retain paint, geometry and text placement", async () => {
  const { server, sections } = await serveDocx(FIXTURE, `<!doctype html>
    <style>html,body{margin:0}*{box-sizing:border-box}img{display:block;max-width:100%}#host{width:1000px;zoom:1.5}</style>
    <main id="host"></main><script type="module">
    import { render } from "/docx.mjs";
    window.paint = async model => {
      window.handle = await render({ document:model, target:document.querySelector('#host'),
        mediaResolver:async id => (await fetch('/media/'+id.split('/').at(-1))).arrayBuffer() });
      await document.fonts.ready;
      await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
    };
    window.inspect = specs => specs.map(spec => {
      const normalize = text => text.replace(/\\s+/g,' ').trim();
      const candidates = [...document.querySelectorAll('[data-aared-element="text-box"]')];
      const box = candidates.find(node => normalize(node.textContent).includes(spec.text));
      if (!box) return {text:spec.text,missing:true};
      const style = getComputedStyle(box);
      const surface = box.querySelector('[data-aared-element="text-box-surface"]') ?? box;
      const content = box.querySelector('[data-aared-element="text-box-content"]') ?? box;
      const range = document.createRange();
      range.selectNodeContents(content);
      const rect = box.getBoundingClientRect();
      const text = range.getBoundingClientRect();
      return { text:spec.text, missing:false, width:parseFloat(style.width), height:parseFloat(style.height),
        overflow:getComputedStyle(surface).overflow, frameOverflow:style.overflow,
        border:[...box.querySelectorAll('path')].some(el=>el.getAttribute('stroke-dasharray')!==null)
          ? getComputedStyle(surface).borderTopStyle : style.borderTopStyle,
        radius:style.borderRadius,
        transforms:[box,...box.querySelectorAll('*')].map(el=>getComputedStyle(el).transform).filter(value=>value!=='none'),
        whitespace:[...content.querySelectorAll('p,h1,h2,h3,h4,h5,h6')].map(el=>getComputedStyle(el).whiteSpace),
        topGap:(text.top-rect.top)/1.5, bottomGap:(rect.bottom-text.bottom)/1.5,
        paths:[...box.querySelectorAll('path')].map(el=>({d:el.getAttribute('d'),fill:el.getAttribute('fill')})),
        gradients:box.querySelectorAll('linearGradient,radialGradient').length,
        patterns:box.querySelectorAll('pattern').length,
        pictures:[...box.querySelectorAll('svg image')].map(el=>el.getAttribute('href')),
        fallback:box.dataset.aaredFillFallback ?? surface.dataset.aaredFillFallback ?? null,
        canvas:box.querySelectorAll('canvas').length };
    });
    </script>`);
  const boxes = textBoxes(sections);
    const specimens = boxes.filter((box) => box.fill?.kind !== "solid" && box.fill !== undefined
      || box.geometry?.preset === "roundRect" || box.geometry?.preset === "ellipse"
      || box.transform?.rotation || box.verticalAlign === "bottom" || box.verticalAlign === "middle"
      || box.wrapping === "none" || box.border?.style === "dashed");
    expect(specimens.some((box) => box.fill?.kind === "linear-gradient")).toBe(true);
    expect(specimens.some((box) => box.fill?.kind === "pattern")).toBe(true);
    expect(specimens.some((box) => box.fill?.kind === "picture")).toBe(true);
    expect(specimens.some((box) => box.geometry?.preset === "ellipse")).toBe(true);
    expect(specimens.some((box) => box.transform?.rotation)).toBe(true);
    await withVisualPage({ server, width: 1200, height: 900 }, async (view) => {
    await view.evaluate(`paint(${JSON.stringify(sections)})`);
    // Picture paint is asynchronous; wait on its actual SVG href, not an intermediate URL map.
    await view.evaluate(`new Promise((resolve,reject)=>{
      const deadline=Date.now()+6000;
      const poll=()=>{ const imgs=[...document.querySelectorAll('svg image')];
        if(imgs.some(img=>(img.getAttribute('href')??'').startsWith('blob:')))resolve();
        else if(Date.now()>deadline)reject(new Error('picture fill did not paint'));else setTimeout(poll,25); };poll(); })`);
    type Result = { missing: boolean; width: number; height: number; overflow: string; frameOverflow: string; border: string; radius: string;
      transforms: string[]; whitespace: string[]; topGap: number; bottomGap: number; paths: Array<{d: string; fill: string}>; gradients: number;
      patterns: number; pictures: string[]; fallback: string | null; canvas: number };
    const painted = await view.evaluate(`inspect(${JSON.stringify(specimens.map((box) => ({ text: boxText(box) })))})`) as Result[];
    for (const [index, output] of painted.entries()) {
      const model = specimens[index]!;
      expect(output.missing).toBe(false);
      expect(output.canvas).toBe(0);
      expect(output.fallback).toBeNull();
      if (model.fill?.kind === "linear-gradient" || model.fill?.kind === "path-gradient") expect(output.gradients).toBeGreaterThan(0);
      if (model.fill?.kind === "pattern") expect(output.patterns).toBeGreaterThan(0);
      if (model.fill?.kind === "picture") expect(output.pictures.some((href) => href?.startsWith("blob:"))).toBe(true);
      if (model.transform?.rotation) expect(output.transforms.length).toBeGreaterThan(0);
      if (model.geometry?.preset === "ellipse" || model.geometry?.preset === "roundRect") {
        // The outline is a real preset path with arcs, not a CSS rectangle.
        expect(output.paths.some((path) => path.d.includes("A"))).toBe(true);
      }
      // Bottom alignment sits near the lower edge; middle sits inside the frame
      // with comparable space above and below the text (Word's own probe).
      if (model.verticalAlign === "bottom") expect(output.topGap).toBeGreaterThan(output.bottomGap);
      if (model.verticalAlign === "middle") expect(Math.abs(output.topGap - output.bottomGap)).toBeLessThan(24);
      if (model.wrapping === "none") expect(output.whitespace.every((value) => value === "pre" || value === "nowrap")).toBe(true);
      if (model.border?.style === "dashed" || model.border?.style === "dotted") expect(output.paths.length).toBeGreaterThan(0);
      // Text stays inside its frame. A rotated box is exempt: its painted text
      // is transformed, so its axis-aligned rect legitimately leaves the frame.
      if (!model.transform?.rotation) {
        expect(output.topGap).toBeGreaterThanOrEqual(-1);
        expect(output.bottomGap).toBeGreaterThanOrEqual(-1);
      }
    }
    await view.evaluate("(() => { window.handle.destroy(); window.handle.destroy(); return document.querySelector('[data-aared-element=\"document\"]') !== null; })()");
  });
}, 30000);

visual("visual DOM output: text-box resize and clipping use the painted frame without mutating resolved data", async () => {
  const { server, sections } = await serveDocx(FIXTURE, `<!doctype html><style>html,body{margin:0}#host{width:900px}</style><main id="host"></main><script type="module">
    import { render } from '/docx.mjs';
    window.paint = async model => {
      const before=JSON.stringify(model);
      window.handle=await render({document:model,target:document.querySelector('#host'),pagination:{mode:'layout'}});
      await document.fonts.ready;
      await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
      const boxes=[...document.querySelectorAll('[data-aared-element="text-box"]')];
      return {unchanged:before===JSON.stringify(model),pages:window.handle.pageCount,
        boxes:boxes.map(box=>{
          const content=box.querySelector('[data-aared-element="text-box-content"]')??box;
          const surface=box.querySelector('[data-aared-element="text-box-surface"]')??box;
          const style=getComputedStyle(box),rect=box.getBoundingClientRect();
          const range=document.createRange();range.selectNodeContents(content);
          const text=range.getBoundingClientRect();
          return {height:parseFloat(style.height),overflow:getComputedStyle(surface).overflow,text:box.textContent,
            textBottom:text.bottom,frameBottom:rect.bottom,paddingBottom:parseFloat(style.paddingBottom),
            pageBottom:box.closest('[data-aared-element="page"]').getBoundingClientRect().bottom};
        })};
    };</script>`);
  const source = textBoxes(sections).find((box) => box.autofit === "resizeShape" && boxText(box).length > 50);
    expect(source).toBeDefined();
    const original = source!;
    // Keep real Word-authored text/formatting; stale cached extents exercise both resize directions.
    const unanchored = (over: Partial<DocxTextBoxNode>): DocxTextBoxNode => {
      const { anchor: _drop, ...rest } = original;
      return { ...rest, ...over };
    };
    const grow = unanchored({ height: 0.2 });
    const shrink = unanchored({ height: 4 });
    const clipped = unanchored({ autofit: "none", height: 0.2 });
    const { header: _header, footer: _footer, ...page } = sections.at(-1)!;
    const model: DocumentSection[] = [{ ...page, root: [grow, shrink, clipped] }];
    await withVisualPage({ server, width: 1100, height: 1400 }, async (view) => {
    const output = await view.evaluate(`paint(${JSON.stringify(model)})`) as {
      unchanged: boolean; pages: number;
      boxes: Array<{height:number;overflow:string;text:string;textBottom:number;frameBottom:number;paddingBottom:number;pageBottom:number}>;
    };
    expect(output.unchanged).toBe(true);
    expect(output.boxes).toHaveLength(3);
    const [grown, shrunk, fixed] = output.boxes;
    expect(grown!.height).toBeGreaterThan(0.2 * 96 + 20);
    expect(shrunk!.height).toBeLessThan(4 * 96 - 20);
    expect(Math.abs(grown!.height - shrunk!.height)).toBeLessThan(1);
    expect(grown!.textBottom).toBeLessThanOrEqual(grown!.frameBottom + 1);
    expect(shrunk!.textBottom).toBeLessThanOrEqual(shrunk!.frameBottom + 1);
    expect(fixed!.height).toBeCloseTo(0.2 * 96, 0);
    expect(["clip", "hidden"]).toContain(fixed!.overflow);
    expect(fixed!.textBottom).toBeGreaterThan(fixed!.frameBottom);
    expect(fixed!.text).toContain(boxText(original).split(" ").at(-1)!);
    for (const box of output.boxes) expect(box.frameBottom).toBeLessThan(box.pageBottom);
    await view.evaluate("window.handle.destroy()");
  });
}, 30000);
