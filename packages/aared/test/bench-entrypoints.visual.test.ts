import { expect } from "bun:test";
import { resolve, join } from "node:path";
import { visual, within, withVisualPage } from "./harness.js";

visual("production bench measurement imports never execute the React app bootstrap", async () => {
  const root = resolve(import.meta.dir, "../../..");
  const dist = join(root, "bench/dist");
  const page = await Bun.file(join(dist, "measure.html")).text();
  const loaded: string[] = [];
  const app = await Bun.file(join(dist, "index.html")).text();
  const appEntry = /<script[^>]+src="([^"]+)"/.exec(app)![1]!;
  await withVisualPage({ page, onRequest: async url => {
    if (url.pathname.startsWith("/assets/")) {
      loaded.push(url.pathname);
      const file = Bun.file(join(dist, url.pathname));
      return await file.exists() ? new Response(file) : new Response("missing asset", { status: 404 });
    }
    if (url.pathname.startsWith("/f/")) {
      const file = Bun.file(join(root, "fixtures", url.pathname.slice(3)));
      return await file.exists() ? new Response(file) : new Response("missing fixture", { status: 404 });
    }
  } }, async view => {
    await within("production measurement bootstrap", view.evaluate(`new Promise(resolve => {
      const check = () => window.__m ? resolve(true) : requestAnimationFrame(check); check();
    })`), 15000);
    expect(await view.evaluate(`document.getElementById('root')`)).toBeNull();
    const selectors = { pptx: "slide", xlsx: "grid", docx: "page" };
    for (const format of ["pptx", "xlsx", "docx"] as const) {
      await within(`production ${format} measurement render`, view.evaluate(`__m.open('${format}', 'aared', '${format}-coverage.${format}')`), 30000);
      const painted = await view.evaluate<{ count: number; text: string }>(`({
        count: document.querySelectorAll('[data-aared-element="${selectors[format]}"]').length,
        text: document.querySelector('#host').textContent
      })`);
      expect(painted.count).toBeGreaterThan(0);
      expect(painted.text).toMatch(/[\u0600-\u06ff]/);
      await view.evaluate(`__m.destroy()`);
      expect(await view.evaluate<number>(`document.querySelector('#host').childElementCount`)).toBe(0);
    }
    expect(loaded).not.toContain(appEntry);
  });
}, 120000);
