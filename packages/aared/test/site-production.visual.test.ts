import { expect } from "bun:test";
import { join, resolve } from "node:path";
import { serveVisual, visual, within, withVisualPage } from "./harness.js";

visual("production project Pages keeps navigation, assets and every demo under its base", async () => {
  const dist = resolve(import.meta.dir, "../../../site/dist");
  if (!await Bun.file(join(dist, "index.html")).exists()) throw new Error("Build the site with bun run docs:build before production site tests.");
  const missing: string[] = [];
  const requests: string[] = [];
  const server = serveVisual({ page: "", onRequest: async url => {
    if (url.pathname === "/") return new Response(null, { status: 302, headers: { location: "/aared/" } });
    requests.push(url.pathname);
    if (!url.pathname.startsWith("/aared/")) {
      missing.push(url.pathname);
      return new Response("outside the site base", { status: 404 });
    }
    const path = url.pathname.slice("/aared/".length);
    const file = Bun.file(join(dist, path.endsWith("/") || path === "" ? `${path}index.html` : path));
    if (!await file.exists()) {
      missing.push(url.pathname);
      return new Response("not found", { status: 404 });
    }
    return new Response(file);
  } });
  await withVisualPage({ server, width: 1440, height: 1100 }, async view => {
    const waitReady = async (selector: string, format: string): Promise<void> => {
      const status = `${selector} [data-state][data-format]`;
      const result = await within(`production ${selector} ${format} painted`, view.evaluate(`new Promise(resolve => {
        const check = () => {
          const status = document.querySelector(${JSON.stringify(status)});
          if (status?.dataset.state === 'error') return resolve({error:status.textContent});
          const stage = document.querySelector(${JSON.stringify(`${selector} [data-stage]`)});
          const root = stage?.querySelector('aared-viewer')?.shadowRoot ?? stage;
          const painted = root && [...root.querySelectorAll('[data-aared-element]')].some(node => node.textContent.trim().length > 0);
          // Hidden demos can report ready before the viewport's next-frame paint.
          // Re-read output after making a panel visible, not merely its status.
          if (status?.dataset.state === 'ready' && status.dataset.format === ${JSON.stringify(format)} && painted) return resolve({ready:true,painted:true});
          requestAnimationFrame(check);
        }; check();
      })`), 30000);
      expect(result).toEqual({ ready: true, painted: true });
    };
    await waitReady("#light-table-frame", "pptx");
    for (const format of ["xlsx", "docx"]) {
      await view.evaluate(`document.querySelector('#light-table-frame [data-format="${format}"]').click()`);
      await waitReady("#light-table-frame", format);
    }
    for (const route of ["start", "api", "playground"]) {
      await view.navigate(`${server.url.origin}/aared/${route}/`);
      expect(await view.evaluate<boolean>(`document.querySelector('h1').textContent.length > 0`)).toBe(true);
      expect(await view.evaluate<string>(`document.querySelector('.site-nav a.active').pathname`)).toBe(`/aared/${route}/`);
      const links = await view.evaluate<string[]>(`[...document.querySelectorAll('a[href], img[src], script[src], link[href], video[src], video[poster]')]
        .flatMap(node => ['href','src','poster'].map(name => node.getAttribute(name)).filter(Boolean))
        .map(path => new URL(path, location.href)).filter(url => url.origin === location.origin).map(url => url.pathname)`);
      expect(links.length).toBeGreaterThan(5);
      for (const path of links) expect(path.startsWith("/aared/"), path).toBe(true);
    }
    // The same production asset/Worker route serves all four integrations.
    for (const panel of ["component", "react", "facade", "headless"]) {
      await view.evaluate(`(() => {document.querySelector('#tab-${panel}').click();document.querySelector('#panel-${panel}').scrollIntoView();})()`);
      for (const format of ["pptx", "xlsx", "docx"]) {
        await view.evaluate(`document.querySelector('[data-pg-format="${format}"]').click()`);
        await waitReady(`#panel-${panel}`, format);
      }
    }
    expect(requests.some(path => /worker.*\.js$/.test(path))).toBe(true);
    for (const format of ["pptx", "xlsx", "docx"]) expect(requests.some(path => path.endsWith(`.${format}`))).toBe(true);
    expect(missing.filter(path => path !== "/favicon.ico")).toEqual([]);
  });
}, 180000);
