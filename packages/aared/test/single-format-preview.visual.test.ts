import { expect, test } from "bun:test";
import { browserModule } from "./harness.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;

visual("single-format previews share navigation, search, cancellation and owned cleanup", async () => {
  const formats = ["pptx", "xlsx", "docx"] as const;
  const entries = new Map(await Promise.all(formats.map(async (format) => [format,
    await browserModule(`${import.meta.dir}/../src/${format}/index.ts`)] as const)));
  const fixtures = new Map(await Promise.all(formats.map(async (format) => [format,
    await Bun.file(`${import.meta.dir}/../../../fixtures/${format}-coverage.${format}`).arrayBuffer()] as const)));
  const worker = await browserModule(`${import.meta.dir}/../src/core/worker.ts`);
  const server = Bun.serve({ port: 0, fetch(request) {
    const path = new URL(request.url).pathname;
    if (path === "/worker.mjs") return new Response(worker, { headers: { "content-type": "text/javascript" } });
    for (const format of formats) {
      if (path === `/${format}.mjs`) return new Response(entries.get(format), { headers: { "content-type": "text/javascript" } });
      if (path === `/fixture.${format}`) return new Response(fixtures.get(format));
    }
    return new Response(`<!doctype html><main id="host" style="width:1100px;height:650px"></main><script type="module">
      const host = document.querySelector('#host');
      window.run = async (format) => {
        const { preview } = await import('/' + format + '.mjs');
        const viewer = await preview({ source: '/fixture.' + format, target: host });
        const kinds = [];
        viewer.addEventListener('indexchange', e => kinds.push(e.detail.index));
        if (format !== 'docx') {
          await viewer.next();
          await viewer.previous();
          await Promise.all([viewer.goTo(1), viewer.goTo(0)]);
        } else await viewer.docx.goToHeading(0);
        viewer.setZoom(1.25);
        viewer.setAppearance({ theme: 'dark' });
        const query = format === 'docx' ? 'nile' : 'الانتاج';
        const matches = await viewer.search(query);
        await viewer.searchNavigation.next();
        const highlights = [...CSS.highlights.values()].reduce((n, h) => n + h.size, 0);
        const state = { format: viewer.format, index: viewer.index, zoom: viewer.zoom,
          theme: host.querySelector('[data-aared-theme]')?.dataset.aaredTheme,
          count: matches.length, highlights, kinds };
        await viewer.search(query, { highlight: false });
        const cleared = CSS.highlights.size === 0;
        viewer.destroy(); viewer.destroy();
        const empty = host.childElementCount === 0 && CSS.highlights.size === 0;
        const closed = await viewer.search(query).then(() => 'resolved', e => e.code);
        const cancelled = await preview({ source: '/fixture.' + format, target: host, signal: AbortSignal.abort() }).then(() => 'resolved', e => e.code);
        const wrong = format === 'pptx' ? 'xlsx' : 'pptx';
        const mismatch = await preview({ source: '/fixture.' + wrong, target: host }).then(() => 'resolved', e => e.code);
        return { ...state, cleared, empty, closed, cancelled, mismatch };
      };
    </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
  } });
  try {
    await using view = new Bun.WebView({ width: 1200, height: 800 });
    await view.navigate(server.url.href);
    for (const format of formats) {
      const result = await view.evaluate(`run(${JSON.stringify(format)})`) as {
        format: string; index: number; zoom: number; theme: string; count: number;
        highlights: number; kinds: number[]; cleared: boolean; empty: boolean;
        closed: string; cancelled: string; mismatch: string;
      };
      expect(result.format).toBe(format);
      expect(result.index).toBe(0);
      expect(result.zoom).toBe(1.25);
      expect(result.theme).toBe("dark");
      expect(result.count).toBeGreaterThan(0);
      expect(result.highlights).toBeGreaterThan(0);
      expect(result.cleared).toBe(true);
      expect(result.empty).toBe(true);
      expect(result.closed).toBe("closed");
      expect(result.cancelled).toBe("aborted");
      expect(result.mismatch).toBe("unsupported-format");
      if (format !== "docx") expect(result.kinds.slice(0, 2)).toEqual([1, 0]);
    }
  } finally { server.stop(); }
}, 60000);
