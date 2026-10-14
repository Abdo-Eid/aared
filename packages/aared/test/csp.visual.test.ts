import { expect } from "bun:test";
import { browserModule, visual, withVisualPage } from "./harness.js";


const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;
const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

/** The page's own look, from a file: under the policy below the page itself may not use inline style either. */
const PAGE_CSS = "body { margin: 0 } aared-viewer { display: flex; height: 600px } #paged { width: 900px; height: 600px; overflow: auto }";

/**
 * A page whose policy forbids inline style (`style-src 'self'`, no
 * `'unsafe-inline'`) shows every format, through `<aared-viewer>` and through
 * a headless layout-mode render, with no violation: renderer defaults and the
 * element's chrome are constructed sheets, and inline style is only ever set
 * through the CSSOM. A `<style>` written into a shadow root is the control:
 * it must be refused and reported, or the check proves nothing.
 */
visual("visual CSP: every format and the element render under style-src without 'unsafe-inline'", async () => {
  const [elementModule, previewModule, workerModule, pptx, xlsx, docx] = await Promise.all([
    browserModule(`${import.meta.dir}/../src/element/index.ts`),
    browserModule(`${import.meta.dir}/../src/preview/index.ts`),
    browserModule(`${import.meta.dir}/../src/core/worker.ts`),
    Bun.file(PPTX_FIXTURE).arrayBuffer(),
    Bun.file(XLSX_FIXTURE).arrayBuffer(),
    Bun.file(DOCX_FIXTURE).arrayBuffer(),
  ]);
  const scripts: Record<string, string> = { "/element.mjs": elementModule, "/preview.mjs": previewModule, "/worker.mjs": workerModule };
  const files: Record<string, ArrayBuffer> = { "/deck.pptx": pptx, "/book.xlsx": xlsx, "/doc.docx": docx };
  // Every request that reaches the server. The policy below blocks any other
  // origin (`default-src 'self'`), so this is the whole network of the run.
  const requests: string[] = [];
  await withVisualPage({
    modules: scripts,
    files,
    page: `<!doctype html><link rel="stylesheet" href="/page.css"><script>
        window.violations = [];
        document.addEventListener("securitypolicyviolation", (event) => {
          window.violations.push({ directive: event.violatedDirective, sample: event.sample });
        }, true);
      </script><aared-viewer id="viewer" controls></aared-viewer><div id="paged"></div><div id="control"></div><script type="module">
        import { preview } from "/preview.mjs";
        import "/element.mjs";
        const viewer = document.querySelector("#viewer");
        const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        const show = (src) => new Promise((resolve, reject) => {
          viewer.addEventListener("load", () => resolve(), { once: true });
          viewer.addEventListener("error", (event) => reject(new Error(event.detail.message)), { once: true });
          viewer.src = src;
        });
        window.run = async () => {
          const out = { shown: [] };
          for (const src of ["/deck.pptx", "/book.xlsx", "/doc.docx"]) {
            await show(src);
            await frame();
            out.shown.push(viewer.viewer?.format);
          }
          const stage = viewer.shadowRoot.querySelector('[part="stage"]');
          out.stage = getComputedStyle(stage).backgroundColor;
          const paged = await preview({ source: "/doc.docx", target: document.querySelector("#paged"), pagination: { mode: "layout" } });
          await frame();
          out.pages = paged.docx.pageCount;
          paged.destroy();
          // Search highlighting must also stay clean: its ::highlight rule
          // lives in a constructed stylesheet, never in a <style> element.
          const probed = await preview({ source: "/doc.docx", target: document.querySelector("#paged") });
          await frame();
          await probed.search("nile");
          await frame();
          const probedTarget = document.querySelector("#paged");
          const highlightName = probedTarget.dataset.aaredSearchHighlight;
          out.searchHighlights = [...(CSS.highlights.get(highlightName) ?? [])].length;
          out.searchRule = document.adoptedStyleSheets.some((sheet) => [...sheet.cssRules].some((rule) => rule.cssText.includes("::highlight(" + highlightName + ")")));
          out.searchLegacyStyle = document.querySelectorAll('style[data-aared-search-highlight-style]').length;
          probed.destroy();
          out.aared = window.violations.splice(0);
          // The control: inline style text in a shadow root, as the element used to write.
          const shadow = document.querySelector("#control").attachShadow({ mode: "open" });
          const style = document.createElement("style");
          style.textContent = ":host { display: block }";
          shadow.append(style);
          await frame();
          out.control = window.violations.splice(0).map((violation) => violation.directive);
          return out;
        };
      </script>`,
    pageHeaders: {
      "content-security-policy":
        "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self'; worker-src 'self' blob:; img-src 'self' blob: data:; font-src 'self' blob: data:",
    },
    onRequest: (url, request) => {
      requests.push(`${request.method} ${url.pathname}`);
      if (url.pathname === "/page.css") return new Response(PAGE_CSS, { headers: { "content-type": "text/css" } });
      return undefined;
    },
    width: 1280,
    height: 900,
  }, async (view) => {
    const out = await view.evaluate(`run()`) as {
      shown: string[];
      stage: string;
      pages: number;
      searchHighlights: number;
      searchRule: boolean;
      searchLegacyStyle: number;
      aared: Array<{ directive: string; sample: string }>;
      control: string[];
    };
    expect(out.shown).toEqual(["pptx", "xlsx", "docx"]);
    // The element's chrome still styles the stage (the light theme's own colour).
    expect(out.stage).toBe("rgb(242, 242, 242)");
    expect(out.pages).toBeGreaterThan(1);
    // Highlighted search ran under the same policy: matches found, its rule
    // in a constructed sheet, no legacy <style> element, no violation.
    expect(out.searchHighlights).toBeGreaterThan(0);
    expect(out.searchRule).toBe(true);
    expect(out.searchLegacyStyle).toBe(0);
    expect(out.aared).toEqual([]);
    expect(out.control.length).toBe(1);
    expect(out.control[0]).toStartWith("style-src");
    // The documents stay in the page: no other origin was reachable (no
    // violation above), and the server saw only reads of its own files, never
    // an upload.
    const known = new Set(["/", "/preview.mjs", "/element.mjs", "/worker.mjs", "/page.css", "/deck.pptx", "/book.xlsx", "/doc.docx", "/favicon.ico"]);
    expect(requests.filter((request) => !request.startsWith("GET "))).toEqual([]);
    expect(requests.filter((request) => !known.has(request.slice(4)))).toEqual([]);
    for (const path of ["/deck.pptx", "/book.xlsx", "/doc.docx"]) expect(requests).toContain(`GET ${path}`);
  });
}, 60000);
