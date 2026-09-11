import { expect, test } from "bun:test";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import { OpcPackage } from "../src/core/opc.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/arabic-docx-bidi-coverage.docx`;

/** Build the browser renderer bundle, serve it with fixture media, and expose flow + layout render entry points. */
async function serveDocxHarness(): Promise<{ server: ReturnType<typeof Bun.serve>; sections: unknown }> {
  const bytes = await Bun.file(FIXTURE_PATH).arrayBuffer();
  const pkg = await OpcPackage.open(bytes);
  const sections = await parseDocxSections(pkg, await indexDocx(pkg));
  const media = new Map<string, Uint8Array>();
  for (const part of pkg.partNames()) {
    if (part.startsWith("word/media/")) media.set(part, await pkg.readPart(part));
  }
  const build = await Bun.build({
    entrypoints: [`${import.meta.dir}/../src/docx/index.ts`],
    target: "browser",
    format: "esm",
  });
  if (!build.success) throw new Error(build.logs.map((log) => log.message).join("\n"));
  const module = await build.outputs[0]!.text();
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/docx.mjs") {
        return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      if (url.pathname.startsWith("/media/")) {
        const id = `word/media/${url.pathname.slice("/media/".length)}`;
        const data = media.get(id);
        if (data === undefined) return new Response("missing media", { status: 404 });
        const bytes = data.slice().buffer as ArrayBuffer;
        return new Response(bytes, { headers: { "content-type": id.endsWith(".png") ? "image/png" : "image/jpeg" } });
      }
      return new Response(`<!doctype html><style>html,body{margin:0;font-family:Manrope,ui-sans-serif,system-ui,sans-serif}#host{width:1200px}</style><main id="host"></main><script type="module">
        import { render } from "/docx.mjs";
        window.renderedFixture = null;
        const wiring = (model) => ({
          mediaResolver: async (mediaId) => {
            const name = mediaId.split("/").at(-1);
            const response = await fetch("/media/" + name);
            if (!response.ok) throw new Error("missing media " + mediaId);
            return response.arrayBuffer();
          },
        });
        window.renderFixture = async (model) => {
          window.renderedFixture?.destroy();
          window.renderedFixture = await render({ document: model, target: document.querySelector("#host"), ...wiring(model) });
        };
        window.renderLayoutFixture = async (model) => {
          window.renderedFixture?.destroy();
          window.renderedFixture = await render({
            document: model,
            target: document.querySelector("#host"),
            ...wiring(model),
            pagination: { mode: "layout" },
          });
          return {
            pageCount: window.renderedFixture.pageCount,
            pages: window.renderedFixture.pages,
          };
        };
        window.destroyFixture = () => window.renderedFixture?.destroy();
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  return { server, sections };
}

async function waitForImages(view: Bun.WebView): Promise<void> {
  // Images resolve asynchronously through the media resolver; poll for
  // blob URLs instead of sleeping a fixed interval.
  await view.evaluate(`new Promise((resolve, reject) => {
    const deadline = Date.now() + 4000;
    const poll = () => {
      const images = [...document.querySelectorAll('[data-aared-element="image"] img')];
      if (images.length >= 2 && images.every((img) => img.src.startsWith("blob:"))) resolve(undefined);
      else if (Date.now() > deadline) reject(new Error("timed out waiting for document images"));
      else setTimeout(poll, 50);
    };
    poll();
  })`);
}

visual("visual DOM output: DOCX flow HTML keeps bidi, outline, tables, lists, images, and footnotes", async () => {
  const { server, sections } = await serveDocxHarness();
  try {
    await using view = new Bun.WebView({ width: 1000, height: 900 });
    await view.navigate(server.url.href);
    await view.evaluate(`renderFixture(${JSON.stringify(sections)})`);
    await waitForImages(view);
    const output = await view.evaluate(`(() => {
      const paragraphs = [...document.querySelectorAll('[data-aared-element="paragraph"]')];
      const mixed = paragraphs.find((line) => line.textContent.includes("Nile Heritage Foundation"));
      const rtl = paragraphs.find((line) => line.textContent.includes("صوت النهر — سلسلة وثائقية"));
      const quote = paragraphs.find((line) => line.textContent.includes("Cut on the movement"));
      const headings = [...document.querySelectorAll('[data-aared-element="heading"]')];
      const tables = [...document.querySelectorAll('[data-aared-element="table"]')];
      const firstTable = tables[0];
      const lists = [...document.querySelectorAll('[data-aared-element="list"]')];
      const markers = [...document.querySelectorAll('[data-aared-element="list-marker"]')].map((marker) => marker.dataset.aaredMarker);
      const figures = [...document.querySelectorAll('[data-aared-element="image"]')];
      const images = [...document.querySelectorAll('[data-aared-element="image"] img')];
      const refs = [...document.querySelectorAll('[data-aared-element="footnote-ref"] a')];
      const notes = [...document.querySelectorAll('[data-aared-element="footnote"]')];
      const header = document.querySelector('[data-aared-element="doc-header"]');
      const footer = document.querySelector('[data-aared-element="doc-footer"]');
      const runStyle = (line, text) => {
        const leaves = [...line.querySelectorAll("span")].filter((candidate) => candidate.querySelector("span") === null);
        const span = leaves.find((candidate) => candidate.textContent.includes(text));
        return span ? { fontFamily: getComputedStyle(span).fontFamily, text: span.textContent } : null;
      };
      return {
        canvasCount: document.querySelectorAll("canvas").length,
        mixed: mixed ? {
          direction: mixed.dataset.aaredDirection,
          computed: getComputedStyle(mixed).direction,
          latin: runStyle(mixed, "Nile Heritage"),
          arabic: runStyle(mixed, "مؤسسة"),
        } : null,
        rtl: rtl ? { direction: rtl.dataset.aaredDirection, align: rtl.dataset.aaredAlign, computed: getComputedStyle(rtl).direction } : null,
        quote: quote ? { runs: quote.querySelectorAll("span span").length } : null,
        headings: headings.map((heading) => ({ tag: heading.tagName, level: heading.dataset.aaredLevel, text: heading.textContent.slice(0, 30) })),
        tables: tables.length,
      firstTable: firstTable ? {
        direction: firstTable.dataset.aaredTableDirection,
        cells: [...firstTable.querySelectorAll('[data-aared-element="table-cell"]')].slice(0, 2).map((cell) => cell.textContent),
        spanned: tables.some((table) => table.querySelector('[style*="span 2"]') !== null),
      } : null,
        lists: lists.length,
        markers: markers.slice(0, 14),
        figures: figures.map((figure) => figure.dataset.aaredMediaId),
        images: images.map((img) => ({ src: img.src.slice(0, 5), alt: img.alt })),
        refs: refs.map((link) => link.getAttribute("href")),
        notes: notes.map((note) => ({ id: note.id, label: note.dataset.aaredLabel })),
        header: header ? { count: document.querySelectorAll('[data-aared-element="doc-header"]').length, text: header.textContent.slice(0, 40) } : null,
        footer: footer ? { count: document.querySelectorAll('[data-aared-element="doc-footer"]').length, text: footer.textContent.slice(0, 60) } : null,
      };
    })()`) as {
      canvasCount: number;
      mixed: { direction?: string; computed: string; latin: { fontFamily: string } | null; arabic: { fontFamily: string } | null } | null;
      rtl: { direction?: string; align?: string; computed: string } | null;
      quote: { runs: number } | null;
      headings: Array<{ tag: string; level?: string; text: string }>;
      tables: number;
      firstTable: { direction?: string; cells: string[]; spanned: boolean } | null;
      lists: number;
      markers: string[];
      figures: string[];
      images: Array<{ src: string; alt: string }>;
      refs: string[];
      notes: Array<{ id: string; label?: string }>;
      header: { count: number; text: string } | null;
      footer: { count: number; text: string } | null;
    };
    expect(output.canvasCount).toBe(0);
    // Mixed-script paragraph: LTR flow, per-script resolved families.
    expect(output.mixed?.direction).toBe("ltr");
    expect(output.mixed?.computed).toBe("ltr");
    expect(output.mixed?.latin?.fontFamily).toContain("Arial");
    expect(output.mixed?.arabic?.fontFamily).toContain("Arial");
    expect(output.rtl).toEqual({ direction: "rtl", align: "right", computed: "rtl" });
    expect(output.quote?.runs).toBeGreaterThanOrEqual(3);
    // Heading outline keeps levels and order.
    expect(output.headings.map((heading) => `${heading.tag}:L${heading.level}`)).toEqual([
      "H1:L1", "H2:L2", "H2:L2", "H1:L1", "H1:L1", "H1:L1", "H1:L1",
    ]);
    // Tables keep logical XML order; direction mirrors display.
    expect(output.tables).toBe(3);
    expect(output.firstTable?.direction).toBe("ltr");
    expect(output.firstTable?.cells[0]).toContain("Client");
    expect(output.firstTable?.cells[1]).toContain("Nile Heritage Foundation");
    expect(output.firstTable?.spanned).toBe(true);
    // Lists resolve bullet and decimal markers with levels.
    expect(output.lists).toBe(2);
    expect(output.markers.slice(0, 7)).toEqual(["—", "—", "◦", "—", "—", "—", "—"]);
    expect(output.markers.slice(7, 12)).toEqual(["1.", "2.", "a)", "b)", "3."]);
    // Images load through the media resolver into real <img> elements.
    expect(output.figures).toHaveLength(3);
    expect(output.images.length).toBeGreaterThanOrEqual(2);
    expect(output.images.every((img) => img.src.startsWith("blob:"))).toBe(true);
    // Footnote references link to end-of-flow notes rendered once.
    expect(output.refs).toContain("#aared-footnote-f1");
    expect(output.notes).toEqual([
      { id: "aared-footnote-f1", label: "1" },
      { id: "aared-footnote-f2", label: "2" },
    ]);
    expect(output.header?.count).toBe(1);
    expect(output.header?.text).toContain("NILE FRAME STUDIO");
    expect(output.footer?.count).toBe(1);
    expect(output.footer?.text).toContain("Page 1 of 1");

    // Authored relative spacing is preserved but clamped to a safe line box,
    // so Arabic glyph ink cannot clip the way a raw 0.9 box would (see
    // HANDOFF's title correction). The clamp is proven font-independently:
    // sub-floor multipliers produce identical boxes (the floor engaged),
    // while above-floor multipliers keep their authored height.
    await view.evaluate(`renderFixture(${JSON.stringify([{
      properties: {
        breakType: "nextPage",
        page: {
          widthTwips: 12240, heightTwips: 15840, orientation: "portrait",
          margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 },
        },
        headerRefs: [], footerRefs: [],
      },
      root: [0.9, 1, 1.5].map((lineSpacing) => ({
        type: "paragraph",
        runs: [{ text: "صوت النهر", fontFamily: "Arial", fontSizePt: 40 }],
        direction: "rtl",
        fontSizePt: 40,
        lineSpacing,
      })),
    }])})`);
    const spacing = await view.evaluate(`(() => [...document.querySelectorAll('[data-aared-element="paragraph"]')].map((line) => ({
      lineHeight: line.style.lineHeight,
      source: line.dataset.aaredLineSpacing,
      box: line.getBoundingClientRect().height,
    })))()`) as Array<{ lineHeight?: string; source?: string; box: number }>;
    expect(spacing.map((line) => [line.lineHeight, line.source])).toEqual([
      ["1.1", "0.9"],
      ["1.1", "1"],
      ["1.5", "1.5"],
    ]);
    expect(spacing[0]?.box).toBeGreaterThan(0);
    expect(spacing[1]?.box).toBeGreaterThan(0);
    expect(spacing[2]?.box).toBeGreaterThan(0);
    expect(Math.abs(spacing[0]!.box - spacing[1]!.box)).toBeLessThanOrEqual(1);
    expect(spacing[2]!.box).toBeGreaterThan(spacing[1]!.box * 1.2);

    // destroy() is idempotent and leaves re-rendering intact.
    const destroyed = await view.evaluate(`(() => {
      destroyFixture();
      destroyFixture();
      return document.querySelector('[data-aared-element="document"]') !== null;
    })()`);
    expect(destroyed).toBe(true);
  } finally {
    server.stop();
  }
}, 30000);

visual("visual DOM output: DOCX layout mode paginates the gate fixture to four structural pages", async () => {
  const { server, sections } = await serveDocxHarness();
  try {
    await using view = new Bun.WebView({ width: 1300, height: 900 });
    await view.navigate(server.url.href);
    const handle = await view.evaluate(`renderLayoutFixture(${JSON.stringify(sections)})`) as {
      pageCount: number;
      pages: Array<{ index: number; width: number; height: number; orientation: string }>;
    };
    expect(handle.pageCount).toBe(4);
    expect(handle.pages.map((page) => page.orientation)).toEqual(["portrait", "portrait", "portrait", "landscape"]);
    await waitForImages(view);
    const output = await view.evaluate(`(() => {
      const pages = [...document.querySelectorAll('[data-aared-element="page"]')];
      const headers = [...document.querySelectorAll('[data-aared-element="page-header"]')];
      const footers = [...document.querySelectorAll('[data-aared-element="page-footer"]')];
      const paragraphs = [...document.querySelectorAll('[data-aared-element="paragraph"]')];
      const quote = paragraphs.find((line) => (line.textContent ?? "").includes("Cut on the movement"));
      const quoteStyle = quote ? getComputedStyle(quote) : null;
      const orderedItem = [...document.querySelectorAll('[data-aared-element="list-item"]')]
        .find((item) => (item.textContent ?? "").includes("Broadcast master"));
      const orderedPara = orderedItem?.querySelector('[data-aared-element="paragraph"]');
      const secondPage = pages[1];
      const specTable = secondPage ? [...secondPage.querySelectorAll('[data-aared-element="table"]')].find((table) =>
        (table.textContent ?? "").includes("Parameter")) : null;
      return {
        canvasCount: document.querySelectorAll("canvas").length,
        pages: pages.map((page) => ({
          page: page.dataset.page,
          orientation: page.dataset.aaredOrientation,
          width: page.style.width,
          height: page.style.height,
          text: (page.textContent ?? "").slice(0, 4000),
        })),
        headers: headers.length,
        footers: footers.map((footer) => footer.textContent),
        notes: [...document.querySelectorAll('[data-aared-element="footnote"]')].map((note) => ({
          id: note.id,
          page: note.closest('[data-aared-element="page"]')?.dataset.page,
        })),
        quote: quoteStyle ? {
          borderWidth: quoteStyle.borderLeftWidth,
          borderColor: quoteStyle.borderLeftColor,
          borderStyle: quoteStyle.borderLeftStyle,
          padding: quoteStyle.paddingLeft,
        } : null,
        orderedItem: orderedItem && orderedPara ? {
          itemMargin: getComputedStyle(orderedItem).marginLeft,
          paraMargin: getComputedStyle(orderedPara).marginLeft,
        } : null,
        tableFit: secondPage && specTable ? (() => {
          const pageBox = secondPage.getBoundingClientRect();
          const tableBox = specTable.getBoundingClientRect();
          const style = getComputedStyle(secondPage);
          const contentBottom = pageBox.bottom - parseFloat(style.paddingBottom || "0");
          return { tableBottom: tableBox.bottom, contentBottom };
        })() : null,
        figures: [...document.querySelectorAll('[data-aared-element="page"]')].map((page) =>
          [...page.querySelectorAll('[data-aared-element="image"] img')].map((img) => {
            const style = getComputedStyle(img);
            return { left: style.marginLeft, right: style.marginRight };
          })),
      };
    })()`) as {
      canvasCount: number;
      pages: Array<{ page?: string; orientation?: string; width: string; height: string; text: string }>;
      headers: number;
      footers: string[];
      notes: Array<{ id: string; page?: string }>;
      quote: { borderWidth: string; borderColor: string; borderStyle: string; padding: string } | null;
      orderedItem: { itemMargin: string; paraMargin: string } | null;
      tableFit: { tableBottom: number; contentBottom: number } | null;
      figures: Array<Array<{ left: string; right: string }>>;
    };
    expect(output.canvasCount).toBe(0);
    expect(output.pages.map((page) => page.page)).toEqual(["1", "2", "3", "4"]);
    expect(output.pages.map((page) => page.orientation)).toEqual(["portrait", "portrait", "portrait", "landscape"]);
    // Real section geometry drives the sheets, not a fixed body padding.
    expect(parseFloat(output.pages[0]?.width ?? "0")).toBeCloseTo(793.73, 1);
    expect(parseFloat(output.pages[0]?.height ?? "0")).toBeCloseTo(1122.53, 1);
    expect(parseFloat(output.pages[3]?.width ?? "0")).toBeCloseTo(1122.53, 1);
    expect(parseFloat(output.pages[3]?.height ?? "0")).toBeCloseTo(793.73, 1);
    // Content maps to the Word reference structure.
    expect(output.pages[0]?.text).toContain("Contents");
    expect(output.pages[1]?.text).toContain("2. Technical Specifications");
    expect(output.pages[1]?.text).toContain("Parameter");
    expect(output.pages[2]?.text).toContain("3. Schedule");
    expect(output.pages[2]?.text).toContain("4. Contacts");
    expect(output.pages[2]?.text).not.toContain("Appendix A");
    expect(output.pages[3]?.text).toContain("Appendix A — Shot List, Episode 04");
    // Headers and footers repeat on every physical page...
    expect(output.headers).toBe(4);
    expect(output.footers).toHaveLength(4);
    // ...with PAGE/NUMPAGES evaluated against the finalized layout.
    expect(output.footers[1]).toContain("Page 2 of 4");
    expect(output.footers[3]).toContain("Page 4 of 4");
    // Footnotes follow their references onto the owning pages.
    expect(output.notes).toEqual([
      { id: "aared-footnote-f1", page: "2" },
      { id: "aared-footnote-f2", page: "3" },
    ]);
    // The pull quote paints its gold `w:pBdr` left border, not a bare line.
    expect(output.quote).toMatchObject({ borderStyle: "solid" });
    expect(parseFloat(output.quote?.borderWidth ?? "0")).toBeGreaterThan(0);
    expect(output.quote?.borderColor).toContain("184, 127, 30");
    expect(parseFloat(output.quote?.padding ?? "0")).toBeGreaterThan(0);
    // Numbering indents position list text once: the item carries the
    // numbering indent, the inner paragraph carries none.
    expect(output.orderedItem).not.toBeNull();
    expect(parseFloat(output.orderedItem?.itemMargin ?? "0")).toBeCloseTo(33.33, 0);
    expect(parseFloat(output.orderedItem?.paraMargin ?? "0")).toBeCloseTo(0, 0);
    // The spec table fits inside page 2's content box — no mid-row clip —
    // with its last row present on the same page as its caption.
    expect(output.tableFit).not.toBeNull();
    expect(output.tableFit!.tableBottom).toBeLessThanOrEqual(output.tableFit!.contentBottom + 1);
    expect(output.pages[1]?.text).toContain("Subtitles");
    // Centered source paragraphs center their images: the appendix figure
    // splits leftover space equally on both sides.
    const appendixFigures = output.figures[3] ?? [];
    expect(appendixFigures.length).toBeGreaterThan(0);
    const appendixFigure = appendixFigures[appendixFigures.length - 1]!;
    expect(parseFloat(appendixFigure.left)).toBeGreaterThan(0);
    expect(appendixFigure.left).toBe(appendixFigure.right);
  } finally {
    server.stop();
  }
}, 30000);
