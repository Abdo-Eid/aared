import { expect, test } from "bun:test";
import { serveDocx } from "./docx-harness.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

/** Serve the fixture with a page that exposes flow + layout render entry points. */
async function serveDocxHarness(): ReturnType<typeof serveDocx> {
  return serveDocx(FIXTURE_PATH, `<!doctype html><style>html,body{margin:0;font-family:Manrope,ui-sans-serif,system-ui,sans-serif}#host{width:1200px}</style><main id="host"></main><script type="module">
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
      </script>`);
}

async function waitForImages(view: Bun.WebView): Promise<void> {
  // Images resolve asynchronously through the media resolver; poll for
  // blob URLs instead of sleeping a fixed interval.
  await view.evaluate(`new Promise((resolve, reject) => {
    const deadline = Date.now() + 4000;
    const poll = () => {
      const images = [...document.querySelectorAll('[data-aared-element="image"] img, [data-aared-element="inline-image"] img')];
      if (images.length >= 3 && images.every((img) => img.src.startsWith("blob:"))) resolve(undefined);
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
      const images = [...document.querySelectorAll('[data-aared-element="image"] img, [data-aared-element="inline-image"] img')];
      const inlineImages = document.querySelectorAll('[data-aared-element="inline-image"]').length;
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
        // Runs sit directly in the paragraph (no wrapper); tab spacers and markers aren't runs.
        quote: quote ? {
          runs: [...quote.querySelectorAll("span")].filter((span) => !span.hasAttribute("data-aared-tab") && !span.hasAttribute("data-aared-element")).length,
          wrappers: document.querySelectorAll('[data-aared-element="paragraph-text"]').length,
        } : null,
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
        inlineImages,
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
      quote: { runs: number; wrappers: number } | null;
      headings: Array<{ tag: string; level?: string; text: string }>;
      tables: number;
      firstTable: { direction?: string; cells: string[]; spanned: boolean } | null;
      lists: number;
      markers: string[];
      figures: string[];
      inlineImages: number;
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
    expect(output.quote?.wrappers).toBe(0);
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
    // The two picture-only paragraphs are block figures; the header logo sits
    // on its wordmark's line, so it is an inline image, not a third figure.
    expect(output.figures).toHaveLength(2);
    expect(output.inlineImages).toBe(1);
    expect(output.images.length).toBeGreaterThanOrEqual(3);
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
    // Flow has no pagination, so each field shows its cached result once.
    expect(output.footer?.text).toContain("Page 4 of 4");

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
      root: [0.8, 0.9, 1, 1.5].map((lineSpacing) => ({
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
    // Word's multiple scales its ~1.15em single-line box, so 1 paints 1.15 and
    // 1.5 paints 1.725; only multiples below ~0.96 reach the safety floor.
    expect(spacing.map((line) => [line.lineHeight, line.source])).toEqual([
      ["1.1", "0.8"],
      ["1.1", "0.9"],
      ["1.15", "1"],
      ["1.725", "1.5"],
    ]);
    for (const line of spacing) expect(line.box).toBeGreaterThan(0);
    expect(Math.abs(spacing[0]!.box - spacing[1]!.box)).toBeLessThanOrEqual(1);
    expect(spacing[3]!.box).toBeGreaterThan(spacing[2]!.box * 1.2);

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
    // Numbering indents position list text once: the paragraph carries the
    // indent from its start edge (so RTL items indent from the right), and the
    // marker hangs inside it — the item box itself adds nothing.
    expect(output.orderedItem).not.toBeNull();
    expect(parseFloat(output.orderedItem?.itemMargin ?? "0")).toBeCloseTo(0, 0);
    expect(parseFloat(output.orderedItem?.paraMargin ?? "0")).toBeCloseTo(33.33, 0);
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

visual("run vertical-align maps OOXML super/subscript to the CSS keywords", async () => {
  const { server, sections } = await serveDocxHarness();
  try {
    await using view = new Bun.WebView({ width: 1000, height: 900 });
    await view.navigate(server.url.href);
    await view.evaluate(`renderFixture(${JSON.stringify(sections)})`);
    const output = await view.evaluate(`(() => {
      // The "25pPAL" run carries w:vertAlign superscript; footnote markers
      // share the same paint path.
      const spans = [...document.querySelectorAll("span")].filter((s) => s.textContent === "PAL" || /^[0-9]+$/.test(s.textContent ?? ""));
      return spans.map((s) => {
        const cs = getComputedStyle(s);
        return { text: s.textContent, inlineVa: s.style.verticalAlign, computedVa: cs.verticalAlign, inlineFs: s.style.fontSize };
      });
    })()`) as Array<{ text: string; inlineVa: string; computedVa: string; inlineFs: string }>;
    const pal = output.find((s) => s.text === "PAL");
    expect(pal).toBeDefined();
    // `superscript` is the OOXML name, dropped by every browser — the renderer
    // must state the shift in CSS. Word raises a superscript by a third of its
    // 10.5pt run and sets it at two thirds (read from a native export's PDF).
    expect(pal!.inlineVa).toBe("3.5pt");
    expect(parseFloat(pal!.computedVa)).toBeCloseTo(4.667, 2);
    expect(pal!.inlineFs).toBe("7pt");
  } finally {
    server.stop();
  }
}, 30000);

/**
 * Layout pages paint lazily (`content-visibility: auto`): a page off screen
 * skips style and layout until it is shown, and its tab stops lay out then —
 * never by forcing layout of every page at render time.
 */
visual("visual DOM output: DOCX layout pages paint lazily and lay out tabs when shown", async () => {
  const { server } = await serveDocx(FIXTURE_PATH, `<!doctype html><style>html,body{margin:0}#host{width:900px;height:600px;overflow:auto}</style><main id="host"></main><script type="module">
        import { render } from "/docx.mjs";
        window.renderPages = async (model) => {
          window.handle = await render({ document: model, target: document.querySelector("#host"), pagination: { mode: "layout" }, zoom: 1 });
          return window.handle.pageCount;
        };
      </script>`);
  try {
    await using view = new Bun.WebView({ width: 1000, height: 700 });
    await view.navigate(server.url.href);
    const page = { widthTwips: 12240, heightTwips: 15840, orientation: "portrait", margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 } };
    const root = Array.from({ length: 12 }, (_, index) => [
      ...(index === 0 ? [] : [{ type: "pageBreak", strength: "mandatory" }]),
      { type: "paragraph", runs: [{ text: `Page ${index + 1}\tRight` }], tabStops: [{ pos: 6, align: "right" }], direction: "ltr" },
    ]).flat();
    const pageCount = await view.evaluate(`renderPages(${JSON.stringify([{ properties: { breakType: "nextPage", page, headerRefs: [], footerRefs: [] }, root }])})`) as number;
    expect(pageCount).toBe(12);
    await view.evaluate("new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))");
    const read = `(() => {
      const pages = [...document.querySelectorAll('[data-aared-element="page"]')];
      const spacer = (page) => page.querySelector('[data-aared-tab]');
      const tabEnd = (page) => {
        const range = document.createRange();
        const tab = spacer(page);
        range.setStartAfter(tab);
        range.setEndAfter(tab.parentElement.lastChild);
        const body = page.querySelector('[data-aared-element="page-body"]').getBoundingClientRect();
        return range.getBoundingClientRect().right - body.left;
      };
      return {
        lazy: pages.every((element) => getComputedStyle(element).contentVisibility === "auto"),
        firstShown: pages[0].querySelector('[data-aared-element="page-body"]').checkVisibility({ contentVisibilityAuto: true }),
        lastShown: pages.at(-1).querySelector('[data-aared-element="page-body"]').checkVisibility({ contentVisibilityAuto: true }),
        firstWidth: spacer(pages[0]).style.width,
        lastWidth: spacer(pages.at(-1)).style.width,
        firstEnd: tabEnd(pages[0]),
      };
    })()`;
    const before = await view.evaluate(read) as { lazy: boolean; firstShown: boolean; lastShown: boolean; firstWidth: string; lastWidth: string; firstEnd: number };
    expect(before.lazy).toBe(true);
    expect(before.firstShown).toBe(true);
    expect(before.lastShown).toBe(false);
    // The shown page's right tab ends its text on the 6in stop; the skipped page still has its default spacer.
    expect(before.firstWidth).not.toBe("48px");
    expect(Math.abs(before.firstEnd - 6 * 96)).toBeLessThanOrEqual(1);
    expect(before.lastWidth).toBe("48px");

    await view.evaluate("window.handle.goToPage(11)");
    const after = await view.evaluate(`new Promise((resolve) => {
      const deadline = Date.now() + 4000;
      const poll = () => {
        const pages = [...document.querySelectorAll('[data-aared-element="page"]')];
        const last = pages.at(-1);
        const width = last.querySelector('[data-aared-tab]').style.width;
        if (width !== "48px" || Date.now() > deadline) {
          resolve({ width, shown: last.querySelector('[data-aared-element="page-body"]').checkVisibility({ contentVisibilityAuto: true }) });
        } else requestAnimationFrame(poll);
      };
      poll();
    })`) as { width: string; shown: boolean };
    expect(after.shown).toBe(true);
    expect(after.width).not.toBe("48px");
  } finally {
    server.stop();
  }
}, 30000);

const OUTLINE_FIXTURE = `${import.meta.dir}/../../../fixtures/docx-outline.docx`;

/**
 * `docx-outline.docx` (Word-authored) puts a heading-styled paragraph in the
 * page header and a Heading 2 in a table cell around three body headings. Only
 * the body headings are the outline, so `goToHeading(i)` must land on
 * `headings.list()[i]` in both modes — the painter used to count the header
 * and table headings too, and restart at 0 inside tables.
 */
visual("visual DOM output: DOCX heading and page navigation use the outline's own index space", async () => {
  const { server, sections } = await serveDocx(OUTLINE_FIXTURE, `<!doctype html><style>html,body{margin:0}#host{width:900px;height:600px;overflow:auto}</style><main id="host"></main><script type="module">
        import { render } from "/docx.mjs";
        const outlineOf = (model) => model.flatMap((section, index) =>
          section.root.filter((node) => node.type === "heading").map((node) => ({ title: node.text, section: index })));
        window.probe = async (model, mode) => {
          const host = document.querySelector("#host");
          const handle = await render({ document: model, target: host, pagination: { mode } });
          const landed = [];
          const original = Element.prototype.scrollIntoView;
          Element.prototype.scrollIntoView = function () { landed.push(this); };
          const outline = outlineOf(model);
          const byIndex = [];
          const byTitle = [];
          try {
            for (const [index, heading] of outline.entries()) {
              await handle.goToHeading(index);
              byIndex.push(landed.at(-1)?.textContent.trim() ?? null);
              await handle.goToHeading(heading.title.toUpperCase());
              byTitle.push(landed.at(-1)?.textContent.trim() ?? null);
            }
            let pageLanding = null;
            let pageError = null;
            try {
              await handle.goToPage(handle.pageCount - 1);
              pageLanding = landed.at(-1)?.dataset.page ?? null;
            } catch (error) {
              pageError = error.code;
            }
            let missing = null;
            try { await handle.goToHeading(outline.length); } catch (error) { missing = error.code; }
            const headings = [...host.querySelectorAll('[data-aared-element="heading"]')];
            const indexed = headings.filter((element) => element.dataset.aaredHeadingIndex !== undefined);
            const ids = indexed.map((element) => element.id);
            return {
              outline: outline.map((heading) => heading.title),
              byIndex,
              byTitle,
              painted: headings.length,
              indexed: indexed.map((element) => element.textContent.trim()),
              uniqueIds: new Set(ids).size === ids.length,
              pageCount: handle.pageCount,
              pageLanding,
              pageError,
              missing,
            };
          } finally {
            Element.prototype.scrollIntoView = original;
            handle.destroy();
          }
        };
      </script>`);
  try {
    await using view = new Bun.WebView({ width: 1000, height: 800 });
    await view.navigate(server.url.href);
    const model = JSON.stringify(sections);
    for (const mode of ["flow", "layout"] as const) {
      const result = await view.evaluate(`probe(${model}, "${mode}")`) as {
        outline: string[];
        byIndex: Array<string | null>;
        byTitle: Array<string | null>;
        painted: number;
        indexed: string[];
        uniqueIds: boolean;
        pageCount: number;
        pageLanding: string | null;
        pageError: string | null;
        missing: string | null;
      };
      expect(result.outline).toEqual(["Introduction / مقدمة", "Details / التفاصيل", "Closing / الخاتمة"]);
      // goToHeading(i) and goToHeading(title) both land on outline entry i.
      expect(result.byIndex).toEqual(result.outline);
      expect(result.byTitle).toEqual(result.outline);
      // The header and table headings still paint as headings, just unindexed.
      expect(result.painted).toBeGreaterThan(result.outline.length);
      expect(result.indexed).toEqual(result.outline);
      expect(result.uniqueIds).toBe(true);
      expect(result.missing).toBe("not-found");
      if (mode === "layout") {
        // Pages exist only in layout mode; the page break gives this file two.
        expect(result.pageCount).toBeGreaterThanOrEqual(2);
        expect(result.pageLanding).toBe(String(result.pageCount));
        expect(result.pageError).toBeNull();
      } else {
        expect(result.pageCount).toBe(0);
        expect(result.pageError).toBe("unsupported-feature");
      }
    }
  } finally {
    server.stop();
  }
}, 30000);
