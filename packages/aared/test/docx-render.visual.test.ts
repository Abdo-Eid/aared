import { expect, test } from "bun:test";
import { fontAlias } from "../src/core/typography.js";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import { OpcPackage } from "../src/core/opc.js";
import { browserModule } from "./harness.js";
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
        headings: headings.map((heading) => ({ tag: heading.tagName, level: heading.dataset.aaredLevel, text: heading.textContent.slice(0, 30), weight: getComputedStyle(heading.querySelector("span") ?? heading).fontWeight })),
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
      headings: Array<{ tag: string; level?: string; text: string; weight: string }>;
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
    // Flow paints the header first (its Heading 1 styled mark), and the
    // tracker's Heading 2 total row paints inside its table cell.
    expect(output.headings.map((heading) => `${heading.tag}:L${heading.level}`)).toEqual([
      "H1:L1", "H1:L1", "H2:L2", "H2:L2", "H1:L1", "H1:L1", "H1:L1", "H1:L1", "H2:L2", "H1:L1",
    ]);
    // Weight comes from the runs, never the browser's <h1> bold: the header
    // mark is a Heading 1 set not bold in Word; the body headings are bold.
    expect(output.headings[0]).toMatchObject({ text: "Appendix / ملحق", weight: "400" });
    expect(output.headings[1]?.weight).toBe("700");
    // Tables keep logical XML order; direction mirrors display.
    expect(output.tables).toBe(4);
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
    // so Arabic glyph ink cannot clip the way a raw 0.9 box would
    // (ARCHITECTURE § Format rules). The clamp is proven font-independently:
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

visual("visual DOM output: DOCX layout mode paginates the gate fixture to its six structural pages", async () => {
  const { server, sections } = await serveDocxHarness();
  try {
    await using view = new Bun.WebView({ width: 1300, height: 900 });
    await view.navigate(server.url.href);
    const handle = await view.evaluate(`renderLayoutFixture(${JSON.stringify(sections)})`) as {
      pageCount: number;
      pages: Array<{ index: number; width: number; height: number; orientation: string }>;
    };
    // Word's export: three portrait pages, then the landscape appendix and its two pages of anchored pictures.
    expect(handle.pageCount).toBe(6);
    expect(handle.pages.map((page) => page.orientation)).toEqual(["portrait", "portrait", "portrait", "landscape", "landscape", "landscape"]);
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
    expect(output.pages.map((page) => page.page)).toEqual(["1", "2", "3", "4", "5", "6"]);
    expect(output.pages.map((page) => page.orientation)).toEqual(["portrait", "portrait", "portrait", "landscape", "landscape", "landscape"]);
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
    expect(output.headers).toBe(6);
    expect(output.footers).toHaveLength(6);
    // ...with PAGE/NUMPAGES evaluated against the finalized layout.
    expect(output.footers[1]).toContain("Page 2 of 6");
    expect(output.footers[5]).toContain("Page 6 of 6");
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

/**
 * The gate fixture has a Heading 1 styled paragraph in the appendix header
 * and a Heading 2 styled total row in the delivery-tracker table, around the
 * body headings (all Word-authored). Only
 * the body headings are the outline, so `goToHeading(i)` must land on
 * `headings.list()[i]` in both modes — the painter used to count the header
 * and table headings too, and restart at 0 inside tables.
 */
visual("visual DOM output: DOCX heading and page navigation use the outline's own index space", async () => {
  const { server, sections } = await serveDocx(FIXTURE_PATH, `<!doctype html><style>html,body{margin:0}#host{width:900px;height:600px;overflow:auto}</style><main id="host"></main><script type="module">
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
      expect(result.outline).toEqual([
        "1. Project Overview",
        "1.1 Objectives   الأهداف",
        "1.2 Deliverables   المخرجات",
        "2. Technical Specifications",
        "3. Schedule",
        "4. Contacts",
        "5. Delivery tracker / متابعة التسليم",
        "Appendix A — Shot List, Episode 04",
      ]);
      // goToHeading(i) and goToHeading(title) both land on outline entry i.
      expect(result.byIndex).toEqual(result.outline);
      expect(result.byTitle).toEqual(result.outline);
      // The header and table headings still paint as headings, just unindexed.
      expect(result.painted).toBeGreaterThan(result.outline.length);
      expect(result.indexed).toEqual(result.outline);
      expect(result.uniqueIds).toBe(true);
      expect(result.missing).toBe("not-found");
      if (mode === "layout") {
        // Pages exist only in layout mode.
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

/**
 * The delivery tracker (section 5) is a Word-authored `w:bidiVisual` table. Word draws
 * its first logical column on the right, its `w:left` border (3pt) on the
 * right edge as the start side, its `w:right` border (0.5pt) on the left, and
 * keeps the `w:left` cell margin (0.3in) on the physical left (probe
 * docx-bidi-visual, measured in Word's export). Read back from the DOM.
 */
visual("visual DOM output: DOCX right-to-left tables mirror columns and borders, not cell margins", async () => {
  const { server, sections } = await serveDocx(FIXTURE_PATH, `<!doctype html><style>html,body{margin:0}#host{width:900px}</style><main id="host"></main><script type="module">
        import { render } from "/docx.mjs";
        window.probe = async (model) => {
          const handle = await render({ document: model, target: document.querySelector("#host") });
          const table = [...document.querySelectorAll('[data-aared-element="table"]')].find((grid) => grid.textContent.includes("موقع التصوير"));
          const cells = [...table.querySelectorAll('[data-aared-element="table-cell"]')];
          const box = (cell) => cell.getBoundingClientRect();
          const style = (cell) => getComputedStyle(cell);
          const [first, second, , last] = cells;
          const result = {
            direction: table.dataset.aaredTableDirection,
            firstText: first.textContent.trim(),
            firstRightOfSecond: box(first).left > box(second).left,
            firstTouchesTableRight: Math.abs(box(first).right - table.getBoundingClientRect().right) < 1,
            startBorder: { right: style(first).borderRightWidth, left: style(first).borderLeftWidth },
            endBorder: { left: style(last).borderLeftWidth, right: style(last).borderRightWidth },
            padding: { left: style(first).paddingLeft, right: style(first).paddingRight },
            indent: { right: getComputedStyle(table).marginRight, left: getComputedStyle(table).marginLeft },
            // The tracker update line: "71%" is a left-to-right run in a
            // right-to-left paragraph, so "%" sits right of the digits, as in Word.
            ltrRun: (() => {
              const span = [...document.querySelectorAll("span")].find((el) => el.textContent === "71%");
              const glyph = (index) => { const range = document.createRange(); range.setStart(span.firstChild, index); range.setEnd(span.firstChild, index + 1); return range.getBoundingClientRect().left; };
              return { mark: getComputedStyle(span, "::before").content, percentRightOfDigits: glyph(2) > glyph(0) };
            })(),
            canvasCount: document.querySelectorAll("canvas").length,
          };
          handle.destroy();
          return result;
        };
      </script>`);
  try {
    await using view = new Bun.WebView({ width: 1000, height: 700 });
    await view.navigate(server.url.href);
    const result = await view.evaluate(`probe(${JSON.stringify(sections)})`) as {
      direction: string;
      firstText: string;
      firstRightOfSecond: boolean;
      firstTouchesTableRight: boolean;
      startBorder: { right: string; left: string };
      endBorder: { left: string; right: string };
      padding: { left: string; right: string };
      indent: { right: string; left: string };
      ltrRun: { mark: string; percentRightOfDigits: boolean };
      canvasCount: number;
    };
    expect(result.direction).toBe("rtl");
    expect(result.firstText).toBe("الحلقة");
    // The first logical column sits on the right, as in Word.
    expect(result.firstRightOfSecond).toBe(true);
    expect(result.firstTouchesTableRight).toBe(true);
    // `w:left` is the start side: 3pt (4px) on the right edge of the first column.
    expect(result.startBorder).toEqual({ right: "4px", left: "0px" });
    // `w:right` is the end side: 0.5pt (1px) on the left edge of the last column.
    expect(result.endBorder.left).toBe("1px");
    // Cell margins stay physical: 0.3in on the left, 0.05in on the right. Word
    // keeps the text a margin beyond the inner half of the 3pt start border
    // (probe docx-cell-border), and the cell paints all 4px of it: 4.8 + 2 - 4.
    expect(result.padding).toEqual({ left: "28.8px", right: "2.8px" });
    // `w:tblInd` indents from the start side: the right of this table.
    expect(result.indent.left).toBe("0px");
    expect(Number.parseFloat(result.indent.right)).toBeCloseTo(25 / 15, 1);
    expect(result.ltrRun).toEqual({ mark: '"‎"', percentRightOfDigits: true });
    expect(result.canvasCount).toBe(0);
  } finally {
    server.stop();
  }
}, 30000);

/**
 * Three Word layout rules, read back from the DOM of the coverage fixture:
 * an exact 18pt line that mixes Arial Arabic with Calibri runs stays 24px
 * (CSS stretched it to 25px); justified paragraphs of a Word 2013 document
 * break with slightly narrower spaces (probe docx-justify); and a cell's side
 * padding keeps its text a margin beyond half of each side border, including
 * the inside divider its neighbour paints (probe docx-cell-border).
 */
visual("visual DOM output: DOCX exact line pitch, justified spacing and cell border insets follow Word", async () => {
  const { server, sections } = await serveDocx(FIXTURE_PATH, `<!doctype html><style>html,body{margin:0}#host{width:900px}</style><main id="host"></main><script type="module">
        import { render } from "/docx.mjs";
        window.probe = async (model) => {
          const handle = await render({ document: model, target: document.querySelector("#host"), zoom: 1 });
          await document.fonts.ready;
          const paragraph = (prefix) => [...document.querySelectorAll("p")].find((el) => el.textContent.startsWith(prefix));
          const exact = paragraph("آخر تحديث");
          const table = [...document.querySelectorAll('[data-aared-element="table"]')].find((grid) => grid.textContent.includes("Nile Heritage Foundation"));
          const [first, second] = table.querySelectorAll('[data-aared-element="table-cell"]');
          const pad = (cell) => [getComputedStyle(cell).paddingLeft, getComputedStyle(cell).paddingRight].map(Number.parseFloat);
          const result = {
            exact: {
              lineHeight: getComputedStyle(exact).lineHeight,
              height: exact.getBoundingClientRect().height,
              fonts: [...new Set([...exact.querySelectorAll("span")].map((span) => getComputedStyle(span).fontFamily))],
              runLineHeights: [...new Set([...exact.querySelectorAll("span")].map((span) => getComputedStyle(span).lineHeight))],
            },
            justifiedSpacing: Number.parseFloat(getComputedStyle(paragraph("Sawt al-Nahr follows")).wordSpacing),
            leftSpacing: getComputedStyle(exact).wordSpacing,
            first: pad(first),
            second: pad(second),
          };
          handle.destroy();
          return result;
        };
      </script>`);
  try {
    await using view = new Bun.WebView({ width: 1000, height: 700 });
    await view.navigate(server.url.href);
    const result = await view.evaluate(`probe(${JSON.stringify(sections)})`) as {
      exact: { lineHeight: string; height: number; fonts: string[]; runLineHeights: string[] };
      justifiedSpacing: number;
      leftSpacing: string;
      first: [number, number];
      second: [number, number];
    };
    // One line at exactly 18pt, though its runs mix two fonts.
    expect(result.exact.lineHeight).toBe("24px");
    expect(result.exact.height).toBe(24);
    expect(result.exact.fonts.length).toBeGreaterThan(1);
    expect(result.exact.runLineHeights).toEqual(["0px"]);
    // Justified text breaks with narrower spaces; other paragraphs keep the font's.
    expect(result.justifiedSpacing).toBeLessThan(0);
    expect(result.leftSpacing).toBe("0px");
    // 120-twip (8px) margins; 0.5pt outer edges and a 0.25pt inside divider,
    // each painted at 1px by the cell that draws it.
    const outer = 4 / 6;
    const divider = 2 / 6;
    // First column: paints its outer left edge; the divider is its neighbour's.
    expect(result.first[0]).toBeCloseTo(8 + outer / 2 - 1, 2);
    expect(result.first[1]).toBeCloseTo(8 + divider / 2, 2);
    // Last column: paints the divider on its left and the outer right edge.
    expect(result.second[0]).toBeCloseTo(8 + divider / 2 - 1, 2);
    expect(result.second[1]).toBeCloseTo(8 + outer / 2 - 1, 2);
  } finally {
    server.stop();
  }
}, 30000);

/**
 * A delivered font substitute actually gets chosen, on a machine that lacks
 * the requested family entirely — not just registered and never selected
 * (DECISION_LOG §45; this was an open question in ARCHITECTURE.md before the
 * fix). Renames one run to a family name guaranteed not to exist anywhere,
 * then registers a `FontFace` under that name's alias (`fontAlias`, never
 * the bare requested name) sourced from a real, dramatically different
 * system font (`local('Courier New')`, monospace) before rendering, and
 * checks the computed style names the alias and the painted width matches
 * Courier's, not the generic sans-serif fallback a silently-ignored
 * registration would paint instead.
 */
visual("visual DOM output: a delivered font substitute is actually painted, not just registered", async () => {
  const NONEXISTENT_FAMILY = "Definitely Not A Real Font Zzqx";
  const ALIAS = fontAlias(NONEXISTENT_FAMILY);
  const { server, sections } = await serveDocxHarness();
  try {
    await using view = new Bun.WebView({ width: 1000, height: 900 });
    await view.navigate(server.url.href);
    const renamed = structuredClone(sections);
    let targetText: string | undefined;
    outer: for (const section of renamed) {
      for (const node of section.root) {
        if (node.type === "paragraph" && node.runs[0] !== undefined && node.runs[0].text.trim().length > 20) {
          node.runs[0].fontFamily = NONEXISTENT_FAMILY;
          targetText = node.runs[0].text;
          break outer;
        }
      }
    }
    if (targetText === undefined) throw new Error("fixture has no paragraph run to rename");
    await view.evaluate(`(async () => {
      const face = new FontFace(${JSON.stringify(ALIAS)}, "local('Courier New')");
      await face.load();
      document.fonts.add(face);
    })()`);
    await view.evaluate(`renderFixture(${JSON.stringify(renamed)})`);
    const result = await view.evaluate(`(() => {
      const span = [...document.querySelectorAll("span")].find((el) => el.textContent === ${JSON.stringify(targetText)});
      const size = getComputedStyle(span).fontSize;
      const courier = document.createElement("span");
      courier.style.cssText = "position:absolute;visibility:hidden;font:" + size + " 'Courier New';white-space:pre";
      courier.textContent = span.textContent;
      document.body.append(courier);
      const sansSerif = document.createElement("span");
      sansSerif.style.cssText = "position:absolute;visibility:hidden;font:" + size + " sans-serif;white-space:pre";
      sansSerif.textContent = span.textContent;
      document.body.append(sansSerif);
      return {
        computedFamily: getComputedStyle(span).fontFamily,
        spanWidth: span.getBoundingClientRect().width,
        courierWidth: courier.getBoundingClientRect().width,
        sansSerifWidth: sansSerif.getBoundingClientRect().width,
      };
    })()`) as { computedFamily: string; spanWidth: number; courierWidth: number; sansSerifWidth: number };
    // The stack names the requested family, its alias, then the generic —
    // in that order, so a real "Definitely Not A Real Font Zzqx" would still win.
    expect(result.computedFamily).toBe(`"${NONEXISTENT_FAMILY}", "${ALIAS}", sans-serif`);
    // Nothing answers to the requested name or its alias by coincidence: the
    // painted width is Courier's (monospace, so unmistakably different from
    // any sans-serif fallback), not the generic sans-serif's.
    const gap = Math.abs(result.courierWidth - result.sansSerifWidth);
    expect(gap).toBeGreaterThan(10);
    expect(result.spanWidth).toBeCloseTo(result.courierWidth, 0);
  } finally {
    server.stop();
  }
}, 30000);

/**
 * A Latin-only substitute (Arimo for Arial) has no Arabic glyphs, so `aared-fonts`
 * adds a second, `unicode-range`-limited Arabic face to the same alias
 * (DECISION_LOG §45). Registered exactly the way the resolver does it — added
 * to `document.fonts`, never `.load()`ed — under an alias no font answers to:
 * a Courier face for non-Arabic text and a much wider, calibrated face
 * (`size-adjust`) for Arabic. Arabic paints in the Arabic face and Latin in
 * the Latin face; the Arabic face is fetched, and waited for before layout
 * measures, only when the document has Arabic text — a Latin-only document
 * never downloads it.
 */
visual("visual DOM output: an Arabic-limited fallback face paints only Arabic and is fetched only when the document has Arabic", async () => {
  const NONEXISTENT_FAMILY = "Definitely Not A Real Font Zzqx";
  const ALIAS = fontAlias(NONEXISTENT_FAMILY);
  const ARABIC_RANGE = "U+0600-06FF, U+0750-077F, U+0870-08FF, U+FB50-FDFF, U+FE70-FEFF";
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const source = await parseDocxSections(pkg, await indexDocx(pkg));
  // Any paragraph, at any depth (tables, lists), with a run of only Latin or only Arabic text:
  // reduced to that one run, so the rendered span's text is exactly the run's.
  type Paragraph = { type: "paragraph"; runs: Array<{ text: string; fontFamily?: string }> };
  const found: Paragraph[] = [];
  const walk = (value: unknown): void => {
    if (value === null || typeof value !== "object") return;
    if (Array.isArray(value)) {
      for (const item of value) walk(item);
      return;
    }
    const record = value as Record<string, unknown>;
    if (record.type === "paragraph" && Array.isArray(record.runs)) found.push(record as unknown as Paragraph);
    for (const key in record) walk(record[key]);
  };
  walk(source);
  const onlyRun = (paragraph: Paragraph, pattern: RegExp): Paragraph | undefined => {
    const run = paragraph.runs.find((candidate) => pattern.test(candidate.text));
    return run === undefined ? undefined : { ...paragraph, runs: [run] };
  };
  const latin = found.map((paragraph) => onlyRun(paragraph, /^[A-Za-z ,.]{30,}$/)).find((paragraph) => paragraph !== undefined);
  const arabic = found.map((paragraph) => onlyRun(paragraph, /^[\u0600-\u06FF\u060C ,.]{20,}$/)).find((paragraph) => paragraph !== undefined);
  if (latin === undefined || arabic === undefined) throw new Error("fixture has no Latin-only and Arabic-only run");
  const runText = (paragraph: Paragraph): string => paragraph.runs[0]!.text;
  const document = (nodes: unknown[]): unknown => {
    const section = structuredClone({ properties: source[0]!.properties, root: nodes });
    renameAllFonts(section, NONEXISTENT_FAMILY);
    return [section];
  };
  const latinOnly = JSON.stringify(document([latin]));
  const both = JSON.stringify(document([latin, arabic]));
  const renderer = await browserModule(`${import.meta.dir}/../src/docx/index.ts`);
  const arabicBytes = await Bun.file("C:/Windows/Fonts/simpo.ttf").arrayBuffer();
  let arabicFetches = 0;
  const page = `<!doctype html><style>html,body{margin:0}#host{width:1200px}</style><main id="host"></main><script type="module">
        import { render } from "/docx.mjs";
        window.renderedFixture = null;
        window.renderFlowFixture = async (model) => {
          window.renderedFixture?.destroy();
          window.renderedFixture = await render({ document: model, target: document.querySelector("#host") });
        };
      </script>`;
  const server = Bun.serve({
    port: 0,
    async fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/docx.mjs") return new Response(renderer, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      if (url.pathname === "/arabic-face") {
        arabicFetches++;
        // A real, delayed network fetch, so measuring before it lands would read the fallback.
        await new Promise((resolve) => setTimeout(resolve, 200));
        return new Response(arabicBytes, { headers: { "content-type": "font/ttf" } });
      }
      return new Response(page, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1000, height: 900 });
    await view.navigate(server.url.href);
    await view.evaluate(`(async () => {
      const latinFace = new FontFace(${JSON.stringify(ALIAS)}, "local('Courier New')", { unicodeRange: "U+0000-05FF" });
      await latinFace.load();
      document.fonts.add(latinFace);
      // Left unloaded on purpose: this is how the resolver leaves an Arabic-limited face.
      document.fonts.add(new FontFace(${JSON.stringify(ALIAS)}, "url(/arabic-face)", { unicodeRange: ${JSON.stringify(ARABIC_RANGE)}, sizeAdjust: "150%" }));
    })()`);

    // A document with no Arabic text never fetches the Arabic face.
    await view.evaluate(`renderFlowFixture(${latinOnly})`);
    expect(arabicFetches).toBe(0);

    // With Arabic text, the render waits for the face before it settles.
    await view.evaluate(`renderFlowFixture(${both})`);
    expect(arabicFetches).toBeGreaterThan(0);
    const result = await view.evaluate(`(async () => {
      const control = new FontFace("ctl", "url(/arabic-face)", { sizeAdjust: "150%" });
      document.fonts.add(control);
      await control.load();
      // Each control is measured at the size of the span it is compared with: the two runs
      // come from different paragraphs.
      const measure = (span, family, text = span.textContent) => {
        const probe = document.createElement("span");
        probe.style.cssText = "position:absolute;visibility:hidden;white-space:pre;font:" + getComputedStyle(span).fontSize + " " + family;
        probe.textContent = text;
        document.body.append(probe);
        return probe.getBoundingClientRect().width;
      };
      // Under the alias, spaces and punctuation belong to the Latin face's range (Courier here,
      // Arimo in the real thing — the same as Arial's own spaces in Arabic text), so the expected
      // Arabic width is the Arabic letters in the Arabic face plus everything else in the Latin one.
      const expectedArabic = (span) => span.textContent.match(/[\\u0600-\\u06FF]+|[^\\u0600-\\u06FF]+/g)
        .reduce((sum, segment) => sum + measure(span, /[\\u0600-\\u06FF]/.test(segment) ? "ctl" : "'Courier New'", segment), 0);
      const spans = [...document.querySelectorAll("span")];
      const latinSpan = spans.find((el) => el.textContent === ${JSON.stringify(runText(latin))});
      const arabicSpan = spans.find((el) => el.textContent === ${JSON.stringify(runText(arabic))});
      return {
        latinWidth: latinSpan.getBoundingClientRect().width,
        latinControl: measure(latinSpan, "'Courier New'"),
        arabicWidth: arabicSpan.getBoundingClientRect().width,
        arabicControl: expectedArabic(arabicSpan),
        arabicGeneric: measure(arabicSpan, "sans-serif"),
      };
    })()`) as { latinWidth: number; latinControl: number; arabicWidth: number; arabicControl: number; arabicGeneric: number };
    expect(result.latinWidth).toBeCloseTo(result.latinControl, 0);
    // The painted Arabic is the calibrated Arabic face, not the generic fallback a face that
    // was registered but never waited for would leave.
    expect(Math.abs(result.arabicControl - result.arabicGeneric)).toBeGreaterThan(20);
    expect(result.arabicWidth).toBeCloseTo(result.arabicControl, 0);
  } finally {
    server.stop();
  }
}, 30000);

/**
 * The fixture's anchored pictures (`wp:anchor`, pages 5 and 6), read back from
 * the layout DOM against Word's own export of the same file
 * (tools/word-render, word-spans.json). Word centres a column-centred picture
 * on the text column; runs text down both sides of a square or through
 * picture with room on each; keeps a top-and-bottom picture on a band of its
 * own; and starts the paragraph after a picture that hangs below its text at
 * the ordinary paragraph step (Word's line tops 187 and 268 on page 6: 81px).
 */
visual("visual DOM output: DOCX anchored pictures wrap text as Word places them", async () => {
  const { server, sections } = await serveDocx(FIXTURE_PATH, `<!doctype html><style>html,body{margin:0}#host{width:1300px}</style><main id="host"></main><script type="module">
        import { render } from "/docx.mjs";
        window.probe = async (model) => {
          const handle = await render({
            document: model, target: document.querySelector("#host"), pagination: { mode: "layout" }, zoom: 1,
            mediaResolver: async (id) => (await fetch("/media/" + id.split("/").at(-1))).arrayBuffer(),
          });
          await document.fonts.ready;
          await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
          const pages = [...document.querySelectorAll('[data-aared-element="page"]')];
          const read = (page) => {
            const origin = page.getBoundingClientRect();
            const body = page.querySelector('[data-aared-element="page-body"]').getBoundingClientRect();
            return [...page.querySelectorAll('[data-aared-element="anchored-image"]')].map((element) => {
              const picture = element.matches("img, [data-aared-image-pending]") ? element : element.querySelector("img, [data-aared-image-pending]");
              const box = picture.getBoundingClientRect();
              const paragraph = element.closest("p");
              // Every piece of text in the paragraph beside the picture's rows, as painted.
              const text = [];
              const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
              for (let node = walker.nextNode(); node !== null; node = walker.nextNode()) {
                if (node.parentElement.closest('[hidden], [aria-hidden="true"]') !== null || node.textContent.trim() === "") continue;
                const range = document.createRange();
                range.selectNodeContents(node);
                for (const rect of range.getClientRects()) if (rect.width > 0) text.push({ left: rect.left, right: rect.right, top: rect.top, bottom: rect.bottom });
              }
              const besides = text.filter((rect) => rect.bottom > box.top + 2 && rect.top < box.bottom - 2);
              const next = paragraph.nextElementSibling;
              return {
                placement: element.dataset.aaredPlacement,
                wrap: element.dataset.aaredWrap,
                centreOffset: (box.left + box.right) / 2 - (body.left + body.right) / 2,
                overlapsText: besides.some((rect) => rect.right > box.left + 1 && rect.left < box.right - 1),
                textLeft: besides.some((rect) => rect.right <= box.left),
                textRight: besides.some((rect) => rect.left >= box.right),
                layer: picture.dataset.aaredLayer ?? null,
                zIndex: getComputedStyle(picture).zIndex,
                nextStep: next === null ? null : next.getBoundingClientRect().top - paragraph.getBoundingClientRect().top,
              };
            });
          };
          const result = { pageCount: handle.pageCount, five: read(pages[4]), six: read(pages[5]), canvasCount: document.querySelectorAll("canvas").length };
          handle.destroy();
          return result;
        };
      </script>`);
  try {
    await using view = new Bun.WebView({ width: 1300, height: 900 });
    await view.navigate(server.url.href);
    type Read = { placement: string; wrap: string; centreOffset: number; overlapsText: boolean; textLeft: boolean; textRight: boolean; layer: string | null; zIndex: string; nextStep: number | null };
    const result = await view.evaluate(`probe(${JSON.stringify(sections)})`) as { pageCount: number; five: Read[]; six: Read[]; canvasCount: number };
    expect(result.pageCount).toBe(6);
    expect(result.canvasCount).toBe(0);
    // Page 5: square at each column edge, top and bottom, behind text.
    expect(result.five.map((read) => [read.wrap, read.placement])).toEqual([["square", "wrap"], ["square", "wrap"], ["topAndBottom", "band"], ["none", "overlay"]]);
    // Page 6: tight at the edge, then square and through in the centre, text down both sides.
    expect(result.six.map((read) => [read.wrap, read.placement])).toEqual([["tight", "wrap"], ["square", "split"], ["through", "split"]]);
    for (const read of [...result.five, ...result.six]) {
      if (read.placement !== "overlay") expect(read.overlapsText).toBe(false);
    }
    // A wrapped picture at the left column edge has text on its right only, and the other way round.
    expect([result.five[0]!.textLeft, result.five[0]!.textRight]).toEqual([false, true]);
    expect([result.five[1]!.textLeft, result.five[1]!.textRight]).toEqual([true, false]);
    // The band keeps its rows to itself; the split runs text on both sides.
    expect([result.five[2]!.textLeft, result.five[2]!.textRight]).toEqual([false, false]);
    for (const read of result.six.slice(1)) expect([read.textLeft, read.textRight]).toEqual([true, true]);
    // Column-centred pictures sit on the column's centre, as in Word.
    for (const read of [result.five[2]!, result.five[3]!, result.six[1]!, result.six[2]!]) expect(Math.abs(read.centreOffset)).toBeLessThanOrEqual(1);
    // Behind text: under the paragraph's text, which runs straight across it.
    expect(result.five[3]).toMatchObject({ layer: "behind", zIndex: "-1", overlapsText: true });
    // The paragraph after the centred square starts where Word starts it (81px on), not below the picture's wrap distance.
    expect(Math.abs(result.six[1]!.nextStep! - 81)).toBeLessThanOrEqual(3);
  } finally {
    server.stop();
  }
}, 30000);

/** Every run's `fontFamily` set to `family`, deep-walked across sections including tables/lists/footnotes/headers/footers. */
function renameAllFonts(sections: unknown, family: string): void {
  const visit = (value: unknown): void => {
    if (value === null || typeof value !== "object") return;
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    const record = value as Record<string, unknown>;
    if (typeof record.fontFamily === "string") record.fontFamily = family;
    for (const key in record) visit(record[key]);
  };
  visit(sections);
}

/**
 * Layout pagination waits for a delivered substitute's `FontFace` to finish
 * loading before it measures — not just registers it and measures whatever
 * was painted first (DECISION_LOG §45). `aared-fonts`'s resolver sources a
 * substitute from a `url()` (Google Fonts), a genuine network fetch, so this
 * serves a real font from a deliberately delayed same-origin route and
 * renames every run in the document to a nonexistent family. Paginating
 * immediately after registering the face (never awaiting its `.load()`,
 * exactly how the resolver leaves it — DECISION_LOG §45) must still measure
 * the *loaded* face once the render settles, matching a render started only
 * once the face was already fully loaded.
 */
visual("visual DOM output: layout pagination waits for a font to finish loading before it measures", async () => {
  const NONEXISTENT_FAMILY = "Definitely Not A Real Font Zzqx";
  const ALIAS = fontAlias(NONEXISTENT_FAMILY);
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const sections = await parseDocxSections(pkg, await indexDocx(pkg));
  const renderer = await browserModule(`${import.meta.dir}/../src/docx/index.ts`);
  const fontBytes = await Bun.file("C:/Windows/Fonts/Cour.ttf").arrayBuffer();
  const page = `<!doctype html><style>html,body{margin:0}#host{width:1200px}</style><main id="host"></main><script type="module">
        import { render } from "/docx.mjs";
        window.renderedFixture = null;
        window.renderLayoutFixture = async (model) => {
          window.renderedFixture?.destroy();
          window.renderedFixture = await render({ document: model, target: document.querySelector("#host"), pagination: { mode: "layout" } });
          return { pageCount: window.renderedFixture.pageCount };
        };
      </script>`;
  const server = Bun.serve({
    port: 0,
    async fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/docx.mjs") return new Response(renderer, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      if (url.pathname === "/font") {
        // A real network fetch, deliberately delayed: long enough that a
        // synchronous, pre-fix measurement (which runs with zero await
        // before this resolves) would reliably see the font still loading.
        await new Promise((resolve) => setTimeout(resolve, 200));
        return new Response(fontBytes, { headers: { "content-type": "font/ttf" } });
      }
      return new Response(page, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1000, height: 900 });
    await view.navigate(server.url.href);
    const renamed = structuredClone(sections);
    renameAllFonts(renamed, NONEXISTENT_FAMILY);
    const model = JSON.stringify(renamed);

    const noFontPages = ((await view.evaluate(`renderLayoutFixture(${model})`)) as { pageCount: number }).pageCount;

    // Registers the face and triggers the render in the same synchronous
    // tick, `.load()` not awaited first — mirrors the resolver exactly.
    const coldPages = ((await view.evaluate(`(async () => {
      const face = new FontFace(${JSON.stringify(ALIAS)}, "url(/font)");
      document.fonts.add(face);
      face.load();
      return renderLayoutFixture(${model});
    })()`)) as { pageCount: number }).pageCount;

    // A fresh face, this time `.load()` awaited to completion before the render starts.
    const warmPages = ((await view.evaluate(`(async () => {
      const face = new FontFace(${JSON.stringify(ALIAS)}, "url(/font)");
      document.fonts.add(face);
      await face.load();
      return renderLayoutFixture(${model});
    })()`)) as { pageCount: number }).pageCount;

    // Meaningful test: a monospace substitute across the whole document
    // really does change the page count from the generic-fallback baseline.
    expect(coldPages).not.toBe(noFontPages);
    // The point: whether the caller awaited `.load()` must not matter.
    expect(coldPages).toBe(warmPages);
  } finally {
    server.stop();
  }
}, 30000);

/**
 * `DocxRenderHandle.warnings` names a family that painted with nothing of
 * its own — the browser's generic fallback, not even a delivered substitute
 * (DECISION_LOG §45). A working, real font and a working, delivered
 * substitute both produce no warning: only genuine unavailability does.
 */
visual("visual DOM output: DocxRenderHandle.warnings names fonts that painted with the generic fallback", async () => {
  const NONEXISTENT_FAMILY = "Definitely Not A Real Font Zzqx";
  const ALIAS = fontAlias(NONEXISTENT_FAMILY);
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const sections = await parseDocxSections(pkg, await indexDocx(pkg));
  const renderer = await browserModule(`${import.meta.dir}/../src/docx/index.ts`);
  const fontBytes = await Bun.file("C:/Windows/Fonts/Cour.ttf").arrayBuffer();
  const page = `<!doctype html><style>html,body{margin:0}#host{width:1200px}</style><main id="host"></main><script type="module">
        import { render } from "/docx.mjs";
        window.renderedFixture = null;
        window.renderFlowFixture = async (model) => {
          window.renderedFixture?.destroy();
          window.renderedFixture = await render({ document: model, target: document.querySelector("#host") });
          return window.renderedFixture.warnings;
        };
      </script>`;
  const server = Bun.serve({
    port: 0,
    async fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/docx.mjs") return new Response(renderer, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      if (url.pathname === "/font") return new Response(fontBytes, { headers: { "content-type": "font/ttf" } });
      return new Response(page, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1000, height: 900 });
    await view.navigate(server.url.href);

    // The unmodified fixture: every family is real, nothing warns.
    const realFontWarnings = (await view.evaluate(`renderFlowFixture(${JSON.stringify(sections)})`)) as Array<{ code: string }>;
    expect(realFontWarnings.filter((w) => w.code === "font-unavailable")).toEqual([]);

    // Renamed to a family nobody has, with nothing delivered for it: warns.
    const renamed = structuredClone(sections);
    renameAllFonts(renamed, NONEXISTENT_FAMILY);
    const missingWarnings = (await view.evaluate(`renderFlowFixture(${JSON.stringify(renamed)})`)) as Array<{ code: string; message: string }>;
    const missing = missingWarnings.find((w) => w.code === "font-unavailable");
    expect(missing).toBeDefined();
    expect(missing!.message).toContain(NONEXISTENT_FAMILY);

    // The same rename, this time with a substitute registered and loaded first: no warning.
    const substitutedWarnings = (await view.evaluate(`(async () => {
      const face = new FontFace(${JSON.stringify(ALIAS)}, "url(/font)");
      document.fonts.add(face);
      await face.load();
      return renderFlowFixture(${JSON.stringify(renamed)});
    })()`)) as Array<{ code: string }>;
    expect(substitutedWarnings.filter((w) => w.code === "font-unavailable")).toEqual([]);
  } finally {
    server.stop();
  }
}, 30000);
