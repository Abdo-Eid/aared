import { afterEach, expect, spyOn, test } from "bun:test";
import { open as openDocx } from "../src/docx/index.js";
import { open as openPptx } from "../src/pptx/index.js";
import { open as openXlsx } from "../src/xlsx/index.js";
import { WorkerSession } from "../src/core/session.js";
import type { DocumentSection, FlowNode } from "../src/core/model.js";
import { headingOutline } from "../src/docx/outline.js";
import { buildDocxSearchIndex } from "../src/docx/search.js";

const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;
const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;

const spies: Array<{ mockRestore(): void }> = [];
afterEach(() => {
  for (const spy of spies.splice(0)) spy.mockRestore();
});

/** Spies on both reads: search should only ever ask the Worker for words, never for a model. */
function spyReads(): { text: ReturnType<typeof spyOn>; unit: ReturnType<typeof spyOn> } {
  const text = spyOn(WorkerSession.prototype, "text");
  const unit = spyOn(WorkerSession.prototype, "unit");
  spies.push(text, unit);
  return { text, unit };
}

test("PPTX search reads each slide's words once, never its model, then answers from its text index", async () => {
  const deck = await openPptx(await Bun.file(PPTX_FIXTURE).arrayBuffer());
  try {
    const { text, unit } = spyReads();
    const first = await deck.search("الإنتاج");
    expect(first.length).toBeGreaterThan(0);
    expect(text).toHaveBeenCalledTimes(deck.slides.list().length);

    const second = await deck.search("الإنتاج");
    expect(second).toEqual(first);
    // A different query reuses the same index too.
    await deck.search("no-such-text-anywhere");
    expect(text).toHaveBeenCalledTimes(deck.slides.list().length);
    expect(unit).not.toHaveBeenCalled();
  } finally {
    deck.close();
  }
});

test("XLSX search reads only the words of the sheet it is asked about, once, then answers from its cell index", async () => {
  const workbook = await openXlsx(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const { text, unit } = spyReads();
    const summary = workbook.sheets.get("Summary");
    const first = await summary.search("budget");
    expect(first.length).toBeGreaterThan(0);
    // One sheet read of the four; the other three stay unparsed.
    expect(text).toHaveBeenCalledTimes(1);
    expect(text.mock.calls[0]![0]).toBe(summary.info.index);

    expect(await summary.search("budget")).toEqual(first);
    expect(await summary.search("a different query")).toEqual([]);
    expect(text).toHaveBeenCalledTimes(1);

    // A second sheet is its own read, and its text is not Summary's.
    expect(await workbook.sheets.get("Shot Log").search("budget")).toEqual([]);
    expect(text).toHaveBeenCalledTimes(2);
    expect(unit).not.toHaveBeenCalled();
  } finally {
    workbook.close();
  }
});

test("a DOCX session refuses the text read: its page already holds the whole document", async () => {
  const session = await WorkerSession.openAs("docx", await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const error = await session.text(0).catch((e: unknown) => e);
    expect((error as { code?: string }).code).toBe("unsupported-feature");
  } finally {
    session.close();
  }
});

test("XLSX search of the sheet on screen takes its words from the sheet already parsed, with the same matches", async () => {
  const fresh = await openXlsx(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  const viewed = await openXlsx(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    // Reading the sheet (as a render does) leaves it in the Worker's cache; search then reuses it.
    await viewed.sheets.get(0).read();
    expect(await viewed.sheets.get(0).search("الإنتاج")).toEqual(await fresh.sheets.get(0).search("الإنتاج"));
  } finally {
    fresh.close();
    viewed.close();
  }
});

test("PPTX search returns each occurrence with a short snippet and its slide", async () => {
  const deck = await openPptx(await Bun.file(PPTX_FIXTURE).arrayBuffer());
  try {
    const matches = await deck.search("الإنتاج");
    expect(matches.length).toBeGreaterThan(0);
    for (const match of matches) {
      expect(match.text.length).toBeLessThanOrEqual(120);
      // A match names its slide the way every unit is named: goTo key plus label.
      expect(match.slide.title).toBe(deck.slides.get(match.slide.index).info.title);
    }
    expect(await deck.search("")).toEqual([]);
  } finally {
    deck.close();
  }
});

test("PPTX search returns all five Arabic bullet occurrences on the same slide", async () => {
  const deck = await openPptx(await Bun.file(PPTX_FIXTURE).arrayBuffer());
  try {
    const matches = await deck.search("نقطة");
    expect(matches).toHaveLength(5);
    expect(matches.every((match) => match.slide.index === 4)).toBe(true);
    expect(matches.every((match) => match.text.includes("نقطة"))).toBe(true);
    expect(await deck.search("نُقْـطَة")).toEqual(matches);
    // A repeated letter in one paragraph is also multiple occurrences.
    const letters = await deck.search("ا");
    const texts = letters.map((match) => `${match.slide.index}:${match.text}`);
    expect(new Set(texts).size).toBeLessThan(texts.length);
  } finally { deck.close(); }
});

test("PPTX search ignores alef forms, tashkeel and digit script, and still reports the slide's own text", async () => {
  const deck = await openPptx(await Bun.file(PPTX_FIXTURE).arrayBuffer());
  try {
    const exact = await deck.search("الإنتاج");
    expect(exact).toEqual([{ slide: { index: 0, title: "Slide 1" }, text: "تقرير الإنتاج الأسبوعي" }]);
    // Bare alef, and the same word carrying tashkeel and a tatweel.
    expect(await deck.search("الانتاج")).toEqual(exact);
    expect(await deck.search("الإِنْتَاجـ")).toEqual(exact);
    // "لقطة ١٢" is written with Arabic-Indic digits on slide 1; "2.1%" with Western ones on slide 10.
    const shot = await deck.search("لقطة 12");
    expect(shot.map((match) => match.slide.index)).toEqual([1]);
    expect(shot[0]!.text).toBe("Shot 12 — لقطة ١٢");
    expect((await deck.search("٢٫١%")).map((match) => match.slide.index)).toContain(10);
    // Folding widens matches; it does not make unrelated text match.
    expect(await deck.search("الانتاجات")).toEqual([]);
    expect(await deck.search("َ")).toEqual([]);
  } finally {
    deck.close();
  }
});

test("XLSX search ignores alef forms, tashkeel and digit script, and reports the cell's own text", async () => {
  const workbook = await openXlsx(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const sheet = workbook.sheets.get(0);
    const exact = await sheet.search("الإنتاج");
    expect(exact.map((match) => `${match.sheet.title}!${match.address}`)).toEqual(["الميزانية!A1", "الميزانية!B11"]);
    expect(await sheet.search("الانتاج")).toEqual(exact);
    expect(await sheet.search("الإِنْتَاج")).toEqual(exact);
    expect(exact[0]!.text).toBe("ميزانية الإنتاج الشهرية");
    // "91,200 EGP" is Western digits and a Western comma; the query uses Arabic-Indic ones and the Arabic thousands sign.
    const budget = await sheet.search("٩١٬٢٠٠ egp");
    expect(budget.map((match) => `${match.sheet.title}!${match.address}`)).toEqual(["الميزانية!C5"]);
    expect(budget[0]!.text).toBe("91,200 EGP");
  } finally {
    workbook.close();
  }
});

test("DOCX search ignores alef forms, tashkeel and digit script, and reports the block's own text", async () => {
  const document = await openDocx(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const exact = await document.search("مقابلة");
    expect(exact).toHaveLength(1);
    // The same word with tashkeel and a tatweel finds the same table cell.
    expect(await document.search("مُقَابَلَـة")).toEqual(exact);
    // "H.264, 24 Mb/s" is a table cell written with Western digits.
    const spec = await document.search("H.264, ٢٤ Mb/s");
    expect(spec).toHaveLength(1);
    expect(spec[0]).toMatchObject({ kind: "table", text: "H.264, 24 Mb/s" });
    expect(await document.search("َ")).toEqual([]);
  } finally {
    document.close();
  }
});

test("DOCX search returns each distinct header story once, even when inherited by later sections", async () => {
  const document = await openDocx(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const matches = (await document.search("NILE")).filter((match) => match.kind === "header");
    expect(matches).toHaveLength(2);
    expect(matches.map(({ section, heading, text, story }) => ({ section, heading, text, story }))).toEqual([
      {
        section: 0,
        heading: null,
        text: "NILE FRAME STUDIO\tاستوديو نايل فريم",
        story: { part: "word/header2.xml", variant: "default" },
      },
      {
        section: 1,
        heading: null,
        text: "NILE FRAME STUDIO\tاستوديو نايل فريم",
        story: { part: "word/header4.xml", variant: "default" },
      },
    ]);
    const footers = (await document.search("confidential")).filter((match) => match.kind === "footer");
    expect(footers.length).toBeGreaterThan(0);
    expect(new Set(footers.map((match) => match.story?.part)).size).toBe(footers.length);
    expect(footers.every((match) => match.heading === null && match.story !== undefined)).toBe(true);
  } finally {
    document.close();
  }
});

test("DOCX search indexes distinct first/even stories and nested header/footer content", async () => {
  const document = await openDocx(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const base = document.sections.get(0);
    const paragraph = (text: string): FlowNode => ({ type: "paragraph", runs: [{ text }], direction: "ltr" });
    const sharedDefault: FlowNode[] = [paragraph("shared-default-marker")];
    const first: FlowNode[] = [paragraph("first-variant-marker")];
    const even: FlowNode[] = [
      { type: "textBox", width: 2, height: 1, direction: "ltr", content: [paragraph("even-textbox-marker")] },
      {
        type: "table",
        direction: "ltr",
        columns: [1],
        rows: [{ cells: [{ nodes: [paragraph("even-table-marker")] }] }],
      },
    ];
    const footer: FlowNode[] = [paragraph("footer-story-marker")];
    const sections: DocumentSection[] = [
      {
        ...base,
        properties: { ...base.properties, headerParts: { default: "word/header-shared.xml", first: "word/header-first.xml", even: "word/header-even.xml" }, footerParts: { default: "word/footer-shared.xml" } },
        root: [],
        header: sharedDefault,
        firstHeader: first,
        evenHeader: even,
        footer,
      },
      {
        ...base,
        properties: { ...base.properties, headerParts: { default: "word/header-new.xml" }, footerParts: {} },
        root: [],
        header: [paragraph("new-default-marker")],
        firstHeader: first,
        evenHeader: even,
        footer,
      },
    ];
    const index = buildDocxSearchIndex(sections, headingOutline(sections));
    const matches = (query: string) => index.filter((entry) => entry.key.includes(query)).map((entry) => entry.match);
    expect(matches("shared-default-marker")).toMatchObject([
      { kind: "header", section: 0, heading: null, story: { part: "word/header-shared.xml", variant: "default" } },
    ]);
    expect(matches("new-default-marker")).toMatchObject([
      { kind: "header", section: 1, story: { part: "word/header-new.xml", variant: "default" } },
    ]);
    expect(matches("first-variant-marker")).toMatchObject([
      { kind: "header", section: 0, story: { part: "word/header-first.xml", variant: "first" } },
    ]);
    expect(matches("even-textbox-marker")).toMatchObject([
      { kind: "header", section: 0, story: { part: "word/header-even.xml", variant: "even" } },
    ]);
    expect(matches("even-table-marker")).toMatchObject([
      { kind: "header", section: 0, story: { part: "word/header-even.xml", variant: "even" } },
    ]);
    expect(matches("footer-story-marker")).toMatchObject([
      { kind: "footer", section: 0, story: { part: "word/footer-shared.xml", variant: "default" } },
    ]);
  } finally {
    document.close();
  }
});

test("an aborted search leaves no partial index behind", async () => {
  const deck = await openPptx(await Bun.file(PPTX_FIXTURE).arrayBuffer());
  try {
    const controller = new AbortController();
    controller.abort();
    const error = await deck.search("الإنتاج", { signal: controller.signal }).catch((e: unknown) => e);
    expect((error as Error).name).toBe("AbortError");
    // The next, un-aborted search still reads and finds everything.
    expect((await deck.search("الإنتاج")).length).toBeGreaterThan(0);
  } finally {
    deck.close();
  }
});
