import { expect, test } from "bun:test";
import { isAbort, isAaredError } from "../src/core/index.js";
import { open } from "../src/docx/index.js";
import type { DocxRenderHandle } from "../src/docx/index.js";

const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

test("search finds table-cell text with cell coordinates", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const [match, ...rest] = await document.search("H.264, 24 Mb/s");
    expect(rest).toEqual([]);
    expect(match).toEqual({
      kind: "table",
      section: 0,
      heading: { index: 3, title: "2. Technical Specifications" },
      text: "H.264, 24 Mb/s",
      table: { index: 1, row: 2, col: 2 },
    });

    // Arabic cell text in the landscape appendix (section 1, its own table index).
    const arabic = await document.search("مقابلة");
    expect(arabic).toHaveLength(1);
    expect(arabic[0]).toMatchObject({
      kind: "table",
      section: 1,
      heading: { index: 6, title: "Appendix A — Shot List, Episode 04" },
      table: { index: 0, row: 3, col: 2 },
    });
  } finally {
    document.close();
  }
});

test("a match's heading index is the outline index goToHeading takes", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const outline = document.headings.list();
    for (const match of await document.search("a")) {
      if (match.heading === null) continue;
      expect(outline[match.heading.index]?.title).toBe(match.heading.title);
    }
  } finally {
    document.close();
  }
});

test("search finds footnote text attributed to the referencing section and heading", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const [loudness] = await document.search("EBU R 128");
    expect(loudness).toMatchObject({
      kind: "footnote",
      section: 0,
      heading: { title: "2. Technical Specifications" },
      note: { id: "f1", label: "1" },
    });
    expect(loudness!.table).toBeUndefined();

    const [schedule] = await document.search("archive clearance");
    expect(schedule).toMatchObject({
      kind: "footnote",
      section: 0,
      heading: { title: "3. Schedule" },
      note: { id: "f2", label: "2" },
    });
    // Notes sit at the end of the last section's root, but never report it.
    expect(document.sections.list().length).toBeGreaterThan(1);
  } finally {
    document.close();
  }
});

test("search results are document-ordered, notes last, and safe to mutate", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const first = await document.search("a");
    const kinds = first.map((match) => match.kind);
    const isNote = (kind: string): boolean => kind === "footnote" || kind === "endnote";
    const firstNote = kinds.findIndex(isNote);
    expect(firstNote).toBeGreaterThan(0);
    expect(kinds.slice(firstNote).every(isNote)).toBe(true);

    // The index is cached, so results must be fresh copies each call.
    for (const match of first) {
      match.text = "mutated";
      if (match.table !== undefined) match.table.index = -1;
      if (match.heading !== null) match.heading.title = "mutated";
    }
    const second = await document.search("a");
    expect(second.map((match) => match.text)).not.toContain("mutated");
    expect(second.map((match) => match.table?.index)).not.toContain(-1);
    expect(second.map((match) => match.heading?.title)).not.toContain("mutated");
    expect(second).toHaveLength(first.length);
  } finally {
    document.close();
  }
});

test("search truncates long matches to 120 characters", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const matches = await document.search("the");
    for (const match of matches) {
      expect(match.text.length).toBeLessThanOrEqual(120);
    }
  } finally {
    document.close();
  }
});

test("search results include heading context", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const matches = await document.search("Project Overview");
    expect(matches.length).toBeGreaterThan(0);
    // The first match is the contents-page entry (before any heading), the
    // second is the heading itself — neither has a preceding heading.
    expect(matches[0]!.heading).toBeNull();
    expect(matches[1]!.heading).toBeNull();
    // A body match reports the nearest preceding heading as its context.
    const body = await document.search("Hold the audience");
    expect(body).toHaveLength(1);
    expect(body[0]!.heading?.title).toContain("1.1 Objectives");
  } finally {
    document.close();
  }
});

test("a document is one unit: one title, no sizes; the outline lives on document.headings", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    expect(document.info.count).toBe(1);
    // The unit's label is the file's own title (core properties `dc:title`).
    expect(document.info.titles).toEqual(["Production Brief — Sawt al-Nahr / موجز الإنتاج"]);
    expect(document.info.sizes).toBeUndefined();
    // Sections and headings are resolved separately via their collections.
    expect(document.sections.list().length).toBeGreaterThan(1);
    expect(document.headings.list().length).toBeGreaterThan(1);
  } finally {
    document.close();
  }
});

test("using a closed document rejects closed — a usage error, never a silent cancellation", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  document.close();
  let caught: unknown;
  try {
    await document.search("a");
  } catch (error) {
    caught = error;
  }
  expect(isAaredError(caught) && caught.code === "closed").toBe(true);
  expect(isAbort(caught)).toBe(false);
});

// Type-level: a render handle's pages are read-only to its consumers.
// `test/` is typechecked, so `bun run typecheck` enforces these.
function readonlyPages(handle: DocxRenderHandle): void {
  // @ts-expect-error pageCount is read-only
  handle.pageCount = 2;
  // @ts-expect-error pages is a read-only array
  handle.pages.push({ index: 0, width: 1, height: 1, orientation: "portrait" });
}
void readonlyPages;
