import { expect, test } from "bun:test";
import { AaredError } from "../src/core/index.js";
import { open } from "../src/docx/index.js";
import { rejection } from "./harness.js";

const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;
const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

test("opens a document with sections and headings resolved together", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    expect(document.info.format).toBe("docx");
    const sections = document.sections.list();
    expect(sections.length).toBeGreaterThan(0);
    expect(document.sections.get(0)).toBe(sections[0]!);
    for (const section of sections) {
      expect(Array.isArray(section.root)).toBe(true);
      expect(section.properties.page.orientation).toMatch(/portrait|landscape/);
    }
    const headings = document.headings.list();
    expect(headings.length).toBeGreaterThan(0);
    expect(headings[0]).toMatchObject({ index: 0, level: 1, section: 0 });
    expect(headings[0]?.title).toContain("Project Overview");
    // Every entry's index is its own position: the one index space
    // goToHeading() and search() matches share.
    expect(headings.map((heading) => heading.index)).toEqual(headings.map((_, index) => index));
  } finally {
    document.close();
  }
});

test("per-section page setup stays on the sections; info names the one unit", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const [body, appendix] = document.sections.list();
    // Portrait A4 body, landscape appendix.
    expect(body!.properties.page.orientation).toBe("portrait");
    expect(appendix!.properties.page.orientation).toBe("landscape");
    // A section is not a unit, and a document has no single size.
    expect(document.info.count).toBe(1);
    expect(document.info.titles).toHaveLength(1);
    expect(document.info.sizes).toBeUndefined();
  } finally {
    document.close();
  }
});

test("finds a heading by index or by title — exact, then ignoring case, never a substring", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const [first] = document.headings.list();
    expect(document.headings.get(0)).toBe(first!);
    expect(document.headings.get(first!.title)).toBe(first!);
    expect(document.headings.get(first!.title.toUpperCase())).toBe(first!);
    expect(() => document.headings.get("Project")).toThrow("no heading titled");
  } finally {
    document.close();
  }
});

test("reports unknown sections and headings with not-found errors", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    expect(() => document.sections.get(99)).toThrow("outside this");
    expect(() => document.headings.get(-1)).toThrow();
  } finally {
    document.close();
  }
});

test("searches document text with heading context", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    expect(await document.search("")).toEqual([]);
    const matches = await document.search("Objectives");
    expect(matches.length).toBeGreaterThan(0);
    for (const match of matches) {
      expect(match.section).toBeGreaterThanOrEqual(0);
      expect(match.text.toLowerCase()).toContain("objectives");
    }
  } finally {
    document.close();
  }
});

test("serves image media through the media collection", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const mediaIds: string[] = [];
    const visit = (value: unknown): void => {
      if (Array.isArray(value)) {
        for (const item of value) visit(item);
        return;
      }
      if (value === null || typeof value !== "object") return;
      for (const [key, child] of Object.entries(value)) {
        if (key === "mediaId" && typeof child === "string" && child !== "") mediaIds.push(child);
        else visit(child);
      }
    };
    for (const section of document.sections.list()) visit(section.root);
    expect(mediaIds.length).toBeGreaterThan(0);
    const bytes = await document.media.get(mediaIds[0]!);
    expect(bytes.byteLength).toBeGreaterThan(0);
    const missing = await rejection(document.media.get("word/media/not-present.png"));
    expect((missing as AaredError).code).toBe("media-error");
  } finally {
    document.close();
  }
});

test("rejects non-DOCX files and closes cleanly", async () => {
  const error = await rejection(open(await Bun.file(PPTX_FIXTURE).arrayBuffer()));
  expect((error as AaredError).code).toBe("unsupported-format");
  expect(error.message).toContain("aared/pptx");

  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    document.close();
    // Cached sections stay readable; session-backed work rejects.
    expect(document.sections.list().length).toBeGreaterThan(0);
    const mediaError = await rejection(document.media.get("word/media/image1.png"));
    expect(mediaError.message).toContain("document is closed");
  } finally {
    document.close();
  }
});
