import { expect, test } from "bun:test";
import { AaredError } from "../src/core/index.js";
import { open } from "../src/docx/index.js";

const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-docx-bidi-coverage.docx`;
const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-pptx-rtl-coverage.pptx`;

async function rejection(promise: Promise<unknown>): Promise<Error> {
  try {
    await promise;
  } catch (error) {
    return error instanceof Error ? error : new Error(String(error));
  }
  throw new Error("expected the request to reject");
}

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
    expect(headings[0]?.text).toContain("Project Overview");
    expect(headings[0]?.level).toBe(1);
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
    expect(document.search("")).toEqual([]);
    const matches = document.search("Objectives");
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

test("rejects non-DOCX files and exposes document capabilities", async () => {
  const error = await rejection(open(await Bun.file(PPTX_FIXTURE).arrayBuffer()));
  expect((error as AaredError).code).toBe("unsupported-format");
  expect(error.message).toContain("aared/pptx");

  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    expect(document.capabilities).toEqual({
      textSelection: true,
      media: true,
      document: { flow: true, layout: true, footnotes: true },
      editing: false,
    });
    document.close();
    // Cached sections stay readable; session-backed work rejects.
    expect(document.sections.list().length).toBeGreaterThan(0);
    const mediaError = await rejection(document.media.get("word/media/image1.png"));
    expect(mediaError.message).toContain("document is closed");
  } finally {
    document.close();
  }
});
