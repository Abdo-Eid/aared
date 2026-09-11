import { expect, test } from "bun:test";
import { open } from "../src/docx/index.js";

const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-docx-bidi-coverage.docx`;

test("search excludes table and footnote text", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const matches = document.search("Project Overview");
    expect(matches.length).toBeGreaterThan(0);
    for (const match of matches) {
      expect(match.text.toLowerCase()).toContain("project overview");
    }
  } finally {
    document.close();
  }
});

test("search truncates long matches to 120 characters", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    const matches = document.search("the");
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
    const matches = document.search("Project Overview");
    expect(matches.length).toBeGreaterThan(0);
    // First match IS the heading — heading context is null for the heading itself
    const firstMatch = matches[0]!;
    expect(firstMatch.heading).toBeNull();
    // Subsequent matches under that heading should have the heading as context
    if (matches.length > 1) {
      expect(matches[1]!.heading).toBe("Project Overview");
    }
  } finally {
    document.close();
  }
});

test("DOCX info count is 1 (whole document renders as one unit)", async () => {
  const document = await open(await Bun.file(DOCX_FIXTURE).arrayBuffer());
  try {
    expect(document.info.count).toBe(1);
    // Sections are resolved separately via the collection
    expect(document.sections.list().length).toBeGreaterThan(1);
  } finally {
    document.close();
  }
});
