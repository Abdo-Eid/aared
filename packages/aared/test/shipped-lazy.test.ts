import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { basename, dirname, resolve } from "node:path";

const DIST = resolve(import.meta.dir, "../dist");

/** Walk tsdown's emitted static imports/re-exports, deliberately excluding import(...). */
function eagerGraph(entry: string): string[] {
  const seen = new Set<string>();
  const walk = (file: string): void => {
    if (seen.has(file)) return;
    seen.add(file);
    const text = readFileSync(file, "utf8");
    for (const match of text.matchAll(/(?:from|import)\s*["'`](\.{1,2}\/[^"'`]+)["'`]/g)) {
      walk(resolve(dirname(file), match[1]!));
    }
  };
  walk(resolve(DIST, entry));
  return [...seen];
}

test("shipped format entries defer the embed controller and preview adapters", () => {
  for (const format of ["pptx", "xlsx", "docx"]) {
    const graph = eagerGraph(`${format}/index.mjs`);
    expect(graph.length).toBeGreaterThan(1);
    expect(graph.some((file) => /^(?:embed-viewer|auto-viewer|viewer)-/.test(basename(file)))).toBe(false);
    expect(graph.map((file) => readFileSync(file, "utf8")).join("\n")).not.toContain("goTo() was called after destroy()");
  }
});

test("generic preview, hooks and the element also defer the controller", () => {
  for (const entry of ["preview", "react", "element"]) {
    expect(eagerGraph(`${entry}/index.mjs`).some((file) => basename(file).startsWith("embed-viewer-"))).toBe(false);
  }
});

test("shipped XLSX and DOCX keep effects lazy while synchronous PPTX retains its painter", () => {
  for (const format of ["pptx", "xlsx", "docx"]) {
    const graph = eagerGraph(`${format}/index.mjs`);
    expect(graph.some((file) => basename(file).startsWith("effect-paint-"))).toBe(format === "pptx");
    if (format === "xlsx") expect(graph.some((file) => basename(file).startsWith("transform-paint-"))).toBe(true);
  }
});
