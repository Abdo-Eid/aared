import { expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import { join, relative, resolve } from "node:path";

/**
 * Shipped-output checks. `bun run build` must precede these: they read `dist/`
 * as a consumer receives it, not the sources.
 */

const DIST = resolve(import.meta.dir, "../dist");

function distFiles(dir = DIST): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) =>
    entry.isDirectory() ? distFiles(join(dir, entry.name)) : [join(dir, entry.name)],
  );
}

function read(...parts: string[]): string {
  return readFileSync(join(DIST, ...parts), "utf8");
}

test("shipped JavaScript carries no development diagnostics", () => {
  const scripts = distFiles().filter((file) => file.endsWith(".mjs"));
  expect(scripts.length).toBeGreaterThan(20);
  for (const file of scripts) {
    const text = readFileSync(file, "utf8");
    // `MARK_FALLBACKS` gates development chrome; tsdown defines it away.
    expect(text.includes("__AARED_MARK_FALLBACKS__"), relative(DIST, file)).toBe(false);
    expect(text.includes("MARK_FALLBACKS"), relative(DIST, file)).toBe(false);
  }
});

test("the PPTX-only preset catalogue stays out of the flat Worker", () => {
  // A payload literal from core/presets.ts: if this reaches the Worker, every
  // consumer pays for shapes it cannot draw.
  const marker = "actionButtonBackPrevious";
  const worker = read("core", "worker.mjs");
  expect(worker.includes(marker), "Worker must not contain the preset catalogue").toBe(false);
  const withCatalogue = distFiles().filter((file) => file.endsWith(".mjs") && readFileSync(file, "utf8").includes(marker));
  expect(withCatalogue.length, "no emitted module carries the catalogue").toBeGreaterThan(0);
  for (const file of withCatalogue) {
    expect(file.replaceAll("\\", "/")).toContain("/dist/core/renderer-");
  }
});

test("shipped errors name the published package entrypoint", () => {
  // Chunks are content-hashed, so find the error text rather than a filename.
  const scripts = distFiles().filter((file) => file.endsWith(".mjs"));
  const errors = scripts.filter((file) => readFileSync(file, "utf8").includes("@aared/viewer/"));
  expect(errors.length, "no shipped module carries the package's error prefix").toBeGreaterThan(0);
  for (const file of scripts) {
    const text = readFileSync(file, "utf8");
    expect(text.includes("`aared/"), `${relative(DIST, file)} has a stale unscoped prefix`).toBe(false);
  }
});