import { expect, test } from "bun:test";
import { open as openDocx } from "../src/docx/index.js";
import { open as openPptx } from "../src/pptx/index.js";
import { open as openXlsx } from "../src/xlsx/index.js";
import { drainWorkerPool } from "../src/core/session.js";

/**
 * The single-file worker contract.
 *
 * `dist/core/worker.mjs` is self-contained: every format parser is bundled
 * flat into the one file, so there are no sibling chunks to deploy, no
 * relative import that a consumer bundler could drop, and the same file
 * backs Blob-URL workers directly (see "Worker override" in the API
 * reference). Both halves are pinned here: the artifact must stay
 * self-contained *and* actually open every format through a Blob-URL
 * worker — the consumer recipe, exercised for real.
 *
 * These tests read `dist`, so they need a prior
 * `bun run build` (workspace root); a missing artifact fails loudly rather
 * than skipping, so the guarantee can't rot silently.
 */

const DIST = `${import.meta.dir}/../dist/core`;
const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/post-production-review.pptx`;
const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/post-production-tracker.xlsx`;
const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

async function distText(name: string): Promise<string> {
  const file = Bun.file(`${DIST}/${name}`);
  if (!(await file.exists())) {
    throw new Error(
      `dist/core/${name} is missing — run \`bun run build\` from the workspace root, then re-run this test.`,
    );
  }
  return file.text();
}

test("worker artifact is self-contained", async () => {
  const source = await distText("worker.mjs");
  // A Blob worker (or a verbatim-copied one) has no sibling files: any
  // relative import — static or dynamic — would 404 at first open(). Minified
  // output drops the space in `from "./x"`, so the pattern allows none.
  expect(source).not.toMatch(/from\s*["']\.\.?\//);
  expect(source).not.toMatch(/import\s*\(\s*["']\./);
});

test("the shipped Worker constructor stays the literal consumer bundlers detect", async () => {
  // Vite bundles the worker only when `new URL(...)` sits literally inside
  // `new Worker(...)` — its own detection pattern, copied from Vite 7.3.6
  // (vite/dist/node/chunks/config.js). Minification must not reshape it.
  const VITE_WORKER = /\bnew\s+(?:Worker|SharedWorker)\s*\(\s*(new\s+URL\s*\(\s*('[^']+'|"[^"]+"|`[^`]+`)\s*,\s*import\.meta\.url\s*\))/dg;
  const glob = new Bun.Glob("*.mjs");
  const hits: string[] = [];
  for (const name of glob.scanSync(DIST)) {
    const text = await distText(name);
    for (const match of text.matchAll(VITE_WORKER)) {
      const call = text.slice(match.index, match.index + 120);
      hits.push(`${match[2]!.slice(1, -1)} module:${/type\s*:\s*["'`]module["'`]/.test(call)}`);
    }
  }
  expect(hits).toEqual(["./worker.mjs module:true"]);
});

test("the documented aared/worker subpath is the single-file artifact", () => {
  const resolved = Bun.resolveSync("aared/worker", import.meta.dir).replaceAll("\\", "/");
  expect(resolved.endsWith("/dist/core/worker.mjs")).toBe(true);
});

test("inline Worker source opens PPTX, XLSX, and DOCX on one pooled Worker", async () => {
  // The consumer recipe, as documented: import the script by its package
  // subpath (`aared/worker?raw` in a bundler), hand its source to aared.
  await distText("worker.mjs");
  const source = await Bun.file(Bun.resolveSync("aared/worker", import.meta.dir)).text();
  // Count Workers: one closed document's Worker serves the next open, whatever its format.
  const original = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  const Native = globalThis.Worker;
  let started = 0;
  Object.defineProperty(globalThis, "Worker", {
    configurable: true,
    value: class CountingWorker extends Native {
      constructor(...args: ConstructorParameters<typeof Worker>) {
        super(...args);
        started += 1;
      }
    },
  });
  try {
    const deck = await openPptx(await Bun.file(PPTX_FIXTURE).arrayBuffer(), { worker: { source } });
    try {
      const slides = deck.slides.list();
      expect(slides.length).toBeGreaterThan(0);
      expect(await slides[0]!.read()).toBeDefined();
    } finally {
      deck.close();
    }

    const workbook = await openXlsx(await Bun.file(XLSX_FIXTURE).arrayBuffer(), { worker: { source } });
    try {
      const sheets = workbook.sheets.list();
      expect(sheets.length).toBeGreaterThan(0);
      expect(await sheets[0]!.read()).toBeDefined();
    } finally {
      workbook.close();
    }

    const document = await openDocx(await Bun.file(DOCX_FIXTURE).arrayBuffer(), { worker: { source } });
    try {
      expect(document.sections.list().length).toBeGreaterThan(0);
    } finally {
      document.close();
    }
    expect(started).toBe(1);
  } finally {
    drainWorkerPool();
    if (original === undefined) delete (globalThis as { Worker?: unknown }).Worker;
    else Object.defineProperty(globalThis, "Worker", original);
  }
});
