import { expect, test } from "bun:test";
import { BULLET, LEFT_TO_RIGHT_MARK, NON_BREAKING_HYPHEN, RETURN_ARROW, RIGHT_TO_LEFT_MARK } from "../src/core/chars.js";

/**
 * Shipped JS is Latin-1. V8 stores a script one byte per character only
 * while every character fits in Latin-1; one wider character makes the whole
 * script two bytes per character — for the ~128 KB Worker, in every Worker.
 * Wider characters aared needs at runtime come from `src/core/chars.ts`.
 *
 * Reads `dist`, so it needs a prior `bun run build`; a missing artifact fails
 * loudly. The scan also catches `\u` escapes of wider characters, because a
 * consumer minifier printing with `charset: "utf8"` turns them back into the
 * raw character.
 */

const DIST = `${import.meta.dir}/../dist`;
const WIDE_ESCAPE = /(?<!\\)(?:\\\\)*\\u(?:\{0*[1-9a-f][0-9a-f]{2,}\}|(?!00)[0-9a-f]{4})/gi;

async function wideCharacters(path: string): Promise<string[]> {
  const text = await Bun.file(path).text();
  const hits: string[] = [];
  const lineOf = (offset: number): number => text.slice(0, offset).split("\n").length;
  for (let i = 0; i < text.length; i += 1) {
    const code = text.charCodeAt(i);
    if (code > 0xff) hits.push(`${path.slice(DIST.length + 1)}:${lineOf(i)} U+${code.toString(16).toUpperCase().padStart(4, "0")} ${JSON.stringify(text.slice(Math.max(0, i - 30), i + 10))}`);
  }
  for (const match of text.matchAll(WIDE_ESCAPE)) {
    hits.push(`${path.slice(DIST.length + 1)}:${lineOf(match.index)} escape ${match[0]}`);
  }
  return hits;
}

async function distFiles(): Promise<string[]> {
  const worker = Bun.file(`${DIST}/core/worker.mjs`);
  if (!(await worker.exists())) {
    throw new Error("dist/core/worker.mjs is missing — run `bun run build` from the workspace root, then re-run this test.");
  }
  return [...new Bun.Glob("**/*.mjs").scanSync(DIST)].map((name) => `${DIST}/${name.replaceAll("\\", "/")}`);
}

test("the worker artifact is one-byte: no character above U+00FF", async () => {
  expect(await wideCharacters(`${DIST}/core/worker.mjs`)).toEqual([]);
});

test("every shipped chunk is one-byte: no character above U+00FF", async () => {
  const hits: string[] = [];
  for (const path of await distFiles()) hits.push(...(await wideCharacters(path)));
  expect(hits).toEqual([]);
});

test("the characters built from char codes keep their values", () => {
  expect([BULLET, NON_BREAKING_HYPHEN, LEFT_TO_RIGHT_MARK, RIGHT_TO_LEFT_MARK, RETURN_ARROW].map((c) => c.codePointAt(0))).toEqual([
    0x2022, 0x2011, 0x200e, 0x200f, 0x21a9,
  ]);
});
