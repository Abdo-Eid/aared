/**
 * Compare Word and aared token positions produced by `rasterize.py` and
 * `capture.ts`. Matching is by normalized full token and occurrence order,
 * never by a truncated prefix or same-page key.
 *
 * Usage: bun tools/review/docx/compare-pagination.ts tools/out/docx/docx-coverage
 */
import { readFile, writeFile } from "node:fs/promises";
import { join, resolve } from "node:path";

interface WordBox {
  page: number;
  text: string;
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

interface Pair {
  text: string;
  wordPage: number;
  aaredPage: number;
  dx: number;
  dy: number;
}

const [outputArg] = process.argv.slice(2);
if (outputArg === undefined) {
  throw new Error("Usage: bun tools/review/docx/compare-pagination.ts <review-output-dir>");
}
const outputDir = resolve(outputArg);
const readBoxes = async (name: string): Promise<WordBox[]> => {
  const path = join(outputDir, name);
  try {
    return await Bun.file(path).json() as WordBox[];
  } catch {
    throw new Error(`aared Word review: missing ${path}; rerun Word export and aared capture`);
  }
};

const [word, aared] = await Promise.all([readBoxes("word-words.json"), readBoxes("aared-words.json")]);
const normalize = (text: string): string => text.normalize("NFC").replace(/\s+/gu, " ").trim();
const order = (a: WordBox, b: WordBox): number => a.page - b.page || a.y0 - b.y0 || a.x0 - b.x0;
const groupByText = (items: WordBox[]): Map<string, WordBox[]> => {
  const groups = new Map<string, WordBox[]>();
  for (const item of items) {
    const key = normalize(item.text);
    if (key === "") continue;
    const group = groups.get(key) ?? [];
    group.push(item);
    groups.set(key, group);
  }
  for (const group of groups.values()) group.sort(order);
  return groups;
};

const wordGroups = groupByText(word);
const aaredGroups = groupByText(aared);
const pairs: Pair[] = [];
const missingFromAared: Array<{ text: string; count: number }> = [];
const extraInAared: Array<{ text: string; count: number }> = [];
let repeatedGroups = 0;
for (const key of new Set([...wordGroups.keys(), ...aaredGroups.keys()])) {
  const wordItems = wordGroups.get(key) ?? [];
  const aaredItems = aaredGroups.get(key) ?? [];
  if (Math.max(wordItems.length, aaredItems.length) > 1) repeatedGroups++;
  const count = Math.min(wordItems.length, aaredItems.length);
  for (let i = 0; i < count; i++) {
    const w = wordItems[i]!;
    const a = aaredItems[i]!;
    pairs.push({ text: key, wordPage: w.page, aaredPage: a.page, dx: (a.x0 + a.x1 - w.x0 - w.x1) / 2, dy: (a.y0 + a.y1 - w.y0 - w.y1) / 2 });
  }
  if (wordItems.length > count) missingFromAared.push({ text: key, count: wordItems.length - count });
  if (aaredItems.length > count) extraInAared.push({ text: key, count: aaredItems.length - count });
}

const median = (values: number[]): number => {
  if (values.length === 0) return Number.NaN;
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)]!;
};
const vertical = pairs.map((pair) => pair.dy);
const meanAbsDy = vertical.reduce((sum, value) => sum + Math.abs(value), 0) / Math.max(1, vertical.length);
const pageMismatches = pairs.filter((pair) => pair.wordPage !== pair.aaredPage);
const byWordPage = new Map<number, Pair[]>();
for (const pair of pairs) {
  const group = byWordPage.get(pair.wordPage) ?? [];
  group.push(pair);
  byWordPage.set(pair.wordPage, group);
}
const pages = [...byWordPage.entries()].sort(([a], [b]) => a - b).map(([page, items]) => ({
  page,
  matched: items.length,
  pageMismatches: items.filter((item) => item.wordPage !== item.aaredPage).length,
  medianDy: median(items.map((item) => item.dy)),
  maxAbsDy: Math.max(...items.map((item) => Math.abs(item.dy))),
}));
const report = {
  wordTokens: word.length,
  aaredTokens: aared.length,
  matched: pairs.length,
  repeatedTokenGroups: repeatedGroups,
  pageMismatches: pageMismatches.length,
  missingFromAared,
  extraInAared,
  medianDx: median(pairs.map((pair) => pair.dx)),
  medianDy: median(vertical),
  meanAbsDy,
  within4px: vertical.filter((value) => Math.abs(value) <= 4).length,
  over12px: vertical.filter((value) => Math.abs(value) > 12).length,
  pages,
};
await writeFile(join(outputDir, "pagination-comparison.json"), JSON.stringify(report, null, 2));
console.log(`Word tokens=${word.length}; aared tokens=${aared.length}; matched=${pairs.length}; repeated token groups=${repeatedGroups}`);
console.log(`page mismatches=${pageMismatches.length}; missing from aared=${missingFromAared.length}; extra in aared=${extraInAared.length}`);
console.log(`median dx=${report.medianDx.toFixed(1)}px; median dy=${report.medianDy.toFixed(1)}px; mean |dy|=${meanAbsDy.toFixed(1)}px; within 4px=${report.within4px}; over 12px=${report.over12px}`);
for (const page of pages) {
  console.log(` page ${page.page}: matched=${page.matched}; page mismatches=${page.pageMismatches}; median dy=${page.medianDy.toFixed(1)}px; max |dy|=${page.maxAbsDy.toFixed(1)}px`);
}
