/**
 * A side-by-side review page for the DOCX pairs: Word's export beside aared's
 * capture, one row per physical page, at the same display width so the two can
 * be judged against each other.
 *
 *   bun tools/review/docx/review.ts tools/out/docx/docx-coverage
 *   bun tools/review/docx/review.ts tools/out/docx/review
 *
 * The PPTX and XLSX capture tests write the same page from inside `bun test`
 * (`test/pptx-capture.visual.test.ts`, `test/xlsx-capture.visual.test.ts`). The
 * DOCX review is run by hand rather than by a test (AGENTS.md "Testing rule"),
 * so it gets a command instead. Writes `review.html` into the directory it is
 * given, next to the `word-page-N.png` / `aared-page-N.png` pairs, and reads
 * `word-export.json` and `pagination-comparison.json` when they are there.
 */
import { readdir, readFile, writeFile } from "node:fs/promises";
import { join, resolve } from "node:path";
import { reviewPage, escapeHtml } from "../page.js";

const [directory] = process.argv.slice(2);
if (directory === undefined) {
  throw new Error("Usage: bun tools/review/docx/review.ts <output-dir>");
}
const dir = resolve(directory);

const names = await readdir(dir);
const numbered = (pattern: RegExp): number[] =>
  names
    .map((name) => pattern.exec(name)?.[1])
    .filter((value): value is string => value !== undefined)
    .map(Number)
    .sort((left, right) => left - right);

const wordPages = numbered(/^word-page-(\d+)\.png$/);
const aaredPages = numbered(/^aared-page-(\d+)\.png$/);
const absentPages = numbered(/^aared-absent-page-(\d+)\.png$/);
if (wordPages.length === 0) throw new Error(`aared Word review: no word-page-N.png in ${dir}`);
if (aaredPages.length === 0) throw new Error(`aared Word review: no aared-page-N.png in ${dir}`);

async function readJson<T>(name: string): Promise<T | undefined> {
  try {
    return JSON.parse(await readFile(join(dir, name), "utf8")) as T;
  } catch {
    return undefined;
  }
}

const exported = await readJson<{ wordVersion?: string; pageCount?: number }>("word-export.json");
const compared = await readJson<{
  wordTokens?: number;
  aaredTokens?: number;
  matched?: number;
  pageMismatches?: number;
  medianDy?: number;
  within4px?: number;
  over12px?: number;
}>("pagination-comparison.json");

const hasAbsent = absentPages.length > 0;
// Page-count differences are review evidence, not a reason to suppress the report.
const allPages = [...new Set([...wordPages, ...aaredPages, ...absentPages])].sort((a, b) => a - b);
const cards = allPages
  .map((page) => {
    const columns = [wordPages.includes(page)
      ? `<figure><figcaption>Word page ${page}</figcaption><img src="word-page-${page}.png" alt="Word page ${page}"></figure>`
      : "<figure><figcaption>Word</figcaption><p>No corresponding page.</p></figure>"];
    columns.push(aaredPages.includes(page)
      ? `<figure><figcaption>aared page ${page}</figcaption><img src="aared-page-${page}.png" alt="aared page ${page}"></figure>`
      : "<figure><figcaption>aared</figcaption><p>No corresponding page.</p></figure>");
    if (hasAbsent && absentPages.indexOf(page) !== -1) {
      const absent = page;
      columns.push(`<figure><figcaption>aared, Office fonts absent</figcaption><img src="aared-absent-page-${absent}.png" alt="aared page ${absent} without Office fonts"></figure>`);
    }
    return `<section class="pair">${columns.join("")}</section>`;
  })
  .join("\n");

const round = (value: number | undefined): string | undefined =>
  value === undefined ? undefined : (Number.isInteger(value) ? String(value) : value.toFixed(1));

const summary = [
  exported?.wordVersion === undefined ? undefined : `Word ${exported.wordVersion}`,
  `pages: Word ${wordPages.length}, aared ${aaredPages.length}`,
  compared === undefined
    ? undefined
    : `tokens: Word ${compared.wordTokens}, aared ${compared.aaredTokens}, matched ${compared.matched}`
      + `; page mismatches ${compared.pageMismatches}; median dy ${round(compared.medianDy)}px`
      + `; within 4px ${compared.within4px}, over 12px ${compared.over12px}`,
].filter((part): part is string => part !== undefined).join(" · ");

const html = reviewPage("aared DOCX review", `<p class="summary">${escapeHtml(summary)}</p>
<p class="summary">Read a layout number from <code>word-spans.json</code> rather than judging a screenshot; this page is for typography, bidi placement and wrapping.</p>
${cards}
`);

await writeFile(join(dir, "review.html"), html);
console.log(`${join(dir, "review.html")} (${wordPages.length} page pair${wordPages.length === 1 ? "" : "s"}${hasAbsent ? ", with the absent-fonts column" : ""})`);
