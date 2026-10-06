import { expect, test } from "bun:test";
import { mkdirSync, rmSync, writeFileSync } from "node:fs";
import { join } from "node:path";

const ROOT = join(import.meta.dir, "..", "..", "..");
// Loaded at run time by a computed path, so the site's module never joins the
// package's TypeScript program (a declaration build would emit beside it).
const PLAYGROUND_SNIPPETS = join(ROOT, "site", "src", "lib", "playground-snippets.ts");
const DOCS_DIR = join(ROOT, "site", "src", "content", "docs");
const SRC = join(ROOT, "packages", "aared", "src");
const TSC = join(ROOT, "node_modules", ".bin", process.platform === "win32" ? "tsc.exe" : "tsc");
// Inside the package, so `react` and `@types/bun` resolve the way a consumer's would.
const OUT = join(import.meta.dir, ".api-snippets");

/** The names the examples use as an app would already have them: a file, a host element. */
const PRELUDE = [
  "declare const file: File;",
  "declare const myFile: File;",
  "declare const container: HTMLElement;",
  "declare const el: HTMLElement;",
  "export {};",
  "",
].join("\n");

interface Snippet {
  name: string;
  lang: "ts" | "tsx";
  code: string;
}

/**
 * A docs page's runnable examples: fenced ```ts check / ```tsx check. The
 * page's signature listings (interfaces, pseudo-declarations) carry no marker
 * and are not compiled; an example a reader can paste whole carries it.
 */
export function checkedBlocks(mdx: string, page: string): Snippet[] {
  const blocks: Snippet[] = [];
  for (const match of mdx.matchAll(/^```(tsx?) check[^\n]*\n([\s\S]*?)^```/gm)) {
    const line = mdx.slice(0, match.index).split("\n").length;
    blocks.push({ name: `${page}-line-${line}`, lang: match[1] as "ts" | "tsx", code: match[2]! });
  }
  return blocks;
}

/**
 * docs-parity checks that every export is named; this checks that what the
 * API page, the Start page and the playground tell a reader to write
 * compiles, under the repo's strict settings, against the package source (no
 * build needed).
 */
test("the docs pages' examples and the playground snippets compile against the package", async () => {
  const api = checkedBlocks(await Bun.file(join(DOCS_DIR, "api.mdx")).text(), "api");
  expect(api.length).toBeGreaterThanOrEqual(8);
  const start = checkedBlocks(await Bun.file(join(DOCS_DIR, "start.mdx")).text(), "start");
  expect(start.length).toBeGreaterThanOrEqual(6);
  type Group = Record<string, { title: string; lang: string; code: string }>;
  const site = (await import(PLAYGROUND_SNIPPETS)) as { reactSnippets: Group; facadeSnippets: Group; headlessSnippets: Group };
  const playground: Snippet[] = [site.reactSnippets, site.facadeSnippets, site.headlessSnippets].flatMap((group) =>
    Object.entries(group).map(([format, snippet]) => ({
      name: `playground-${snippet.title.replace(/\W/g, "-")}-${format}`,
      lang: snippet.lang === "tsx" ? "tsx" : "ts",
      code: snippet.code,
    } satisfies Snippet)),
  );
  rmSync(OUT, { recursive: true, force: true });
  mkdirSync(OUT, { recursive: true });
  try {
    for (const snippet of [...api, ...start, ...playground]) {
      writeFileSync(join(OUT, `${snippet.name}.${snippet.lang}`), PRELUDE + snippet.code);
    }
    writeFileSync(join(OUT, "tsconfig.json"), JSON.stringify({
      extends: join(ROOT, "tsconfig.base.json"),
      compilerOptions: {
        noEmit: true,
        declaration: false,
        declarationMap: false,
        jsx: "react-jsx",
        // `aared/<entry>` is the package's source, as the exports map points a consumer at its build.
        paths: { "aared/*": [join(SRC, "*", "index.ts")], aared: [join(SRC, "core", "index.ts")] },
      },
      include: ["*.ts", "*.tsx"],
    }));
    const result = Bun.spawnSync([TSC, "-p", join(OUT, "tsconfig.json")], { cwd: OUT });
    const output = `${result.stdout.toString()}${result.stderr.toString()}`;
    expect(output).toBe("");
    expect(result.exitCode).toBe(0);
  } finally {
    rmSync(OUT, { recursive: true, force: true });
  }
}, 120000);
