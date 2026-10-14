import { expect, test } from "bun:test";
import { mkdtempSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { generate } from "../scripts/generate-llms-txt.js";
import { unified } from "unified";
import remarkParse from "remark-parse";
import remarkMdx from "remark-mdx";

const PACKAGE = resolve(import.meta.dir, "../../packages/aared");

function codeBlocks(source: string): { lang: string | null; meta: string | null; value: string }[] {
  const tree = unified().use(remarkParse).use(remarkMdx).parse(source.replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n/, ""));
  const blocks: { lang: string | null; meta: string | null; value: string }[] = [];
  const walk = (node: { type: string; children?: readonly unknown[]; value?: string; lang?: string | null; meta?: string | null }): void => {
    if (node.type === "code") blocks.push({ lang: node.lang ?? null, meta: node.meta ?? null, value: node.value! });
    for (const child of node.children ?? []) walk(child as Parameters<typeof walk>[0]);
  };
  walk(tree);
  return blocks;
}

/** The same public-export extraction `docs-parity.test.ts` uses. */
function exportedNames(source: string): string[] {
  const names = new Set<string>();
  for (const match of source.matchAll(/^export\s+(?:async\s+)?(?:type|interface|class|function|const|enum)\s+(\w+)/gm)) names.add(match[1]!);
  for (const match of source.matchAll(/export\s+type\s*{([\s\S]*?)}\s+from/g)) {
    for (const entry of match[1]!.split(",")) {
      const name = entry.trim().match(/^(\w+)/)?.[1];
      if (name !== undefined) names.add(name);
    }
  }
  for (const match of source.matchAll(/export\s*(type\s*)?{([\s\S]*?)}\s+from/g)) {
    if (match[1] !== undefined) continue;
    for (const entry of match[2]!.split(",")) {
      const name = entry.trim().match(/^(\w+)/)?.[1];
      if (name !== undefined) names.add(name);
    }
  }
  return [...names].sort();
}

test("generated agent docs cover every public export, resolve their links and stay deterministic", () => {
  const first = mkdtempSync(join(tmpdir(), "aared-llms-a-"));
  const second = mkdtempSync(join(tmpdir(), "aared-llms-b-"));
  const a = generate(first);
  const b = generate(second);
  expect(Object.keys(a).sort()).toEqual(["api.md", "llms-full.txt", "llms.txt", "start.md"]);
  // Byte-stable: agents and the site must not see churn between builds.
  expect(b).toEqual(a);
  expect(readFileSync(join(first, "llms-full.txt"), "utf8")).toBe(`${a["start.md"]}\n---\n\n${a["api.md"]}`);

  // The real contract: every public name of every entry is named, the same way
  // docs-parity guards api.mdx against the sources.
  const api = a["api.md"]!;
  // Generated recipes retain every fenced example, including relative imports
  // such as the Worker factory. Only MDX module imports may be removed.
  for (const page of ["start", "api"]) {
    const source = readFileSync(resolve(import.meta.dir, `../src/content/docs/${page}.mdx`), "utf8");
    const original = codeBlocks(source);
    expect(original.length).toBeGreaterThan(10);
    expect(codeBlocks(a[`${page}.md`]!), `${page}.md changed a code sample`).toEqual(original);
  }
  expect(api).toContain('import AaredWorker from "./worker?worker";');
  const entries = {
    "@aared/viewer/core": "core/index.ts",
    "@aared/viewer/pptx": "pptx/index.ts",
    "@aared/viewer/xlsx": "xlsx/index.ts",
    "@aared/viewer/docx": "docx/index.ts",
    "@aared/viewer/preview": "preview/index.ts",
    "@aared/viewer/react": "react/index.ts",
    "@aared/viewer/element": "element/index.ts",
  };
  let checked = 0;
  for (const [entry, file] of Object.entries(entries)) {
    for (const name of exportedNames(readFileSync(join(PACKAGE, "src", file), "utf8"))) {
      checked += 1;
      expect(new RegExp(`\\b${name}\\b`).test(api), `api.md omits ${name} from ${entry}`).toBe(true);
    }
  }
  expect(checked).toBeGreaterThan(120);
  for (const name of ["open", "render", "preview", "detect", "usePreview", "AaredError", "PreviewHandle"]) {
    expect(new RegExp(`\\b${name}\\b`).test(api), `api.md omits ${name}`).toBe(true);
  }

  // Partials are inlined at the position they are used, and every in-site link
  // is fully absolute: these files are read off the deployment, where
  // `../api/` or `/aared/api/` resolve to nothing.
  const start = a["start.md"]!;
  const intro = start.indexOf("Get one file on screen");
  const heading = start.indexOf("## Choose your integration");
  const chooser = start.indexOf("You want to");
  expect(intro).toBeGreaterThan(-1);
  expect(heading).toBeGreaterThan(intro);
  // The partial's table belongs under its own heading, not above the intro,
  // and the heading must not be left empty by a dropped component tag.
  expect(chooser).toBeGreaterThan(heading);
  expect(start.slice(heading, chooser)).not.toContain("\n## ");
  expect(start).toContain("@aared/viewer/react");
  for (const [name, text] of [["start.md", a["start.md"]!], ["api.md", a["api.md"]!]] as const) {
    const relative = [...text.matchAll(/\]\(([^)]+)\)/g)].map(match => match[1]!)
      .filter(url => !/^(?:[a-z][a-z0-9+.-]*:|#)/i.test(url));
    for (const url of relative) expect(url, `${name} link is not absolute`).toMatch(/^https:\/\//);
  }
  // Cross-page navigation exists in both documents and must resolve to the site.
  expect(a["api.md"]).toContain("https://abdo-eid.github.io/aared/start/");
  // No interactive components survive, and no internal render model leaks.
  for (const text of [a["start.md"]!, a["api.md"]!]) {
    expect(text.includes("mdxJsx")).toBe(false);
    expect(/PlaygroundCode|ElementDemo|FacadeDemo|HooksDemo|HeadlessDemo/.test(text)).toBe(false);
    for (const leaked of ["PaintSlideInput", "ChartSpecModel", "WorkerSession", "note-"]) {
      expect(text, `generated doc leaked ${leaked}`).not.toContain(leaked);
    }
  }
  // llms.txt is the index: it links the other three under the real base.
  const index = a["llms.txt"]!;
  expect(index).toStartWith("# aared (عارض) — @aared/viewer");
  expect(index).toContain("TypeScript Office viewer, published on npm as @aared/viewer");
  expect(index).toContain("npm install @aared/viewer");
  expect(index).toContain("https://www.npmjs.com/package/@aared/viewer");
  expect(index).toContain("docs/COMPETITORS.md");
  expect(index).toContain("docs/BENCHMARKS.md");
  expect(index).toContain("native PowerPoint, Excel and Word output");
  for (const link of ["start.md", "api.md", "llms-full.txt"]) {
    expect(index).toContain(`https://abdo-eid.github.io/aared/${link}`);
  }
});
