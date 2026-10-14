/**
 * Generate the agent-facing plain-Markdown entry points from the site's own
 * docs sources, after `astro build`:
 *
 *   dist/llms.txt      a short header plus annotated links
 *   dist/start.md      start.mdx with its shared partials inlined
 *   dist/api.md        api.mdx
 *   dist/llms-full.txt start.md + api.md, concatenated
 *
 * Deriving these from the same MDX the site renders keeps them honest: those
 * sources are already covered by `api-snippets.test.ts` and
 * `docs-parity.test.ts`, so a contract cannot be right on the site and wrong
 * here. Notes, ARCHITECTURE and internal models are never inputs.
 *
 *   bun run --cwd site llms   (also run by `bun run docs:build`)
 */
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import remarkParse from "remark-parse";
import remarkMdx from "remark-mdx";
import remarkStringify from "remark-stringify";
import { unified } from "unified";
import type { Root as MdastRoot } from "mdast";

const SITE_DIR = join(dirname(fileURLToPath(import.meta.url)), "..");
const DIST = join(SITE_DIR, "dist");
const DOCS = join(SITE_DIR, "src", "content", "docs");
const START = join(DOCS, "start.mdx");
const API = join(DOCS, "api.mdx");

/** The deployed documentation root, so relative MDX links become absolute. */
const SITE_URL = readSiteUrl();
function readSiteUrl(): string {
  const config = readFileSync(join(SITE_DIR, "astro.config.mjs"), "utf8");
  const site = /^const site = '([^']+)';$/m.exec(config)?.[1];
  const base = /^const base = '([^']+)';$/m.exec(config)?.[1];
  if (site === undefined || base === undefined) throw new Error("astro.config.mjs no longer declares site/base literals");
  // `site` is the full page URL and already ends with the deployment path.
  return `${new URL(site).origin}${base}`;
}

/** Every doc node is something remark can print; these carry no output. */
type AnyNode = {
  type: string;
  name?: string;
  url?: string;
  children?: AnyNode[];
  value?: string;
  data?: { estree?: { body: {
    type: string;
    source?: { value?: unknown };
    specifiers?: { type: string; local: { name: string } }[];
  }[] } };
};
const DROPPED = new Set(["mdxFlowExpression", "mdxTextExpression", "mdxjsEsm", "yaml"]);

/**
 * MDX components (`PlaygroundCode`, `ElementDemo`, …) render interactive UI and
 * already receive their code as fenced children, so keep the children and drop
 * the wrapper. Braced comment expressions render nothing and are dropped
 * too. Returning rebuilt child lists avoids splicing a tree while visiting it.
 */
function flatten(nodes: AnyNode[], partials: Map<string, AnyNode[]>, seen: Set<string>): AnyNode[] {
  const out: AnyNode[] = [];
  for (const node of nodes) {
    if (DROPPED.has(node.type)) continue;
    if (node.type === "mdxJsxFlowElement" || node.type === "mdxJsxTextElement") {
      const name = typeof node.name === "string" ? node.name : undefined;
      // An MDX partial's content belongs where it is used, not where it was
      // imported: `<ChoosingALayer />` under its own heading carries the real
      // prose, and inlining at the import line would move it above the
      // introduction and leave that section empty.
      if (name !== undefined && partials.has(name)) {
        if (seen.has(name)) continue; // a partial that includes itself
        out.push(...partials.get(name)!);
        continue;
      }
      out.push(...flatten(node.children ?? [], partials, seen));
      continue;
    }
    if (node.children !== undefined) node.children = flatten(node.children, partials, seen);
    out.push(node);
  }
  return out;
}

/**
 * Rewrite in-site links to absolute documentation URLs. These files are read
 * off the deployment by agents, where a root-relative or `../` link resolves
 * against nothing, so links stay fully qualified. `pageUrl` is where the source
 * file is published, which is what a relative link resolves against on the
 * site. External URLs and bare fragments are untouched.
 */
function absolutize(nodes: AnyNode[], pageUrl: string): void {
  for (const node of nodes) {
    if ((node.type === "link" || node.type === "definition") && node.url !== undefined) {
      const url = node.url;
      if (/^[a-z][a-z0-9+.-]*:/i.test(url) || url.startsWith("//") || url.startsWith("#")) continue;
      const hash = url.indexOf("#");
      const path = hash === -1 ? url : url.slice(0, hash);
      const fragment = hash === -1 ? "" : url.slice(hash);
      if (path === "") continue;
      node.url = path.startsWith("/")
        ? `${SITE_URL}${path.slice(1)}${fragment}`
        : `${new URL(path, pageUrl).href}${fragment}`;
    }
    if (node.children !== undefined) absolutize(node.children, pageUrl);
  }
}

const processor = unified().use(remarkParse).use(remarkMdx);

/** One partial, converted to nodes with its own imports already bound. */
function parseToNodes(file: string, pageUrl: string, seen: Set<string>): AnyNode[] {
  const path = resolve(file);
  if (seen.has(path)) return [];
  seen.add(path);
  const raw = readFileSync(path, "utf8").replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n/, "");
  const tree = processor.parse(raw);
  const imports = new Map<string, string>();
  // Only real MDX module imports bind partials. Fenced examples may contain the
  // same import syntax, but their code nodes must survive byte-for-byte.
  for (const node of tree.children as AnyNode[]) {
    if (node.type !== "mdxjsEsm") continue;
    for (const statement of node.data?.estree?.body ?? []) {
      const specifier = statement.source?.value;
      if (statement.type !== "ImportDeclaration" || typeof specifier !== "string"
        || !specifier.startsWith(".") || !specifier.endsWith(".mdx")) continue;
      for (const binding of statement.specifiers ?? []) {
        if (binding.type === "ImportDefaultSpecifier") {
          imports.set(binding.local.name, resolve(dirname(path), specifier));
        }
      }
    }
  }
  const partials = new Map<string, AnyNode[]>();
  for (const [name, target] of imports) {
    partials.set(name, parseToNodes(target, pageUrl, seen));
  }
  const nodes = flatten(tree.children as AnyNode[], partials, new Set());
  absolutize(nodes, pageUrl);
  return nodes;
}

/**
 * Render one MDX source to Markdown. A local `import X from "./y"` binds `X` to
 * that file, and `<X />` is replaced in place by the converted partial, so the
 * prose keeps the position the author gave it.
 */
function toMarkdown(file: string, pageUrl: string, heading: string, seen = new Set<string>()): string {
  const children = parseToNodes(resolve(file), pageUrl, seen);
  const body = String(unified().use(remarkStringify, { bullet: "-", fences: true }).stringify({ type: "root", children } as MdastRoot)).trim();
  return heading === "" ? `${body}\n` : `# ${heading}\n\n${body}\n`;
}

const startUrl = `${SITE_URL}start/`;
const apiUrl = `${SITE_URL}api/`;

/** Convert both pages plus the agent index, and write them into `outDir`. */
export function generate(outDir: string = DIST): Record<string, string> {
  const start = toMarkdown(START, startUrl, "Getting started");
  const api = toMarkdown(API, apiUrl, "API reference");
  const full = `${start}\n---\n\n${api}`;

  // Guard the conversion itself: a truncated or component-only document would
  // still "succeed", and these files are the package's agent entry points.
  for (const [name, text] of [["start.md", start], ["api.md", api]] as const) {
    if (text.length < (name === "api.md" ? 8_000 : 1_500)) throw new Error(`${name} conversion produced only ${text.length} characters`);
    if (text.includes("mdxJsx")) throw new Error(`${name} still contains unflattened MDX`);
    // A component tag outside code means the flatten step missed one; inside a
    // fence or a code span the same text is a legitimate generic or signature.
    const prose = text.replace(/```[\s\S]*?```/g, "").replace(/`[^`\n]*`/g, "");
    const leftover = /<\/?[A-Z][\w.]*/.exec(prose);
    if (leftover !== null) throw new Error(`${name} still contains the tag ${leftover[0]}`);
  }
  // These pages legitimately link the published ARCHITECTURE/Notes documents, so
  // guard the real contract instead: internal rendering models must not surface.
  for (const leaked of ["note-", "PaintSlideInput", "ChartSpecModel", "WorkerSession"]) {
    if (full.includes(leaked)) throw new Error(`generated docs leaked internal reference "${leaked}"`);
  }

  const llms = [
    "# aared",
    "",
    "Browser-only, view-only previewer for PPTX, XLSX and DOCX, with Arabic and other",
    "RTL text rendered as real DOM/SVG rather than canvas pixels. Parsing runs in a",
    "Web Worker; rendering requires a browser.",
    "",
    "## Plain Markdown",
    "",
    `- [Getting started](${SITE_URL}start.md) — install, host sizing, integration choices, recipes`,
    `- [API reference](${SITE_URL}api.md) — supported signatures, types, events, errors, fidelity limits`,
    `- [Full document](${SITE_URL}llms-full.txt) — both files concatenated`,
    `- [Rendered documentation](${SITE_URL}) — the same content as the site`,
    "",
    "## Integration notes",
    "",
    "- Import a documented subpath such as `@aared/viewer/pptx`; the package root is not importable.",
    "- A slide sizes itself; a sheet or a document needs an explicit host height.",
    "- Destroy each render handle, then close the document. Hooks and the element own teardown.",
    "- Exact signatures are in the installed declarations under `dist/`.",
    "",
  ].join("\n");

  mkdirSync(outDir, { recursive: true });
  const outputs: Record<string, string> = { "llms.txt": llms, "start.md": start, "api.md": api, "llms-full.txt": full };
  for (const [name, text] of Object.entries(outputs)) writeFileSync(join(outDir, name), text);
  return outputs;
}

if (import.meta.main) {
  const outputs = generate();
  console.log(`generated ${Object.entries(outputs).map(([name, text]) => `${name} (${text.length}b)`).join(", ")} from ${relative(SITE_DIR, DOCS)}`);
}
