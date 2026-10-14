/**
 * Word review transforms private models (including absent-font simulations),
 * so it uses the shared internal painter harness, not the public handle facade.
 * The previous entry had a Windows inferred-root EBADF trap; the relocated
 * renderer bundles successfully through browserModule, with production flags.
 */
import { readdir } from "node:fs/promises";
import { join, resolve } from "node:path";
import { indexDocx, parseDocxSections } from "../../../packages/aared/src/core/docx/document.js";
import { OpcPackage } from "../../../packages/aared/src/core/opc.js";
import { browserModule } from "../../../packages/aared/test/harness.js";

const AARED_DIST = resolve(import.meta.dir, "../../../packages/aared/dist");
const FONTS_DIST = resolve(import.meta.dir, "../../../packages/aared-fonts/dist");

/** Validate capture's actual module files before native export replaces artifacts. */
export async function assertDocxReviewBuild(dist = AARED_DIST): Promise<void> {
  const core = join(dist, "core");
  let files: string[];
  try {
    // Bun.file().exists() is false for directories, even in a complete build.
    files = (await readdir(core, { withFileTypes: true }))
      .filter(entry => entry.isFile() && entry.name.endsWith(".mjs"))
      .map(entry => entry.name);
  } catch (error) {
    if (!["ENOENT", "ENOTDIR"].includes((error as NodeJS.ErrnoException).code ?? "")) throw error;
    throw new Error("Built package chunks are missing; run bun run build first.");
  }
  if (!files.includes("index.mjs")) {
    throw new Error("Built core entry is missing; run bun run build first.");
  }
  // A leftover entry alone is not a complete build: every sibling import in
  // the directory capture serves must resolve, including content-hashed chunks.
  for (const name of files) {
    const source = await Bun.file(join(core, name)).text();
    for (const match of source.matchAll(/(?:from|import)\s*\(?\s*["'`](\.{1,2}\/[^"'`]+)["'`]/g)) {
      if (!await Bun.file(resolve(core, match[1]!)).exists()) {
        throw new Error(`Built core chunk is missing (${name} -> ${match[1]}); run bun run build first.`);
      }
    }
  }
}

export async function serveDocxReview(
  fixturePath: string,
  page: string,
  modules: Record<string, string> = {},
): Promise<{ server: ReturnType<typeof Bun.serve>; sections: Awaited<ReturnType<typeof parseDocxSections>> }> {
  const docxEntry = await browserModule(resolve(import.meta.dir, "../../../packages/aared/src/docx/renderer.ts"), { production: true });
  const coreFiles = new Set((await readdir(join(AARED_DIST, "core"))).filter((file) => file.endsWith(".mjs")));

  const pkg = await OpcPackage.open(await Bun.file(fixturePath).arrayBuffer());
  const sections = await parseDocxSections(pkg, await indexDocx(pkg));
  const media = new Map<string, Uint8Array>();
  for (const part of pkg.partNames()) {
    if (part.startsWith("word/media/")) media.set(part, await pkg.readPart(part));
  }

  const fontsEntry = Bun.file(join(FONTS_DIST, "index.mjs"));
  const fontsModule = (await fontsEntry.exists())
    ? (await fontsEntry.text()).replaceAll('from "@aared/viewer/core"', 'from "/core/index.mjs"')
    : undefined;

  const server = Bun.serve({
    port: 0,
    async fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/docx.mjs") {
        return new Response(docxEntry, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      if (url.pathname.startsWith("/core/")) {
        const name = url.pathname.slice("/core/".length);
        if (coreFiles.has(name)) {
          return new Response(Bun.file(join(AARED_DIST, "core", name)), { headers: { "content-type": "text/javascript; charset=utf-8" } });
        }
        return new Response("missing renderer chunk", { status: 404 });
      }
      if (url.pathname === "/fonts/index.mjs" && fontsModule !== undefined) {
        return new Response(fontsModule, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      const module = modules[url.pathname];
      if (module !== undefined) return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      if (url.pathname.startsWith("/media/")) {
        const id = `word/media/${url.pathname.slice("/media/".length)}`;
        const data = media.get(id);
        if (data === undefined) return new Response("missing media", { status: 404 });
        return new Response(data.slice().buffer as ArrayBuffer, {
          headers: { "content-type": id.endsWith(".png") ? "image/png" : "image/jpeg" },
        });
      }
      return new Response(page, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  return { server, sections };
}

/** Transpile the optional browser helper without bundling its filesystem imports. */
export async function absentFontsBrowserModule(): Promise<string> {
  if (!(await Bun.file(join(FONTS_DIST, "index.mjs")).exists())) {
    throw new Error("aared Word review: built font resolver is missing; run `bun run build` first");
  }
  const sourcePath = resolve(import.meta.dir, "../../../packages/aared-fonts/test/absent-fonts.browser.ts");
  const source = (await Bun.file(sourcePath).text())
    .replace('from "@aared/viewer/core"', 'from "/core/index.mjs"')
    .replace('from "../src/index.js"', 'from "/fonts/index.mjs"');
  return new Bun.Transpiler({ loader: "ts", target: "browser" }).transformSync(source);
}
