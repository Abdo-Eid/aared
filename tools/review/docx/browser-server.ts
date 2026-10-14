/**
 * Serve the built DOCX renderer to Word-review WebViews without calling
 * `Bun.build()`. That bundler's inferred root-directory lookup fails with
 * `EBADF` on this Windows checkout. The package build (`tsdown`) emits the
 * renderer and its shared chunks under `packages/aared/dist/`.
 */
import { readdir } from "node:fs/promises";
import { join, resolve } from "node:path";
import { indexDocx, parseDocxSections } from "../../../packages/aared/src/core/docx/document.js";
import { OpcPackage } from "../../../packages/aared/src/core/opc.js";

const AARED_DIST = resolve(import.meta.dir, "../../../packages/aared/dist");
const FONTS_DIST = resolve(import.meta.dir, "../../../packages/aared-fonts/dist");

export async function serveDocxFromDist(
  fixturePath: string,
  page: string,
  modules: Record<string, string> = {},
): Promise<{ server: ReturnType<typeof Bun.serve>; sections: Awaited<ReturnType<typeof parseDocxSections>> }> {
  const docxEntry = Bun.file(join(AARED_DIST, "docx/index.mjs"));
  if (!(await docxEntry.exists())) {
    throw new Error("aared Word review: built renderer is missing; run `bun run build` first");
  }
  const coreFiles = new Set((await readdir(join(AARED_DIST, "core"))).filter((file) => file.endsWith(".mjs")));

  const pkg = await OpcPackage.open(await Bun.file(fixturePath).arrayBuffer());
  const sections = await parseDocxSections(pkg, await indexDocx(pkg));
  const media = new Map<string, Uint8Array>();
  for (const part of pkg.partNames()) {
    if (part.startsWith("word/media/")) media.set(part, await pkg.readPart(part));
  }

  const fontsEntry = Bun.file(join(FONTS_DIST, "index.mjs"));
  const fontsModule = (await fontsEntry.exists())
    ? (await fontsEntry.text()).replaceAll('from "aared/core"', 'from "/core/index.mjs"')
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
    .replace('from "aared/core"', 'from "/core/index.mjs"')
    .replace('from "../src/index.js"', 'from "/fonts/index.mjs"');
  return new Bun.Transpiler({ loader: "ts", target: "browser" }).transformSync(source);
}
