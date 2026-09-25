/**
 * Serves a DOCX fixture to a browser page, for the visual tests and for
 * `tools/word-render/capture.ts`. Not a test itself (no `.test.` in the name).
 * The page is the caller's: this only provides what every DOCX page needs.
 */

import type { DocumentSection } from "../src/core/model.js";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import { OpcPackage } from "../src/core/opc.js";
import { browserModule } from "./harness.js";

const DOCX_ENTRY = `${import.meta.dir}/../src/docx/index.ts`;

/**
 * Parse the fixture, bundle the renderer, and serve them: the renderer at
 * `/docx.mjs`, the fixture's media parts at `/media/<file>`, and `page` (an
 * HTML document that imports `/docx.mjs`) for every other path. The caller
 * stops the server and hands the returned sections to its page to render.
 */
export async function serveDocx(
  fixturePath: string,
  page: string,
): Promise<{ server: ReturnType<typeof Bun.serve>; sections: DocumentSection[] }> {
  const pkg = await OpcPackage.open(await Bun.file(fixturePath).arrayBuffer());
  const sections = await parseDocxSections(pkg, await indexDocx(pkg));
  const media = new Map<string, Uint8Array>();
  for (const part of pkg.partNames()) {
    if (part.startsWith("word/media/")) media.set(part, await pkg.readPart(part));
  }
  const renderer = await browserModule(DOCX_ENTRY);
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/docx.mjs") {
        return new Response(renderer, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
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
