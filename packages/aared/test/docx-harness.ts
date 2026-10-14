/**
 * Serves a DOCX fixture to a browser page, for the visual tests and for
 * `tools/review/docx/capture.ts`. Not a test itself (no `.test.` in the name).
 * The page is the caller's: this only provides what every DOCX page needs.
 */

import type { DocumentSection } from "../src/core/model.js";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import { OpcPackage } from "../src/core/opc.js";
import { browserModule, serveVisual } from "./harness.js";

const DOCX_ENTRY = `${import.meta.dir}/../src/docx/index.ts`;

/**
 * Parse the fixture, bundle the renderer, and serve them: the renderer at
 * `/docx.mjs`, the fixture's media parts at `/media/<file>`, and `page` (an
 * HTML document that imports `/docx.mjs`) for every other path, plus any
 * extra `modules` (path to JavaScript source). The caller owns the server:
 * it either stops the server itself and hands the sections to its own page,
 * or passes the server to `withVisualPage`, which drives and stops it.
 */
export async function serveDocx(
  fixturePath: string,
  page: string,
  modules: Record<string, string> = {},
): Promise<{ server: ReturnType<typeof Bun.serve>; sections: DocumentSection[] }> {
  const pkg = await OpcPackage.open(await Bun.file(fixturePath).arrayBuffer());
  const sections = await parseDocxSections(pkg, await indexDocx(pkg));
  const media = new Map<string, Uint8Array>();
  for (const part of pkg.partNames()) {
    if (part.startsWith("word/media/")) media.set(part, await pkg.readPart(part));
  }
  const renderer = await browserModule(DOCX_ENTRY);
  // Static handling (renderer, extra modules, fallback page) is shared with
  // every other visual test; only the media route keeps its own content-type
  // sniffing and missing-media 404, unchanged from the direct server.
  const server = serveVisual({
    // The renderer wins its own path: extra `modules` are additional scripts.
    modules: { ...modules, "/docx.mjs": renderer },
    page,
    onRequest: (url) => {
      if (url.pathname.startsWith("/media/")) {
        const id = `word/media/${url.pathname.slice("/media/".length)}`;
        const data = media.get(id);
        if (data === undefined) return new Response("missing media", { status: 404 });
        return new Response(data.slice().buffer as ArrayBuffer, {
          headers: { "content-type": id.endsWith(".png") ? "image/png" : "image/jpeg" },
        });
      }
      return undefined;
    },
  });
  return { server, sections };
}
