import { describe, expect, test } from "bun:test";
import { existsSync } from "node:fs";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import { OpcPackage } from "../src/core/opc.js";

/**
 * The parsed sections are a pure function of the file: how the Worker reads
 * `document.xml` (whole tree or streamed `w:body` children) and whether it
 * keeps anything afterwards must not change a single field. The committed
 * fixtures are pinned by snapshot; the larger local files (generated or
 * Word-exported, not committed) by digest, when they are present.
 */

const ROOT = `${import.meta.dir}/../../..`;

async function sectionsJson(path: string): Promise<string> {
  const pkg = await OpcPackage.open(await Bun.file(path).arrayBuffer());
  return JSON.stringify(await parseDocxSections(pkg, await indexDocx(pkg)), null, 1);
}

describe("DOCX sections do not depend on how document.xml is read", () => {
  for (const name of ["docx-coverage.docx", "docx-outline.docx"]) {
    test(name, async () => {
      const first = await sectionsJson(`${ROOT}/fixtures/${name}`);
      expect(first).toMatchSnapshot();
      // A second parse of the same package is the same model: nothing cached changes it.
      expect(await sectionsJson(`${ROOT}/fixtures/${name}`)).toBe(first);
    });
  }

  const LOCAL: Record<string, string> = {
    "tools/bench-search/out/large-40x.docx": "b25e3312fd4e5295ce0d171c7a9b1e5ab0a31454fe98ecd8bd17658788d52991",
    "tools/word-render/out/arabic-docx-bidi-coverage/word-source.docx": "a1e7bf569e452544698826b3d5e310116394f35116f9f0489245ced50f9f26de",
    "tools/word-render/out/probe-arab-matrix/word-source.docx": "ae6813c2283d2cc7ae001d3f3225a4d14a878dff3594a916d599054ed00f4caf",
    "tools/word-render/out/probe-arab-theme/word-source.docx": "5aeacab796a235200bdeb5bdb818b7324d1a740e426a20a38899b151b2bee16b",
  };
  for (const [path, digest] of Object.entries(LOCAL)) {
    test.skipIf(!existsSync(`${ROOT}/${path}`) || digest === "")(path, async () => {
      const json = await sectionsJson(`${ROOT}/${path}`);
      expect(new Bun.CryptoHasher("sha256").update(json).digest("hex")).toBe(digest);
    });
  }
});
