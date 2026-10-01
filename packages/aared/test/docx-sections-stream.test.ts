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
  for (const name of ["docx-coverage.docx"]) {
    test(name, async () => {
      const first = await sectionsJson(`${ROOT}/fixtures/${name}`);
      expect(first).toMatchSnapshot();
      // A second parse of the same package is the same model: nothing cached changes it.
      expect(await sectionsJson(`${ROOT}/fixtures/${name}`)).toBe(first);
    });
  }

  const LOCAL: Record<string, string> = {
    "tools/bench-search/out/large-40x.docx": "b59d90d1cb0f6a1c61835574b2b2fa7938abdbd0f98a4b1d89b083f1a7576dfb",
    "tools/word-render/out/arabic-docx-bidi-coverage/word-source.docx": "38a6df39fb9fa20ae3dc05dd7c530f0c0ed5ca0282e25a2a3f4639d10f1e38dd",
    "tools/word-render/out/probe-arab-matrix/word-source.docx": "93e3c62a9cd82d046909d06922f500001c422a4892c717b7b7307f495b7308fd",
    "tools/word-render/out/probe-arab-theme/word-source.docx": "ed5ffa5d3e4b6039861c56417c91ca1f8f10c89a2ac29777b5caf5a5aa3a4bc2",
  };
  for (const [path, digest] of Object.entries(LOCAL)) {
    test.skipIf(!existsSync(`${ROOT}/${path}`) || digest === "")(path, async () => {
      const json = await sectionsJson(`${ROOT}/${path}`);
      expect(new Bun.CryptoHasher("sha256").update(json).digest("hex")).toBe(digest);
    });
  }
});
