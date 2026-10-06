/**
 * Not a test (no `.test.` in the name): reopens a real fixture with one XML part
 * rewritten, so a case can state exactly what differs from the Office-authored
 * file (a hostile count, an extra rule) without a new fixture file.
 */

import { OpcPackage } from "../src/core/opc.js";
import { storedZip } from "./zip.js";

export const FIXTURES = `${import.meta.dir}/../../../fixtures`;

/** A copy of `fixtures/<file>` with `part` rewritten by `rewrite`, opened as a package. */
export async function withPart(file: string, part: string, rewrite: (xml: string) => string): Promise<OpcPackage> {
  const source = await OpcPackage.open(await Bun.file(`${FIXTURES}/${file}`).arrayBuffer());
  const parts: Record<string, Uint8Array> = {};
  for (const name of source.partNames()) parts[name] = await source.readPart(name);
  parts[part] = new TextEncoder().encode(rewrite(new TextDecoder().decode(parts[part])));
  return OpcPackage.open(storedZip(parts).slice().buffer as ArrayBuffer);
}
