/**
 * Browser side of the no-Office-fonts review (`packages/aared/test/
 * absent-fonts.ts` bundles it for the capture pages): make a
 * model paint the way it would on a device with none of Office's fonts (a
 * phone, Linux, macOS). An installed font cannot be hidden from a page (a
 * registered face of the same name that fails to load falls back to the
 * installed one), so every family the model names is renamed to one no
 * machine has, and the real `aared-fonts` resolver delivers each substitute
 * under the renamed family's alias. A family with no substitute falls to the
 * generic, as it would there.
 */
import { fontAlias } from "aared/core";
import { createResolver } from "../src/index.js";

const ABSENT = " absent";
/** Every model key the renderers paint a family name from. */
const FONT_KEYS = new Set(["fontFamily", "typeface"]);

export async function simulateAbsentFonts<T>(model: T): Promise<T> {
  const renamed = structuredClone(model);
  const families = new Set<string>();
  const visit = (value: unknown): void => {
    if (value === null || typeof value !== "object") return;
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    const record = value as Record<string, unknown>;
    for (const key in record) {
      const child = record[key];
      if (FONT_KEYS.has(key) && typeof child === "string" && child !== "" && !child.endsWith(ABSENT)) {
        families.add(child);
        record[key] = child + ABSENT;
      } else visit(child);
    }
  };
  visit(renamed);
  const aliasOf = new Map([...families].map((family) => [fontAlias(family), fontAlias(family + ABSENT)]));
  const resolver = createResolver({
    FontFaceCtor: class extends FontFace {
      constructor(family: string, source: string | BufferSource, descriptors?: FontFaceDescriptors) {
        super(aliasOf.get(family) ?? family, source, descriptors);
      }
    },
  });
  await Promise.all([...families].map((family) => resolver(family)));
  return renamed;
}

(globalThis as { simulateAbsentFonts?: typeof simulateAbsentFonts }).simulateAbsentFonts = simulateAbsentFonts;
