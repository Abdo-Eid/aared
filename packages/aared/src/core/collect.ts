/**
 * Model walks shared by the worker (which `mediaId`s may the media route
 * serve?) and the main thread (which `fontFamily`s should the font resolver
 * load?). Internal; not re-exported.
 *
 * One walk visits each object of a unit with `for…in` (`Object.entries` would
 * allocate a throwaway array per object: millions for a large sheet). The one
 * model where size explodes is a sheet: its cells can only carry the key
 * through their (shared) `style` objects, so each distinct style is visited
 * once instead of once per cell.
 */

import type { ShapeGeometryModel, SheetModel } from "./model.js";

/** Call `onRecord` for every object under `value`; it returns true to skip that object's children. */
function walk(value: unknown, onRecord: (record: Record<string, unknown>) => boolean | void): void {
  if (value === null || typeof value !== "object") return;
  if (Array.isArray(value)) {
    for (const item of value) walk(item, onRecord);
    return;
  }
  const record = value as Record<string, unknown>;
  if (onRecord(record) === true) return;
  for (const name in record) {
    const child = record[name];
    if (child !== null && typeof child === "object") walk(child, onRecord);
  }
}

function isSheet(model: object): model is SheetModel {
  return "rows" in model && "columnWidths" in model && Array.isArray((model as SheetModel).rows);
}

/** Add every non-empty string stored under `key` anywhere in `model` to `into`. */
export function collectStrings(model: unknown, key: string, into: Set<string>): void {
  const onRecord = (record: Record<string, unknown>): void => {
    const value = record[key];
    if (typeof value === "string" && value !== "") into.add(value);
  };
  if (model === null || typeof model !== "object" || !isSheet(model)) {
    walk(model, onRecord);
    return;
  }
  const styles = new Set<object>();
  for (const row of model.rows) {
    for (const entry of row.cells) {
      const style = entry.cell.style;
      if (style !== undefined && !styles.has(style)) {
        styles.add(style);
        walk(style, onRecord);
      }
    }
  }
  for (const name in model) {
    if (name !== "rows") walk((model as unknown as Record<string, unknown>)[name], onRecord);
  }
}

// Arabic script blocks (Arabic, Supplement, Extended-B/A, Presentation Forms A/B). Built from a
// string, not written as a regex literal: a minifier prints a literal's `\u` escapes back as raw
// characters, and a single character above U+00FF makes the whole shipped script two bytes per
// character (see `chars.ts`). The pattern's own `\u` escapes stay as text and RegExp reads them.
const ARABIC_TEXT = /* @__PURE__ */ new RegExp("[\\u0600-\\u06FF\\u0750-\\u077F\\u0870-\\u08FF\\uFB50-\\uFDFF\\uFE70-\\uFEFF]");

/** The CSS `font` shorthand prefix of one style a family paints in. */
export type FontVariant = "" | "bold " | "italic " | "italic bold ";

/**
 * Every distinct `fontFamily` in a DOCX or PPTX `model` with the variants it
 * paints in (a run's own `bold`/`italic`), and whether any of its `text` runs
 * contains Arabic script — one walk, for the font wait that must know all
 * three: a delivered face downloads only when something loads it, so the wait
 * loads exactly the faces the document paints, and an Arabic-limited fallback
 * face only when the document has Arabic. Not for spreadsheets: no format
 * that measures text off screen reads a sheet.
 */
export function collectFontUse(model: unknown): { families: Map<string, Set<FontVariant>>; arabic: boolean } {
  const families = new Map<string, Set<FontVariant>>();
  let arabic = false;
  walk(model, (record) => {
    const family = record.fontFamily;
    if (typeof family === "string" && family !== "") {
      const variant: FontVariant = record.italic === true
        ? (record.bold === true ? "italic bold " : "italic ")
        : (record.bold === true ? "bold " : "");
      let variants = families.get(family);
      if (variants === undefined) families.set(family, (variants = new Set()));
      variants.add(variant);
    }
    if (!arabic && typeof record.text === "string") arabic = ARABIC_TEXT.test(record.text);
  });
  return { families, arabic };
}

/**
 * Every shape geometry in a PPTX model (slide, layout, nested groups) that the
 * worker left for the session to evaluate. See `ShapeGeometryModel.deferred`.
 */
export function collectDeferredGeometry(model: unknown): ShapeGeometryModel[] {
  const found: ShapeGeometryModel[] = [];
  walk(model, (record) => {
    if (typeof record.path !== "string" || record.deferred === undefined || typeof record.viewBox !== "object") return;
    found.push(record as unknown as ShapeGeometryModel);
    return true;
  });
  return found;
}
