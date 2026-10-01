/**
 * Model walks shared by the worker (which `mediaId`s may the media route
 * serve?) and the main thread (which `fontFamily`s should the font resolver
 * load?). Internal; not re-exported.
 *
 * Both walks used to visit every object of a unit through `Object.entries`,
 * which for a million-cell sheet meant millions of throwaway arrays. This
 * walk uses `for…in`, and knows the one model where size explodes: a sheet's
 * cells can only carry the key through their (shared) `style` objects, so
 * each distinct style is visited once instead of once per cell.
 */

import type { ShapeGeometryModel, SheetModel } from "./model.js";

function isSheet(model: object): model is SheetModel {
  return "rows" in model && "columnWidths" in model && Array.isArray((model as SheetModel).rows);
}

/** Add every non-empty string stored under `key` anywhere in `model` to `into`. */
export function collectStrings(model: unknown, key: string, into: Set<string>): void {
  const visit = (value: unknown): void => {
    if (value === null || typeof value !== "object") return;
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    const record = value as Record<string, unknown>;
    for (const name in record) {
      const child = record[name];
      if (name === key && typeof child === "string") {
        if (child !== "") into.add(child);
      } else if (child !== null && typeof child === "object") {
        visit(child);
      }
    }
  };
  if (model === null || typeof model !== "object" || !isSheet(model)) {
    visit(model);
    return;
  }
  const styles = new Set<object>();
  for (const row of model.rows) {
    for (const entry of row.cells) {
      const style = entry.cell.style;
      if (style !== undefined && !styles.has(style)) {
        styles.add(style);
        visit(style);
      }
    }
  }
  for (const name in model) {
    if (name !== "rows") visit((model as unknown as Record<string, unknown>)[name]);
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
  const visit = (value: unknown): void => {
    if (value === null || typeof value !== "object") return;
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    const record = value as Record<string, unknown>;
    for (const name in record) {
      const child = record[name];
      if (typeof child === "string") {
        if (name === "fontFamily" && child !== "") {
          const variant: FontVariant = record.italic === true
            ? (record.bold === true ? "italic bold " : "italic ")
            : (record.bold === true ? "bold " : "");
          let variants = families.get(child);
          if (variants === undefined) families.set(child, (variants = new Set()));
          variants.add(variant);
        } else if (!arabic && name === "text") arabic = ARABIC_TEXT.test(child);
      } else if (child !== null && typeof child === "object") {
        visit(child);
      }
    }
  };
  visit(model);
  return { families, arabic };
}

/**
 * Every shape geometry in a PPTX model (slide, layout, nested groups) that the
 * worker left for the session to evaluate. See `ShapeGeometryModel.deferred`.
 */
export function collectDeferredGeometry(model: unknown): ShapeGeometryModel[] {
  const found: ShapeGeometryModel[] = [];
  const visit = (value: unknown): void => {
    if (value === null || typeof value !== "object") return;
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    const record = value as Record<string, unknown>;
    if (typeof record.path === "string" && record.deferred !== undefined && typeof record.viewBox === "object") {
      found.push(record as unknown as ShapeGeometryModel);
      return;
    }
    for (const name in record) {
      const child = record[name];
      if (child !== null && typeof child === "object") visit(child);
    }
  };
  visit(model);
  return found;
}
