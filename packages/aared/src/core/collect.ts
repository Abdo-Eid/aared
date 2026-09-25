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
