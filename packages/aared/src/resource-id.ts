/**
 * Names for DOM paint resources and CSS highlights, not source Office IDs.
 * A module-local counter alone collides when two aared bundles share a page;
 * Office part/shape IDs also repeat across unrelated files. A lazy random
 * namespace separates bundle copies, while the counter keeps allocation cheap.
 * No document, model or per-resource entry is retained here.
 */
let namespace: string | undefined;
let serial = 0;

export function resourceId(kind: string): string {
  namespace ??= Array.from(crypto.getRandomValues(new Uint32Array(4)), (word) => word.toString(36)).join("-");
  return `aared-${kind}-${namespace}-${++serial}`;
}
