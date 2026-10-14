import type { Collection, MediaCollection } from "./core/index.js";

// Renderer-only associations. A facade has no own/symbol/prototype property
// leading to its backing handle, document or session. Not a public entry.
const sources = new WeakMap<object, { source: () => object; signal: AbortSignal }>();

export function borrowContent<T extends object>(source: object, view: T, assertOpen: () => void, signal: AbortSignal): T {
  sources.set(view, { source: () => { assertOpen(); return source; }, signal });
  return Object.freeze(view);
}

/** Only renderers unwrap a borrowed input; stale inputs fail before painting. */
export function contentSource<T extends object>(input: unknown): T | undefined {
  if (typeof input !== "object" || input === null) return undefined;
  return sources.get(input)?.source() as T | undefined;
}

/** Cancel pending secondary renders on owner teardown, without owning their output handles. */
export function contentSignal(input: unknown, signal?: AbortSignal): AbortSignal | undefined {
  const owner = typeof input === "object" && input !== null ? sources.get(input)?.signal : undefined;
  if (owner === undefined) return signal;
  return signal === undefined ? owner : AbortSignal.any([owner, signal]);
}

export function contentCollection<T, K, V>(source: Collection<T, K>, map: (item: T) => V,
  assertOpen: () => void): Collection<V, K> {
  const mapped = new Map<T, V>();
  const get = (item: T): V => {
    if (!mapped.has(item)) mapped.set(item, map(item));
    return mapped.get(item)!;
  };
  return Object.freeze({
    list: (): readonly V[] => { assertOpen(); return Object.freeze(source.list().map(get)); },
    get: (key: K): V => { assertOpen(); return get(source.get(key)); },
  });
}

export function contentMedia(source: MediaCollection, assertOpen: () => void): MediaCollection {
  return Object.freeze({
    get: async (...args: Parameters<MediaCollection["get"]>) => { assertOpen(); return source.get(...args); },
  });
}
