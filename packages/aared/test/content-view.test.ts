import { expect, mock, test } from "bun:test";
import { AaredError } from "../src/core/errors.js";
import { borrowContent, contentCollection, contentMedia, contentSignal, contentSource } from "../src/content-view.js";

function lifetime() {
  const owner = new AbortController();
  let closed = false;
  return {
    signal: owner.signal,
    assertOpen: () => { if (closed) throw new AaredError("closed", "owner closed"); },
    close: () => { closed = true; owner.abort(); },
  };
}

test("borrowed content is frozen and unwraps only while its owner is open", () => {
  const owner = lifetime();
  const source = { title: "source" };
  const view = borrowContent(source, { info: { count: 1 } }, owner.assertOpen, owner.signal);
  expect(Object.isFrozen(view)).toBe(true);
  expect(Object.getOwnPropertySymbols(view)).toEqual([]);
  expect(contentSource(view)).toBe(source);
  for (const input of [undefined, null, 1, "x", {}]) expect(contentSource(input)).toBeUndefined();
  owner.close();
  expect(() => contentSource(view)).toThrow(AaredError);
});

test("borrowed render cancellation follows either owner or caller without aborting the other", () => {
  for (const cancel of ["owner", "caller"] as const) {
    const owner = lifetime();
    const caller = new AbortController();
    const view = borrowContent({}, {}, owner.assertOpen, owner.signal);
    expect(contentSignal(view)).toBe(owner.signal);
    expect(contentSignal(null, caller.signal)).toBe(caller.signal);
    expect(contentSignal({})).toBeUndefined();
    const joined = contentSignal(view, caller.signal)!;
    expect(joined.aborted).toBe(false);
    if (cancel === "owner") owner.close(); else caller.abort();
    expect(joined.aborted).toBe(true);
    expect(cancel === "owner" ? caller.signal.aborted : owner.signal.aborted).toBe(false);
  }
});

test("borrowed collections memoize inputs, freeze lists and guard both operations after close", () => {
  const owner = lifetime();
  const items = [{ id: 0 }, { id: 1 }];
  const list = mock(() => items);
  const get = mock((key: number) => items[key]!);
  const map = mock((item: { id: number }) => Object.freeze({ id: item.id }));
  const view = contentCollection({ list, get }, map, owner.assertOpen);
  const first = view.get(0);
  const all = view.list();
  expect(all[0]).toBe(first);
  expect(view.get(1)).toBe(all[1]!);
  expect(view.list()).toEqual(all);
  expect(map).toHaveBeenCalledTimes(2);
  expect(Object.isFrozen(view)).toBe(true);
  expect(Object.isFrozen(all)).toBe(true);
  const reads = [list.mock.calls.length, get.mock.calls.length];
  owner.close();
  expect(() => view.list()).toThrow(AaredError);
  expect(() => view.get(0)).toThrow(AaredError);
  expect([list.mock.calls.length, get.mock.calls.length]).toEqual(reads);
});

test("borrowed media forwards the call unchanged and checks owner lifetime before resolving bytes", async () => {
  const owner = lifetime();
  const bytes = new ArrayBuffer(3);
  const get = mock(async (_id: string, _options?: { signal?: AbortSignal }) => bytes);
  const view = contentMedia({ get }, owner.assertOpen);
  const options = { signal: new AbortController().signal };
  expect(await view.get("media", options)).toBe(bytes);
  expect(get.mock.calls).toEqual([["media", options]]);
  expect(Object.isFrozen(view)).toBe(true);
  owner.close();
  await expect(view.get("media")).rejects.toMatchObject({ code: "closed" });
  expect(get).toHaveBeenCalledTimes(1);
});
