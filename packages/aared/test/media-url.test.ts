import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { MediaUrls } from "../src/media-url.js";

// The owner of one render's object URLs. `URL.createObjectURL` is replaced with
// a counter so leaks (a URL created and never revoked) are observable.
let created: string[];
let revoked: string[];
const real = { create: URL.createObjectURL, revoke: URL.revokeObjectURL };

beforeEach(() => {
  created = [];
  revoked = [];
  URL.createObjectURL = () => {
    const url = `blob:test/${created.length}`;
    created.push(url);
    return url;
  };
  URL.revokeObjectURL = (url: string) => {
    revoked.push(url);
  };
});

afterEach(() => {
  URL.createObjectURL = real.create;
  URL.revokeObjectURL = real.revoke;
});

const bytes = (): ArrayBuffer => new Uint8Array([1, 2, 3]).buffer;

function countingResolver(): { resolve: (id: string) => Promise<ArrayBuffer>; calls: string[] } {
  const calls: string[] = [];
  return { calls, resolve: async (id) => (calls.push(id), bytes()) };
}

describe("MediaUrls", () => {
  test("a media part is fetched and held once, however many places use it", async () => {
    const { resolve, calls } = countingResolver();
    const urls = new MediaUrls(resolve);
    const [first, second, third] = await Promise.all([
      urls.url("word/media/logo.png", "image/png"),
      urls.url("word/media/logo.png", "image/png"),
      urls.url("word/media/other.png", "image/png"),
    ]);
    expect(first).toBe(second);
    expect(third).not.toBe(first);
    expect(calls).toEqual(["word/media/logo.png", "word/media/other.png"]);
    expect(created).toHaveLength(2);
  });

  test("close revokes every URL exactly once and is idempotent", async () => {
    const urls = new MediaUrls(countingResolver().resolve);
    await urls.url("a", "image/png");
    await urls.url("b", "image/png");
    expect(urls.closed).toBe(false);
    urls.close();
    urls.close();
    expect(urls.closed).toBe(true);
    expect(revoked.sort()).toEqual([...created].sort());
    expect(revoked).toHaveLength(2);
  });

  test("bytes that arrive after close create no URL, so nothing leaks", async () => {
    let release!: (value: ArrayBuffer) => void;
    const pending = new Promise<ArrayBuffer>((resolve) => (release = resolve));
    const urls = new MediaUrls(() => pending);
    const late = urls.url("slow", "image/png");
    urls.close();
    release(bytes());
    expect(await late).toBeUndefined();
    expect(created).toEqual([]);
  });

  test("a request after close resolves to nothing without fetching", async () => {
    const { resolve, calls } = countingResolver();
    const urls = new MediaUrls(resolve);
    urls.close();
    expect(await urls.url("a", "image/png")).toBeUndefined();
    expect(calls).toEqual([]);
  });

  test("a failed fetch rejects every user of that part, and only that part", async () => {
    const failure = new Error("media unavailable");
    const calls: string[] = [];
    const urls = new MediaUrls(async (id) => {
      calls.push(id);
      if (id === "bad") throw failure;
      return bytes();
    });
    await expect(urls.url("bad", "image/png")).rejects.toBe(failure);
    await expect(urls.url("bad", "image/png")).rejects.toBe(failure);
    expect(await urls.url("good", "image/png")).toStartWith("blob:");
    expect(calls).toEqual(["bad", "good"]);
  });

  test("a resolver that throws synchronously is a rejection, not an escape", async () => {
    const urls = new MediaUrls(() => {
      throw new Error("sync failure");
    });
    await expect(urls.url("a", "image/png")).rejects.toThrow("sync failure");
  });
});
