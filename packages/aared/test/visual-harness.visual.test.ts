/**
 * The shared page runner's lifecycle, which the headless harness tests cannot
 * cover because they never launch a browser: each call gets its own server and
 * its own page, and the server is stopped on the way out -- including one the
 * caller supplied, and including when the body throws. Without this, a
 * converted test that lost its `finally { server.stop(); }` would still pass
 * while leaking a listening socket for the whole run.
 */
import { expect } from "bun:test";
import { serveVisual, visual, withVisualPage } from "./harness.js";

const PAGE = `<!doctype html><p id="marker"></p><script type="module">
  document.querySelector('#marker').textContent = 'page';
</script>`;

visual("visual harness: each page is isolated and its server is stopped on return", async () => {
  const reached: string[] = [];
  const first = await withVisualPage({ page: PAGE, width: 400, height: 300 }, async (view) => {
    reached.push(await view.evaluate<string>(`document.querySelector('#marker').textContent`));
    return view.url;
  });
  reached.push(await withVisualPage({ page: PAGE, width: 400, height: 300 }, async (view) =>
    view.evaluate<string>(`document.querySelector('#marker').textContent`)));
  expect(reached).toEqual(["page", "page"]);
  // Each call served its own origin, and both are gone now.
  expect(first.startsWith("http://")).toBe(true);
  await expect(fetch(first).catch((error: Error) => error)).resolves.toBeInstanceOf(Error);
});

visual("visual harness: a caller-supplied server is stopped unless it must outlive the call", async () => {
  const server = serveVisual({ page: PAGE });
  const url = server.url.href;
  await withVisualPage({ server, width: 400, height: 300 }, async (view) => {
    expect(await view.evaluate<string>(`document.querySelector('#marker').textContent`)).toBe("page");
  });
  await expect(fetch(url).catch((error: Error) => error)).resolves.toBeInstanceOf(Error);

  // `keepServer` is the documented escape hatch for a server shared across
  // calls (the lazy-chunk test's request log lives in the server's closure).
  const shared = serveVisual({ page: PAGE });
  const sharedUrl = shared.url.href;
  try {
    for (const expected of ["page", "page"]) {
      expect(await withVisualPage({ server: shared, keepServer: true, width: 400, height: 300 }, async (view) =>
        view.evaluate<string>(`document.querySelector('#marker').textContent`))).toBe(expected);
    }
  } finally {
    shared.stop(true);
  }
  await expect(fetch(sharedUrl).catch((error: Error) => error)).resolves.toBeInstanceOf(Error);
});

visual("visual harness: a failing page still tears its server down", async () => {
  const server = serveVisual({ page: PAGE });
  const url = server.url.href;
  await expect(withVisualPage({ server, width: 400, height: 300 }, async () => {
    throw new Error("visual harness: deliberate failure");
  })).rejects.toThrow("deliberate failure");
  await expect(fetch(url).catch((error: Error) => error)).resolves.toBeInstanceOf(Error);
});

visual("visual harness: withVisualPage needs a page or a server", async () => {
  // Both are optional in the type (a test supplies one or the other), so the
  // guard is what rejects a call that supplies neither.
  await expect(withVisualPage({ width: 400, height: 300 }, async () => undefined)).rejects.toThrow("needs a page or a server");
});