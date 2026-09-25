/**
 * Object URLs for the media one render paints (pictures, picture bullets,
 * backgrounds). Renderer-side: it creates `Blob` URLs, which `core` never does.
 *
 * One owner per render, so a URL is created once per media part and shared: a
 * logo repeated on every page is fetched and held once, not once per use. The
 * owner revokes every URL together when the render ends or is replaced, and a
 * request that arrives after that resolves to `undefined` — how a late fetch
 * learns to paint nothing — instead of leaking a URL nobody will revoke.
 */

import type { MediaResolver } from "./core/index.js";

export class MediaUrls {
  readonly #resolver: MediaResolver;
  readonly #pending = new Map<string, Promise<string | undefined>>();
  readonly #created = new Set<string>();
  #closed = false;

  constructor(resolver: MediaResolver) {
    this.#resolver = resolver;
  }

  /** True once `close()` has run: later work should paint nothing. */
  get closed(): boolean {
    return this.#closed;
  }

  /**
   * The shared object URL for a media part. Rejects when the resolver fails (a
   * failure is shared by every user of the part); resolves to `undefined` when
   * the owner closed before the bytes arrived.
   */
  url(mediaId: string, contentType: string): Promise<string | undefined> {
    if (this.#closed) return Promise.resolve(undefined);
    let url = this.#pending.get(mediaId);
    if (url === undefined) {
      url = Promise.resolve()
        .then(() => this.#resolver(mediaId))
        .then((bytes) => {
          if (this.#closed) return undefined;
          const created = URL.createObjectURL(new Blob([bytes as BlobPart], { type: contentType }));
          this.#created.add(created);
          return created;
        });
      this.#pending.set(mediaId, url);
    }
    return url;
  }

  /** Revoke every URL and refuse new ones. Idempotent. */
  close(): void {
    if (this.#closed) return;
    this.#closed = true;
    for (const url of this.#created) {
      try {
        URL.revokeObjectURL(url);
      } catch {
        // Revocation is best-effort; a failed revoke cannot affect rendering.
      }
    }
    this.#created.clear();
    this.#pending.clear();
  }
}
