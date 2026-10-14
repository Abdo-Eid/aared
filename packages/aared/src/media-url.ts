/**
 * Object URLs for the media one render paints (pictures, picture bullets,
 * backgrounds). Renderer-side: it creates `Blob` URLs, which `core` never does.
 *
 * One owner per render, so a URL is created once per media part and shared: a
 * logo repeated on every page is fetched and held once, not once per use. The
 * owner revokes every URL together when the render ends or is replaced, and a
 * request that arrives after that resolves to `undefined` — how a late fetch
 * learns to paint nothing — instead of leaking a URL nobody will revoke.
 *
 * An Office SVG picture keeps its raster beside it (`asvg:svgBlip` next to
 * `a:blip`): `ImageSource` carries both, `image()` resolves both URLs, and
 * `paintImageSource()` paints the SVG, falling back to the raster when the
 * SVG cannot be read or decoded.
 */

import type { MediaResolver } from "./core/index.js";

/**
 * One picture as the Worker resolved it: the preferred source plus Office's
 * own raster fallback beside an SVG (`asvg:svgBlip` next to `a:blip`).
 */
export interface ImageSource {
  mediaId: string;
  contentType: string;
  fallbackMediaId?: string;
  fallbackContentType?: string;
}

/** Resolved URLs for an `ImageSource`: the source, and its fallback when it also read. */
export interface ImageUrls {
  url: string;
  fallbackUrl?: string;
}

/**
 * Point an `<img>` at an image source: the primary first, the raster fallback
 * when the primary cannot be read or decoded, `unavailable` when neither
 * paints. At most one decode retry, so a broken fallback ends in
 * `unavailable`, never a loop. Late completions (torn-down render, detached
 * image) paint nothing. The caller owns `unavailable` (its placeholder or
 * marked fallback); load listeners stay the caller's.
 */
export function paintImageSource(
  img: HTMLImageElement,
  urls: MediaUrls,
  source: ImageSource,
  unavailable: () => void,
): void {
  let fellBack = false;
  img.addEventListener("error", () => {
    if (urls.closed || !img.isConnected) return;
    const fallback = !fellBack ? img.dataset.aaredImageFallback : undefined;
    if (fallback !== undefined) {
      fellBack = true;
      delete img.dataset.aaredImageFallback;
      img.src = fallback;
      return;
    }
    unavailable();
  });
  void urls.image(source).then(
    (resolved) => {
      if (resolved === undefined || !img.isConnected) {
        if (resolved === undefined && img.isConnected && !urls.closed) unavailable();
        return;
      }
      if (resolved.fallbackUrl !== undefined) img.dataset.aaredImageFallback = resolved.fallbackUrl;
      img.src = resolved.url;
    },
    () => {
      if (img.isConnected && !urls.closed) unavailable();
    },
  );
}

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

  /**
   * URLs for one image source: the primary, plus the raster fallback when it
   * also reads. A primary that cannot be read promotes a readable fallback to
   * the source itself. `undefined` when the owner closed or nothing readable
   * arrived — the caller then paints its unavailable state.
   */
  async image(source: ImageSource): Promise<ImageUrls | undefined> {
    const url = await this.url(source.mediaId, source.contentType).catch(() => undefined);
    if (url !== undefined) {
      let fallbackUrl: string | undefined;
      if (source.fallbackMediaId !== undefined && source.fallbackContentType !== undefined) {
        fallbackUrl = await this.url(source.fallbackMediaId, source.fallbackContentType).catch(() => undefined);
      }
      return { url, ...(fallbackUrl === undefined ? {} : { fallbackUrl }) };
    }
    if (source.fallbackMediaId === undefined || source.fallbackContentType === undefined) return undefined;
    const fallback = await this.url(source.fallbackMediaId, source.fallbackContentType).catch(() => undefined);
    return fallback === undefined ? undefined : { url: fallback };
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
