/** Prefix site-owned URLs only; example code and external URLs are unchanged. */
export function withSiteBase(path: string, base: string): string {
  if (!path.startsWith("/") || path.startsWith("//")) return path;
  const prefix = `${base.replace(/\/$/, "")}/`;
  if (path.startsWith(prefix)) return path;
  return prefix + path.slice(1);
}

export function sitePath(path: string): string {
  return withSiteBase(path, import.meta.env.BASE_URL ?? "/");
}
