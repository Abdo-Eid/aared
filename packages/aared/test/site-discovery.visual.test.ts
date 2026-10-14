import { expect } from "bun:test";
import { visual } from "./harness.js";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

const DIST = resolve(import.meta.dir, "../../../site/dist");
const SITE = "https://abdo-eid.github.io/aared/";
const ROUTES = ["", "start/", "api/", "playground/"];

/** Read real emitted HTML, independent of Astro's component source. */
function attribute(html: string, tag: string, key: string, value: string, name: string): string | undefined {
  const tags = html.match(new RegExp(`<${tag}\\b[^>]*>`, "g")) ?? [];
  const node = tags.find(text => text.includes(`${key}="${value}"`));
  return node?.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1];
}

visual("production pages connect the brand to npm and emit one canonical URL under the project base", () => {
  for (const route of ROUTES) {
    const html = readFileSync(resolve(DIST, route, "index.html"), "utf8");
    const canonical = `${SITE}${route}`;
    expect(attribute(html, "link", "rel", "canonical", "href"), route).toBe(canonical);
    expect((html.match(/rel="canonical"/g) ?? []).length).toBe(1);
    expect(attribute(html, "link", "rel", "sitemap", "href")).toBe("/aared/sitemap-index.xml");
    const title = html.match(/<title>([^<]+)<\/title>/)?.[1];
    expect(title).toContain("aared");
    expect(title).toContain("عارض");
    expect(title).toContain("@aared/viewer");
    const description = attribute(html, "meta", "name", "description", "content");
    for (const term of ["aared", "@aared/viewer", "PPTX", "XLSX", "DOCX"]) {
      expect(description, `${route} description omits ${term}`).toContain(term);
    }
    expect(attribute(html, "meta", "property", "og:url", "content")).toBe(canonical);
    expect(attribute(html, "meta", "property", "og:title", "content")).toBe(title);
    expect(attribute(html, "meta", "property", "og:description", "content")).toBe(description);
    expect(attribute(html, "meta", "name", "twitter:card", "content")).toBe("summary");
    expect(html).not.toContain("/aared/aared/");
    expect(html).not.toMatch(/name="robots"[^>]*content="[^"]*noindex/);
  }
});

visual("production sitemap lists all public HTML pages once, with no doubled base or Markdown duplicates", () => {
  const index = readFileSync(resolve(DIST, "sitemap-index.xml"), "utf8");
  const sitemapUrls = [...index.matchAll(/<loc>([^<]+)<\/loc>/g)].map(match => match[1]!);
  expect(sitemapUrls).toHaveLength(1);
  expect(sitemapUrls[0]).toBe(`${SITE}sitemap-0.xml`);
  const sitemap = readFileSync(resolve(DIST, "sitemap-0.xml"), "utf8");
  const pages = [...sitemap.matchAll(/<loc>([^<]+)<\/loc>/g)].map(match => match[1]!);
  expect(pages.sort()).toEqual(ROUTES.map(route => `${SITE}${route}`).sort());
  expect(new Set(pages).size).toBe(pages.length);
  expect(sitemap).not.toContain("/aared/aared/");
  expect(sitemap).not.toMatch(/\.md<|llms.*\.txt</);
});

visual("production landing page exposes installation and fidelity evidence without requiring client JavaScript", () => {
  const html = readFileSync(resolve(DIST, "index.html"), "utf8");
  expect(html).toContain('href="https://www.npmjs.com/package/@aared/viewer"');
  expect(html).toContain("published on npm");
  expect(html).toContain("TypeScript Office viewer");
  expect(html).toContain("native PowerPoint, Excel and Word output");
  expect(html).toContain("docs/COMPETITORS.md");
  expect(html).toContain("/api/#fidelity-is-part-of-the-contract");
});
