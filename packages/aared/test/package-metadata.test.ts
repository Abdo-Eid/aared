import { expect, test } from "bun:test";
import { join } from "node:path";

const ASTRO_CONFIG = join(import.meta.dir, "../../../site/astro.config.mjs");

/** The site's origin plus `base`, so README URLs cannot drift from deployment. */
async function siteUrl(): Promise<string> {
  const config = await Bun.file(ASTRO_CONFIG).text();
  const site = /^const site = '([^']+)';$/m.exec(config)?.[1];
  const base = /^const base = '([^']+)';$/m.exec(config)?.[1];
  if (site === undefined || base === undefined) throw new Error("site/astro.config.mjs no longer declares site/base literals");
  // `site` is the full page URL, which already ends with the deployment path.
  const origin = new URL(site).origin;
  return `${origin}${base}`;
}

test("npm metadata identifies the public viewer package and includes its license and installation guide", async () => {
  const dir = join(import.meta.dir, "..");
  const pkg = await Bun.file(join(dir, "package.json")).json();
  expect(pkg.name).toBe("@aared/viewer");
  expect(pkg.private).toBeUndefined();
  expect(pkg.publishConfig).toEqual({ access: "public" });
  expect(pkg.files).toEqual(["dist"]);
  expect(Object.keys(pkg.exports).sort()).toEqual(["core", "docx", "element", "pptx", "preview", "react", "worker", "xlsx"].map(name => `./${name}`).sort());
  expect(pkg.repository.directory).toBe("packages/aared");
  expect(pkg.homepage).toBe("https://abdo-eid.github.io/aared/");
  const normalize = (text: string): string => text.replaceAll("\r\n", "\n").trim();
  expect(normalize(await Bun.file(join(dir, "LICENSE")).text())).toBe(normalize(await Bun.file(join(dir, "../../LICENSE")).text()));
  expect(await Bun.file(join(dir, "README.md")).text()).toContain("npm install @aared/viewer");
});

test("the shipped README's imports, layers and doc links match the package and the deployed site", async () => {
  const dir = join(import.meta.dir, "..");
  const pkg = await Bun.file(join(dir, "package.json")).json();
  const readme = await Bun.file(join(dir, "README.md")).text();
  const exported = new Set(Object.keys(pkg.exports).map(name => `@aared/viewer/${name.slice(2)}`));

  // Every subpath the README tells an agent to import must be real. The bare
  // package name is not importable, so only subpaths are collected.
  const specifiers = new Set([...readme.matchAll(/@aared\/viewer\/[\w-]+/g)].map(match => match[0]));
  expect(specifiers.size).toBeGreaterThanOrEqual(7);
  for (const specifier of specifiers) expect(exported, `README imports ${specifier}`).toContain(specifier);

  // One recipe per advertised layer; `worker` is deployment input, not a layer.
  for (const layer of ["element", "preview", "react", "pptx", "xlsx", "docx"]) {
    expect(readme, `README is missing a recipe for ${layer}`).toContain(`@aared/viewer/${layer}`);
  }
  for (const boundary of ["View-only", "recalculat", "Word", "fonts"]) {
    expect(readme, `README omits the "${boundary}" boundary`).toContain(boundary);
  }

  // Doc links resolve under the deployed base, including the agent entry points.
  const site = await siteUrl();
  for (const link of ["start/", "api/", "llms.txt"]) {
    expect(readme, `README does not link ${link}`).toContain(`${site}${link}`);
  }
  expect(readme).toContain("dist/**/*.d.mts");
});
