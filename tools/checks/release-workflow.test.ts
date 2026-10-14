import { expect, test } from "bun:test";
import { join } from "node:path";

test("Pages deploys release tags only and validates the tagged package before upload", async () => {
  const workflows = join(import.meta.dir, "../../.github/workflows");
  const workflow = Bun.YAML.parse(await Bun.file(join(workflows, "deploy.yml")).text()) as {
    on: Record<string, unknown>;
    permissions: Record<string, string>;
    jobs: Record<string, { if?: string; needs?: string; permissions?: Record<string, string>; steps: {
      uses?: string; run?: string; with?: Record<string, unknown>; env?: Record<string, string>;
    }[] }>;
  };
  expect(workflow.on).toEqual({ release: { types: ["published"] } });
  expect(workflow.permissions).toEqual({ contents: "read" });
  expect(await Bun.file(join(workflows, "ci.yml")).exists()).toBe(false);
  const build = workflow.jobs.build!;
  expect(build.if).toContain("!github.event.release.prerelease");
  expect(build.permissions).toEqual({ contents: "read", pages: "read" });
  const checkout = build.steps.find(step => step.uses?.startsWith("actions/checkout@"))!;
  expect(checkout.with).toEqual({ ref: "${{ github.event.release.tag_name }}", "fetch-depth": 0 });
  expect(build.steps.find(step => step.uses?.startsWith("oven-sh/setup-bun@"))!.with).toEqual({ "bun-version": "1.4.3" });
  const guard = build.steps.find(step => step.env?.RELEASE_TAG)!;
  expect(guard.env!.RELEASE_TAG).toBe("${{ github.event.release.tag_name }}");
  expect(guard.run).toContain('test "$RELEASE_TAG" = "v$version"');
  expect(guard.run).toContain('git cat-file -t "refs/tags/$RELEASE_TAG"');
  expect(guard.run).toContain("git merge-base --is-ancestor HEAD refs/remotes/origin/main");
  expect(build.steps.filter(step => step.run?.startsWith("bun ")).map(step => step.run)).toEqual([
    "bun install --frozen-lockfile", "bun run build", "bun run typecheck", "bun test", "bun run docs:build",
  ]);
  expect(build.steps.findIndex(step => step.run === "bun test")).toBeLessThan(build.steps.findIndex(step => step.uses?.startsWith("actions/upload-pages-artifact@")));
  expect(workflow.jobs.deploy!.needs).toBe("build");
  expect(workflow.jobs.deploy!.permissions).toEqual({ pages: "write", "id-token": "write" });
});
