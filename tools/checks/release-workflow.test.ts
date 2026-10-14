import { expect, test } from "bun:test";
import { dirname, join, resolve } from "node:path";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";

const WORKFLOWS = join(import.meta.dir, "../../.github/workflows");
interface Workflow {
  on: Record<string, unknown>;
  permissions: Record<string, string>;
  jobs: Record<string, { if?: string; needs?: string; permissions?: Record<string, string>; steps: {
    name?: string; if?: string; uses?: string; run?: string; with?: Record<string, unknown>; env?: Record<string, string>;
  }[] }>;
}
async function readWorkflow(): Promise<Workflow> {
  return Bun.YAML.parse(await Bun.file(join(WORKFLOWS, "deploy.yml")).text()) as Workflow;
}

test("Pages validates published package identity before either deployment route uploads", async () => {
  const workflow = await readWorkflow();
  expect(workflow.on).toEqual({ release: { types: ["published"] }, workflow_dispatch: null });
  expect(workflow.permissions).toEqual({ contents: "read" });
  expect(await Bun.file(join(WORKFLOWS, "ci.yml")).exists()).toBe(false);
  const build = workflow.jobs.build!;
  expect(build.if).toContain("!github.event.release.prerelease");
  expect(build.if).toContain("github.event_name == 'workflow_dispatch'");
  expect(build.permissions).toEqual({ contents: "read", pages: "read" });
  const checkout = build.steps.find(step => step.uses?.startsWith("actions/checkout@"))!;
  expect(checkout.with).toEqual({ ref: "${{ github.event_name == 'release' && github.event.release.tag_name || github.ref }}", "fetch-depth": 0 });
  expect(build.steps.find(step => step.uses?.startsWith("oven-sh/setup-bun@"))!.with).toEqual({ "bun-version": "1.4.3" });
  const guard = build.steps.find(step => step.env?.RELEASE_TAG)!;
  expect(guard.if).toBe("${{ github.event_name == 'release' }}");
  expect(guard.env!.RELEASE_TAG).toBe("${{ github.event.release.tag_name }}");
  expect(guard.run).toContain('test "$RELEASE_TAG" = "v$version"');
  expect(guard.run).toContain('git cat-file -t "refs/tags/$RELEASE_TAG"');
  expect(guard.run).toContain("git merge-base --is-ancestor HEAD refs/remotes/origin/main");
  const docsGuard = build.steps.find(step => step.env?.DEPLOY_REF)!;
  expect(docsGuard.if).toBe("${{ github.event_name == 'workflow_dispatch' }}");
  expect(docsGuard.env).toEqual({ GH_TOKEN: "${{ github.token }}", DEPLOY_REF: "${{ github.ref }}" });
  expect(docsGuard.run).toContain('test "$DEPLOY_REF" = refs/heads/main');
  expect(docsGuard.run).toContain('releases/latest" --jq .tag_name');
  expect(docsGuard.run).toContain('test "$release_tag" = "v$version"');
  expect(docsGuard.run).toContain('git cat-file -t "refs/tags/$release_tag"');
  expect(docsGuard.run).toContain('git merge-base --is-ancestor "$release_commit" HEAD');
  expect(docsGuard.run).toContain('npm view "@aared/viewer@$version" gitHead --registry=https://registry.npmjs.org/');
  expect(docsGuard.run).toContain('test "$published_commit" = "$release_commit"');
  expect(docsGuard.run).toContain("-- packages/aared ':!packages/aared/test' package.json bun.lock tsconfig.base.json");
  expect(build.steps.filter(step => step.run?.startsWith("bun ")).map(step => step.run)).toEqual([
    "bun install --frozen-lockfile", "bun run build", "bun run typecheck", "bun test", "bun run docs:build",
  ]);
  expect(build.steps.findIndex(step => step.run === "bun test")).toBeLessThan(build.steps.findIndex(step => step.uses?.startsWith("actions/upload-pages-artifact@")));
  expect(workflow.jobs.deploy!.needs).toBe("build");
  expect(workflow.jobs.deploy!.permissions).toEqual({ pages: "write", "id-token": "write" });
});

interface GuardCase {
  name: string;
  changes?: Record<string, string>;
  ref?: string;
  releaseTag?: string;
  publishedCommit?: string;
  lightweightTag?: boolean;
  accepted: boolean;
}

const CASES: GuardCase[] = [
  { name: "site and regression-test updates", changes: { "site/src/pages/index.astro": "verification metadata", "packages/aared/test/site.test.ts": "regression" }, accepted: true },
  { name: "unpublished renderer changes", changes: { "packages/aared/src/index.ts": "changed library" }, accepted: false },
  { name: "unpublished package README changes", changes: { "packages/aared/README.md": "changed npm material" }, accepted: false },
  { name: "changed dependency lockfile", changes: { "bun.lock": "new dependencies" }, accepted: false },
  { name: "changed root build manifest", changes: { "package.json": '{"private":true,"scripts":{"build":"different"}}' }, accepted: false },
  { name: "changed shared TypeScript config", changes: { "tsconfig.base.json": '{"compilerOptions":{"target":"ES2020"}}' }, accepted: false },
  { name: "unpublished manifest version", changes: { "packages/aared/package.json": '{"name":"@aared/viewer","version":"0.1.2"}' }, accepted: false },
  { name: "a non-main ref", ref: "refs/heads/experiment", accepted: false },
  { name: "a different latest release", releaseTag: "v0.1.0", accepted: false },
  { name: "npm source differing from the release", publishedCommit: "0".repeat(40), accepted: false },
  { name: "an unannotated release tag", lightweightTag: true, accepted: false },
];

/** Execute the workflow's actual Bash guard; only remote API responses are stubbed. */
async function exerciseGuard(options: GuardCase): Promise<{ code: number; output: string }> {
  const scratch = join(tmpdir(), "opencode");
  await mkdir(scratch, { recursive: true });
  const cwd = await mkdtemp(join(scratch, "aared-deploy-guard-"));
  const git = (...args: string[]): string => {
    const result = Bun.spawnSync(["git", ...args], { cwd, stdout: "pipe", stderr: "pipe" });
    if (result.exitCode !== 0) throw new Error(`Fixture git failed: ${result.stderr.toString()}`);
    return result.stdout.toString().trim();
  };
  const write = async (files: Record<string, string>): Promise<void> => {
    for (const [path, content] of Object.entries(files)) {
      const file = join(cwd, path);
      await mkdir(dirname(file), { recursive: true });
      await writeFile(file, content);
    }
  };
  try {
    git("init", "--initial-branch=main");
    git("config", "user.name", "Deployment guard fixture");
    git("config", "user.email", "guard@example.invalid");
    await write({
      "packages/aared/package.json": '{"name":"@aared/viewer","version":"0.1.1"}',
      "packages/aared/src/index.ts": "published library",
      "packages/aared/README.md": "published package material",
      "package.json": '{"private":true}',
      "bun.lock": "published dependencies",
      "tsconfig.base.json": "{}",
    });
    git("add", ".");
    git("commit", "-m", "Published package");
    const releaseCommit = git("rev-parse", "HEAD");
    if (options.lightweightTag) git("tag", "v0.1.1");
    else git("tag", "-a", "v0.1.1", "-m", "Published version");
    git("remote", "add", "origin", ".");
    await write({ "site/src/pages/index.astro": "documentation update", ...options.changes });
    git("add", ".");
    git("commit", "-m", "Documentation update");
    const version = (await Bun.file(join(cwd, "packages/aared/package.json")).json() as { version: string }).version;
    const workflow = await readWorkflow();
    const guard = workflow.jobs.build!.steps.find(step => step.env?.DEPLOY_REF)!.run!;
    // Git for Windows includes the same shell used by Actions; no network,
    // personal gh/npm login or global git configuration participates here.
    const bash = process.platform === "win32" ? resolve(git("--exec-path"), "../../..", "bin/bash.exe") : "bash";
    const stubs = `bun() { printf '%s\\n' "$TEST_VERSION"; }
gh() { printf '%s\\n' "$TEST_TAG"; }
npm() { printf '%s\\n' "$TEST_COMMIT"; }
`;
    const result = Bun.spawnSync([bash, "--noprofile", "--norc", "-c", stubs + guard], {
      cwd, stdout: "pipe", stderr: "pipe",
      env: {
        ...process.env,
        DEPLOY_REF: options.ref ?? "refs/heads/main",
        GITHUB_REPOSITORY: "fixture/viewer",
        GH_TOKEN: "",
        TEST_VERSION: version,
        TEST_TAG: options.releaseTag ?? "v0.1.1",
        TEST_COMMIT: options.publishedCommit ?? releaseCommit,
      },
    });
    return { code: result.exitCode, output: result.stdout.toString() + result.stderr.toString() };
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
}

for (const scenario of CASES) {
  test(`documentation deployment ${scenario.accepted ? "accepts" : "rejects"} ${scenario.name}`, async () => {
    const result = await exerciseGuard(scenario);
    if (scenario.accepted) expect(result.code, result.output).toBe(0);
    else expect(result.code, result.output).not.toBe(0);
  }, 30000);
}
