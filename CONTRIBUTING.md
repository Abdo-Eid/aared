# Contributing to aared

Thanks for helping. This page only says where to look; the rules themselves
live in one place each, so they don't drift.

## Before you start

- **Check the gap inventory before opening an issue.** If the difference is
  already in [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md#gaps-limits-and-fidelity-gaps)
  — open, deferred, accepted or by design — it is known; don't file it again.
  Accepted limits are in the same file (Known limits).
- **Open an issue first** for anything bigger than a small fix, so we agree
  on the change before you build it.
- **Read [`AGENTS.md`](AGENTS.md).** It is the working charter for humans and
  agents alike: the rules, commands, package map, testing rule, commit
  conventions and the ground-truth probe for OOXML questions.
- Some things are out of scope by design (editing, canvas text, PDF); see
  the first rules in `AGENTS.md` and the
  [architecture boundaries](docs/ARCHITECTURE.md#purpose-and-boundaries).

## Set up and check your work

Setup and commands are in the README's [Develop](README.md#develop) section
(Bun required). Build first with `bun run build`, then run `bun run typecheck`
and `bun test`. Browser regressions also need production site/bench builds and
the applicable visual checks; prerequisites live in [Tools](tools/README.md).
There is no separate push/PR CI, so run the applicable checks locally before
pushing or opening a pull request. Docs-only changes are exempt; see
[`AGENTS.md § Testing rule`](AGENTS.md#testing-rule).

## Docs and the agent-facing copies

The docs site also emits plain Markdown (`llms.txt`, `start.md`, `api.md`,
`llms-full.txt`) and the npm package ships a README. These are generated or
mirrored from `site/src/content/docs/` — edit the MDX, then rebuild the site;
never hand-edit a generated file. The package README is the first thing an
agent reads after `package.json`, so keep its recipes and boundaries honest;
`packages/aared/test/package-metadata.test.ts` checks its imports against the
real export map.

Discovery metadata uses the same identity: aared (عارض), published as
`@aared/viewer`, a TypeScript Office viewer for PPTX, XLSX and DOCX in the
browser. Keep README introductions, npm description/keywords, GitHub topics and
site metadata consistent without rewriting the product story. The site build
emits a sitemap; `packages/aared/test/site-discovery.visual.test.ts` checks produced HTML/XML, and the
production browser test checks the `/aared/` deployment base. Canonical URLs and
the sitemap must never double that prefix.

The maintainer can submit `https://abdo-eid.github.io/aared/sitemap-index.xml`
to [Google Search Console](https://search.google.com/search-console/about) or
[Bing Webmaster Tools](https://www.bing.com/webmasters/about) after the release deploys, using
an account that can verify the URL-prefix property. Indexing is external and can
lag publication; llms.txt is an agent reading index, not a crawler registration.
Do not add `/aared/robots.txt` as if it controlled the host's root crawler policy.
Assess discovery with fresh, name-only queries and record actual returned sources,
not an LLM's inferred familiarity. npm metadata changes need a new package version;
only the maintainer publishes with 2FA.

## Fixtures

Fixtures are authored in real Office, never generated. Follow the rules in
[`fixtures/README.md`](fixtures/README.md), and ask before adding a new file.

## Pull requests

Follow [`AGENTS.md § Commits`](AGENTS.md#commits). In the description, say
what was missing, what a consumer can do now, which formats it touches, and
how you proved it ([`AGENTS.md § Capability delivery`](AGENTS.md#capability-delivery)).
Update the doc that owns each fact you changed (the "Where to read" table in
`AGENTS.md`).
Follow [Documentation discipline](AGENTS.md#documentation-discipline):

- [`CHANGELOG.md`](CHANGELOG.md) has no entries through the first release.
  Start using Unreleased for user-visible changes after that release; do not
  backfill development history.
- [`docs/Notes.md`](docs/Notes.md) preserves experience that is expensive to
  rediscover, not a journal of contributions. Add evidence-backed Office
  findings, subtle failures or non-obvious rejected approaches when useful;
  routine changes, status reports and validation logs do not need an entry.

## Releases

Follow [Published releases](AGENTS.md#published-releases). Only the maintainer
publishes npm with 2FA. Pages deploys on published stable GitHub Releases from
annotated version tags on `main`, not ordinary pushes; the deployment validates
the tagged checkout before publishing the site.

## License

Contributions are licensed under the project's [Apache-2.0](LICENSE) license.
