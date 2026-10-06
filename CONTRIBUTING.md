# Contributing to aared

Thanks for helping. This page only says where to look; the rules themselves
live in one place each, so they don't drift.

## Before you start

- **Check the roadmap before opening an issue.** If the gap is already in
  [`docs/ROADMAP.md`](docs/ROADMAP.md), it is known; don't file it again.
  Accepted limits are in [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md)
  (Known limits).
- **Open an issue first** for anything bigger than a small fix, so we agree
  on the change before you build it.
- **Read [`AGENTS.md`](AGENTS.md).** It is the working charter for humans and
  agents alike: the rules, commands, package map, testing rule, commit
  conventions and the ground-truth probe for OOXML questions.
- Some things are out of scope by design (editing, canvas text, PDF); see
  the first rules in `AGENTS.md` and
  [`docs/DECISION_LOG.md`](docs/DECISION_LOG.md).

## Set up and check your work

Setup and commands are in the README's [Develop](README.md#develop) section
(Bun required). Before opening a pull request, run `bun run typecheck`,
`bun test` and `bun run build`. Docs-only changes are exempt; see
[`AGENTS.md § Testing rule`](AGENTS.md#testing-rule).

## Fixtures

Fixtures are authored in real Office, never generated. Follow the rules in
[`fixtures/README.md`](fixtures/README.md), and ask before adding a new file.

## Pull requests

Follow [`AGENTS.md § Commits`](AGENTS.md#commits). In the description, say
what was missing, what a consumer can do now, which formats it touches, and
how you proved it ([`AGENTS.md § Capability delivery`](AGENTS.md#capability-delivery)).
Update the doc that owns each fact you changed (the Docs map in `AGENTS.md`).

## License

Contributions are licensed under the project's [Apache-2.0](LICENSE) license.
