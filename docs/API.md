# API

The public API reference lives in the docs site:
[`site/src/content/docs/api.mdx`](../site/src/content/docs/api.mdx)
(`bun run docs:dev`, then open `/api`). It is kept honest by the
`docs-parity` test (`packages/aared/test/docs-parity.test.ts`): every public
export must be named there.
