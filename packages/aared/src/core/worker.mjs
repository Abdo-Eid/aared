// Source-mode worker entry: `WorkerSession` starts its worker with
// `new URL("./worker.mjs", import.meta.url)`, so when tests (or any
// source-mode consumer) run from `src/`, that URL resolves here. Bun's
// Worker runtime transpiles the `.ts` import on the fly; the published
// `dist/core/worker.mjs` is the bundled equivalent (see tsdown.config.ts).
import "./worker.ts";
