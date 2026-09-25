/**
 * Bench-only stub for Node-native canvas packages (`canvas`, `@napi-rs/canvas`
 * and its platform subpaths). Nothing in the browser needs them: Konva is
 * aliased to its browser build (see vite.config.ts) and unpdf reaches canvas
 * only through a guarded Node branch. If a renderer ever executes a stubbed
 * import in the browser, that column — and only that column — fails visibly
 * instead of breaking the whole bundle.
 */
export default {};
