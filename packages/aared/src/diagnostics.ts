declare const __AARED_MARK_FALLBACKS__: boolean | undefined;

/**
 * Development diagnostics — PPTX dashed fallback outlines, the XLSX sheet
 * notice bar — paint when running from source (unit/visual tests,
 * harnesses) and never in the shipped bundles: tsdown `define`s the flag
 * as `false`, so consumers and builds like the bench see neither. The
 * `typeof` guard keeps diagnostics on wherever the build flag is absent.
 */
export const MARK_FALLBACKS = typeof __AARED_MARK_FALLBACKS__ === "undefined" ? true : __AARED_MARK_FALLBACKS__;
