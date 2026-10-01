/**
 * The no-Office-fonts review: with `AARED_VISUAL_ABSENT_FONTS=1` (`bun run
 * test:visual:artifacts --absent-fonts`), the paired capture suites also
 * render every unit the way a device without Office's fonts would, through
 * the real `aared-fonts` resolver and Google Fonts (network), and write those
 * captures beside the others as a third review column. The browser side
 * lives with the resolver it drives: `packages/aared-fonts/test/
 * absent-fonts.browser.ts`.
 */
export const ABSENT_FONTS = process.env.AARED_VISUAL_ABSENT_FONTS === "1";

/** The browser module that defines `window.simulateAbsentFonts(model)`. */
export async function absentFontsModule(): Promise<string> {
  const result = await Bun.build({ entrypoints: [`${import.meta.dir}/../../aared-fonts/test/absent-fonts.browser.ts`], target: "browser", format: "esm" });
  if (!result.success) throw new Error(result.logs.map((log) => log.message).join("\n"));
  return result.outputs[0]!.text();
}
