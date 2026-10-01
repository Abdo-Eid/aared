// --artifacts keeps the paired captures and review pages; --absent-fonts also
// renders each capture as a device without Office's fonts would, through the
// real aared-fonts resolver (network; see packages/aared/test/absent-fonts.ts).
const child = Bun.spawn(["bun", "test", "--test-name-pattern", "visual"], {
  cwd: import.meta.dir + "/..",
  env: {
    ...process.env,
    AARED_VISUAL: "1",
    ...(process.argv.includes("--artifacts") ? { AARED_VISUAL_ARTIFACTS: "1" } : {}),
    ...(process.argv.includes("--absent-fonts") ? { AARED_VISUAL_ABSENT_FONTS: "1" } : {}),
  },
  stdin: "inherit",
  stdout: "inherit",
  stderr: "inherit",
});

process.exit(await child.exited);
