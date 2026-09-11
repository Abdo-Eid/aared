const child = Bun.spawn(["bun", "test", "--test-name-pattern", "visual"], {
  cwd: import.meta.dir + "/..",
  env: {
    ...process.env,
    AARED_VISUAL: "1",
    ...(process.argv.includes("--artifacts") ? { AARED_VISUAL_ARTIFACTS: "1" } : {}),
  },
  stdin: "inherit",
  stdout: "inherit",
  stderr: "inherit",
});

process.exit(await child.exited);
