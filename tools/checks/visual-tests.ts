// --artifacts keeps the paired captures and review pages; --absent-fonts also
// renders each capture as a device without Office's fonts would, through the
// real aared-fonts resolver (network; see packages/aared/test/absent-fonts.ts).
import { parseArgs } from "node:util";

const { values } = parseArgs({ args: process.argv.slice(2), options: {
  artifacts: { type: "boolean" }, "absent-fonts": { type: "boolean" },
} });
// Select files, not titles: borrowed-content/search/lifecycle tests need not
// spell "visual" in their names, but must run under the documented command.
const child = Bun.spawn([process.execPath, "test", "visual.test"], {
  cwd: import.meta.dir + "/../..",
  env: {
    ...process.env,
    AARED_VISUAL: "1",
    ...(values.artifacts ? { AARED_VISUAL_ARTIFACTS: "1" } : {}),
    ...(values["absent-fonts"] ? { AARED_VISUAL_ABSENT_FONTS: "1" } : {}),
  },
  stdin: "inherit",
  stdout: "inherit",
  stderr: "inherit",
});

process.exit(await child.exited);
