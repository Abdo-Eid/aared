import { join } from "node:path";
import { parseToolCommand } from "./cli/options.js";

const ROOT = join(import.meta.dir, "..");

async function runScript(file: string, args: string[]): Promise<number> {
  const child = Bun.spawn([process.execPath, join(import.meta.dir, file), ...args], {
    cwd: ROOT, stdin: "inherit", stdout: "inherit", stderr: "inherit",
  });
  return child.exited;
}

export async function runTools(argv: string[]): Promise<number> {
  const command = parseToolCommand(argv);
  switch (command.kind) {
    case "help": console.log(command.text); return 0;
    case "review": {
      const { reviewFile } = await import("./review/index.js");
      const result = await reviewFile(command.options);
      console.log(`Review: ${result.review}\nArtifacts: ${result.outputDir}`);
      return 0;
    }
    case "bench": return runScript("benchmarks/measure.ts", command.args);
    case "inspect": return runScript("inspect/core.ts", [command.file, ...(command.output === undefined ? [] : [command.output]), ...(command.bench ? ["bench"] : [])]);
    case "check": return runScript("checks/consumer-build.ts", []);
    case "doctor": {
      const { doctor } = await import("./cli/doctor.js");
      await doctor();
      return 0;
    }
  }
}

if (import.meta.main) {
  try {
    process.exitCode = await runTools(process.argv.slice(2));
  } catch (error) {
    console.error(`aared tools: ${error instanceof Error ? error.message : String(error)}`);
    process.exitCode = 1;
  }
}
