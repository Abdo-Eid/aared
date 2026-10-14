import { extname, resolve } from "node:path";
import { parseArgs } from "node:util";
import { parseParts, type ExcelPartSpec } from "../review/office-export.js";
import { validateBenchmarkArgs } from "../benchmarks/measure.js";

export type ReviewFormat = "pptx" | "xlsx" | "docx";
export interface ReviewOptions {
  file: string;
  format: ReviewFormat;
  absentFonts: boolean;
  pairFull: boolean;
  parts: ExcelPartSpec[];
  width: number;
}

export type ToolCommand =
  | { kind: "help"; text: string }
  | { kind: "review"; options: ReviewOptions }
  | { kind: "bench"; args: string[] }
  | { kind: "inspect"; file: string; output?: string; bench: boolean }
  | { kind: "check" }
  | { kind: "doctor" };

const HELP: Record<string, string> = {
  tools: `Usage: bun run tools <command> [options]
  review <file>     Native Office export + aared capture + review.html (all three formats)
  bench <action>    run, size, search, calibrate, report
  inspect <file>    Dump resolved Worker models/media hashes, or parser timings
  check consumer-build   Production Vite/webpack packaging checks
  doctor           Report local prerequisites without launching Office
Use <command> --help for options. See tools/README.md for specialist tools.`,
  review: `Usage: bun run tools review <file.pptx|xlsx|docx> [options]
  --absent-fonts    Add a third capture with Office fonts absent (network)
  --width <pixels>  PPTX export/capture width (default 1280)
  --part <sheet:name=A1:G10>  XLSX row band; repeatable, visible sheets numbered from 1
  --pair-full       XLSX: also write full-sheet side-by-side PNGs (requires Pillow)
Outputs: tools/out/<format>/<file-stem>/review.html. An export replaces that file's output folder.
DOCX requires bun run build first; run doctor to check prerequisites.`,
  bench: `Usage: bun run tools bench <run|size|search|calibrate|report> [options]
  run [fixture:renderer ...] [--all]   Open/navigation/memory (aared by default)
  size [--all]                        Download size
  search                              Browser search on generated large fixtures
  calibrate                           Verify page/Worker memory readings
  report [--md] [--site]               Regenerate docs/site from saved runs (both by default)
Browser measurements require bun run build and bun run --cwd bench build.
Benchmarks are manual, on an idle machine on mains power; not ordinary tests.`,
  inspect: `Usage: bun run tools inspect <file> [--out <json>] [--bench]
  --out <json>    Write the resolved model/layout/media-hash dump
  --bench         Parser timings and heap/RSS only (not browser performance)
Default: print timings; --out and --bench are mutually exclusive.`,
  check: "Usage: bun run tools check consumer-build\nRequires bun run build first; builds scratch consumers and runs real browser probes.",
  doctor: "Usage: bun run tools doctor\nReports Bun, builds, Python packages and registered Office applications. Never launches Office or installs anything.",
};

/** Parse without side effects: bad arguments must fail before export cleanup or browser launch. */
export function parseToolCommand(argv: string[], cwd = process.cwd()): ToolCommand {
  const [command, ...args] = argv;
  if (command === undefined || command === "--help" || command === "-h") return { kind: "help", text: HELP.tools! };
  if (!Object.hasOwn(HELP, command) || command === "tools") throw new Error(`Unknown tools command '${command}'. Run bun run tools --help.`);
  if (args.includes("--help") || args.includes("-h")) return { kind: "help", text: HELP[command]! };
  if (command === "review") {
    const { values, positionals } = parseArgs({ args, allowPositionals: true, options: {
      "absent-fonts": { type: "boolean" }, "pair-full": { type: "boolean" },
      width: { type: "string" }, part: { type: "string", multiple: true },
    } });
    if (positionals.length !== 1) throw new Error(HELP.review);
    const file = resolve(cwd, positionals[0]!);
    const format = extname(file).slice(1).toLowerCase();
    if (format !== "pptx" && format !== "xlsx" && format !== "docx") throw new Error("Review requires a .pptx, .xlsx or .docx file.");
    if (format !== "pptx" && values.width !== undefined) throw new Error("--width is only supported for PPTX review.");
    if (format !== "xlsx" && (values.part !== undefined || values["pair-full"])) throw new Error("--part and --pair-full are only supported for XLSX review.");
    const width = values.width === undefined ? 1280 : Number(values.width);
    if (!Number.isInteger(width) || width <= 0) throw new Error("--width requires a positive integer pixel value.");
    const parts = parseParts((values.part ?? []).flatMap((part) => ["--part", part]));
    for (const part of parts) {
      const match = /^A(\d+):([A-Z]+)(\d+)$/.exec(part.range);
      if (match === null || Number(match[1]) <= 0 || Number(match[3]) < Number(match[1])) {
        throw new Error("XLSX review parts require an uppercase row-band range starting at column A, such as A1:G10.");
      }
    }
    return { kind: "review", options: { file, format, width, parts, absentFonts: values["absent-fonts"] ?? false, pairFull: values["pair-full"] ?? false } };
  }
  if (command === "bench") {
    const [action, ...rest] = args;
    const { values, positionals } = parseArgs({ args: rest, allowPositionals: true, options: {
      all: { type: "boolean" }, md: { type: "boolean" }, site: { type: "boolean" },
    } });
    let legacy: string[];
    if (action === "report") {
      if (values.all || positionals.length !== 0) throw new Error(HELP.bench);
      legacy = [...(values.md || !values.site ? ["--render-md"] : []), ...(values.site || !values.md ? ["--render-site"] : [])];
    } else {
      if (action !== "run" && action !== "size" && action !== "search" && action !== "calibrate") throw new Error(HELP.bench);
      if (values.md || values.site) throw new Error("--md and --site are only supported for bench report.");
      legacy = [`--${action}`, ...(values.all ? ["--all"] : []), ...positionals];
    }
    validateBenchmarkArgs(legacy);
    return { kind: "bench", args: legacy };
  }
  if (command === "inspect") {
    const { values, positionals } = parseArgs({ args, allowPositionals: true, options: { out: { type: "string" }, bench: { type: "boolean" } } });
    if (positionals.length !== 1 || (values.out !== undefined && values.bench)) throw new Error(HELP.inspect);
    return { kind: "inspect", file: resolve(cwd, positionals[0]!), bench: values.bench ?? false,
      ...(values.out === undefined ? {} : { output: resolve(cwd, values.out) }) };
  }
  if (command === "check") {
    if (args.length !== 1 || args[0] !== "consumer-build") throw new Error(HELP.check);
    return { kind: "check" };
  }
  if (args.length !== 0) throw new Error(HELP.doctor);
  return { kind: "doctor" };
}
