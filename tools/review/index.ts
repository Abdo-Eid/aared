import { join } from "node:path";
import { pathToFileURL } from "node:url";
import type { ReviewOptions } from "../cli/options.js";
import { command } from "../shared/office.js";
import { renderPowerPoint, renderWord } from "./office-export.js";

/** Existing format backends keep their own layout policy and artifact directories. */
export async function reviewFile(options: ReviewOptions): Promise<{ review: string; outputDir: string }> {
  if (!(await Bun.file(options.file).exists())) throw new Error(`Review input does not exist: ${options.file}`);
  let outputDir: string;
  switch (options.format) {
    case "pptx": {
      console.log("PPTX: native export and DOM/SVG slide capture");
      const { capturePptx } = await import("./pptx.js");
      const reference = await renderPowerPoint(options.file, options.width);
      outputDir = reference.outputDir;
      await capturePptx(options.file, reference, { absentFonts: options.absentFonts });
      break;
    }
    case "xlsx": {
      console.log("XLSX: native A4 print and DOM sheet/part capture");
      const { captureXlsx } = await import("./xlsx.js");
      const result = await captureXlsx(options.file, options);
      outputDir = result.outputDir;
      break;
    }
    case "docx": {
      // Preflight before the native exporter replaces existing artifacts. The
      // capture owns the requirements, including the shared chunk graph.
      const { assertDocxReviewBuild } = await import("./docx/browser-server.js");
      await assertDocxReviewBuild();
      console.log("DOCX: native export, layout capture and token comparison");
      const reference = await renderWord(options.file);
      outputDir = reference.outputDir;
      const wordTool = (name: string, args: string[]) => command("Word review", process.execPath, [join(import.meta.dir, "docx", name), ...args]);
      console.log((await wordTool("capture.ts", [options.file, outputDir])).trim());
      if (options.absentFonts) console.log((await wordTool("capture.ts", [options.file, outputDir, "--absent-fonts"])).trim());
      console.log((await wordTool("compare-pagination.ts", [outputDir])).trim());
      await wordTool("review.ts", [outputDir]);
      break;
    }
  }
  const review = join(outputDir, "review.html");
  console.log(pathToFileURL(review).href);
  return { review, outputDir };
}
