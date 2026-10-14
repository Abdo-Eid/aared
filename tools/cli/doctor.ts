import { join } from "node:path";
import { command } from "../shared/office.js";
import { REPO_ROOT } from "../shared/paths.js";

/** Capability report only: never create a COM application or install dependencies. */
export async function doctor(): Promise<void> {
  console.log(`Bun ${process.versions.bun}; ${process.platform} ${process.arch}; WebView ${typeof Bun.WebView === "function" ? "available" : "unavailable (Bun 1.4+ required)"}`);
  const root = REPO_ROOT;
  for (const file of ["packages/aared/dist/docx/index.mjs", "packages/aared/dist/core/worker.mjs", "bench/dist/measure.html"]) {
    console.log(`${await Bun.file(join(root, file)).exists() ? "OK" : "MISSING"} ${file}`);
  }
  const python = Bun.which("python") ?? Bun.which("python3");
  if (python === null) console.log("MISSING Python (DOCX/XLSX native rasterization)");
  else {
    const packages = await command("prerequisite check", python, ["-c",
      "import importlib.util,json,sys; print(json.dumps({'python':sys.version.split()[0], **{m:importlib.util.find_spec(m) is not None for m in ['fitz','PIL','win32com','fontTools','pptx','openpyxl','lxml']}}))",
    ]);
    console.log(`Python packages: ${packages.trim()}`);
    console.log("fitz (PyMuPDF): DOCX/XLSX review; PIL (Pillow): XLSX PNG pairs; other packages: specialist research/probes.");
  }
  const pwsh = Bun.which("pwsh");
  if (process.platform !== "win32" || pwsh === null) console.log("Native review requires Windows, pwsh and the matching Office application.");
  else {
    const office = await command("prerequisite check", pwsh, ["-NoProfile", "-NonInteractive", "-Command",
      "@('PowerPoint.Application','Excel.Application','Word.Application') | ForEach-Object { $_ + ': ' + ([type]::GetTypeFromProgID($_) -ne $null) }",
    ]);
    console.log(`Registered Office applications (not launched):\n${office.trim()}`);
  }
  console.log("Build missing entries with bun run build / bun run --cwd bench build. No files or applications changed.");
}
