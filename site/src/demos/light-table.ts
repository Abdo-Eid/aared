import type { DemoOutcome } from "../lib/demo-frame.js";

function nameOf(source: string | File): string {
  return typeof source === "string" ? source : source.name;
}

function formatOf(source: string | File): "pptx" | "xlsx" | "docx" {
  const name = nameOf(source).toLowerCase();
  if (name.endsWith(".xlsx")) return "xlsx";
  if (name.endsWith(".docx")) return "docx";
  return "pptx";
}

export async function runLightTable(stage: HTMLElement, source: string | File, signal: AbortSignal): Promise<DemoOutcome> {
  const renderPanel = stage.querySelector<HTMLElement>('[data-panel="render"]');
  if (!renderPanel) throw new Error("light table: stage is missing [data-panel='render']");

  const format = formatOf(source);

  const viewer = format === "xlsx"
    ? await (await import("@aared/viewer/xlsx")).preview({ source, target: renderPanel, zoom: 0.7, signal })
    : format === "docx"
      ? await (await import("@aared/viewer/docx")).preview({ source, target: renderPanel, zoom: 0.65, signal })
      : await (await import("@aared/viewer/pptx")).preview({ source, target: renderPanel, signal });
  return {
    summary: { format: viewer.format, units: viewer.info.count },
    dispose: () => viewer.destroy(),
  };
}
