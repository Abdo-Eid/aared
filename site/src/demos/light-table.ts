import { defaultResolver } from "aared-fonts";
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

export async function runLightTable(stage: HTMLElement, source: string | File): Promise<DemoOutcome> {
  const renderPanel = stage.querySelector<HTMLElement>('[data-panel="render"]');
  if (!renderPanel) throw new Error("light table: stage is missing [data-panel='render']");

  const format = formatOf(source);

  if (format === "xlsx") {
    const { open, render } = await import("aared/xlsx");
    const workbook = await open(source);
    try {
      const sheet = workbook.sheets.get(0);
      const viewer = await render({
        sheet,
        target: renderPanel,
        zoom: 0.7,
        appearance: { spreadsheet: { sheetTabs: true } },
      });
      return {
        summary: { format: workbook.info.format, units: workbook.info.count },
        dispose: () => { viewer.destroy(); workbook.close(); },
      };
    } catch (error) {
      workbook.close();
      throw error;
    }
  }

  if (format === "docx") {
    const { open, render } = await import("aared/docx");
    const document = await open(source);
    try {
      const viewer = await render({ document, target: renderPanel, zoom: 0.65 });
      return {
        summary: { format: document.info.format, units: document.info.count },
        dispose: () => { viewer.destroy(); document.close(); },
      };
    } catch (error) {
      document.close();
      throw error;
    }
  }

  const { open, render } = await import("aared/pptx");
  const deck = await open(source, { fontResolver: defaultResolver });
  try {
    const slide = deck.slides.get(0);
    const viewer = await render({ slide, target: renderPanel });
    return {
      summary: { format: deck.info.format, units: deck.info.count },
      dispose: () => { viewer.destroy(); deck.close(); },
    };
  } catch (error) {
    deck.close();
    throw error;
  }
}
