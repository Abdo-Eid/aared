import { expect } from "bun:test";
import { visual } from "./harness.js";
import { rm } from "node:fs/promises";
import { renderPowerPoint } from "../../../tools/review/office-export.js";
import { capturePptx } from "../../../tools/review/pptx.js";
import { ABSENT_FONTS } from "./absent-fonts.js";

const RETAIN_ARTIFACTS = process.env.AARED_VISUAL_ARTIFACTS === "1";
const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

visual("visual review: PPTX produces same-size PowerPoint and aared slide captures", async () => {
  const reference = await renderPowerPoint(FIXTURE_PATH);
  try {
    const captures = await capturePptx(FIXTURE_PATH, reference, {
      absentFonts: ABSENT_FONTS, retainArtifacts: RETAIN_ARTIFACTS,
    });
    expect(captures).toHaveLength(reference.slideCount);
    for (const capture of captures) {
      expect(capture.canvasCount).toBe(0);
      expect(capture.width).toBe(reference.width);
      expect(capture.height).toBe(reference.height);
      expect(capture.png).toEqual({ width: reference.width, height: reference.height });
      expect(capture.slide.width).toBeGreaterThan(0);
      expect(capture.slide.height).toBeGreaterThan(0);
      expect(capture.slide.width / capture.slide.height).toBeCloseTo(capture.expectedSize.width / capture.expectedSize.height, 2);
    }
  } finally {
    if (!RETAIN_ARTIFACTS) await rm(reference.outputDir, { recursive: true, force: true });
  }
}, ABSENT_FONTS ? 300_000 : 120_000);
