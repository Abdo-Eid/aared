import type { ShapeModel, SlideElementModel, SlideModel, TableModel, TextBodyModel } from "../src/core/model.js";

const MANIFEST_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.json`;

interface TextSelector {
  hasText: string;
}

interface CoverageManifest {
  id: string;
  slides: Record<string, TextSelector>;
}

export function joined(body: TextBodyModel): string {
  return body.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join("")).join("\n");
}

/** Depth-first flatten of a slide's elements, expanding groups so selectors see nested children. */
export function flattenElements(elements: readonly SlideElementModel[]): SlideElementModel[] {
  const flat: SlideElementModel[] = [];
  for (const element of elements) {
    if (element.kind === "group") flat.push(...flattenElements(element.children));
    else flat.push(element);
  }
  return flat;
}

function elementText(element: SlideElementModel): string {
  if (element.kind === "shape") return element.text === undefined ? "" : joined(element.text);
  if (element.kind === "table") return element.rows.flatMap((row) => row.cells.map((cell) => joined(cell.text))).join("\n");
  if (element.kind === "group") return element.children.map(elementText).join("\n");
  return "";
}

function slideSummary(slide: SlideModel): string {
  return flattenElements(slide.elements).map((element) => `${element.kind}:${element.name}:${elementText(element).slice(0, 80)}`).join(" | ");
}

export async function coverageManifest(): Promise<CoverageManifest> {
  return JSON.parse(await Bun.file(MANIFEST_PATH).text()) as CoverageManifest;
}

export async function findCoverageSlide(slides: SlideModel[], id: string): Promise<{ index: number; slide: SlideModel }> {
  const manifest = await coverageManifest();
  const selector = manifest.slides[id];
  if (selector === undefined) throw new Error(`fixture ${manifest.id}: unknown slide selector "${id}"`);
  const matches = slides
    .map((slide, index) => ({ index, slide }))
    .filter(({ slide }) => flattenElements(slide.elements).some((element) => elementText(element).includes(selector.hasText)));
  if (matches.length === 1) return matches[0]!;
  const candidates = slides.map((slide, index) => `${index}: ${slideSummary(slide)}`).join("\n");
  throw new Error(`fixture ${manifest.id}: selector "${id}" expected one slide containing "${selector.hasText}", found ${matches.length}\n${candidates}`);
}

function findOne<T extends SlideElementModel>(
  slide: SlideModel,
  description: string,
  predicate: (element: SlideElementModel) => element is T,
): T {
  const matches = flattenElements(slide.elements).filter(predicate);
  if (matches.length === 1) return matches[0]!;
  throw new Error(`fixture selector ${description} expected one element, found ${matches.length}: ${slideSummary(slide)}`);
}

export function findShapeWithText(slide: SlideModel, text: string): ShapeModel {
  return findOne(slide, `shape text "${text}"`, (element): element is ShapeModel => element.kind === "shape" && element.text !== undefined && joined(element.text).includes(text));
}

export function findShapesWithPreset(slide: SlideModel, preset: string): ShapeModel[] {
  const matches = flattenElements(slide.elements).filter((element): element is ShapeModel => element.kind === "shape" && element.geometry?.preset === preset);
  if (matches.length > 0) return matches;
  throw new Error(`fixture selector preset "${preset}" found no elements: ${slideSummary(slide)}`);
}

export function findTableWithText(slide: SlideModel, text: string): TableModel {
  return findOne(slide, `table text "${text}"`, (element): element is TableModel => element.kind === "table" && elementText(element).includes(text));
}
