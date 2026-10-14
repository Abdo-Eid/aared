/** Cheap DrawingML transform painting, independent of optional visual effects. */
import type { TransformModel } from "./core/model.js";

/** Apply rotation/flips to a positioned paint box around its center. */
export function paintTransform(element: HTMLElement, transform: TransformModel | undefined): void {
  if (transform === undefined) return;
  const parts: string[] = [];
  if (transform.rotation !== undefined) parts.push(`rotate(${transform.rotation}deg)`);
  if (transform.flipHorizontal === true || transform.flipVertical === true) {
    parts.push(`scale(${transform.flipHorizontal === true ? -1 : 1}, ${transform.flipVertical === true ? -1 : 1})`);
  }
  if (parts.length === 0) return;
  element.style.transformOrigin = "center";
  element.style.transform = parts.join(" ");
}
