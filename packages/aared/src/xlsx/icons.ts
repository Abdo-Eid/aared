/**
 * XLSX icon-set catalogue: compact SVG inner content for every built-in
 * family, sharing geometry and tier colours. Renderer-only (never the
 * Worker): the Worker resolves a compact `{ family, icon, showValue }` and
 * this module draws it. Icons are recognisable at normal use, not
 * pixel-identical to Excel; paired captures judge them.
 *
 * Catalogue order is lowest tier first (index 0 is the down/red end, the
 * last index the up/green end), so a tier maps directly to an icon unless
 * `reverse` flips it (handled in `core`, before the model is built).
 */

const LOW = "#F8696B";
const MID = "#FFEB84";
const HIGH = "#63BE7B";
const ORANGE = "#FFB84D";
const LIGHT_GREEN = "#A8D86B";
const GREY_LOW = "#BDBDBD";
const GREY_MID = "#9E9E9E";
const GREY_HIGH = "#616161";

function polygon(points: string, fill: string): string {
  return `<polygon points="${points}" fill="${fill}"/>`;
}

function circle(cx: number, cy: number, r: number, fill: string): string {
  return `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${fill}"/>`;
}

function rect(x: number, y: number, w: number, h: number, fill: string): string {
  return `<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="${fill}"/>`;
}

/** Triangle pointing `dir`, centred in the 16px box. */
function arrow(dir: "up" | "right" | "down" | "left" | "upRight" | "downRight", fill: string): string {
  switch (dir) {
    case "up": return polygon("8,2 14,12 2,12", fill);
    case "down": return polygon("2,4 14,4 8,14", fill);
    case "right": return polygon("4,2 12,8 4,14", fill);
    case "left": return polygon("12,2 4,8 12,14", fill);
    case "upRight": return polygon("3,13 13,3 13,10 10,13", fill);
    case "downRight": return polygon("3,3 13,13 6,13 3,10", fill);
  }
}

function flag(fill: string): string {
  return rect(3, 2, 1.6, 12, "#616161") + polygon("5,2 14,6 5,10", fill);
}

function diamond(fill: string): string {
  return polygon("8,2 14,8 8,14 2,8", fill);
}

function check(fill: string): string {
  return `<polyline points="3,8.5 6.5,12 13,4.5" fill="none" stroke="${fill}" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>`;
}

function cross(fill: string): string {
  return `<path d="M4,4 L12,12 M12,4 L4,12" stroke="${fill}" stroke-width="2.4" stroke-linecap="round"/>`;
}

function bang(fill: string): string {
  return rect(7, 3, 2, 7, fill) + circle(8, 12.4, 1.3, fill);
}

/** Five-point star centred at (cx,cy) with outer radius r. */
function star(cx: number, cy: number, r: number, fill: string): string {
  const inner = r * 0.42;
  let points = "";
  for (let i = 0; i < 10; i++) {
    const radius = i % 2 === 0 ? r : inner;
    const angle = (Math.PI / 5) * i - Math.PI / 2;
    points += `${(cx + radius * Math.cos(angle)).toFixed(1)},${(cy + radius * Math.sin(angle)).toFixed(1)} `;
  }
  return polygon(points.trim(), fill);
}

/** Circle with a pie wedge filled clockwise from 12 o'clock: 0, 1/4, 1/2, 3/4, full. */
function quarter(filled: number): string {
  const back = circle(8, 8, 6, "#E0E0E0");
  if (filled === 0) return back;
  if (filled === 4) return circle(8, 8, 6, "#616161");
  const wedge = filled === 1
    ? `<path d="M8,8 L8,2 A6,6 0 0 1 14,8 Z" fill="#616161"/>`
    : filled === 2
      ? `<path d="M8,8 L8,2 A6,6 0 0 1 8,14 L8,8 Z" fill="#616161"/>`
      : `<path d="M8,8 L8,2 A6,6 0 0 1 14,8 L14,14 A6,6 0 0 1 8,14 Z" fill="#616161"/>`;
  return back + wedge;
}

/** Four small boxes in a row, the first `filled` of them dark. */
function boxes(filled: number): string {
  let out = "";
  for (let i = 0; i < 4; i++) {
    out += rect(1 + i * 3.8, 5, 3, 6, i < filled ? "#616161" : "#E0E0E0");
  }
  return out;
}

/** `count` stars in a row, filled yellow (ratings and star tiers). */
function starRow(count: number): string {
  let out = "";
  for (let i = 0; i < count; i++) {
    out += star(3.2 + i * 3.2, 8, 1.9, "#E8A838");
  }
  return out;
}

function traffic(count: number, index: number): string {
  const colors = count === 3 ? [LOW, MID, HIGH] : count === 4 ? [LOW, ORANGE, MID, HIGH] : [LOW, ORANGE, MID, LIGHT_GREEN, HIGH];
  return circle(8, 8, 5.4, colors[index]!);
}

function arrowTier(family: string, index: number): string {
  const grey = family.endsWith("Gray");
  const low = grey ? GREY_LOW : LOW;
  const mid = grey ? GREY_MID : MID;
  const high = grey ? GREY_HIGH : HIGH;
  const mid2 = grey ? GREY_MID : ORANGE;
  const mid3 = grey ? GREY_LOW : LIGHT_GREEN;
  if (family.startsWith("5")) {
    const dirs: Array<"down" | "downRight" | "right" | "upRight" | "up"> = ["down", "downRight", "right", "upRight", "up"];
    const fills = [low, mid2, mid, mid3, high];
    return arrow(dirs[index]!, fills[index]!);
  }
  if (family.startsWith("4")) {
    const dirs: Array<"down" | "right" | "upRight" | "up"> = ["down", "right", "upRight", "up"];
    const fills = [low, mid2, mid, high];
    return arrow(dirs[index]!, fills[index]!);
  }
  const dirs: Array<"down" | "right" | "up"> = ["down", "right", "up"];
  return arrow(dirs[index]!, [low, mid, high][index]!);
}

/**
 * Inner SVG content for `family` icon `index` (0 is the lowest tier).
 * Every branch is a few elements; complex tiers (ratings, quarters, boxes)
 * repeat one shared shape rather than carrying unique paths.
 */
export function iconSvg(family: string, index: number): string {
  switch (family) {
    case "3Arrows":
    case "3ArrowsGray":
    case "4Arrows":
    case "4ArrowsGray":
    case "5Arrows":
    case "5ArrowsGray":
      return arrowTier(family, index);
    case "3TrafficLights1":
    case "3TrafficLights2":
      return traffic(3, index);
    case "4TrafficLights":
      return traffic(4, index);
    case "3Flags":
      return flag([LOW, MID, HIGH][index]!);
    case "3Signs":
      return [cross(LOW), bang(MID), check(HIGH)][index]!;
    case "3Symbols":
      return [circle(8, 8, 5.4, LOW), circle(8, 8, 5.4, MID), circle(8, 8, 5.4, HIGH)][index]!;
    case "3Symbols2":
      return [diamond(LOW), diamond(MID), diamond(HIGH)][index]!;
    case "3Triangles":
      return [polygon("8,13 2,5 14,5", LOW), polygon("8,13 2,5 14,5", MID), polygon("8,13 2,5 14,5", HIGH)][index]!;
    case "4RedToBlack":
      return circle(8, 8, 5.4, ["#E53935", "#757575", "#424242", "#000000"][index]!);
    case "4Rating":
      return starRow(index + 1);
    case "5Rating":
      return starRow(index + 1);
    case "5Quarters":
      return quarter(index);
    case "3Stars":
      return starRow(index + 1);
    case "5Boxes":
      return boxes(index);
    default:
      return circle(8, 8, 5.4, MID);
  }
}

/** How many icons `family` holds (0 for unknown families, which never reach the renderer). */
export function iconCount(family: string): number {
  if (family.startsWith("3")) return 3;
  if (family.startsWith("4")) return 4;
  if (family.startsWith("5")) return 5;
  return 0;
}
