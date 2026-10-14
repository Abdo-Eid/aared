import { expect } from "bun:test";
import { visual } from "../../packages/aared/test/harness.js";
import { reviewPage } from "./page.js";

visual("visual review shell reads equal-width columns and stacks them on narrow screens", async () => {
  const html = reviewPage("Review", '<section class="pair"><figure>Office</figure><figure>aared</figure><figure>Absent fonts</figure></section><section class="pair"><figure>Office</figure><figure>aared</figure></section>');
  await using view = new Bun.WebView({ width: 1200, height: 800 });
  await view.navigate(`data:text/html,${encodeURIComponent(html)}`);
  await view.resize(1200, 800);
  const widths = await view.evaluate(`[...document.querySelectorAll('.pair')].map(pair => [...pair.children].map(child => child.getBoundingClientRect().width))`) as number[][];
  for (const group of widths) for (const width of group) expect(width).toBeCloseTo(group[0]!, 1);
  await view.resize(500, 800);
  const narrow = await view.evaluate(`({ flow: getComputedStyle(document.querySelector('.pair')).flexDirection, width: document.body.scrollWidth, viewport: innerWidth })`) as { flow: string; width: number; viewport: number };
  expect(narrow.flow).toBe("column");
  expect(narrow.width).toBeLessThanOrEqual(narrow.viewport);
}, 30_000);
