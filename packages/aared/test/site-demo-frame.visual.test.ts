import { expect } from "bun:test";
import { browserModule, visual, withVisualPage } from "./harness.js";

/**
 * The playground frames start lazily, so a tab that has just been shown can
 * have its initial run and a picker change overlap. Both go into one stage, and
 * a viewer's `destroy()` empties its target, so a superseded run finishing
 * late used to wipe the newer run's paint (the status said "ready" over an
 * empty stage). `attachDemo` now hands every run a signal and aborts the
 * earlier one before the newer one starts.
 */
visual("visual DOM output: a playground frame cancels a superseded run before the newer one paints", async () => {
  const frame = await browserModule(`${import.meta.dir}/../../../site/src/lib/demo-frame.ts`);
  const page = `<!doctype html><div id="frame" data-initial="slow"><select data-pick><option value="">-</option><option value="fast">fast</option></select><p data-status></p><div data-stage></div></div>
    <script type="module">
      import { attachDemo } from "/frame.mjs";
      window.log = [];
      attachDemo(document.getElementById("frame"), () => (stage, source, signal) => {
        window.log.push("start " + source + (signal === undefined ? " (no signal)" : ""));
        signal?.addEventListener("abort", () => window.log.push("abort " + source));
        if (source === "slow") {
          // A run that is still opening when the newer one starts: it rejects when cancelled, as preview() does.
          return new Promise((resolve, reject) => signal?.addEventListener("abort", () => reject(new Error("cancelled"))));
        }
        stage.textContent = "painted " + source;
        return Promise.resolve({ summary: { format: "xlsx", units: 2 }, dispose: () => { window.log.push("dispose " + source); stage.replaceChildren(); } });
      });
      window.pick = () => {
        const select = document.querySelector("[data-pick]");
        select.value = "fast";
        select.dispatchEvent(new Event("change"));
      };
    </script>`;
  await withVisualPage({ modules: { "/frame.mjs": frame }, page }, async (view) => {
    await view.evaluate("pick()");
    await view.evaluate("new Promise((resolve) => setTimeout(resolve, 100))");
    const result = await view.evaluate(`({
      log: window.log,
      stage: document.querySelector("[data-stage]").textContent,
      state: document.querySelector("[data-status]").dataset.state,
      status: document.querySelector("[data-status]").textContent,
    })`) as { log: string[]; stage: string; state: string; status: string };
    // The earlier run is cancelled, and before the newer one starts.
    expect(result.log).toEqual(["start slow", "abort slow", "start fast"]);
    expect(result.stage).toBe("painted fast");
    expect(result.state).toBe("ready");
    expect(result.status).toBe("xlsx · 2 sheets");
  });
}, 30000);
