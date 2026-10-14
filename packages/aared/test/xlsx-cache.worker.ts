// Test-only observation in a real Worker. Count streaming worksheet parses at
// the OPC boundary; no production diagnostics or RPC messages are added.
import { OpcPackage } from "../src/core/opc.js";
import type { WorkerRequest } from "../src/core/rpc.js";

let pause = false;
let resume: (() => void) | undefined;
const readXml = OpcPackage.prototype.readXml;
OpcPackage.prototype.readXml = async function (part, signal, options) {
  if (options?.streamChildrenOf === "sheetData") {
    postMessage({ type: "test-sheet-parse", part });
    if (pause) {
      pause = false;
      postMessage({ type: "test-paused" });
      await new Promise<void>((resolve) => { resume = resolve; });
    }
  }
  return readXml.call(this, part, signal, options);
};

await import("../src/core/worker.js");
const scope = globalThis as unknown as { onmessage: (event: MessageEvent<WorkerRequest>) => void };
const receive = scope.onmessage;
scope.onmessage = (event) => {
  const type = event.data.type as string;
  if (type === "test-pause") pause = true;
  else if (type === "test-resume") { resume?.(); resume = undefined; }
  else receive(event);
};
