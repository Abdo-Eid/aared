import { AaredError, isAbort } from "./core/errors.js";

/** Only the owner of a built-in action reports it; public calls throw/reject instead. */
export function reportBackgroundError(target: EventTarget, scope: string, failure: unknown): void {
  if (isAbort(failure)) return;
  const cause = failure instanceof AaredError ? failure : new AaredError("internal",
    `${scope}: ${failure instanceof Error ? failure.message : String(failure)}`, { cause: failure });
  target.dispatchEvent(new CustomEvent("error", { detail: { message: cause.message, cause } }));
}
