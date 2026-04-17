import { calculateCagr, CagrResult, VALUE_ERROR } from "../core/cagr";

/**
 * Task pane CAGR calculator (spec §3.13). Validates input before delegating
 * to the pure core. Returns the numeric CAGR or the #VALUE! sentinel.
 */
export function runCagrCalculator(values: number[]): CagrResult {
  if (!Array.isArray(values) || values.length === 0) {
    return VALUE_ERROR;
  }
  for (const v of values) {
    if (typeof v !== "number" || !Number.isFinite(v)) {
      return VALUE_ERROR;
    }
  }
  return calculateCagr(values);
}
