import { describe, expect, it } from "vitest";
import { runCagrCalculator } from "../src/services/cagr.service";
import { VALUE_ERROR } from "../src/core/cagr";

describe("CAGR calculator contract (spec §3.13)", () => {
  it("computes CAGR from a two-value range", () => {
    const result = runCagrCalculator([100, 121]);
    expect(typeof result).toBe("number");
    expect(result as number).toBeCloseTo(0.21, 5);
  });

  it("computes CAGR from a three-value range", () => {
    const result = runCagrCalculator([100, 110, 121]);
    expect(typeof result).toBe("number");
    expect(result as number).toBeCloseTo(0.1, 5);
  });

  it("returns #VALUE! when start is zero", () => {
    expect(runCagrCalculator([0, 121])).toBe(VALUE_ERROR);
  });

  it("returns #VALUE! when only one value is supplied", () => {
    expect(runCagrCalculator([100])).toBe(VALUE_ERROR);
  });

  it("returns #VALUE! when start is negative", () => {
    expect(runCagrCalculator([-100, 121])).toBe(VALUE_ERROR);
  });

  it("returns #VALUE! when end is zero", () => {
    expect(runCagrCalculator([100, 0])).toBe(VALUE_ERROR);
  });

  it("returns #VALUE! for empty input", () => {
    expect(runCagrCalculator([])).toBe(VALUE_ERROR);
  });

  it("rejects non-finite numbers", () => {
    expect(runCagrCalculator([100, NaN])).toBe(VALUE_ERROR);
    expect(runCagrCalculator([100, Infinity])).toBe(VALUE_ERROR);
  });
});
