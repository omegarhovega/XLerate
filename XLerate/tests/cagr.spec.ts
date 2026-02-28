import { describe, expect, it } from "vitest";
import { calculateCagr, VALUE_ERROR } from "../src/core/cagr";

describe("CAGR baseline parity", () => {
  it("calculates two-point CAGR", () => {
    expect(calculateCagr([100, 121])).toBeCloseTo(0.21, 10);
  });

  it("calculates multi-period CAGR", () => {
    expect(calculateCagr([100, 110, 121])).toBeCloseTo(0.1, 10);
  });

  it("returns VALUE error when first value <= 0", () => {
    expect(calculateCagr([0, 121])).toBe(VALUE_ERROR);
  });

  it("returns VALUE error when last value <= 0", () => {
    expect(calculateCagr([100, 0])).toBe(VALUE_ERROR);
  });

  it("returns VALUE error when period count <= 0", () => {
    expect(calculateCagr([100])).toBe(VALUE_ERROR);
  });
});
