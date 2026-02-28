import { describe, expect, it } from "vitest";
import { switchSignCell } from "../src/core/switchSign";

describe("switchSignCell baseline parity", () => {
  it("negates numeric constants", () => {
    expect(switchSignCell({ isFormula: false, value: 10 })).toEqual({ isFormula: false, value: -10 });
    expect(switchSignCell({ isFormula: false, value: -42 })).toEqual({ isFormula: false, value: 42 });
    expect(switchSignCell({ isFormula: false, value: 0 })).toEqual({ isFormula: false, value: 0 });
  });

  it("wraps normal formulas", () => {
    expect(switchSignCell({ isFormula: true, formula: "=A1+B1" })).toEqual({
      isFormula: true,
      formula: "=-(A1+B1)"
    });
  });

  it("handles array formulas", () => {
    expect(
      switchSignCell({
        isFormula: true,
        isArrayFormula: true,
        formula: "{=A1:A3*2}"
      })
    ).toEqual({ isFormula: true, isArrayFormula: true, formula: "{=-(A1:A3*2)}" });
  });

  it("leaves non-numeric constants unchanged", () => {
    expect(switchSignCell({ isFormula: false, value: "hello" })).toEqual({ isFormula: false, value: "hello" });
    expect(switchSignCell({ isFormula: false, value: null })).toEqual({ isFormula: false, value: null });
  });
});
