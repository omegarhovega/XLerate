import { describe, expect, it } from "vitest";
import { wrapFormulaWithError, wrapSelectionFormulas } from "../src/core/errorWrap";

describe("error wrap baseline parity", () => {
  it("wraps one formula with default NA()", () => {
    expect(wrapFormulaWithError("=A1/B1")).toBe("=IFERROR(A1/B1, NA())");
  });

  it("wraps one formula with custom fallback", () => {
    expect(wrapFormulaWithError("=SUM(A1:A3)", "0")).toBe("=IFERROR(SUM(A1:A3), 0)");
  });

  it("allows nested wrapping", () => {
    expect(wrapFormulaWithError("=IFERROR(A1/B1,0)", "NA()")).toBe("=IFERROR(IFERROR(A1/B1,0), NA())");
  });

  it("wraps only formula items in a selection", () => {
    expect(
      wrapSelectionFormulas([
        { isFormula: true, formula: "=A1/B1" },
        { isFormula: false, value: 100 },
        { isFormula: false, value: "x" },
        { isFormula: false, value: "" }
      ])
    ).toEqual([
      { isFormula: true, formula: "=IFERROR(A1/B1, NA())" },
      { isFormula: false, value: 100 },
      { isFormula: false, value: "x" },
      { isFormula: false, value: "" }
    ]);
  });
});
