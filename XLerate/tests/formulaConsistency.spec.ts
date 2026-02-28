import { describe, expect, it } from "vitest";
import {
  analyzeHorizontalFormulaConsistency,
  collectAdjacentEqualFormulas,
  type FormulaConsistencyCell
} from "../src/core/formulaConsistency";

function row(cells: Array<string | null>): FormulaConsistencyCell[] {
  return cells.map((formulaR1C1) =>
    formulaR1C1 === null ? { isFormula: false } : { isFormula: true, formulaR1C1 }
  );
}

function asMarks(
  grid: ReturnType<typeof analyzeHorizontalFormulaConsistency>
): Array<Array<"C" | "I" | ".">> {
  return grid.map((r) =>
    r.map((mark) => {
      if (mark === "consistent") return "C";
      if (mark === "inconsistent") return "I";
      return ".";
    })
  );
}

describe("formula consistency baseline parity", () => {
  it("collects formulas that appear in adjacent equal pairs", () => {
    const formulas = collectAdjacentEqualFormulas([
      row(["=RC[-1]", "=RC[-1]", null, "=SUM(RC[-1]:RC[-2])"]),
      row([null, "=R1C1", "=R1C1", null])
    ]);

    expect(formulas).toEqual(new Set(["=RC[-1]", "=R1C1"]));
  });

  it("marks direct right-neighbor checks as consistent or inconsistent", () => {
    const marks = analyzeHorizontalFormulaConsistency([
      row(["=RC[-1]", "=RC[-1]", "=SUM(RC[-1]:RC[-2])", "=RC[-1]"])
    ]);

    expect(asMarks(marks)).toEqual([["C", "I", "I", "C"]]);
  });

  it("marks formula cells with non-formula right side as consistent only when formula is in a known pair", () => {
    const marks = analyzeHorizontalFormulaConsistency([
      row(["=R1C1", "=R1C1", null, "=R10C10", null]),
      row([null, "=R1C1", null, "=R20C20", null])
    ]);

    expect(asMarks(marks)).toEqual([
      ["C", "C", ".", ".", "."],
      [".", "C", ".", ".", "."]
    ]);
  });

  it("applies last-column rule using adjacent-pair membership", () => {
    const marks = analyzeHorizontalFormulaConsistency([
      row([null, "=R2C2", "=R2C2"]),
      row([null, "=R5C5", "=R6C6"])
    ]);

    expect(asMarks(marks)).toEqual([
      [".", "C", "C"],
      [".", "I", "."]
    ]);
  });

  it("never marks non-formula cells", () => {
    const marks = analyzeHorizontalFormulaConsistency([
      row([null, null, null]),
      row([null, "=RC[-1]", null])
    ]);

    expect(asMarks(marks)).toEqual([
      [".", ".", "."],
      [".", ".", "."]
    ]);
  });
});
