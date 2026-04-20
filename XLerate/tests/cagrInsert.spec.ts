import { describe, expect, it } from "vitest";
import {
  buildCagrInsertFormula,
  CAGR_RESULT_NUMBER_FORMAT,
  findContiguousLeftCagrSeries,
  toA1Address,
} from "../src/core/cagrInsert";

describe("cagr insert helpers", () => {
  it("finds the contiguous numeric run immediately left of the destination", () => {
    const series = findContiguousLeftCagrSeries([
      { row: 4, col: 1, value: 100 },
      { row: 4, col: 2, value: 110 },
      { row: 4, col: 3, value: 121 },
    ]);

    expect(series).toEqual({
      start: { row: 4, col: 1 },
      end: { row: 4, col: 3 },
      values: [100, 110, 121],
      periodCount: 2,
    });
  });

  it("stops at the first empty or non-numeric boundary", () => {
    const blockedByText = findContiguousLeftCagrSeries([
      { row: 0, col: 0, value: 100 },
      { row: 0, col: 1, value: 110 },
      { row: 0, col: 2, value: "n/a" },
      { row: 0, col: 3, value: 121 },
    ]);

    expect(blockedByText).toBeNull();
  });

  it("requires at least two numeric cells", () => {
    expect(
      findContiguousLeftCagrSeries([
        { row: 0, col: 2, value: 121 },
      ])
    ).toBeNull();
  });

  it("converts zero-based coordinates to A1 addresses", () => {
    expect(toA1Address(0, 0)).toBe("A1");
    expect(toA1Address(4, 3)).toBe("D5");
    expect(toA1Address(1, 27)).toBe("AB2");
  });

  it("builds the worksheet formula and result format", () => {
    const series = findContiguousLeftCagrSeries([
      { row: 4, col: 1, value: 100 },
      { row: 4, col: 2, value: 110 },
      { row: 4, col: 3, value: 121 },
    ]);

    expect(series).not.toBeNull();
    expect(buildCagrInsertFormula(series!)).toBe("=POWER(D5/B5,1/2)-1");
    expect(CAGR_RESULT_NUMBER_FORMAT).toBe("0.0%");
  });
});
