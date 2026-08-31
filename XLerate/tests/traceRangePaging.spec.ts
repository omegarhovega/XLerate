import { describe, expect, it } from "vitest";
import { planTraceRangePageBlocks } from "../src/core/traceRangePaging";

describe("trace range page block planning", () => {
  it("loads a 100-cell single-column page as one rectangular block", () => {
    expect(planTraceRangePageBlocks(10_000, 1, 0, 100)).toEqual([
      { rowOffset: 0, columnOffset: 0, rowCount: 100, columnCount: 1, cellCount: 100 },
    ]);
  });

  it("loads complete rows as one block", () => {
    expect(planTraceRangePageBlocks(100, 10, 0, 100)).toEqual([
      { rowOffset: 0, columnOffset: 0, rowCount: 10, columnCount: 10, cellCount: 100 },
    ]);
  });

  it("uses at most three blocks for a page crossing partial rows", () => {
    expect(planTraceRangePageBlocks(100, 12, 100, 100)).toEqual([
      { rowOffset: 8, columnOffset: 4, rowCount: 1, columnCount: 8, cellCount: 8 },
      { rowOffset: 9, columnOffset: 0, rowCount: 7, columnCount: 12, cellCount: 84 },
      { rowOffset: 16, columnOffset: 0, rowCount: 1, columnCount: 8, cellCount: 8 },
    ]);
  });

  it("clips a requested page to the source range", () => {
    expect(planTraceRangePageBlocks(3, 4, 10, 100)).toEqual([
      { rowOffset: 2, columnOffset: 2, rowCount: 1, columnCount: 2, cellCount: 2 },
    ]);
    expect(planTraceRangePageBlocks(3, 4, 12, 100)).toEqual([]);
  });

  it("covers exactly the requested row-major cells without exceeding three blocks", () => {
    for (const columnCount of [1, 2, 7, 12, 101, 250]) {
      const rowCount = 30;
      const start = Math.min(columnCount * rowCount - 1, 37);
      const requested = Math.min(100, columnCount * rowCount - start);
      const blocks = planTraceRangePageBlocks(rowCount, columnCount, start, requested);
      expect(blocks.length).toBeLessThanOrEqual(3);
      expect(blocks.reduce((sum, block) => sum + block.cellCount, 0)).toBe(requested);
    }
  });
});
