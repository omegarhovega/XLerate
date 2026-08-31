import { describe, expect, it } from "vitest";
import {
  buildTraceLoadingPolicy,
  getRangePageBounds,
  shouldPrefetchTraceNodes,
  shouldSummarizeTraceRange,
} from "../src/core/tracePolicy";
import type { TraceNode } from "../src/core/traceGraph";

describe("trace loading policy", () => {
  it("uses the approved default limits", () => {
    expect(buildTraceLoadingPolicy()).toEqual({
      rangeSummaryThreshold: 50,
      rangePageSize: 100,
      visibleNodeLimit: 500,
      hardNodeLimit: 500,
      prefetchMaxCandidates: 20,
      prefetchMaxCells: 50,
    });
  });

  it("uses the configured safety limit without exceeding the product hard cap", () => {
    expect(buildTraceLoadingPolicy({ safetyLimit: 1250 }).hardNodeLimit).toBe(1250);
    expect(buildTraceLoadingPolicy({ safetyLimit: 99999 }).hardNodeLimit).toBe(5000);
    expect(buildTraceLoadingPolicy({ safetyLimit: 20 }).visibleNodeLimit).toBe(20);
  });

  it("summarizes a rectangular area only above the threshold", () => {
    const policy = buildTraceLoadingPolicy();
    expect(shouldSummarizeTraceRange(50, policy)).toBe(false);
    expect(shouldSummarizeTraceRange(51, policy)).toBe(true);
    expect(shouldSummarizeTraceRange(10_000, policy)).toBe(true);
  });

  it("computes bounded row-major range pages", () => {
    const policy = buildTraceLoadingPolicy();
    expect(getRangePageBounds(250, 0, policy)).toEqual({ start: 0, count: 100, hasMore: true });
    expect(getRangePageBounds(250, 1, policy)).toEqual({ start: 100, count: 100, hasMore: true });
    expect(getRangePageBounds(250, 2, policy)).toEqual({ start: 200, count: 50, hasMore: false });
    expect(getRangePageBounds(250, 3, policy)).toEqual({ start: 250, count: 0, hasMore: false });
  });

  it("prefetches only a small frontier of unloaded cell nodes", () => {
    const policy = buildTraceLoadingPolicy();
    const cells = Array.from({ length: 20 }, (_, index) => ({
      kind: "cell" as const,
      id: `cell-${index}`,
      key: `cell-${index}`,
      parentId: "root",
      level: 1,
      worksheetName: "Sheet1",
      rowIndex: index,
      columnIndex: 0,
      address: `Sheet1!A${index + 1}`,
      value: "",
      formula: "=1",
      loadState: "unloaded" as const,
    })) satisfies TraceNode[];

    expect(shouldPrefetchTraceNodes(cells, policy)).toBe(true);
    expect(shouldPrefetchTraceNodes([...cells, cells[0]], policy)).toBe(false);
    expect(
      shouldPrefetchTraceNodes(
        [
          {
            kind: "range",
            id: "range",
            key: "range",
            parentId: "root",
            level: 1,
            worksheetName: "Sheet1",
            rowIndex: 0,
            columnIndex: 0,
            rowCount: 100,
            columnCount: 1,
            cellCount: 100,
            address: "Sheet1!A1:A100",
            loadState: "unloaded",
          },
        ],
        policy
      )
    ).toBe(false);
  });
});
