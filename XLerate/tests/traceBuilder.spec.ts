import { describe, expect, it } from "vitest";
import { buildTrace, type TraceCellInfo } from "../src/core/traceBuilder";

function cell(
  address: string,
  rowIndex: number,
  columnIndex: number,
  value: unknown = null,
  formula: unknown = null,
  worksheetName = "Sheet1"
): TraceCellInfo {
  return { worksheetName, rowIndex, columnIndex, address, value, formula };
}

describe("buildTrace", () => {
  it("returns just the root when no neighbors are produced", async () => {
    const root = cell("Sheet1!A1", 0, 0, 42);
    const result = await buildTrace({
      root,
      maxDepth: 5,
      getNeighbors: async () => [],
    });
    expect(result.rows).toEqual([{ level: 0, address: "Sheet1!A1", value: "42", formula: "" }]);
    expect(result.truncated).toBe(false);
  });

  it("walks a linear chain with correct BFS levels", async () => {
    const a = cell("Sheet1!A1", 0, 0);
    const b = cell("Sheet1!B1", 0, 1);
    const c = cell("Sheet1!C1", 0, 2);
    const d = cell("Sheet1!D1", 0, 3);

    const neighborMap = new Map<string, TraceCellInfo[]>([
      ["Sheet1!A1", [b]],
      ["Sheet1!B1", [c]],
      ["Sheet1!C1", [d]],
      ["Sheet1!D1", []],
    ]);

    const result = await buildTrace({
      root: a,
      maxDepth: 10,
      getNeighbors: async (c) => neighborMap.get(c.address) ?? [],
    });

    expect(result.rows.map((r) => ({ level: r.level, address: r.address }))).toEqual([
      { level: 0, address: "Sheet1!A1" },
      { level: 1, address: "Sheet1!B1" },
      { level: 2, address: "Sheet1!C1" },
      { level: 3, address: "Sheet1!D1" },
    ]);
    expect(result.truncated).toBe(false);
  });

  it("respects maxDepth and does not expand past the cap", async () => {
    // A -> B -> C. Depth 1 should include A and B but not C.
    const a = cell("Sheet1!A1", 0, 0);
    const b = cell("Sheet1!B1", 0, 1);
    const c = cell("Sheet1!C1", 0, 2);
    const neighborMap = new Map<string, TraceCellInfo[]>([
      ["Sheet1!A1", [b]],
      ["Sheet1!B1", [c]],
      ["Sheet1!C1", []],
    ]);

    const result = await buildTrace({
      root: a,
      maxDepth: 1,
      getNeighbors: async (c) => neighborMap.get(c.address) ?? [],
    });

    expect(result.rows.map((r) => r.address)).toEqual(["Sheet1!A1", "Sheet1!B1"]);
  });

  it("avoids cycles via the visited set", async () => {
    // A <-> B: A references B and B references A. Each should appear once.
    const a = cell("Sheet1!A1", 0, 0);
    const b = cell("Sheet1!B1", 0, 1);
    const neighborMap = new Map<string, TraceCellInfo[]>([
      ["Sheet1!A1", [b]],
      ["Sheet1!B1", [a]],
    ]);

    const result = await buildTrace({
      root: a,
      maxDepth: 10,
      getNeighbors: async (c) => neighborMap.get(c.address) ?? [],
    });

    expect(result.rows.map((r) => r.address)).toEqual(["Sheet1!A1", "Sheet1!B1"]);
  });

  it("keys visited by (sheet, row, col), not by address — two cells with the same displayed address on different sheets are distinct", async () => {
    const a = cell("Sheet1!A1", 0, 0);
    const bSheet1 = cell("Sheet1!B1", 0, 1);
    const bSheet2 = cell("Sheet2!B1", 0, 1, null, null, "Sheet2");
    const neighborMap = new Map<string, TraceCellInfo[]>([
      ["Sheet1!A1", [bSheet1, bSheet2]],
    ]);

    const result = await buildTrace({
      root: a,
      maxDepth: 10,
      getNeighbors: async (c) => neighborMap.get(c.address) ?? [],
    });

    expect(result.rows.map((r) => r.address)).toEqual([
      "Sheet1!A1",
      "Sheet1!B1",
      "Sheet2!B1",
    ]);
  });

  it("truncates when maxRows is reached and marks the result", async () => {
    const root = cell("Sheet1!A1", 0, 0);
    const siblings = Array.from({ length: 10 }, (_, i) => cell(`Sheet1!B${i + 1}`, i, 1));

    const result = await buildTrace({
      root,
      maxDepth: 10,
      maxRows: 3, // root + 2 siblings
      getNeighbors: async (c) => (c.address === "Sheet1!A1" ? siblings : []),
    });

    expect(result.rows.map((r) => r.address)).toEqual(["Sheet1!A1", "Sheet1!B1", "Sheet1!B2"]);
    expect(result.truncated).toBe(true);
  });

  it("uses scalarFromMatrix + formatters — Office.js matrix values and formulas round-trip", async () => {
    const root: TraceCellInfo = {
      worksheetName: "Sheet1",
      rowIndex: 0,
      columnIndex: 0,
      address: "Sheet1!A1",
      value: [[42]],
      formula: [["=SUM(B1:B5)"]],
    };

    const result = await buildTrace({ root, maxDepth: 0, getNeighbors: async () => [] });
    expect(result.rows[0]).toEqual({
      level: 0,
      address: "Sheet1!A1",
      value: "42",
      formula: "=SUM(B1:B5)",
    });
  });

  it("treats level >= maxDepth as a leaf (neighbors still discovered at exactly maxDepth are recorded, but not further expanded)", async () => {
    // Root A (level 0) -> B (level 1) -> C (level 2). maxDepth = 1 means:
    // level 0 expands (produces B at level 1), level 1 does NOT expand (so C is never seen).
    const a = cell("Sheet1!A1", 0, 0);
    const b = cell("Sheet1!B1", 0, 1);
    const c = cell("Sheet1!C1", 0, 2);
    const neighborMap = new Map<string, TraceCellInfo[]>([
      ["Sheet1!A1", [b]],
      ["Sheet1!B1", [c]],
    ]);
    const result = await buildTrace({
      root: a,
      maxDepth: 1,
      getNeighbors: async (c) => neighborMap.get(c.address) ?? [],
    });
    expect(result.rows.map((r) => r.address)).toEqual(["Sheet1!A1", "Sheet1!B1"]);
  });
});
