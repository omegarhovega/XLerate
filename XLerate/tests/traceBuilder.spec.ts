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

/**
 * Adapts a per-cell-address neighbor map into the batched
 * `getAllNeighbors(cells[]) => cells[][]` shape. Tests retain readable
 * per-cell wiring while exercising the post-batching builder signature.
 */
function batched(
  map: Map<string, TraceCellInfo[]>
): (cells: TraceCellInfo[]) => Promise<TraceCellInfo[][]> {
  return async (cells) => cells.map((c) => map.get(c.address) ?? []);
}

describe("buildTrace", () => {
  it("returns just the root when no neighbors are produced", async () => {
    const root = cell("Sheet1!A1", 0, 0, 42);
    const result = await buildTrace({
      root,
      maxDepth: 5,
      getAllNeighbors: async () => [[]],
    });
    expect(result.rows).toEqual([
      { level: 0, address: "Sheet1!A1", value: "42", formula: "", parentAddress: null },
    ]);
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
      getAllNeighbors: batched(neighborMap),
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
      getAllNeighbors: batched(neighborMap),
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
      getAllNeighbors: batched(neighborMap),
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
      getAllNeighbors: batched(neighborMap),
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
      getAllNeighbors: async (cells) =>
        cells.map((c) => (c.address === "Sheet1!A1" ? siblings : [])),
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

    const result = await buildTrace({ root, maxDepth: 0, getAllNeighbors: async () => [[]] });
    expect(result.rows[0]).toEqual({
      level: 0,
      address: "Sheet1!A1",
      value: "42",
      formula: "=SUM(B1:B5)",
      parentAddress: null,
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
      getAllNeighbors: batched(neighborMap),
    });
    expect(result.rows.map((r) => r.address)).toEqual(["Sheet1!A1", "Sheet1!B1"]);
  });

  describe("onProgress (progressive loading)", () => {
    it("fires once with isFinal=true when maxDepth=0", async () => {
      const root = cell("Sheet1!A1", 0, 0, 42);
      const emissions: Array<{ level: number; isFinal: boolean; addresses: string[] }> = [];

      await buildTrace({
        root,
        maxDepth: 0,
        getAllNeighbors: async () => [],
        onProgress: async (p) => {
          emissions.push({ level: p.level, isFinal: p.isFinal, addresses: p.rows.map((r) => r.address) });
        },
      });

      expect(emissions).toEqual([{ level: 0, isFinal: true, addresses: ["Sheet1!A1"] }]);
    });

    it("emits once per BFS level with cumulative rows; marks the final emission", async () => {
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
      const emissions: Array<{ level: number; isFinal: boolean; addresses: string[] }> = [];

      await buildTrace({
        root: a,
        maxDepth: 3,
        getAllNeighbors: batched(neighborMap),
        onProgress: async (p) => {
          emissions.push({ level: p.level, isFinal: p.isFinal, addresses: p.rows.map((r) => r.address) });
        },
      });

      expect(emissions).toEqual([
        { level: 0, isFinal: false, addresses: ["Sheet1!A1"] },
        { level: 1, isFinal: false, addresses: ["Sheet1!A1", "Sheet1!B1"] },
        { level: 2, isFinal: false, addresses: ["Sheet1!A1", "Sheet1!B1", "Sheet1!C1"] },
        { level: 3, isFinal: true, addresses: ["Sheet1!A1", "Sheet1!B1", "Sheet1!C1", "Sheet1!D1"] },
      ]);
    });

    it("sets isFinal=true on the emission when BFS exits early because a level produced no new cells", async () => {
      // Root has no precedents. Loop body runs once (level 0), adds nothing,
      // sees nextLevelCells is empty — that emission is the last.
      const a = cell("Sheet1!A1", 0, 0);
      const emissions: Array<{ level: number; isFinal: boolean }> = [];

      await buildTrace({
        root: a,
        maxDepth: 5,
        getAllNeighbors: async () => [[]],
        onProgress: async (p) => {
          emissions.push({ level: p.level, isFinal: p.isFinal });
        },
      });

      // First emit: root with isFinal=false (because maxDepth>0).
      // Second emit: after loop iter 1, nextLevelCells empty → isFinal=true.
      expect(emissions).toEqual([
        { level: 0, isFinal: false },
        { level: 1, isFinal: true },
      ]);
    });

    it("sets isFinal=true on truncation emission", async () => {
      const root = cell("Sheet1!A1", 0, 0);
      const siblings = Array.from({ length: 10 }, (_, i) => cell(`Sheet1!B${i + 1}`, i, 1));
      const emissions: Array<{ level: number; isFinal: boolean; truncated: boolean }> = [];

      await buildTrace({
        root,
        maxDepth: 5,
        maxRows: 3,
        getAllNeighbors: async (cells) =>
          cells.map((c) => (c.address === "Sheet1!A1" ? siblings : [])),
        onProgress: async (p) => {
          emissions.push({ level: p.level, isFinal: p.isFinal, truncated: p.truncated });
        },
      });

      // Level 0 root, then level 1 hits maxRows=3 mid-expansion.
      expect(emissions).toEqual([
        { level: 0, isFinal: false, truncated: false },
        { level: 1, isFinal: true, truncated: true },
      ]);
    });

    it("awaits each onProgress call before starting the next level (back-pressure)", async () => {
      const a = cell("Sheet1!A1", 0, 0);
      const b = cell("Sheet1!B1", 0, 1);
      const c = cell("Sheet1!C1", 0, 2);
      const neighborMap = new Map<string, TraceCellInfo[]>([
        ["Sheet1!A1", [b]],
        ["Sheet1!B1", [c]],
        ["Sheet1!C1", []],
      ]);

      const order: string[] = [];
      await buildTrace({
        root: a,
        maxDepth: 3,
        getAllNeighbors: async (cells) => {
          order.push(`fetch:${cells.map((x) => x.address).join(",")}`);
          return batched(neighborMap)(cells);
        },
        onProgress: async (p) => {
          // Yield one microtask to simulate async work (e.g. messageChild).
          await Promise.resolve();
          order.push(`emit:level=${p.level}`);
        },
      });

      // Expect: emit(root) → fetch(root) → emit(level1) → fetch(level1) → emit(level2) → fetch(level2) → emit(level3)
      expect(order).toEqual([
        "emit:level=0",
        "fetch:Sheet1!A1",
        "emit:level=1",
        "fetch:Sheet1!B1",
        "emit:level=2",
        "fetch:Sheet1!C1",
        "emit:level=3",
      ]);
    });

    it("passing no onProgress is non-breaking (backward-compat)", async () => {
      const a = cell("Sheet1!A1", 0, 0);
      const b = cell("Sheet1!B1", 0, 1);

      const result = await buildTrace({
        root: a,
        maxDepth: 1,
        getAllNeighbors: batched(new Map([["Sheet1!A1", [b]]])),
      });

      expect(result.rows.map((r) => r.address)).toEqual(["Sheet1!A1", "Sheet1!B1"]);
      expect(result.truncated).toBe(false);
    });
  });

  describe("parentAddress (tree structure)", () => {
    it("root has parentAddress=null", async () => {
      const root = cell("Sheet1!A1", 0, 0);
      const result = await buildTrace({
        root,
        maxDepth: 0,
        getAllNeighbors: async () => [[]],
      });
      expect(result.rows[0].parentAddress).toBe(null);
    });

    it("each non-root row records the address of the cell whose expansion discovered it", async () => {
      // A -> [B, C]; B -> [D]; C -> [E, F].
      const a = cell("Sheet1!A1", 0, 0);
      const b = cell("Sheet1!B1", 0, 1);
      const c = cell("Sheet1!C1", 0, 2);
      const d = cell("Sheet1!D1", 0, 3);
      const e = cell("Sheet1!E1", 0, 4);
      const f = cell("Sheet1!F1", 0, 5);
      const neighborMap = new Map<string, TraceCellInfo[]>([
        ["Sheet1!A1", [b, c]],
        ["Sheet1!B1", [d]],
        ["Sheet1!C1", [e, f]],
      ]);
      const result = await buildTrace({
        root: a,
        maxDepth: 3,
        getAllNeighbors: batched(neighborMap),
      });
      const byAddr = new Map(result.rows.map((r) => [r.address, r.parentAddress]));
      expect(byAddr.get("Sheet1!A1")).toBe(null);
      expect(byAddr.get("Sheet1!B1")).toBe("Sheet1!A1");
      expect(byAddr.get("Sheet1!C1")).toBe("Sheet1!A1");
      expect(byAddr.get("Sheet1!D1")).toBe("Sheet1!B1");
      expect(byAddr.get("Sheet1!E1")).toBe("Sheet1!C1");
      expect(byAddr.get("Sheet1!F1")).toBe("Sheet1!C1");
    });

    it("on DAG (shared precedent), first discoverer becomes parent — visited set prevents re-attribution", async () => {
      // A -> [B, C]; both B and C have D as a precedent, but D appears
      // once, attributed to B (BFS visits B before C).
      const a = cell("Sheet1!A1", 0, 0);
      const b = cell("Sheet1!B1", 0, 1);
      const c = cell("Sheet1!C1", 0, 2);
      const d = cell("Sheet1!D1", 0, 3);
      const neighborMap = new Map<string, TraceCellInfo[]>([
        ["Sheet1!A1", [b, c]],
        ["Sheet1!B1", [d]],
        ["Sheet1!C1", [d]],
      ]);
      const result = await buildTrace({
        root: a,
        maxDepth: 5,
        getAllNeighbors: batched(neighborMap),
      });
      const dRows = result.rows.filter((r) => r.address === "Sheet1!D1");
      expect(dRows).toHaveLength(1);
      expect(dRows[0].parentAddress).toBe("Sheet1!B1");
    });
  });

  it("batches: all cells at one level are passed together to getAllNeighbors, then all at the next", async () => {
    // Root -> [B, C, D] at level 1; B -> [E]; C -> [F, G]; D -> [H] at level 2.
    // Expect exactly TWO calls to getAllNeighbors: once with [root], once with [B, C, D].
    const root = cell("Sheet1!A1", 0, 0);
    const b = cell("Sheet1!B1", 0, 1);
    const c = cell("Sheet1!C1", 0, 2);
    const d = cell("Sheet1!D1", 0, 3);
    const e = cell("Sheet1!E1", 0, 4);
    const f = cell("Sheet1!F1", 0, 5);
    const g = cell("Sheet1!G1", 0, 6);
    const h = cell("Sheet1!H1", 0, 7);

    const calls: string[][] = [];
    const neighborMap = new Map<string, TraceCellInfo[]>([
      ["Sheet1!A1", [b, c, d]],
      ["Sheet1!B1", [e]],
      ["Sheet1!C1", [f, g]],
      ["Sheet1!D1", [h]],
    ]);

    const result = await buildTrace({
      root,
      maxDepth: 2,
      getAllNeighbors: async (cells) => {
        calls.push(cells.map((x) => x.address));
        return cells.map((x) => neighborMap.get(x.address) ?? []);
      },
    });

    expect(calls).toEqual([["Sheet1!A1"], ["Sheet1!B1", "Sheet1!C1", "Sheet1!D1"]]);
    expect(result.rows.map((r) => r.address).sort()).toEqual(
      ["Sheet1!A1", "Sheet1!B1", "Sheet1!C1", "Sheet1!D1", "Sheet1!E1", "Sheet1!F1", "Sheet1!G1", "Sheet1!H1"].sort()
    );
  });
});
