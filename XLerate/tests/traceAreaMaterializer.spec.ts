import { describe, expect, it } from "vitest";
import { materializeTraceArea, materializeTraceCells } from "../src/core/traceAreaMaterializer";
import { buildTraceLoadingPolicy } from "../src/core/tracePolicy";

describe("trace area materialization", () => {
  const policy = buildTraceLoadingPolicy();

  it("keeps a 10,000-cell area as one compact range node without matrices", () => {
    const result = materializeTraceArea(
      {
        worksheetName: "Transactions",
        address: "Transactions!A1:A10000",
        rowIndex: 0,
        columnIndex: 0,
        rowCount: 10_000,
        columnCount: 1,
      },
      policy,
      500
    );

    expect(result.nodes).toHaveLength(1);
    expect(result.nodes[0]).toMatchObject({
      kind: "range",
      address: "Transactions!A1:A10000",
      cellCount: 10_000,
    });
    expect(result.materializedCellCount).toBe(0);
  });

  it("materializes a small area from already-loaded matrices without cell proxies", () => {
    const result = materializeTraceArea(
      {
        worksheetName: "Revenue",
        address: "Revenue!B2:C3",
        rowIndex: 1,
        columnIndex: 1,
        rowCount: 2,
        columnCount: 2,
        values: [
          [10, 20],
          [30, 40],
        ],
        formulas: [
          [10, "=B2*2"],
          [30, 40],
        ],
      },
      policy,
      10
    );

    expect(result.nodes.map((node) => node.address)).toEqual([
      "Revenue!B2",
      "Revenue!C2",
      "Revenue!B3",
      "Revenue!C3",
    ]);
    expect(result.nodes[1]).toMatchObject({ kind: "cell", formula: "=B2*2", value: "20" });
    expect(result.materializedCellCount).toBe(4);
  });

  it("enforces the remaining-node budget before materialization", () => {
    const result = materializeTraceArea(
      {
        worksheetName: "Inputs",
        address: "Inputs!A1:A5",
        rowIndex: 0,
        columnIndex: 0,
        rowCount: 5,
        columnCount: 1,
        values: [[1], [2], [3], [4], [5]],
        formulas: [[1], [2], [3], [4], [5]],
      },
      policy,
      3
    );

    expect(result.nodes).toHaveLength(3);
    expect(result.truncated).toBe(true);
    expect(result.materializedCellCount).toBe(3);
  });

  it("materializes a page block as cells even when it exceeds the summary threshold", () => {
    const values = Array.from({ length: 100 }, (_, index) => [index + 1]);
    const formulas = values.map((row) => [...row]);
    const result = materializeTraceCells(
      {
        worksheetName: "Detail",
        address: "Detail!A1:A100",
        rowIndex: 0,
        columnIndex: 0,
        rowCount: 100,
        columnCount: 1,
        values,
        formulas,
      },
      100
    );

    expect(result.nodes).toHaveLength(100);
    expect(result.nodes[0].address).toBe("Detail!A1");
    expect(result.nodes[99].address).toBe("Detail!A100");
    expect(result.materializedCellCount).toBe(100);
    expect(result.truncated).toBe(false);
  });
});
