import { describe, expect, it } from "vitest";
import { materializeTraceArea } from "../src/core/traceAreaMaterializer";
import { buildTraceLoadingPolicy } from "../src/core/tracePolicy";
import { TracePortFake } from "../src/adapters/tracePortFake";
import { TraceSessionService } from "../src/services/traceSession.service";

describe("trace scalability invariants", () => {
  it("does constant materialization work for a very large rectangular precedent", () => {
    const result = materializeTraceArea(
      {
        worksheetName: "Raw Data",
        address: "'Raw Data'!A1:A10000",
        rowIndex: 0,
        columnIndex: 0,
        rowCount: 10_000,
        columnCount: 1,
      },
      buildTraceLoadingPolicy(),
      500
    );

    expect(result.nodes).toHaveLength(1);
    expect(result.materializedCellCount).toBe(0);
  });

  it("does not request unrelated worksheet branches", async () => {
    const root = {
      kind: "cell" as const,
      key: "summary",
      worksheetName: "Summary",
      rowIndex: 0,
      columnIndex: 0,
      address: "Summary!A1",
      value: "1",
      formula: "=Revenue!A1",
    };
    const port = new TracePortFake(root);
    port.setNeighbors("summary", [
      {
        kind: "cell",
        key: "revenue",
        worksheetName: "Revenue",
        rowIndex: 0,
        columnIndex: 0,
        address: "Revenue!A1",
        value: "1",
        formula: "=Drivers!A1",
      },
    ]);
    for (let index = 0; index < 99; index += 1) {
      port.setNeighbors(`unrelated-${index}`, []);
    }
    const service = new TraceSessionService(port, "precedents", buildTraceLoadingPolicy());

    await service.open();
    expect(port.directNeighborRequests).toEqual(["summary"]);
  });

  it("does not query 100 terminal precedent constants", async () => {
    const root = {
      kind: "cell" as const,
      key: "summary",
      worksheetName: "Summary",
      rowIndex: 0,
      columnIndex: 0,
      address: "Summary!A1",
      value: "100",
      formula: "=SUM(Inputs!A1:A100)",
    };
    const port = new TracePortFake(root);
    port.setNeighbors(
      "summary",
      Array.from({ length: 100 }, (_, index) => ({
        kind: "cell" as const,
        key: `input-${index}`,
        worksheetName: "Inputs",
        rowIndex: index,
        columnIndex: 0,
        address: `Inputs!A${index + 1}`,
        value: String(index),
        formula: "",
      }))
    );
    const service = new TraceSessionService(port, "precedents", buildTraceLoadingPolicy());

    const opened = await service.open();
    expect(opened.children).toHaveLength(100);
    expect(service.getPrefetchCandidates(opened.children)).toEqual([]);
    expect(port.directNeighborRequests).toEqual(["summary"]);
  });
});
