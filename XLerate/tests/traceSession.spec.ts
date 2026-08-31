import { describe, expect, it } from "vitest";
import { TracePortFake } from "../src/adapters/tracePortFake";
import { TraceSessionService } from "../src/services/traceSession.service";
import { buildTraceLoadingPolicy } from "../src/core/tracePolicy";

function cell(id: string, formula = "=1") {
  return {
    kind: "cell" as const,
    key: id,
    worksheetName: "Model",
    rowIndex: Number(id.replace(/\D/g, "")) || 0,
    columnIndex: 0,
    address: `Model!${id.toUpperCase()}`,
    value: "1",
    formula,
  };
}

describe("TraceSessionService", () => {
  it("opens with root and direct children only", async () => {
    const port = new TracePortFake(cell("a1"));
    port.setNeighbors("a1", [cell("a2"), cell("a3")]);
    port.setNeighbors("a2", [cell("a4")]);
    const service = new TraceSessionService(port, "precedents", buildTraceLoadingPolicy());

    const result = await service.open();

    expect(result.children.map((node) => node.key)).toEqual(["a2", "a3"]);
    expect(port.directNeighborRequests).toEqual(["a1"]);
  });

  it("loads only the explicitly expanded branch and caches it", async () => {
    const port = new TracePortFake(cell("a1"));
    port.setNeighbors("a1", [cell("a2"), cell("a3")]);
    port.setNeighbors("a2", [cell("a4")]);
    port.setNeighbors("a3", [cell("a5")]);
    const service = new TraceSessionService(port, "precedents", buildTraceLoadingPolicy());
    const opened = await service.open();
    const a2 = opened.children.find((node) => node.key === "a2")!;

    const first = await service.expand(a2.id);
    const second = await service.expand(a2.id);

    expect(first.children.map((node) => node.key)).toEqual(["a4"]);
    expect(second.cached).toBe(true);
    expect(port.directNeighborRequests).toEqual(["a1", "a2"]);
  });

  it("treats precedent constants as leaves without a provider request", async () => {
    const port = new TracePortFake(cell("a1"));
    port.setNeighbors("a1", [cell("a2", "")]);
    const service = new TraceSessionService(port, "precedents", buildTraceLoadingPolicy());
    const opened = await service.open();

    const result = await service.expand(opened.children[0].id);

    expect(result.parentState).toBe("leaf");
    expect(port.directNeighborRequests).toEqual(["a1"]);
  });

  it("prefetches eligible formula nodes even when the frontier also contains constants", async () => {
    const port = new TracePortFake(cell("a1"));
    port.setNeighbors("a1", [cell("a2", ""), cell("a3", "=B1")]);
    const service = new TraceSessionService(port, "precedents", buildTraceLoadingPolicy());
    const opened = await service.open();

    expect(service.getPrefetchCandidates(opened.children).map((node) => node.key)).toEqual(["a3"]);
  });

  it("represents repeated graph nodes as non-expandable references", async () => {
    const port = new TracePortFake(cell("a1"));
    port.setNeighbors("a1", [cell("a2"), cell("a3")]);
    port.setNeighbors("a2", [cell("a4")]);
    port.setNeighbors("a3", [cell("a4")]);
    const service = new TraceSessionService(port, "precedents", buildTraceLoadingPolicy());
    const opened = await service.open();
    const a2 = opened.children.find((node) => node.key === "a2")!;
    const a3 = opened.children.find((node) => node.key === "a3")!;

    const first = await service.expand(a2.id);
    const second = await service.expand(a3.id);

    expect(first.children[0].kind).toBe("cell");
    expect(second.children[0]).toMatchObject({ kind: "reference", key: "a4" });
  });

  it("caches a bounded prefetch but discards a branch that exceeds its budget", async () => {
    const port = new TracePortFake(cell("a1"));
    port.setNeighbors("a1", [cell("a2"), cell("a3")]);
    port.setNeighbors("a2", [cell("a4"), cell("a5")]);
    port.setNeighbors(
      "a3",
      Array.from({ length: 60 }, (_, index) => cell(`b${index + 1}`))
    );
    const service = new TraceSessionService(port, "precedents", buildTraceLoadingPolicy());
    const opened = await service.open();
    const a2 = opened.children.find((node) => node.key === "a2")!;
    const a3 = opened.children.find((node) => node.key === "a3")!;

    expect(await service.prefetch(a2.id, 50)).toBe(2);
    expect((await service.expand(a2.id)).cached).toBe(true);
    expect(await service.prefetch(a3.id, 50)).toBe(50);
    expect((await service.expand(a3.id)).children).toHaveLength(60);
    expect(port.directNeighborRequests).toEqual(["a1", "a2", "a3", "a3"]);
  });

  it("loads summarized ranges in bounded pages", async () => {
    const port = new TracePortFake(cell("a1"));
    const range = {
      kind: "range" as const,
      key: "range-1",
      worksheetName: "Detail",
      rowIndex: 0,
      columnIndex: 0,
      rowCount: 250,
      columnCount: 1,
      cellCount: 250,
      address: "Detail!A1:A250",
    };
    port.setNeighbors("a1", [range]);
    port.setRangeCells(
      range.key,
      Array.from({ length: 250 }, (_, index) => ({
        kind: "cell" as const,
        key: `detail-${index}`,
        worksheetName: "Detail",
        rowIndex: index,
        columnIndex: 0,
        address: `Detail!A${index + 1}`,
        value: String(index),
        formula: "",
      }))
    );
    const service = new TraceSessionService(port, "precedents", buildTraceLoadingPolicy());
    const opened = await service.open();
    const rangeNode = opened.children[0];

    const first = await service.loadRangePage(rangeNode.id, 0);
    const third = await service.loadRangePage(rangeNode.id, 2);

    expect(first.children).toHaveLength(100);
    expect(first.hasMore).toBe(true);
    expect(third.children).toHaveLength(50);
    expect(third.hasMore).toBe(false);
    expect(port.rangePageRequests).toEqual([
      { key: "range-1", page: 0 },
      { key: "range-1", page: 2 },
    ]);
  });

  it("enforces depth and node budgets before requesting more work", async () => {
    const port = new TracePortFake(cell("a1"));
    port.setNeighbors("a1", [cell("a2"), cell("a3"), cell("a4")]);
    port.setNeighbors("a2", [cell("a5")]);
    const service = new TraceSessionService(
      port,
      "precedents",
      buildTraceLoadingPolicy({ safetyLimit: 3 }),
      1
    );

    const opened = await service.open();
    const depthLimited = await service.expand(opened.children[0].id);

    expect(opened.children).toHaveLength(2);
    expect(opened.truncated).toBe(true);
    expect(depthLimited.parentState).toBe("leaf");
    expect(port.directNeighborRequests).toEqual(["a1"]);
  });

  it("stops range paging when the global node budget is exhausted", async () => {
    const port = new TracePortFake(cell("a1"));
    const range = {
      kind: "range" as const,
      key: "range-limited",
      worksheetName: "Detail",
      rowIndex: 0,
      columnIndex: 0,
      rowCount: 10,
      columnCount: 1,
      cellCount: 10,
      address: "Detail!A1:A10",
    };
    port.setNeighbors("a1", [range]);
    port.setRangeCells(
      range.key,
      Array.from({ length: 10 }, (_, index) => ({
        kind: "cell" as const,
        key: `limited-${index}`,
        worksheetName: "Detail",
        rowIndex: index,
        columnIndex: 0,
        address: `Detail!A${index + 1}`,
        value: String(index),
        formula: "",
      }))
    );
    const service = new TraceSessionService(
      port,
      "precedents",
      buildTraceLoadingPolicy({ safetyLimit: 3 })
    );
    const opened = await service.open();
    const page = await service.loadRangePage(opened.children[0].id, 0);

    expect(page.children).toHaveLength(1);
    expect(page.truncated).toBe(true);
    expect(page.hasMore).toBe(false);
  });

  it("suppresses results after cancellation", async () => {
    const port = new TracePortFake(cell("a1"));
    port.setNeighbors("a1", [cell("a2")]);
    port.pauseDirectNeighborRequests();
    const service = new TraceSessionService(port, "precedents", buildTraceLoadingPolicy());
    const pending = service.open();
    service.cancel();
    port.resumeDirectNeighborRequests();

    await expect(pending).rejects.toThrow("Trace session was cancelled");
  });

  it("passes named-range and table navigation targets through to the port", async () => {
    const port = new TracePortFake(cell("a1"));
    port.setNeighbors("a1", []);
    const service = new TraceSessionService(port, "precedents", buildTraceLoadingPolicy());
    await service.open();

    await service.navigate({
      kind: "namedRange",
      name: "Revenue_Growth",
      formulaWorksheetName: "Model",
    });
    await service.navigate({
      kind: "table",
      tableName: "RevenueTable",
      section: "data",
      columnStart: "Revenue",
    });

    expect(port.selections).toEqual([
      { kind: "namedRange", name: "Revenue_Growth", formulaWorksheetName: "Model" },
      {
        kind: "table",
        tableName: "RevenueTable",
        section: "data",
        columnStart: "Revenue",
      },
    ]);
  });

  it("restores the original root selection when requested", async () => {
    const port = new TracePortFake(cell("root"));
    port.setNeighbors("root", []);
    const service = new TraceSessionService(port, "precedents", buildTraceLoadingPolicy());
    const opened = await service.open();

    await service.navigate({ address: "Other!B2" });
    await service.restoreRootSelection();

    expect(port.selections).toEqual([{ address: "Other!B2" }, { address: opened.root.address }]);
  });
});
