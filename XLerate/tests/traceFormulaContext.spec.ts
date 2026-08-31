import { describe, expect, it } from "vitest";
import { tokenizeFormulaReferences } from "../src/core/formulaReferences";
import {
  buildTraceFormulaContext,
  formulaTargetMatchesNode,
  getHighlightedFormulaSegmentIndexes,
} from "../src/core/traceFormulaContext";
import type { TraceDialogState } from "../src/core/traceDialogState";
import type { TraceCellNode, TraceRangeNode } from "../src/core/traceGraph";

function cell(
  id: string,
  parentId: string | null,
  worksheetName: string,
  rowIndex: number,
  columnIndex: number,
  formula: string
): TraceCellNode {
  return {
    kind: "cell",
    id,
    key: id,
    parentId,
    level: parentId ? 1 : 0,
    worksheetName,
    rowIndex,
    columnIndex,
    address: `${worksheetName}!${String.fromCharCode(65 + columnIndex)}${rowIndex + 1}`,
    value: "1",
    formula,
    loadState: "loaded",
  };
}

function state(nodes: TraceCellNode[], children: Array<[string, string[]]>): TraceDialogState {
  return {
    sessionId: "s1",
    rootId: nodes[0]?.id ?? null,
    nodesById: new Map(nodes.map((node) => [node.id, node])),
    childIdsByParent: new Map(children),
    expandedIds: new Set(children.map(([parentId]) => parentId)),
    rangePages: new Map(),
    totalNodeCount: nodes.length,
    truncated: false,
  };
}

describe("trace formula context", () => {
  it("pins the parent formula and builds its breadcrumb", () => {
    const root = cell("root", null, "Summary", 0, 0, "=Inputs!B1+Inputs!C1");
    const first = cell("first", "root", "Inputs", 0, 1, "");
    const second = cell("second", "root", "Inputs", 0, 2, "");
    const context = buildTraceFormulaContext(
      state([root, first, second], [["root", ["first", "second"]]]),
      "first",
      "precedents"
    );

    expect(context?.formulaNode.id).toBe("root");
    expect(context?.relationNode?.id).toBe("first");
    expect(context?.breadcrumb.map((node) => node.id)).toEqual(["root", "first"]);
  });

  it("shows a dependent's formula and highlights its parent relationship", () => {
    const root = cell("root", null, "Inputs", 0, 0, "");
    const dependent = cell("dependent", "root", "Summary", 0, 3, "=Inputs!A1*100");
    const context = buildTraceFormulaContext(
      state([root, dependent], [["root", ["dependent"]]]),
      "dependent",
      "dependents"
    );

    expect(context?.formulaNode.id).toBe("dependent");
    expect(context?.relationNode?.id).toBe("root");
  });

  it("matches selected cells and summarized ranges to A1 formula ranges", () => {
    const selected = cell("selected", "root", "Detail", 9, 0, "");
    const summarized: TraceRangeNode = {
      kind: "range",
      id: "range",
      key: "range",
      parentId: "root",
      level: 1,
      worksheetName: "Detail",
      rowIndex: 0,
      columnIndex: 0,
      rowCount: 10_000,
      columnCount: 1,
      cellCount: 10_000,
      address: "Detail!A1:A10000",
      loadState: "loaded",
    };

    expect(formulaTargetMatchesNode({ address: "Detail!A1:A10000" }, selected)).toBe(true);
    expect(formulaTargetMatchesNode({ address: "Detail!A1:A10000" }, summarized)).toBe(true);
    expect(formulaTargetMatchesNode({ address: "Other!A1:A10000" }, selected)).toBe(false);
  });

  it("highlights exact A1 references and a single unresolved named reference", () => {
    const selected = cell("selected", "root", "Inputs", 0, 1, "");
    const a1Segments = tokenizeFormulaReferences("=Inputs!B1+Inputs!C1", "Summary");
    expect([...getHighlightedFormulaSegmentIndexes(a1Segments, selected)]).toEqual([1]);

    const namedSegments = tokenizeFormulaReferences("=Annual_volume_growth*100", "Summary");
    expect([...getHighlightedFormulaSegmentIndexes(namedSegments, selected)]).toEqual([1]);
  });
});
