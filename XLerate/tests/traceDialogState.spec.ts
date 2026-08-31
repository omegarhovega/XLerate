import { describe, expect, it } from "vitest";
import {
  applyTraceDialogMessage,
  buildInitialTraceDialogState,
  computeVisibleTraceNodes,
  toggleTraceNode,
} from "../src/core/traceDialogState";
import type { ParentToDialogMessage } from "../src/core/traceProtocol";

const rootLoaded: ParentToDialogMessage = {
  action: "rootLoaded",
  sessionId: "s1",
  direction: "precedents",
  root: {
    kind: "cell",
    id: "root",
    key: "root",
    parentId: null,
    level: 0,
    worksheetName: "Summary",
    rowIndex: 0,
    columnIndex: 0,
    address: "Summary!A1",
    value: "1",
    formula: "=B1",
    loadState: "loaded",
  },
  children: [
    {
      kind: "cell",
      id: "child",
      key: "child",
      parentId: "root",
      level: 1,
      worksheetName: "Inputs",
      rowIndex: 0,
      columnIndex: 1,
      address: "Inputs!B1",
      value: "1",
      formula: "=C1",
      loadState: "unloaded",
    },
  ],
  truncated: false,
  totalNodeCount: 2,
};

describe("trace dialog state", () => {
  it("shows root and direct children after the initial delta", () => {
    const state = applyTraceDialogMessage(buildInitialTraceDialogState("s1"), rootLoaded);
    expect(computeVisibleTraceNodes(state).map((node) => node.id)).toEqual(["root", "child"]);
  });

  it("inserts loaded children under only the requested parent", () => {
    let state = applyTraceDialogMessage(buildInitialTraceDialogState("s1"), rootLoaded);
    state = applyTraceDialogMessage(state, {
      action: "childrenLoaded",
      sessionId: "s1",
      requestId: "r1",
      parentId: "child",
      parentState: "loaded",
      children: [
        {
          kind: "cell",
          id: "grandchild",
          key: "grandchild",
          parentId: "child",
          level: 2,
          worksheetName: "Drivers",
          rowIndex: 0,
          columnIndex: 2,
          address: "Drivers!C1",
          value: "1",
          formula: "",
          loadState: "leaf",
        },
      ],
      truncated: false,
      cached: false,
      totalNodeCount: 3,
    });
    state = toggleTraceNode(state, "child", true);
    expect(computeVisibleTraceNodes(state).map((node) => node.id)).toEqual([
      "root",
      "child",
      "grandchild",
    ]);
  });

  it("ignores stale messages from another session", () => {
    const state = buildInitialTraceDialogState("current");
    expect(applyTraceDialogMessage(state, rootLoaded)).toBe(state);
  });
});
