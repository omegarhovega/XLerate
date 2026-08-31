import { describe, expect, it } from "vitest";
import { parseDialogToParentMessage, parseParentToDialogMessage } from "../src/core/traceProtocol";

describe("trace dialog protocol", () => {
  it("parses session-scoped expansion requests", () => {
    expect(
      parseDialogToParentMessage(
        JSON.stringify({ action: "expand", sessionId: "s1", nodeId: "n1", requestId: "r1" })
      )
    ).toEqual({ action: "expand", sessionId: "s1", nodeId: "n1", requestId: "r1" });
  });

  it("rejects malformed or sessionless messages", () => {
    expect(parseDialogToParentMessage("not json")).toBeNull();
    expect(
      parseDialogToParentMessage(JSON.stringify({ action: "expand", nodeId: "n1" }))
    ).toBeNull();
    expect(
      parseDialogToParentMessage(JSON.stringify({ action: "navigate", address: "" }))
    ).toBeNull();
  });

  it("parses named-range and table navigation targets", () => {
    expect(
      parseDialogToParentMessage(
        JSON.stringify({
          action: "navigate",
          sessionId: "s1",
          target: {
            kind: "namedRange",
            name: "Revenue",
            formulaWorksheetName: "Summary",
          },
        })
      )
    ).toMatchObject({
      action: "navigate",
      target: {
        kind: "namedRange",
        name: "Revenue",
        formulaWorksheetName: "Summary",
      },
    });
    expect(
      parseDialogToParentMessage(
        JSON.stringify({
          action: "navigate",
          sessionId: "s1",
          target: {
            kind: "table",
            tableName: "RevenueTable",
            section: "data",
            columnStart: "Revenue",
          },
        })
      )
    ).toMatchObject({ action: "navigate", target: { kind: "table", tableName: "RevenueTable" } });
  });

  it("preserves the Escape request to restore the root selection", () => {
    expect(
      parseDialogToParentMessage(
        JSON.stringify({ action: "close", sessionId: "s1", restoreRoot: true })
      )
    ).toEqual({ action: "close", sessionId: "s1", restoreRoot: true });
    expect(
      parseDialogToParentMessage(
        JSON.stringify({ action: "close", sessionId: "s1", restoreRoot: false })
      )
    ).toEqual({ action: "close", sessionId: "s1" });
  });

  it("parses child deltas without accepting arbitrary rows", () => {
    const raw = JSON.stringify({
      action: "childrenLoaded",
      sessionId: "s1",
      requestId: "r1",
      parentId: "root",
      parentState: "loaded",
      children: [],
      truncated: false,
      cached: false,
      totalNodeCount: 1,
    });
    expect(parseParentToDialogMessage(raw)).toMatchObject({
      action: "childrenLoaded",
      parentId: "root",
    });
    expect(
      parseParentToDialogMessage(JSON.stringify({ action: "childrenLoaded", children: "bad" }))
    ).toBeNull();
    expect(
      parseParentToDialogMessage(
        JSON.stringify({
          action: "childrenLoaded",
          sessionId: "s1",
          requestId: "r1",
          parentId: "root",
          parentState: "loaded",
          children: [],
          truncated: false,
          cached: false,
        })
      )
    ).toBeNull();
  });
});
