import type { TraceDirection } from "./traceUtils";
import type { TraceNode, TraceNodeLoadState, TraceTableSection, TraceTarget } from "./traceGraph";

export type DialogToParentMessage =
  | { action: "ready"; sessionId: string }
  | { action: "expand"; sessionId: string; nodeId: string; requestId: string }
  | {
      action: "loadRangePage";
      sessionId: string;
      nodeId: string;
      requestId: string;
      page: number;
    }
  | { action: "navigate"; sessionId: string; target: TraceTarget }
  | { action: "close"; sessionId: string; restoreRoot?: boolean };

export type ParentToDialogMessage =
  | {
      action: "rootLoaded";
      sessionId: string;
      direction: TraceDirection;
      root: TraceNode;
      children: TraceNode[];
      truncated: boolean;
      totalNodeCount: number;
    }
  | {
      action: "childrenLoaded";
      sessionId: string;
      requestId: string;
      parentId: string;
      parentState: TraceNodeLoadState;
      children: TraceNode[];
      truncated: boolean;
      cached: boolean;
      totalNodeCount: number;
    }
  | {
      action: "rangePageLoaded";
      sessionId: string;
      requestId: string;
      rangeId: string;
      parentState: TraceNodeLoadState;
      page: number;
      children: TraceNode[];
      hasMore: boolean;
      truncated: boolean;
      totalNodeCount: number;
    }
  | {
      action: "nodeFailed";
      sessionId: string;
      requestId: string;
      nodeId: string;
      message: string;
    };

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.length > 0;
}

function parseRaw(raw: unknown): Record<string, unknown> | null {
  if (typeof raw !== "string") return null;
  try {
    const parsed = JSON.parse(raw);
    return isRecord(parsed) ? parsed : null;
  } catch {
    return null;
  }
}

function isTraceNode(value: unknown): value is TraceNode {
  if (!isRecord(value)) return false;
  return (
    (value.kind === "cell" || value.kind === "range" || value.kind === "reference") &&
    isNonEmptyString(value.id) &&
    isNonEmptyString(value.key) &&
    isNonEmptyString(value.address) &&
    typeof value.level === "number" &&
    (value.parentId === null || typeof value.parentId === "string") &&
    isNodeState(value.loadState)
  );
}

function isNodeState(value: unknown): value is TraceNodeLoadState {
  return (
    value === "unloaded" ||
    value === "loading" ||
    value === "loaded" ||
    value === "leaf" ||
    value === "error"
  );
}

function isTraceTableSection(value: unknown): value is TraceTableSection {
  return value === "all" || value === "data" || value === "headers" || value === "totals";
}

function parseTraceTarget(value: unknown): TraceTarget | null {
  if (!isRecord(value)) return null;
  if (isNonEmptyString(value.address)) {
    return { address: value.address };
  }
  if (
    value.kind === "namedRange" &&
    isNonEmptyString(value.name) &&
    isNonEmptyString(value.formulaWorksheetName) &&
    (value.worksheetName === undefined || isNonEmptyString(value.worksheetName))
  ) {
    return {
      kind: "namedRange",
      name: value.name,
      formulaWorksheetName: value.formulaWorksheetName,
      ...(isNonEmptyString(value.worksheetName) ? { worksheetName: value.worksheetName } : {}),
    };
  }
  if (
    value.kind === "table" &&
    isNonEmptyString(value.tableName) &&
    isTraceTableSection(value.section) &&
    (value.columnStart === undefined || typeof value.columnStart === "string") &&
    (value.columnEnd === undefined || typeof value.columnEnd === "string")
  ) {
    return {
      kind: "table",
      tableName: value.tableName,
      section: value.section,
      columnStart: isNonEmptyString(value.columnStart) ? value.columnStart : undefined,
      columnEnd: isNonEmptyString(value.columnEnd) ? value.columnEnd : undefined,
    };
  }
  return null;
}

export function parseDialogToParentMessage(raw: unknown): DialogToParentMessage | null {
  const message = parseRaw(raw);
  if (!message || !isNonEmptyString(message.sessionId)) return null;

  if (message.action === "ready") {
    return { action: "ready", sessionId: message.sessionId };
  }
  if (
    message.action === "expand" &&
    isNonEmptyString(message.nodeId) &&
    isNonEmptyString(message.requestId)
  ) {
    return {
      action: "expand",
      sessionId: message.sessionId,
      nodeId: message.nodeId,
      requestId: message.requestId,
    };
  }
  if (
    message.action === "loadRangePage" &&
    isNonEmptyString(message.nodeId) &&
    isNonEmptyString(message.requestId) &&
    typeof message.page === "number" &&
    Number.isInteger(message.page) &&
    message.page >= 0
  ) {
    return {
      action: "loadRangePage",
      sessionId: message.sessionId,
      nodeId: message.nodeId,
      requestId: message.requestId,
      page: message.page,
    };
  }
  if (message.action === "navigate") {
    const target = parseTraceTarget(message.target);
    if (!target) return null;
    return {
      action: "navigate",
      sessionId: message.sessionId,
      target,
    };
  }
  if (message.action === "close") {
    return {
      action: "close",
      sessionId: message.sessionId,
      ...(message.restoreRoot === true ? { restoreRoot: true } : {}),
    };
  }
  return null;
}

export function parseParentToDialogMessage(raw: unknown): ParentToDialogMessage | null {
  const message = parseRaw(raw);
  if (!message || !isNonEmptyString(message.sessionId)) return null;

  if (
    message.action === "rootLoaded" &&
    (message.direction === "precedents" || message.direction === "dependents") &&
    isTraceNode(message.root) &&
    Array.isArray(message.children) &&
    message.children.every(isTraceNode) &&
    typeof message.truncated === "boolean" &&
    typeof message.totalNodeCount === "number"
  ) {
    return message as unknown as ParentToDialogMessage;
  }
  if (
    message.action === "childrenLoaded" &&
    isNonEmptyString(message.requestId) &&
    isNonEmptyString(message.parentId) &&
    isNodeState(message.parentState) &&
    Array.isArray(message.children) &&
    message.children.every(isTraceNode) &&
    typeof message.truncated === "boolean" &&
    typeof message.cached === "boolean" &&
    typeof message.totalNodeCount === "number"
  ) {
    return message as unknown as ParentToDialogMessage;
  }
  if (
    message.action === "rangePageLoaded" &&
    isNonEmptyString(message.requestId) &&
    isNonEmptyString(message.rangeId) &&
    isNodeState(message.parentState) &&
    typeof message.page === "number" &&
    Array.isArray(message.children) &&
    message.children.every(isTraceNode) &&
    typeof message.hasMore === "boolean" &&
    typeof message.truncated === "boolean" &&
    typeof message.totalNodeCount === "number"
  ) {
    return message as unknown as ParentToDialogMessage;
  }
  if (
    message.action === "nodeFailed" &&
    isNonEmptyString(message.requestId) &&
    isNonEmptyString(message.nodeId) &&
    typeof message.message === "string"
  ) {
    return message as unknown as ParentToDialogMessage;
  }
  return null;
}
