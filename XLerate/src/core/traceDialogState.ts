import type { TraceNode } from "./traceGraph";
import type { ParentToDialogMessage } from "./traceProtocol";

export type TraceDialogState = {
  sessionId: string;
  rootId: string | null;
  nodesById: Map<string, TraceNode>;
  childIdsByParent: Map<string, string[]>;
  expandedIds: Set<string>;
  rangePages: Map<string, { nextPage: number; hasMore: boolean }>;
  totalNodeCount: number;
  truncated: boolean;
};

export function buildInitialTraceDialogState(sessionId: string): TraceDialogState {
  return {
    sessionId,
    rootId: null,
    nodesById: new Map(),
    childIdsByParent: new Map(),
    expandedIds: new Set(),
    rangePages: new Map(),
    totalNodeCount: 0,
    truncated: false,
  };
}

function addNodes(state: TraceDialogState, nodes: TraceNode[]): void {
  for (const node of nodes) {
    state.nodesById.set(node.id, node);
  }
}

export function applyTraceDialogMessage(
  current: TraceDialogState,
  message: ParentToDialogMessage
): TraceDialogState {
  if (message.sessionId !== current.sessionId) return current;

  const next: TraceDialogState = {
    ...current,
    nodesById: new Map(current.nodesById),
    childIdsByParent: new Map(current.childIdsByParent),
    expandedIds: new Set(current.expandedIds),
    rangePages: new Map(current.rangePages),
  };

  if (message.action === "rootLoaded") {
    next.rootId = message.root.id;
    addNodes(next, [message.root, ...message.children]);
    next.childIdsByParent.set(
      message.root.id,
      message.children.map((node) => node.id)
    );
    next.expandedIds.add(message.root.id);
    next.totalNodeCount = message.totalNodeCount;
    next.truncated = message.truncated;
    return next;
  }

  if (message.action === "childrenLoaded") {
    const parent = next.nodesById.get(message.parentId);
    if (parent) next.nodesById.set(parent.id, { ...parent, loadState: message.parentState });
    addNodes(next, message.children);
    next.childIdsByParent.set(
      message.parentId,
      message.children.map((node) => node.id)
    );
    next.expandedIds.add(message.parentId);
    next.totalNodeCount = message.totalNodeCount;
    next.truncated = next.truncated || message.truncated;
    return next;
  }

  if (message.action === "rangePageLoaded") {
    const parent = next.nodesById.get(message.rangeId);
    if (parent) next.nodesById.set(parent.id, { ...parent, loadState: message.parentState });
    addNodes(next, message.children);
    const existing = next.childIdsByParent.get(message.rangeId) ?? [];
    next.childIdsByParent.set(message.rangeId, [
      ...existing,
      ...message.children.map((node) => node.id),
    ]);
    next.expandedIds.add(message.rangeId);
    next.rangePages.set(message.rangeId, {
      nextPage: message.page + 1,
      hasMore: message.hasMore,
    });
    next.totalNodeCount = message.totalNodeCount;
    next.truncated = next.truncated || message.truncated;
    return next;
  }

  if (message.action === "nodeFailed") {
    const node = next.nodesById.get(message.nodeId);
    if (node) next.nodesById.set(node.id, { ...node, loadState: "error" });
  }
  return next;
}

export function toggleTraceNode(
  current: TraceDialogState,
  nodeId: string,
  expanded?: boolean
): TraceDialogState {
  const next = { ...current, expandedIds: new Set(current.expandedIds) };
  const shouldExpand = expanded ?? !next.expandedIds.has(nodeId);
  if (shouldExpand) next.expandedIds.add(nodeId);
  else next.expandedIds.delete(nodeId);
  return next;
}

export function setTraceNodeLoading(current: TraceDialogState, nodeId: string): TraceDialogState {
  const node = current.nodesById.get(nodeId);
  if (!node) return current;
  const nodesById = new Map(current.nodesById);
  nodesById.set(nodeId, { ...node, loadState: "loading" });
  return { ...current, nodesById };
}

export function computeVisibleTraceNodes(state: TraceDialogState): TraceNode[] {
  if (!state.rootId) return [];
  const visible: TraceNode[] = [];
  const walk = (nodeId: string): void => {
    const node = state.nodesById.get(nodeId);
    if (!node) return;
    visible.push(node);
    if (!state.expandedIds.has(nodeId)) return;
    for (const childId of state.childIdsByParent.get(nodeId) ?? []) walk(childId);
  };
  walk(state.rootId);
  return visible;
}
