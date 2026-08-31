/* global Office */

import {
  applyTraceDialogMessage,
  buildInitialTraceDialogState,
  computeVisibleTraceNodes,
  setTraceNodeLoading,
  toggleTraceNode,
  type TraceDialogState,
} from "../core/traceDialogState";
import { isTraceNodeExpandable, type TraceNode } from "../core/traceGraph";
import { parseParentToDialogMessage, type DialogToParentMessage } from "../core/traceProtocol";

const BODY_ID = "trace-dialog-body";
const STATUS_ID = "trace-dialog-status";
const TITLE_ID = "trace-dialog-title";
const ROW_ID_ATTR = "data-trace-node-id";
const ROW_FOCUSED_CLASS = "trace-row-focused";

const params = new URLSearchParams(window.location.search);
const sessionId = params.get("sessionId") ?? "invalid-session";
const direction = params.get("direction") === "dependents" ? "dependents" : "precedents";

let state: TraceDialogState = buildInitialTraceDialogState(sessionId);
let visibleNodes: TraceNode[] = [];
let currentFocusId: string | null = null;
let rowElementById = new Map<string, HTMLTableRowElement>();
let requestSequence = 0;

function nextRequestId(prefix: string): string {
  requestSequence += 1;
  return `${prefix}-${requestSequence}`;
}

function setDialogStatus(message: string): void {
  const element = document.getElementById(STATUS_ID);
  if (element) element.textContent = message;
}

function setDialogTitle(message: string): void {
  const element = document.getElementById(TITLE_ID);
  if (element) element.textContent = message;
}

function sendToParent(message: DialogToParentMessage): void {
  try {
    Office.context.ui.messageParent(JSON.stringify(message));
  } catch {
    // The parent may have closed while an event was being processed.
  }
}

function navigate(node: TraceNode): void {
  sendToParent({
    action: "navigate",
    sessionId,
    target: { address: node.address },
  });
}

function focusNode(nodeId: string, announce = true): void {
  const target = rowElementById.get(nodeId);
  if (!target) return;

  if (currentFocusId && currentFocusId !== nodeId) {
    const previous = rowElementById.get(currentFocusId);
    previous?.classList.remove(ROW_FOCUSED_CLASS);
    previous?.setAttribute("aria-selected", "false");
  }
  target.classList.add(ROW_FOCUSED_CLASS);
  target.setAttribute("aria-selected", "true");
  currentFocusId = nodeId;
  target.focus();
  target.scrollIntoView({ block: "nearest" });

  if (announce) {
    const node = state.nodesById.get(nodeId);
    if (node) navigate(node);
  }
}

function requestExpansion(node: TraceNode): void {
  if (!isTraceNodeExpandable(node) || node.loadState === "loading") return;
  state = setTraceNodeLoading(state, node.id);
  renderTree(node.id);
  const requestId = nextRequestId(node.kind === "range" ? "page" : "expand");
  if (node.kind === "range") {
    const page = state.rangePages.get(node.id)?.nextPage ?? 0;
    sendToParent({
      action: "loadRangePage",
      sessionId,
      nodeId: node.id,
      requestId,
      page,
    });
  } else {
    sendToParent({ action: "expand", sessionId, nodeId: node.id, requestId });
  }
}

function expandOrToggle(node: TraceNode): void {
  if (!isTraceNodeExpandable(node)) return;
  if (state.expandedIds.has(node.id)) {
    state = toggleTraceNode(state, node.id, false);
    renderTree(node.id);
    return;
  }
  const existingChildren = state.childIdsByParent.get(node.id) ?? [];
  if (existingChildren.length > 0) {
    state = toggleTraceNode(state, node.id, true);
    renderTree(node.id);
    return;
  }
  requestExpansion(node);
}

function createChevron(node: TraceNode): HTMLSpanElement {
  const chevron = document.createElement("span");
  chevron.className = "trace-chevron";
  if (!isTraceNodeExpandable(node)) {
    chevron.classList.add("trace-chevron-leaf");
    chevron.textContent = node.kind === "reference" ? "↩" : "•";
    chevron.setAttribute("aria-hidden", "true");
    return chevron;
  }

  const expanded = state.expandedIds.has(node.id);
  chevron.textContent =
    node.loadState === "loading" ? "…" : node.loadState === "error" ? "!" : expanded ? "▼" : "▶";
  chevron.setAttribute("role", "button");
  chevron.setAttribute(
    "aria-label",
    node.loadState === "loading" ? "Loading" : expanded ? "Collapse" : "Expand"
  );
  chevron.addEventListener("click", (event) => {
    event.stopPropagation();
    expandOrToggle(node);
  });
  return chevron;
}

function createTreeRow(node: TraceNode): HTMLTableRowElement {
  const row = document.createElement("tr");
  row.className = "trace-row-clickable";
  row.setAttribute("role", "treeitem");
  row.setAttribute("tabindex", "0");
  row.setAttribute("aria-level", String(node.level + 1));
  row.setAttribute(ROW_ID_ATTR, node.id);
  row.dataset.renderKey = buildRowRenderKey(node);
  row.setAttribute("aria-selected", "false");
  if (isTraceNodeExpandable(node)) {
    row.setAttribute("aria-expanded", state.expandedIds.has(node.id) ? "true" : "false");
  }

  const addressCell = document.createElement("td");
  addressCell.className = "trace-address-cell";
  const indent = document.createElement("span");
  indent.className = "trace-indent";
  indent.style.setProperty("--indent-level", String(node.level));
  const addressText = document.createElement("span");
  addressText.className = "trace-address-text";
  addressText.textContent = node.address;
  addressCell.append(indent, createChevron(node), addressText);

  const valueCell = document.createElement("td");
  valueCell.textContent =
    node.kind === "range" ? `${node.cellCount.toLocaleString()} cells` : node.value;
  const formulaCell = document.createElement("td");
  formulaCell.textContent = node.kind === "range" ? "Range summary" : node.formula;

  if (node.kind === "range") {
    const pageState = state.rangePages.get(node.id);
    if (pageState?.hasMore && state.expandedIds.has(node.id)) {
      const loadMore = document.createElement("button");
      loadMore.type = "button";
      loadMore.className = "trace-load-more";
      loadMore.textContent = "Load more";
      loadMore.addEventListener("click", (event) => {
        event.stopPropagation();
        requestExpansion(node);
      });
      formulaCell.append(" ", loadMore);
    }
  }

  row.append(addressCell, valueCell, formulaCell);
  row.addEventListener("click", () => navigate(node));
  return row;
}

function buildRowRenderKey(node: TraceNode): string {
  const pageState = node.kind === "range" ? state.rangePages.get(node.id) : undefined;
  return [
    node.kind,
    node.loadState,
    state.expandedIds.has(node.id) ? "expanded" : "collapsed",
    pageState?.nextPage ?? 0,
    pageState?.hasMore === true ? "more" : "done",
  ].join(":");
}

function renderTree(preferredFocusId: string | null = currentFocusId): void {
  const body = document.getElementById(BODY_ID);
  if (!(body instanceof HTMLTableSectionElement)) return;
  visibleNodes = computeVisibleTraceNodes(state);

  if (visibleNodes.length === 0) {
    rowElementById = new Map();
    body.textContent = "";
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 3;
    cell.textContent = "Loading trace root…";
    row.appendChild(cell);
    body.appendChild(row);
    return;
  }

  const visibleIds = new Set(visibleNodes.map((node) => node.id));
  for (const child of Array.from(body.children)) {
    const nodeId = child.getAttribute(ROW_ID_ATTR);
    if (!nodeId || !visibleIds.has(nodeId)) child.remove();
  }

  const nextElements = new Map<string, HTMLTableRowElement>();
  let cursor: ChildNode | null = body.firstChild;
  for (const node of visibleNodes) {
    const previous = rowElementById.get(node.id);
    const renderKey = buildRowRenderKey(node);
    let row = previous;
    if (!row || row.dataset.renderKey !== renderKey) {
      const replacement = createTreeRow(node);
      if (row?.isConnected) {
        const replacedCursor = row === cursor;
        row.replaceWith(replacement);
        if (replacedCursor) cursor = replacement;
      }
      row = replacement;
    }
    if (row !== cursor) body.insertBefore(row, cursor);
    cursor = row.nextSibling;
    nextElements.set(node.id, row);
  }
  rowElementById = nextElements;

  const focusId =
    preferredFocusId && rowElementById.has(preferredFocusId) ? preferredFocusId : state.rootId;
  if (focusId) focusNode(focusId, false);
}

function visibleIndex(nodeId: string): number {
  return visibleNodes.findIndex((node) => node.id === nodeId);
}

function handleDialogKeydown(event: KeyboardEvent): void {
  if (visibleNodes.length === 0) return;
  const index = currentFocusId ? visibleIndex(currentFocusId) : 0;
  const current = visibleNodes[Math.max(0, index)];
  if (!current) return;

  if (
    event.key === "ArrowDown" ||
    event.key === "ArrowUp" ||
    event.key === "Home" ||
    event.key === "End"
  ) {
    event.preventDefault();
    const targetIndex =
      event.key === "Home"
        ? 0
        : event.key === "End"
          ? visibleNodes.length - 1
          : event.key === "ArrowDown"
            ? Math.min(visibleNodes.length - 1, index + 1)
            : Math.max(0, index - 1);
    focusNode(visibleNodes[targetIndex].id);
    return;
  }

  if (event.key === "ArrowRight") {
    event.preventDefault();
    if (!state.expandedIds.has(current.id)) {
      expandOrToggle(current);
    } else {
      const firstChild = state.childIdsByParent.get(current.id)?.[0];
      if (firstChild) focusNode(firstChild);
    }
    return;
  }

  if (event.key === "ArrowLeft") {
    event.preventDefault();
    if (state.expandedIds.has(current.id)) {
      state = toggleTraceNode(state, current.id, false);
      renderTree(current.id);
    } else if (current.parentId) {
      focusNode(current.parentId);
    }
    return;
  }

  if (event.key === "Enter" || event.key === "Escape") {
    event.preventDefault();
    sendToParent({ action: "close", sessionId });
  }
}

function handleDialogFocusIn(event: FocusEvent): void {
  const target = event.target;
  if (!(target instanceof Element)) return;
  const row = target.closest(`tr[${ROW_ID_ATTR}]`);
  if (!(row instanceof HTMLTableRowElement)) return;
  const nodeId = row.getAttribute(ROW_ID_ATTR);
  if (nodeId && currentFocusId !== nodeId) focusNode(nodeId, false);
}

function handleParentMessage(arg: { message?: string } | { error: number }): void {
  if (!("message" in arg) || typeof arg.message !== "string") return;
  const message = parseParentToDialogMessage(arg.message);
  if (!message || message.sessionId !== sessionId) return;

  const focusBefore = currentFocusId;
  state = applyTraceDialogMessage(state, message);
  renderTree(focusBefore ?? (message.action === "rootLoaded" ? message.root.id : null));

  if (message.action === "rootLoaded") {
    setDialogTitle(`Trace ${message.direction}`);
    const suffix = message.truncated ? " (truncated)" : "";
    setDialogStatus(
      `${message.totalNodeCount} nodes loaded${suffix}. Expand a row to load deeper levels.`
    );
    return;
  }
  if (message.action === "childrenLoaded") {
    const suffix = message.cached ? " from cache" : "";
    setDialogStatus(
      `${message.children.length} children loaded${suffix}. ${message.totalNodeCount} nodes in this trace.`
    );
    return;
  }
  if (message.action === "rangePageLoaded") {
    setDialogStatus(
      `Range page ${message.page + 1} loaded (${message.children.length} cells). ${message.totalNodeCount} nodes in this trace.`
    );
    return;
  }
  setDialogStatus(`Trace failed: ${message.message}`);
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setDialogStatus("Trace dialog requires Excel.");
    return;
  }
  setDialogTitle(`Trace ${direction}`);
  const body = document.getElementById(BODY_ID);
  if (body instanceof HTMLTableSectionElement) {
    body.addEventListener("keydown", handleDialogKeydown);
    body.addEventListener("focusin", handleDialogFocusIn);
  }
  renderTree();
  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    handleParentMessage,
    () => sendToParent({ action: "ready", sessionId })
  );
});
