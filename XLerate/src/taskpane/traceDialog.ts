/* global Office */

import {
  applyTraceDialogMessage,
  buildInitialTraceDialogState,
  computeVisibleTraceNodes,
  setTraceNodeLoading,
  toggleTraceNode,
  type TraceDialogState,
} from "../core/traceDialogState";
import { tokenizeFormulaReferences } from "../core/formulaReferences";
import {
  buildTraceFormulaContext,
  getHighlightedFormulaSegmentIndexes,
} from "../core/traceFormulaContext";
import { isTraceNodeExpandable, type TraceNode, type TraceTarget } from "../core/traceGraph";
import { parseParentToDialogMessage, type DialogToParentMessage } from "../core/traceProtocol";

const BODY_ID = "trace-dialog-body";
const STATUS_ID = "trace-dialog-status";
const TITLE_ID = "trace-dialog-title";
const CONTEXT_BREADCRUMB_ID = "trace-context-breadcrumb";
const CONTEXT_FORMULA_ID = "trace-context-formula";
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
  navigateTarget({ address: node.address });
}

function navigateTarget(target: TraceTarget): void {
  sendToParent({
    action: "navigate",
    sessionId,
    target,
  });
}

function describeTraceTarget(target: TraceTarget): string {
  if ("address" in target) return target.address;
  if (target.kind === "namedRange") {
    return target.worksheetName
      ? `named range ${target.worksheetName}!${target.name}`
      : `named range ${target.name}`;
  }
  return `table ${target.tableName}`;
}

function focusNode(nodeId: string, announce = true, moveDomFocus = true): void {
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
  renderTraceContext();
  if (moveDomFocus) {
    target.focus();
    target.scrollIntoView({ block: "nearest" });
  }

  if (announce) {
    const node = state.nodesById.get(nodeId);
    if (node) navigate(node);
  }
}

function stopRowKeyboardForButton(button: HTMLButtonElement): void {
  button.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") event.stopPropagation();
  });
}

function appendFormulaContent(
  formulaCell: HTMLElement,
  node: TraceNode,
  highlightedIndexes: ReadonlySet<number> = new Set()
): void {
  if (node.kind === "range") {
    formulaCell.textContent = "Range summary";
    return;
  }

  const segments = tokenizeFormulaReferences(node.formula, node.worksheetName);
  segments.forEach((segment, index) => {
    if (segment.kind === "text") {
      formulaCell.appendChild(document.createTextNode(segment.text));
      return;
    }
    if (segment.kind === "external") {
      const external = document.createElement("span");
      external.className = "trace-formula-external";
      external.textContent = segment.text;
      external.title = `${segment.reason} Target: [${segment.workbookName}]${segment.worksheetName}!${segment.rangeAddress}`;
      external.setAttribute(
        "aria-label",
        `External workbook reference ${segment.workbookName}, ${segment.worksheetName}!${segment.rangeAddress}. Navigation unavailable.`
      );
      formulaCell.appendChild(external);
      return;
    }
    const reference = document.createElement("button");
    reference.type = "button";
    reference.className = "trace-formula-reference";
    if (highlightedIndexes.has(index)) {
      reference.classList.add("trace-formula-reference-active");
      reference.setAttribute("aria-current", "true");
    }
    reference.textContent = segment.text;
    const targetDescription = describeTraceTarget(segment.target);
    reference.title = `Go to ${targetDescription}`;
    reference.setAttribute("aria-label", `Go to ${targetDescription}`);
    reference.addEventListener("click", (event) => {
      event.stopPropagation();
      navigateTarget(segment.target);
    });
    stopRowKeyboardForButton(reference);
    formulaCell.appendChild(reference);
  });
}

function renderTraceBreadcrumb(contextNodeId: string): void {
  const container = document.getElementById(CONTEXT_BREADCRUMB_ID);
  if (!container) return;
  container.textContent = "";
  const context = buildTraceFormulaContext(state, contextNodeId, direction);
  if (!context) return;

  context.breadcrumb.forEach((node, index) => {
    if (index > 0) {
      const separator = document.createElement("span");
      separator.className = "trace-breadcrumb-separator";
      separator.textContent = "›";
      separator.setAttribute("aria-hidden", "true");
      container.appendChild(separator);
    }
    const button = document.createElement("button");
    button.type = "button";
    button.className = "trace-breadcrumb-button";
    button.textContent = node.address;
    button.title = `Go to ${node.address}`;
    const isCurrent = node.id === contextNodeId;
    if (isCurrent) {
      button.setAttribute("aria-current", "location");
      button.disabled = true;
    } else {
      button.addEventListener("click", () => focusNode(node.id));
    }
    stopRowKeyboardForButton(button);
    container.appendChild(button);
  });
}

function renderTraceContext(): void {
  if (!currentFocusId) return;
  const context = buildTraceFormulaContext(state, currentFocusId, direction);
  if (!context) return;

  renderTraceBreadcrumb(currentFocusId);

  const formula = document.getElementById(CONTEXT_FORMULA_ID);
  if (formula) {
    formula.textContent = "";
    if (context.formulaNode.kind === "range") {
      formula.textContent = "Expand this summarized range to inspect an individual formula.";
    } else if (context.formulaNode.formula.length === 0) {
      formula.textContent = "This cell does not contain a formula.";
    } else {
      const segments = tokenizeFormulaReferences(
        context.formulaNode.formula,
        context.formulaNode.worksheetName
      );
      appendFormulaContent(
        formula,
        context.formulaNode,
        getHighlightedFormulaSegmentIndexes(segments, context.relationNode)
      );
    }
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
  formulaCell.className = "trace-formula-cell";
  appendFormulaContent(formulaCell, node);

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
      stopRowKeyboardForButton(loadMore);
      formulaCell.append(" ", loadMore);
    }
  }

  row.append(addressCell, valueCell, formulaCell);
  row.addEventListener("click", () => focusNode(node.id));
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

  if (event.key === "Enter") {
    event.preventDefault();
    sendToParent({ action: "close", sessionId });
    return;
  }

  if (event.key === "Escape") {
    event.preventDefault();
    sendToParent({ action: "close", sessionId, restoreRoot: true });
  }
}

function handleDialogFocusIn(event: FocusEvent): void {
  const target = event.target;
  if (!(target instanceof Element)) return;
  const row = target.closest(`tr[${ROW_ID_ATTR}]`);
  if (!(row instanceof HTMLTableRowElement)) return;
  const nodeId = row.getAttribute(ROW_ID_ATTR);
  if (nodeId && currentFocusId !== nodeId) {
    focusNode(nodeId, false, target.closest("button") === null);
  }
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
    body.addEventListener("focusin", handleDialogFocusIn);
  }
  document.addEventListener("keydown", handleDialogKeydown);
  renderTree();
  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    handleParentMessage,
    () => sendToParent({ action: "ready", sessionId })
  );
});
