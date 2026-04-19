/* global Office */

import type { TraceRow } from "../core/traceBuilder";

// Phase B (revised): dialogs in Office add-ins cannot call Excel.run —
// https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins
// documents the restriction ("cannot use host-specific APIs like Excel.run
// or Word.run to interact with the host document"). So this file is pure
// UI: it receives pre-computed trace rows from the parent runtime
// (taskpane or commands) via DialogParentMessageReceived, renders them
// as a tree, handles keyboard navigation, and sends back navigate / close
// messages.
//
// Tree model (Phase C): rows carry parentAddress, visibility derives from
// an expanded Set<address>. Default is all nodes expanded. User can
// collapse/expand with chevron click or ArrowRight/Left. Progressive
// setRows updates preserve the user's collapsed nodes.

const BODY_ID = "trace-dialog-body";
const STATUS_ID = "trace-dialog-status";
const TITLE_ID = "trace-dialog-title";
const ROW_ADDRESS_ATTR = "data-trace-address";
const ROW_FOCUSED_CLASS = "trace-row-focused";
const CHEVRON_CLASS = "trace-chevron";

// ---- Tree state ----
// `allRows` carries everything received from the parent (regardless of
// current expand/collapse state); `visibleRows` is the flattened list
// of rows currently visible — what keyboard navigation operates on.
// `expanded` tracks which addresses show their children. `childIndex`
// is derived from allRows and used for O(1) parent→children lookups.
let allRows: TraceRow[] = [];
let visibleRows: TraceRow[] = [];
let expanded = new Set<string>();
let childIndex = new Map<string, TraceRow[]>();
let currentFocusIndex: number | null = null;

// ---- Perf instrumentation (remove when diagnosis is done). ----
const perfSession = new URLSearchParams(window.location.search).get("perfSession") ?? "?";
function logTracePerf(label: string): void {
  // eslint-disable-next-line no-console
  console.log(`[trace-perf] session=${perfSession} ${label} @ ${Date.now()}`);
}

function setDialogStatus(message: string): void {
  const el = document.getElementById(STATUS_ID);
  if (el) el.textContent = message;
}

function setDialogTitle(message: string): void {
  const el = document.getElementById(TITLE_ID);
  if (el) el.textContent = message;
}

function readDirectionFromUrl(): "precedents" | "dependents" {
  const params = new URLSearchParams(window.location.search);
  const raw = params.get("direction");
  return raw === "dependents" ? "dependents" : "precedents";
}

function getRowElements(): HTMLTableRowElement[] {
  const body = document.getElementById(BODY_ID);
  if (!(body instanceof HTMLTableSectionElement)) return [];
  return Array.from(body.querySelectorAll<HTMLTableRowElement>(`tr[${ROW_ADDRESS_ATTR}]`));
}

function sendToParent(
  message:
    | { action: "ready" }
    | { action: "navigate"; address: string }
    | { action: "close" }
): void {
  try {
    Office.context.ui.messageParent(JSON.stringify(message));
  } catch {
    // Intentionally swallow. Most likely "parent not listening yet"; the
    // next keystroke (or re-open) re-sends.
  }
}

// ---- Tree helpers ----

function rebuildChildIndex(): void {
  childIndex = new Map();
  for (const row of allRows) {
    if (row.parentAddress !== null) {
      const list = childIndex.get(row.parentAddress) ?? [];
      list.push(row);
      childIndex.set(row.parentAddress, list);
    }
  }
}

function rowHasChildren(row: TraceRow): boolean {
  return (childIndex.get(row.address) ?? []).length > 0;
}

function findRootRow(): TraceRow | null {
  return allRows.find((r) => r.parentAddress === null) ?? null;
}

/**
 * Recompute `visibleRows` by DFS-traversing `allRows` from the root
 * and descending into children only through `expanded` nodes. Stable
 * with respect to BFS discovery order: siblings appear in the order
 * they were first discovered.
 */
function computeVisibleRows(): void {
  visibleRows = [];
  const root = findRootRow();
  if (!root) return;

  const walk = (node: TraceRow): void => {
    visibleRows.push(node);
    if (!expanded.has(node.address)) return;
    const children = childIndex.get(node.address) ?? [];
    for (const child of children) walk(child);
  };
  walk(root);
}

function findVisibleIndexByAddress(addr: string): number {
  return visibleRows.findIndex((r) => r.address === addr);
}

/**
 * Walk parent pointers until we find an ancestor that is currently
 * visible. Used when the focused row is hidden by a collapse (its
 * ancestor collapsed), so focus can land on the nearest visible
 * ancestor instead of vanishing.
 */
function findNearestVisibleAncestor(addr: string): string | null {
  const rowsByAddress = new Map(allRows.map((r) => [r.address, r]));
  let current = rowsByAddress.get(addr);
  while (current && current.parentAddress !== null) {
    const parent = rowsByAddress.get(current.parentAddress);
    if (parent && findVisibleIndexByAddress(parent.address) >= 0) {
      return parent.address;
    }
    current = parent;
  }
  return null;
}

// ---- Rendering ----

function focusRow(targetIndex: number, options: { announce?: boolean } = {}): void {
  const elements = getRowElements();
  if (elements.length === 0) {
    currentFocusIndex = null;
    return;
  }
  const clamped = Math.max(0, Math.min(elements.length - 1, targetIndex));
  elements.forEach((el, i) => {
    el.classList.toggle(ROW_FOCUSED_CLASS, i === clamped);
    el.setAttribute("aria-selected", i === clamped ? "true" : "false");
  });
  currentFocusIndex = clamped;
  const target = elements[clamped];
  target.focus();
  target.scrollIntoView({ block: "nearest" });

  if (options.announce !== false) {
    const row = visibleRows[clamped];
    if (row) sendToParent({ action: "navigate", address: row.address });
  }
}

function createTreeRow(row: TraceRow): HTMLTableRowElement {
  const tr = document.createElement("tr");
  tr.className = "trace-row-clickable";
  tr.setAttribute("role", "treeitem");
  tr.setAttribute("tabindex", "0");
  tr.setAttribute("aria-level", String(row.level + 1));
  tr.setAttribute(ROW_ADDRESS_ATTR, row.address);
  tr.setAttribute("aria-selected", "false");

  const hasChildren = rowHasChildren(row);
  const isExpanded = expanded.has(row.address);
  if (hasChildren) {
    tr.setAttribute("aria-expanded", isExpanded ? "true" : "false");
  }

  const addressTd = document.createElement("td");
  addressTd.className = "trace-address-cell";

  // Indent spacer: CSS reads --indent-level to compute the width.
  const indent = document.createElement("span");
  indent.className = "trace-indent";
  indent.style.setProperty("--indent-level", String(row.level));
  addressTd.appendChild(indent);

  // Chevron (or invisible spacer for leaves, to keep alignment).
  const chevron = document.createElement("span");
  chevron.className = CHEVRON_CLASS;
  if (hasChildren) {
    chevron.textContent = isExpanded ? "▼" : "▶";
    chevron.setAttribute("role", "button");
    chevron.setAttribute("aria-label", isExpanded ? "Collapse" : "Expand");
    chevron.addEventListener("click", (event) => {
      // Don't let the click bubble to the row handler (which would
      // navigate Excel's selection). Chevron clicks only toggle the
      // tree; they don't change the Excel active cell.
      event.stopPropagation();
      toggleExpanded(row.address, { fromUser: true });
    });
  } else {
    chevron.classList.add("trace-chevron-leaf");
    chevron.textContent = "•";
    chevron.setAttribute("aria-hidden", "true");
  }
  addressTd.appendChild(chevron);

  const addressText = document.createElement("span");
  addressText.className = "trace-address-text";
  addressText.textContent = row.address;
  addressTd.appendChild(addressText);

  const valueTd = document.createElement("td");
  valueTd.textContent = row.value;
  const formulaTd = document.createElement("td");
  formulaTd.textContent = row.formula;

  tr.append(addressTd, valueTd, formulaTd);

  tr.addEventListener("click", () => {
    sendToParent({ action: "navigate", address: row.address });
  });

  return tr;
}

function renderTree(options: { preserveFocus?: boolean } = {}): void {
  const body = document.getElementById(BODY_ID);
  if (!(body instanceof HTMLTableSectionElement)) return;

  // Remember the focused row's address so a re-render (progressive
  // update, expand/collapse) can restore focus to the same logical
  // position rather than snapping back to row 0.
  const focusedAddress =
    options.preserveFocus && currentFocusIndex !== null && currentFocusIndex < visibleRows.length
      ? visibleRows[currentFocusIndex].address
      : null;

  computeVisibleRows();

  body.textContent = "";
  currentFocusIndex = null;

  if (visibleRows.length === 0) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 3;
    td.textContent = "No trace results.";
    tr.appendChild(td);
    body.appendChild(tr);
    return;
  }

  for (const row of visibleRows) {
    body.appendChild(createTreeRow(row));
  }

  logTracePerf(`t7b tree.rendered visible=${visibleRows.length} total=${allRows.length}`);

  // Restore focus: prefer the previously-focused address; if collapsed
  // out of view, fall back to nearest visible ancestor; otherwise row 0.
  let focusTarget = 0;
  if (focusedAddress !== null) {
    const directIdx = findVisibleIndexByAddress(focusedAddress);
    if (directIdx >= 0) {
      focusTarget = directIdx;
    } else {
      const ancestor = findNearestVisibleAncestor(focusedAddress);
      if (ancestor !== null) {
        const ancestorIdx = findVisibleIndexByAddress(ancestor);
        if (ancestorIdx >= 0) focusTarget = ancestorIdx;
      }
    }
  }
  // announce=false: expand/collapse and progressive re-render shouldn't
  // fire navigate on their own. Explicit keyboard movements do.
  focusRow(focusTarget, { announce: false });
  if (focusTarget === 0 && focusedAddress === null) {
    logTracePerf("t8 row0.focused");
  }
}

/**
 * Toggle a node's expansion state. `fromUser=true` means it was a
 * direct user action (click or arrow key), so we re-render with focus
 * preserved on the toggled node. Used by both chevron-click and the
 * keyboard handler.
 */
function toggleExpanded(address: string, options: { fromUser?: boolean } = {}): void {
  if (expanded.has(address)) {
    expanded.delete(address);
  } else {
    expanded.add(address);
  }
  // If the user toggled, their conceptual focus is on the toggled
  // node; make sure we land focus there after re-render even if they
  // were focused on a child that's now hidden.
  if (options.fromUser) {
    // Seed focus via the visibleRows index of the toggled node after
    // recompute. Simplest: set currentFocusIndex = -1 before render
    // so preservation picks up the address → ancestor fallback.
    // But we can also just re-render and re-focus explicitly.
  }
  renderTree({ preserveFocus: true });
  // If the user explicitly toggled a node they weren't focused on,
  // preserve their existing focus. If they toggled the currently
  // focused node, focus stays. Either way the preserveFocus path
  // handles it.
  if (options.fromUser) {
    const idx = findVisibleIndexByAddress(address);
    if (idx >= 0) focusRow(idx, { announce: false });
  }
}

// ---- Keyboard ----

function handleDialogKeydown(event: KeyboardEvent): void {
  if (visibleRows.length === 0) return;
  const i = currentFocusIndex ?? 0;
  const focusedRow = visibleRows[i];

  switch (event.key) {
    case "ArrowDown":
      event.preventDefault();
      focusRow(i + 1);
      break;
    case "ArrowUp":
      event.preventDefault();
      focusRow(i - 1);
      break;
    case "Home":
      event.preventDefault();
      focusRow(0);
      break;
    case "End":
      event.preventDefault();
      focusRow(visibleRows.length - 1);
      break;
    case "ArrowRight": {
      event.preventDefault();
      if (!focusedRow) break;
      const hasKids = rowHasChildren(focusedRow);
      if (!hasKids) break; // leaf: no-op
      if (!expanded.has(focusedRow.address)) {
        // Collapsed parent → expand; focus stays on same node.
        expanded.add(focusedRow.address);
        renderTree({ preserveFocus: true });
      } else {
        // Expanded parent → move focus to first child.
        const firstChildIdx = findVisibleIndexByAddress(focusedRow.address) + 1;
        if (firstChildIdx < visibleRows.length) focusRow(firstChildIdx);
      }
      break;
    }
    case "ArrowLeft": {
      event.preventDefault();
      if (!focusedRow) break;
      if (rowHasChildren(focusedRow) && expanded.has(focusedRow.address)) {
        // Expanded parent → collapse; focus stays on same node.
        expanded.delete(focusedRow.address);
        renderTree({ preserveFocus: true });
      } else if (focusedRow.parentAddress !== null) {
        // Leaf or collapsed: move to parent.
        const parentIdx = findVisibleIndexByAddress(focusedRow.parentAddress);
        if (parentIdx >= 0) focusRow(parentIdx);
      }
      break;
    }
    case "Enter":
    case "Escape":
      event.preventDefault();
      sendToParent({ action: "close" });
      break;
    default:
      break;
  }
}

function handleDialogFocusIn(event: FocusEvent): void {
  const target = event.target;
  if (!(target instanceof Element)) return;
  const row = target.closest(`tr[${ROW_ADDRESS_ATTR}]`);
  if (!(row instanceof HTMLTableRowElement)) return;
  const addr = row.getAttribute(ROW_ADDRESS_ATTR);
  if (addr === null) return;
  const idx = findVisibleIndexByAddress(addr);
  if (idx >= 0 && currentFocusIndex !== idx) focusRow(idx);
}

function wireDialogKeyboard(): void {
  const body = document.getElementById(BODY_ID);
  if (!(body instanceof HTMLTableSectionElement)) return;
  body.addEventListener("keydown", handleDialogKeydown);
  body.addEventListener("focusin", handleDialogFocusIn);
}

// ---- Message handling ----

type ParentToDialog =
  | {
      action: "setRows";
      rows: TraceRow[];
      direction: "precedents" | "dependents";
      startAddress: string;
      truncated: boolean;
      isFinal: boolean;
      level: number;
    }
  | { action: "error"; message: string };

function parseParentMessage(raw: unknown): ParentToDialog | null {
  if (typeof raw !== "string") return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return null;
  }
  if (!parsed || typeof parsed !== "object") return null;
  const obj = parsed as { action?: unknown };
  if (obj.action === "setRows") {
    const msg = parsed as Record<string, unknown>;
    if (!Array.isArray(msg.rows)) return null;
    const direction = msg.direction === "dependents" ? "dependents" : "precedents";
    return {
      action: "setRows",
      rows: msg.rows as TraceRow[],
      direction,
      startAddress: typeof msg.startAddress === "string" ? msg.startAddress : "",
      truncated: msg.truncated === true,
      isFinal: msg.isFinal !== false,
      level: typeof msg.level === "number" ? msg.level : 0,
    };
  }
  if (obj.action === "error") {
    const msg = parsed as { message?: unknown };
    return {
      action: "error",
      message: typeof msg.message === "string" ? msg.message : "Unknown error",
    };
  }
  return null;
}

/**
 * Apply a setRows message: update allRows, rebuild the child index,
 * merge expansion state, and re-render. On the first emission all
 * nodes that have children are expanded. On progressive updates the
 * user's existing collapses stick; newly arriving nodes default to
 * expanded so the user sees the full subtree as it streams in.
 */
function applySetRows(newRows: TraceRow[]): void {
  const isFirstEmit = allRows.length === 0;
  const prevAddresses = new Set(allRows.map((r) => r.address));

  allRows = newRows;
  rebuildChildIndex();

  if (isFirstEmit) {
    // Default: every node that has children starts expanded.
    expanded = new Set();
    for (const row of allRows) {
      if (rowHasChildren(row)) expanded.add(row.address);
    }
  } else {
    // Progressive: only auto-expand nodes that are newly present AND
    // have children. Previously-expanded nodes remain expanded; user's
    // collapsed nodes remain collapsed.
    for (const row of allRows) {
      if (!prevAddresses.has(row.address) && rowHasChildren(row)) {
        expanded.add(row.address);
      }
    }
  }

  renderTree({ preserveFocus: !isFirstEmit });
}

function handleParentMessage(arg: { message?: string } | { error: number }): void {
  if (!("message" in arg) || typeof arg.message !== "string") return;
  const parsed = parseParentMessage(arg.message);
  if (!parsed) return;

  if (parsed.action === "setRows") {
    logTracePerf(
      `t7 setRows.received count=${parsed.rows.length} level=${parsed.level} final=${parsed.isFinal}`
    );
    setDialogTitle(`Trace ${parsed.direction}`);
    applySetRows(parsed.rows);

    const count = parsed.rows.length;
    const noun = count === 1 ? "cell" : "cells";
    const addrPart = parsed.startAddress ? ` on ${parsed.startAddress}` : "";
    const truncPart = parsed.truncated ? " (truncated)" : "";
    if (parsed.isFinal) {
      setDialogStatus(`Trace ${parsed.direction}${addrPart}: ${count} ${noun}${truncPart}.`);
    } else {
      setDialogStatus(
        `Trace ${parsed.direction}${addrPart}: loading… ${count} ${noun} so far (depth ${parsed.level})`
      );
    }
    return;
  }

  if (parsed.action === "error") {
    setDialogStatus(`Trace failed: ${parsed.message}`);
  }
}

Office.onReady((info) => {
  logTracePerf("t2 dialog.onReady");
  if (info.host !== Office.HostType.Excel) {
    setDialogStatus("Trace dialog requires Excel.");
    return;
  }

  setDialogTitle(`Trace ${readDirectionFromUrl()}`);
  wireDialogKeyboard();

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    handleParentMessage,
    () => {
      logTracePerf("t2a addHandler.callback");
      sendToParent({ action: "ready" });
    }
  );
});
