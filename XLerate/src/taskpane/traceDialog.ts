/* global Excel, Office */

import { buildTrace, type TraceCellInfo, type TraceRow } from "../core/traceBuilder";
import { sanitizeTraceDepth, type TraceDirection } from "../core/traceUtils";
import {
  getDirectTraceNeighbors,
  loadTraceCellProperties,
  resolveTraceStartCell,
  snapshotRangeForTrace,
} from "./traceExcelNeighbors";

// Phase B.3: the dialog computes its own trace and renders the list.
// Keyboard focus-nav (B.4) and messageParent protocol (B.5) land next.
// This file has its own Excel.run context, independent of the taskpane's.

const BODY_ID = "trace-dialog-body";
const STATUS_ID = "trace-dialog-status";
const TITLE_ID = "trace-dialog-title";
const ROW_INDEX_ATTR = "data-trace-index";
const ROW_FOCUSED_CLASS = "trace-row-focused";

// Keyboard-navigation state. `currentRows` mirrors what's rendered;
// `currentFocusIndex` is the row that arrow keys operate on. Both are
// module-level: the dialog is single-trace, single-window.
let currentRows: TraceRow[] = [];
let currentFocusIndex: number | null = null;

function setDialogStatus(message: string): void {
  const el = document.getElementById(STATUS_ID);
  if (el) el.textContent = message;
}

function setDialogTitle(message: string): void {
  const el = document.getElementById(TITLE_ID);
  if (el) el.textContent = message;
}

type DialogParams = {
  direction: TraceDirection;
  address: string | null;
  maxDepth: number;
};

function parseDialogParams(): DialogParams {
  const params = new URLSearchParams(window.location.search);
  const rawDirection = params.get("direction");
  const direction: TraceDirection = rawDirection === "dependents" ? "dependents" : "precedents";
  const address = params.get("address");
  const depthRaw = params.get("maxDepth");
  const maxDepth = sanitizeTraceDepth(depthRaw === null ? undefined : Number(depthRaw));
  return { direction, address: address && address.length > 0 ? address : null, maxDepth };
}

function getRowElements(): HTMLTableRowElement[] {
  const body = document.getElementById(BODY_ID);
  if (!(body instanceof HTMLTableSectionElement)) return [];
  return Array.from(body.querySelectorAll<HTMLTableRowElement>(`tr[${ROW_INDEX_ATTR}]`));
}

/**
 * Fire-and-forget notification to the taskpane. `messageParent` is
 * synchronous on the dialog side — the dialog's window state isn't
 * affected and the caller doesn't need to await anything. No await, no
 * throw propagation: if messaging fails we still want the ring to move.
 */
function sendToParent(message: { action: "navigate"; address: string } | { action: "close" }): void {
  try {
    Office.context.ui.messageParent(JSON.stringify(message));
  } catch {
    // Intentionally swallow. A failed messageParent is usually "no parent
    // attached yet" during the very first render; the next keystroke will
    // re-send. Surfacing the error to the user would be noise.
  }
}

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

  // Live-nav: tell the taskpane to move Excel's active cell. Skipped on the
  // initial render (announce: false) because the dialog opens on the user's
  // current active cell anyway, and firing before the taskpane has attached
  // its DialogMessageReceived handler would just be lost.
  if (options.announce !== false) {
    const row = currentRows[clamped];
    if (row) sendToParent({ action: "navigate", address: row.address });
  }
}

function handleDialogKeydown(event: KeyboardEvent): void {
  if (currentRows.length === 0) return;
  const i = currentFocusIndex ?? 0;
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
      focusRow(currentRows.length - 1);
      break;
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
  const row = target.closest(`tr[${ROW_INDEX_ATTR}]`);
  if (!(row instanceof HTMLTableRowElement)) return;
  const raw = row.getAttribute(ROW_INDEX_ATTR);
  const idx = raw === null ? NaN : Number(raw);
  if (!Number.isInteger(idx) || idx < 0 || idx >= currentRows.length) return;
  if (currentFocusIndex !== idx) focusRow(idx);
}

function wireDialogKeyboard(): void {
  const body = document.getElementById(BODY_ID);
  if (!(body instanceof HTMLTableSectionElement)) return;
  body.addEventListener("keydown", handleDialogKeydown);
  body.addEventListener("focusin", handleDialogFocusIn);
}

function renderDialogTraceRows(rows: TraceRow[]): void {
  const body = document.getElementById(BODY_ID);
  if (!(body instanceof HTMLTableSectionElement)) return;

  body.textContent = "";
  currentRows = rows;
  currentFocusIndex = null;

  if (rows.length === 0) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 4;
    td.textContent = "No trace results.";
    tr.appendChild(td);
    body.appendChild(tr);
    return;
  }

  rows.forEach((item, index) => {
    const tr = document.createElement("tr");
    tr.className = "trace-row-clickable";
    tr.setAttribute("role", "option");
    tr.setAttribute("tabindex", "0");
    tr.setAttribute(ROW_INDEX_ATTR, String(index));
    tr.setAttribute("aria-selected", "false");

    const level = document.createElement("td");
    const address = document.createElement("td");
    const value = document.createElement("td");
    const formula = document.createElement("td");

    level.textContent = String(item.level);
    address.textContent = item.address;
    value.textContent = item.value;
    formula.textContent = item.formula;

    tr.append(level, address, value, formula);
    body.appendChild(tr);
  });

  // Dialog always claims focus when opened, so this .focus() reliably lands
  // visible focus on row 0 — unlike the taskpane version where the iframe
  // might not have document focus. Arrow keys work immediately.
  //
  // announce:false on the initial call — the dialog opens on the user's
  // current active cell, so Excel's selection already matches row 0.
  // Firing a navigate here would be redundant AND risks losing the message
  // to a race with the parent's DialogMessageReceived handler attachment.
  focusRow(0, { announce: false });
}

async function runDialogTrace(params: DialogParams): Promise<void> {
  setDialogTitle(`Trace ${params.direction}`);
  setDialogStatus(params.address ? `Tracing ${params.direction} from ${params.address}…` : `Tracing ${params.direction}…`);

  try {
    await Excel.run(async (context) => {
      const rootRange = await resolveTraceStartCell(context, params.address);
      if (!rootRange) {
        setDialogStatus("Could not resolve a starting cell for the trace.");
        return;
      }

      const root = snapshotRangeForTrace(rootRange);

      const getNeighbors = async (info: TraceCellInfo): Promise<TraceCellInfo[]> => {
        const worksheet = context.workbook.worksheets.getItem(info.worksheetName);
        const cell = worksheet.getRangeByIndexes(info.rowIndex, info.columnIndex, 1, 1);
        // resolveTraceStartCell/getDirectTraceNeighbors do their own loads; the
        // ephemeral cell above only needs to exist to call getDirectPrecedents/Dependents.
        loadTraceCellProperties(cell);
        const neighbors = await getDirectTraceNeighbors(context, cell, params.direction);
        return neighbors.map(snapshotRangeForTrace);
      };

      const { rows, truncated } = await buildTrace({
        root,
        maxDepth: params.maxDepth,
        getNeighbors,
      });

      renderDialogTraceRows(rows);
      setDialogStatus(
        `Trace ${params.direction} on ${root.address}: ${rows.length} cell${rows.length === 1 ? "" : "s"}${truncated ? " (truncated)" : ""}.`
      );
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setDialogStatus(`Trace failed: ${message}`);
  }
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setDialogStatus("Trace dialog requires Excel.");
    return;
  }
  wireDialogKeyboard();
  const params = parseDialogParams();
  void runDialogTrace(params);
});
