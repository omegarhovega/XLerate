/* global Office */

import type { TraceRow } from "../core/traceBuilder";

// Phase B (revised): dialogs in Office add-ins cannot call Excel.run —
// https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins
// documents the restriction ("cannot use host-specific APIs like Excel.run
// or Word.run to interact with the host document"). So this file is pure
// UI: it receives the pre-computed trace rows from the parent runtime
// (taskpane or commands) via DialogParentMessageReceived, renders them,
// handles keyboard navigation, and sends back navigate / close messages.

const BODY_ID = "trace-dialog-body";
const STATUS_ID = "trace-dialog-status";
const TITLE_ID = "trace-dialog-title";
const ROW_INDEX_ATTR = "data-trace-index";
const ROW_FOCUSED_CLASS = "trace-row-focused";

// Keyboard-navigation state. `currentRows` mirrors what's rendered;
// `currentFocusIndex` is the row that arrow keys operate on. Module-level:
// the dialog is single-trace, single-window.
let currentRows: TraceRow[] = [];
let currentFocusIndex: number | null = null;

// ---- Perf instrumentation (remove when diagnosis is done). ----
// Uses Date.now() so the dialog's timestamps line up with the parent
// runtime's (separate performance.now() origins make relative times
// otherwise incomparable across contexts). The `perfSession` query
// param is set by the parent at click time; including it in each log
// line lets the developer correlate dialog logs with parent logs when
// multiple traces run close together.
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
  return Array.from(body.querySelectorAll<HTMLTableRowElement>(`tr[${ROW_INDEX_ATTR}]`));
}

/**
 * Fire-and-forget notification to the parent runtime. `messageParent` is
 * synchronous on the dialog side — the dialog's window state isn't
 * affected and the caller doesn't need to await anything.
 */
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

  // Live-nav: tell the parent to move Excel's active cell. Skipped on the
  // initial render because the dialog opens on the user's current active
  // cell anyway — redundant AND would fire before the parent has a chance
  // to listen on the very first frame.
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

function renderDialogTraceRows(
  rows: TraceRow[],
  options: { preserveFocus?: boolean } = {}
): void {
  const body = document.getElementById(BODY_ID);
  if (!(body instanceof HTMLTableSectionElement)) return;

  // Remember the focused row's address (not just its index) so a
  // progressive update that happens to shorten or reorder rows still
  // lands focus on the right row. In practice BFS only appends, but
  // this is the robust path.
  const focusedAddress =
    options.preserveFocus && currentFocusIndex !== null && currentFocusIndex < currentRows.length
      ? currentRows[currentFocusIndex].address
      : null;

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

  logTracePerf(`t7b rows.rendered count=${rows.length}`);

  // Restore focus: prefer the previously-focused address (preserve
  // user position during progressive updates), fall back to row 0.
  let focusTarget = 0;
  if (focusedAddress !== null) {
    const found = rows.findIndex((r) => r.address === focusedAddress);
    if (found >= 0) focusTarget = found;
  }
  focusRow(focusTarget, { announce: false });
  if (focusTarget === 0 && focusedAddress === null) {
    logTracePerf("t8 row0.focused");
  }
}

/**
 * Payload shapes sent by the parent via dialog.messageChild. Kept
 * permissive; unknown actions are ignored. `setRows` may be called
 * multiple times per trace — once per BFS level during progressive
 * loading — so `isFinal` marks the last emission.
 */
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
    // Defensive defaults keep backward compatibility with any older
    // parent that omits the progressive-loading fields.
    return {
      action: "setRows",
      rows: msg.rows as TraceRow[],
      direction,
      startAddress: typeof msg.startAddress === "string" ? msg.startAddress : "",
      truncated: msg.truncated === true,
      isFinal: msg.isFinal !== false, // missing → treat as final (old parent)
      level: typeof msg.level === "number" ? msg.level : 0,
    };
  }
  if (obj.action === "error") {
    const msg = parsed as { message?: unknown };
    return { action: "error", message: typeof msg.message === "string" ? msg.message : "Unknown error" };
  }
  return null;
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
    // Preserve existing focus across progressive updates. BFS appends
    // at the end so earlier rows keep their indices — the current
    // focused row doesn't visually jump. Only the very first emission
    // sets initial focus to row 0.
    const isFirstEmit = currentRows.length === 0;
    renderDialogTraceRows(parsed.rows, { preserveFocus: !isFirstEmit });
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

  // Register parent→dialog listener FIRST, then signal ready. The parent
  // computes the trace (Excel.run is not available inside this dialog) and
  // pushes rows via dialog.messageChild once it sees our ready signal.
  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    handleParentMessage,
    () => {
      logTracePerf("t2a addHandler.callback");
      sendToParent({ action: "ready" });
    }
  );
});
