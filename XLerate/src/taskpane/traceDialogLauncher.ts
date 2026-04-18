/* global Excel, Office */

import { buildTrace, type TraceCellInfo, type TraceRow } from "../core/traceBuilder";
import {
  parseWorksheetScopedAddress,
  sanitizeTraceDepth,
  type TraceDirection,
} from "../core/traceUtils";
import {
  getDirectTraceNeighbors,
  loadTraceCellProperties,
  resolveTraceStartCell,
  snapshotRangeForTrace,
} from "./traceExcelNeighbors";

/**
 * Shared dialog-opener used by both the taskpane (secondary entry
 * buttons) and the commands runtime (ribbon button ExecuteFunction
 * actions). Each caller gets its own module-level `activeDialog`
 * because the taskpane and commands runtime are separate JavaScript
 * contexts — imports don't share state between them.
 *
 * **Trace is computed HERE, in the parent runtime, not in the dialog.**
 * Office dialogs cannot call Excel.run (documented restriction:
 * https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins).
 * The rows are pushed to the dialog via `dialog.messageChild` once the
 * dialog signals it has registered its parent-message handler.
 */

let activeDialog: Office.Dialog | null = null;
/** Stashed until the dialog signals "ready". Then sent via messageChild. */
let pendingRowsPayload: string | null = null;

/**
 * Best-effort return of keyboard focus to the Excel grid after dialog
 * close. Office.js has no direct API for "give focus to the grid", so
 * this is a ladder of workarounds:
 *
 * 1. `worksheet.activate()` + `range.select()` — activate() asks Excel
 *    to make the worksheet the active rectangle, which on Desktop
 *    typically pulls grid focus.
 *
 * If (1) doesn't land focus on the grid (known Online/Mac limitation),
 * the user has one cheap recovery: any keypress after Esc goes to
 * Excel's grid because the taskpane iframe has released focus along
 * with the dialog. The cell is already correct; only focus is off.
 */
async function pullFocusToGrid(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const cell = context.workbook.getActiveCell();
      cell.load("worksheet/name");
      await context.sync();
      cell.worksheet.activate();
      cell.select();
      await context.sync();
    });
  } catch {
    // Non-fatal: worst case, user presses one key/click to resume.
  }
}

function closeActiveDialog(): void {
  if (activeDialog) {
    try {
      activeDialog.close();
    } catch {
      // Already closed; Office.js will null out our reference via the
      // DialogEventReceived handler anyway.
    }
    activeDialog = null;
    pendingRowsPayload = null;
    void pullFocusToGrid();
  }
}

async function getActiveCellAddress(): Promise<string | null> {
  try {
    let address: string | null = null;
    await Excel.run(async (context) => {
      const cell = context.workbook.getActiveCell();
      cell.load("address");
      await context.sync();
      address = cell.address;
    });
    return address;
  } catch {
    return null;
  }
}

async function selectAddressInGrid(address: string): Promise<void> {
  const parsed = parseWorksheetScopedAddress(address);
  if (!parsed) return;
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(parsed.worksheetName);
      sheet.load("isNullObject");
      await context.sync();
      if (sheet.isNullObject) return;
      const target = sheet.getRanges(parsed.rangeAddress);
      target.select();
      await context.sync();
    });
  } catch {
    // Live-nav is fire-and-forget. Surfacing mid-navigation errors would
    // fight user intent; next keystroke has another chance.
  }
}

/**
 * Compute precedents or dependents starting at `startAddress` (or the
 * active cell if null). Runs Excel.run in the parent runtime; all
 * Office.js work stays out of the dialog per the API restriction.
 */
async function computeTrace(
  direction: TraceDirection,
  startAddress: string | null,
  maxDepth: number
): Promise<{ rows: TraceRow[]; startAddress: string; truncated: boolean }> {
  const requires = Office.context.requirements.isSetSupported("ExcelApi", "1.12");
  if (!requires) {
    throw new Error("Trace requires ExcelApi 1.12 or later on this Excel host.");
  }

  let resolvedAddress = "";
  let rows: TraceRow[] = [];
  let truncated = false;

  await Excel.run(async (context) => {
    const rootRange = await resolveTraceStartCell(context, startAddress);
    if (!rootRange) {
      throw new Error("Could not resolve a starting cell for the trace.");
    }
    const root = snapshotRangeForTrace(rootRange);
    resolvedAddress = root.address;

    const getNeighbors = async (info: TraceCellInfo): Promise<TraceCellInfo[]> => {
      const worksheet = context.workbook.worksheets.getItem(info.worksheetName);
      const cell = worksheet.getRangeByIndexes(info.rowIndex, info.columnIndex, 1, 1);
      loadTraceCellProperties(cell);
      const neighbors = await getDirectTraceNeighbors(context, cell, direction);
      return neighbors.map(snapshotRangeForTrace);
    };

    const result = await buildTrace({ root, maxDepth, getNeighbors });
    rows = result.rows;
    truncated = result.truncated;
  });

  return { rows, startAddress: resolvedAddress, truncated };
}

type DialogToParent =
  | { action: "ready" }
  | { action: "navigate"; address: string }
  | { action: "close" };

function parseDialogMessage(raw: unknown): DialogToParent | null {
  if (typeof raw !== "string") return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return null;
  }
  if (!parsed || typeof parsed !== "object") return null;
  const obj = parsed as { action?: unknown; address?: unknown };
  if (obj.action === "ready") return { action: "ready" };
  if (obj.action === "navigate" && typeof obj.address === "string" && obj.address.length > 0) {
    return { action: "navigate", address: obj.address };
  }
  if (obj.action === "close") return { action: "close" };
  return null;
}

function handleDialogMessage(arg: { message?: string; origin?: string | undefined } | { error: number }): void {
  if (!("message" in arg) || typeof arg.message !== "string") return;
  const parsed = parseDialogMessage(arg.message);
  if (!parsed) return;

  if (parsed.action === "ready") {
    // Dialog has registered its parent-message handler; safe to push rows.
    if (activeDialog && pendingRowsPayload) {
      try {
        activeDialog.messageChild(pendingRowsPayload);
      } catch {
        // Unlikely (dialog just signaled ready); swallow.
      }
      pendingRowsPayload = null;
    }
    return;
  }

  if (parsed.action === "navigate") {
    void selectAddressInGrid(parsed.address);
    return;
  }

  if (parsed.action === "close") {
    closeActiveDialog();
  }
}

export type OpenTraceDialogOptions = {
  /** Clamp for BFS depth; sanitized via `sanitizeTraceDepth`. */
  maxDepth?: number;
  /** Dialog height in percent of screen (1-100). Defaults to 60. */
  height?: number;
  /** Dialog width in percent of screen (1-100). Defaults to 40. */
  width?: number;
};

/**
 * Open the trace dialog for the given direction. Computes the trace in
 * the parent runtime (dialogs cannot call Excel.run), stashes the rows,
 * then opens the dialog and pushes rows when the dialog signals ready.
 *
 * Caller convention in the commands runtime: after awaiting this, call
 * `event.completed()` so Office.js marks the ribbon action finished.
 */
export async function openTraceDialog(
  direction: TraceDirection,
  options: OpenTraceDialogOptions = {}
): Promise<void> {
  closeActiveDialog();

  const startAddress = await getActiveCellAddress();
  const maxDepth = sanitizeTraceDepth(options.maxDepth);

  // Compute trace BEFORE opening the dialog. If this throws (bad Excel
  // state, unsupported API), we never open the dialog and the caller can
  // surface the error.
  let computed: { rows: TraceRow[]; startAddress: string; truncated: boolean };
  try {
    computed = await computeTrace(direction, startAddress, maxDepth);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    // Re-throw so the caller's catch (e.g. taskpane guardedRun) can surface
    // it. The dialog doesn't open in this path.
    throw new Error(`Trace failed: ${message}`);
  }

  pendingRowsPayload = JSON.stringify({
    action: "setRows",
    rows: computed.rows,
    direction,
    startAddress: computed.startAddress,
    truncated: computed.truncated,
  });

  const url = new URL("traceDialog.html", window.location.href);
  url.searchParams.set("direction", direction);

  const height = typeof options.height === "number" ? options.height : 60;
  const width = typeof options.width === "number" ? options.width : 40;

  return new Promise<void>((resolve) => {
    Office.context.ui.displayDialogAsync(
      url.toString(),
      { height, width, displayInIframe: true },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          pendingRowsPayload = null;
          resolve();
          return;
        }
        activeDialog = result.value;
        activeDialog.addEventHandler(Office.EventType.DialogMessageReceived, handleDialogMessage);
        activeDialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          activeDialog = null;
          pendingRowsPayload = null;
        });
        resolve();
      }
    );
  });
}
