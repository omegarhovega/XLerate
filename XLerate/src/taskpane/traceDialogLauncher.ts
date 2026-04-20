/* global Excel, Office */

import { buildTrace, type TraceCellInfo, type TraceRow } from "../core/traceBuilder";
import {
  parseWorksheetScopedAddress,
  sanitizeTraceDepth,
  sanitizeTraceSafetyLimit,
  type TraceDirection,
} from "../core/traceUtils";
import {
  getAllDirectTraceNeighbors,
  resolveTraceStartCell,
  snapshotRangeForTrace,
} from "./traceExcelNeighbors";

/**
 * Shared dialog-opener used by both the taskpane (secondary entry
 * buttons) and the ribbon `Office.actions.associate` handlers in
 * `ribbonActions.ts`. With shared runtime enabled in the manifest,
 * both entry points live in the same taskpane JS context and share
 * this module's `activeDialog`; `openTraceDialog()` calls
 * `closeActiveDialog()` first so a second open from either entry
 * point closes the previous dialog cleanly.
 *
 * **Trace is computed HERE, in the parent runtime, not in the dialog.**
 * Office dialogs cannot call Excel.run (documented restriction:
 * https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins).
 * The rows are pushed to the dialog via `dialog.messageChild` once the
 * dialog signals it has registered its parent-message handler.
 */

let activeDialog: Office.Dialog | null = null;
/**
 * Stashed until the dialog signals "ready". The compute runs
 * serially after the ready signal. A prior version kicked off the
 * compute at t0 in parallel with displayDialogAsync; sideload timing
 * showed Excel's add-in runtime serializes concurrent API calls
 * during dialog spawn (a getActiveCell sync that was 2 ms in serial
 * mode became ~1 s when run during spawn), so the "parallel" path
 * gained no overlap and slightly regressed due to contention. See
 * the revert commit for the measured numbers.
 */
type PendingCompute = { direction: TraceDirection; maxDepth: number; safetyLimit: number };
let pendingCompute: PendingCompute | null = null;
let activeDialogSessionId = 0;

// ---- Perf instrumentation (remove when diagnosis is done). ----
// Logs boundary timestamps so we can identify which segment of the
// open-trace-dialog flow dominates the perceived latency. Uses
// Date.now() because the parent runtime and dialog window have
// separate performance.now() origins; absolute epoch ms lets the
// developer compute deltas across contexts by subtracting the "t0"
// line. Grep `[trace-perf]` in DevTools to filter. Session tag is
// printed once per flow so interleaved traces stay readable.
let tracePerfSession = 0;
function logTracePerf(label: string): void {
  // eslint-disable-next-line no-console
  console.log(`[trace-perf] session=${tracePerfSession} ${label} @ ${Date.now()}`);
}

/**
 * Return keyboard focus to the Excel grid after dialog close.
 *
 * Excel now exposes `Workbook.focus()` on ExcelApiDesktop 1.1, which is
 * a host-level focus handoff rather than a DOM/webview blur heuristic.
 * We still activate/select the current cell so the right location is
 * visible, then call `workbook.focus()` when supported so keyboard
 * events go back to the grid.
 *
 * This must still run AFTER the browser has returned focus to the
 * taskpane iframe — if we run synchronously right after `dialog.close()`,
 * the browser's return happens later and overrides us. The caller defers
 * via `setTimeout` for this reason.
 */
async function pullFocusToGrid(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const cell = workbook.getActiveCell();
      cell.load("worksheet/name");
      await context.sync();
      cell.worksheet.activate();
      cell.select();
      if (Office.context.requirements.isSetSupported("ExcelApiDesktop", "1.1")) {
        workbook.focus();
      }
      await context.sync();
    });
  } catch (err) {
    // Non-fatal: worst case, user presses one key/click to resume.
    // eslint-disable-next-line no-console
    console.log("[trace-focus] pullFocusToGrid error", err);
  }
}

/**
 * Schedule `pullFocusToGrid` after a short delay so the browser has
 * time to complete its own post-close focus return (which puts focus
 * on the taskpane iframe) BEFORE we try to pull it to the grid. The
 * `activeDialog !== null` guard at fire time skips the case where a
 * new dialog has opened in the meantime — we must never yank focus
 * off a live dialog.
 */
function schedulePullFocusToGrid(): void {
  setTimeout(() => {
    if (activeDialog !== null) return;
    void pullFocusToGrid();
  }, 50);
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
    pendingCompute = null;
    schedulePullFocusToGrid();
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

export type TraceProgress = {
  rows: TraceRow[];
  startAddress: string;
  level: number;
  isFinal: boolean;
  truncated: boolean;
};

/**
 * Compute precedents or dependents starting at `startAddress` (or the
 * active cell if null). Runs Excel.run in the parent runtime; all
 * Office.js work stays out of the dialog per the API restriction.
 *
 * When `onProgress` is provided, the callback fires once per BFS
 * level — the first emission is the root alone (before any neighbor
 * lookup runs), each subsequent emission adds the next level's
 * newly discovered cells, and the final emission has `isFinal=true`.
 * The builder awaits each callback so consumers can `messageChild`
 * over IPC and the browser can paint before the next level's sync
 * starts. Without `onProgress`, behavior is identical to the prior
 * all-at-once return.
 */
async function computeTrace(
  direction: TraceDirection,
  startAddress: string | null,
  maxDepth: number,
  safetyLimit: number,
  onProgress?: (progress: TraceProgress) => void | Promise<void>
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

    // Batched per-level neighbor lookup: one sync pair per BFS level
    // regardless of breadth.
    const getAllNeighbors = async (cells: TraceCellInfo[]): Promise<TraceCellInfo[][]> => {
      const ranges = cells.map((info) => {
        const worksheet = context.workbook.worksheets.getItem(info.worksheetName);
        return worksheet.getRangeByIndexes(info.rowIndex, info.columnIndex, 1, 1);
      });
      const neighborLists = await getAllDirectTraceNeighbors(context, ranges, direction);
      return neighborLists.map((list) => list.map(snapshotRangeForTrace));
    };

    const result = await buildTrace({
      root,
      maxDepth,
      maxRows: safetyLimit,
      getAllNeighbors,
      onProgress: onProgress
        ? async (p) => {
            await onProgress({
              rows: p.rows,
              startAddress: resolvedAddress,
              level: p.level,
              isFinal: p.isFinal,
              truncated: p.truncated,
            });
          }
        : undefined,
    });
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

async function computeAndPushRows(
  request: PendingCompute,
  dialog: Office.Dialog,
  dialogSessionId: number
): Promise<void> {
  try {
    logTracePerf("t4 compute.start");
    const startAddress = await getActiveCellAddress();
    logTracePerf("t4a compute.gotActiveCell");

    // Stream each BFS level to the dialog as it completes. Level 0
    // (just the root) lands before any neighbor lookup, so the user
    // sees a filled dialog almost instantly. Subsequent levels paint
    // in as Excel's precedent-graph compute reaches them.
    let firstEmitLogged = false;
    await computeTrace(
      request.direction,
      startAddress,
      request.maxDepth,
      request.safetyLimit,
      async (progress) => {
      if (!firstEmitLogged) {
        logTracePerf(`t5a compute.firstEmit level=${progress.level} rows=${progress.rows.length}`);
        firstEmitLogged = true;
      }
      if (progress.isFinal) {
        logTracePerf(`t5 compute.end rows=${progress.rows.length} truncated=${progress.truncated}`);
      }
      if (activeDialog !== dialog || activeDialogSessionId !== dialogSessionId) return;
      try {
        dialog.messageChild(
          JSON.stringify({
            action: "setRows",
            rows: progress.rows,
            direction: request.direction,
            startAddress: progress.startAddress,
            truncated: progress.truncated,
            isFinal: progress.isFinal,
            level: progress.level,
          })
        );
        if (progress.isFinal) logTracePerf("t6 messageChild.sent (final)");
      } catch {
        // Dialog already gone; skip rest of streaming.
      }
      }
    );
  } catch (error) {
    if (activeDialog !== dialog || activeDialogSessionId !== dialogSessionId) return;
    const message = error instanceof Error ? error.message : String(error);
    try {
      dialog.messageChild(JSON.stringify({ action: "error", message }));
    } catch {
      // Dialog already gone; nothing to report.
    }
  }
}

function handleDialogMessage(
  dialog: Office.Dialog,
  dialogSessionId: number,
  arg: { message?: string; origin?: string | undefined } | { error: number }
): void {
  if (!("message" in arg) || typeof arg.message !== "string") return;
  const parsed = parseDialogMessage(arg.message);
  if (!parsed) return;

  if (activeDialog !== dialog || activeDialogSessionId !== dialogSessionId) return;

  if (parsed.action === "ready") {
    logTracePerf("t3 parent.readyReceived");
    // Dialog has registered its parent-message handler; now run the BFS
    // serially (concurrent Excel.run during spawn gets starved by
    // Excel's internal queue, so there's no parallelism win).
    if (activeDialog && pendingCompute) {
      const request = pendingCompute;
      pendingCompute = null;
      void computeAndPushRows(request, dialog, dialogSessionId);
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
  /** Hard cap on total rendered rows; sanitized via `sanitizeTraceSafetyLimit`. */
  safetyLimit?: number;
  /** Dialog height in percent of screen (1-100). Defaults to 60. */
  height?: number;
  /** Dialog width in percent of screen (1-100). Defaults to 40. */
  width?: number;
};

/**
 * Open the trace dialog for the given direction. Returns as soon as
 * the dialog handle has been issued. The BFS runs serially after the
 * dialog signals `ready` — an earlier parallel-at-t0 design regressed
 * slightly because Excel's add-in runtime serializes concurrent API
 * calls during displayDialogAsync processing.
 *
 * The dialog cannot call `Excel.run` itself (documented restriction on
 * the Office Dialog API); all host-document work stays on this side of
 * the boundary.
 */
export async function openTraceDialog(
  direction: TraceDirection,
  options: OpenTraceDialogOptions = {}
): Promise<void> {
  const nextDialogSessionId = activeDialogSessionId + 1;
  tracePerfSession = Date.now();
  logTracePerf(`t0 click direction=${direction}`);
  closeActiveDialog();

  pendingCompute = {
    direction,
    maxDepth: sanitizeTraceDepth(options.maxDepth),
    safetyLimit: sanitizeTraceSafetyLimit(options.safetyLimit),
  };

  const url = new URL("traceDialog.html", window.location.href);
  url.searchParams.set("direction", direction);
  url.searchParams.set("perfSession", String(tracePerfSession));

  const height = typeof options.height === "number" ? options.height : 60;
  const width = typeof options.width === "number" ? options.width : 40;

  return new Promise<void>((resolve) => {
    Office.context.ui.displayDialogAsync(
      url.toString(),
      { height, width, displayInIframe: true },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          pendingCompute = null;
          resolve();
          return;
        }
        logTracePerf("t1 displayDialogAsync.callback");
        const thisDialog = result.value;
        activeDialog = thisDialog;
        activeDialogSessionId = nextDialogSessionId;
        thisDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) =>
          handleDialogMessage(thisDialog, nextDialogSessionId, arg)
        );
        thisDialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          // Cleanup path for user-initiated close (X button, system close).
          // closeActiveDialog() handles both the ref-nulling and the focus
          // pull for the Esc-via-messageParent path; this handler covers
          // the case where the dialog closed without us asking.
          // The identity guard skips the case where a newer openTraceDialog
          // superseded this dialog — in that case closeActiveDialog has
          // already run and reassigned activeDialog to the newer instance.
          if (activeDialog !== thisDialog || activeDialogSessionId !== nextDialogSessionId) return;
          activeDialog = null;
          pendingCompute = null;
          schedulePullFocusToGrid();
        });
        resolve();
      }
    );
  });
}
