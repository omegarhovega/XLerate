/* global Excel, Office */

import { parseWorksheetScopedAddress, sanitizeTraceDepth, type TraceDirection } from "../core/traceUtils";

/**
 * Shared dialog-opener used by both the taskpane (temporary dev buttons /
 * future replacement surfaces) and the commands runtime (ribbon button
 * ExecuteFunction actions). Each caller gets its own module-level
 * `activeDialog` because the taskpane and commands runtime are separate
 * JavaScript contexts — imports don't share state between them.
 *
 * No DOM access: the commands runtime has no taskpane DOM, so anything
 * user-facing happens through Excel (cell selection) or through the
 * dialog itself. Status reporting in the taskpane is the taskpane's own
 * responsibility if it chooses to wrap this launcher.
 */

let activeDialog: Office.Dialog | null = null;

/**
 * Best-effort return of keyboard focus to the Excel grid after dialog
 * close. Office.js has no direct API for "give focus to the grid", so
 * this is a ladder of workarounds in descending order of reliability:
 *
 * 1. `worksheet.activate()` + `range.activate()` — activate() (distinct
 *    from .select()) asks Excel to make the worksheet and cell the
 *    active rectangle, which on Desktop typically pulls grid focus.
 *
 * 2. If the above doesn't land focus on the grid, the user has one
 *    cheap recovery: any key press after Esc will go to Excel's grid
 *    because the taskpane iframe has given up its focus claim along
 *    with the dialog. The cell is already correct; only focus is off.
 *
 * Documented as a known limitation on Online / Mac where (1) may not
 * pull focus. Do not expand this ladder without sideload evidence that
 * the additions actually help — each attempt adds code that can fail
 * in its own ways.
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
    // Fire-and-forget: ask Excel to bring the active cell's worksheet
    // into focus, best-effort. See `pullFocusToGrid` for the full
    // rationale and the known Online/Mac limitations.
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

type DialogToParent = { action: "navigate"; address: string } | { action: "close" };

function parseMessage(raw: unknown): DialogToParent | null {
  if (typeof raw !== "string") return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return null;
  }
  if (!parsed || typeof parsed !== "object") return null;
  const obj = parsed as { action?: unknown; address?: unknown };
  if (obj.action === "navigate" && typeof obj.address === "string" && obj.address.length > 0) {
    return { action: "navigate", address: obj.address };
  }
  if (obj.action === "close") {
    return { action: "close" };
  }
  return null;
}

function handleDialogMessage(arg: { message?: string; origin?: string | undefined } | { error: number }): void {
  if (!("message" in arg) || typeof arg.message !== "string") return;
  const parsed = parseMessage(arg.message);
  if (!parsed) return;
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
 * Open the trace dialog for the given direction. Reads the active cell to
 * pass its address to the dialog (so the dialog's row 0 matches the grid's
 * starting cell). Idempotent: if a previous dialog is still open, it is
 * closed first — Office.js permits only one add-in dialog at a time per
 * host, so attempting to open a second throws otherwise.
 *
 * Caller convention in the commands runtime: after awaiting this, call
 * `event.completed()` so Office.js knows the ribbon action has finished.
 */
export async function openTraceDialog(
  direction: TraceDirection,
  options: OpenTraceDialogOptions = {}
): Promise<void> {
  closeActiveDialog();

  const startAddress = await getActiveCellAddress();
  const maxDepth = sanitizeTraceDepth(options.maxDepth);

  const url = new URL("traceDialog.html", window.location.href);
  url.searchParams.set("direction", direction);
  if (startAddress) url.searchParams.set("address", startAddress);
  url.searchParams.set("maxDepth", String(maxDepth));

  const height = typeof options.height === "number" ? options.height : 60;
  const width = typeof options.width === "number" ? options.width : 40;

  return new Promise<void>((resolve) => {
    Office.context.ui.displayDialogAsync(
      url.toString(),
      { height, width, displayInIframe: true },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          // Resolve rather than reject: the caller (command runtime) can't
          // surface an error anyway, and the taskpane caller doesn't need
          // to distinguish.
          resolve();
          return;
        }
        activeDialog = result.value;
        activeDialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          handleDialogMessage
        );
        activeDialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          activeDialog = null;
        });
        resolve();
      }
    );
  });
}
