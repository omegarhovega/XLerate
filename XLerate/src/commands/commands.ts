/* global Office */

import { openTraceDialog } from "../taskpane/traceDialogLauncher";

Office.onReady(() => {
  // Office.js is ready; no eager work required. Ribbon buttons invoke the
  // `Office.actions.associate`d functions below when the user clicks them.
});

/**
 * Ribbon → Trace Precedents (Dialog). Opens the trace dialog pre-populated
 * with the active cell's precedent tree; the dialog lives in a separate
 * window from this commands runtime, and the messageParent→navigate
 * protocol is wired inside `openTraceDialog`.
 *
 * IMPORTANT: every Office.actions.associate'd function MUST call
 * event.completed() before returning (even on error), or Office.js keeps
 * the ribbon button in a "busy" state.
 */
async function runOpenTracePrecedentsDialog(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await openTraceDialog("precedents");
  } finally {
    event.completed();
  }
}

async function runOpenTraceDependentsDialog(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await openTraceDialog("dependents");
  } finally {
    event.completed();
  }
}

Office.actions.associate("openTracePrecedentsDialog", runOpenTracePrecedentsDialog);
Office.actions.associate("openTraceDependentsDialog", runOpenTraceDependentsDialog);
