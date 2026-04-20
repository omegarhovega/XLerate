/* global Office */

/**
 * Ribbon ExecuteFunction handlers, registered with Office.actions.associate
 * at module-load time. In shared-runtime mode these run in the same JS
 * context as the taskpane — no separate commands iframe, no IPC, no cold
 * start. Importing this module from taskpane.ts is what triggers the
 * registrations.
 */

import { ExcelPortLive } from "../adapters/excelPortLive";
import {
  autoColorProbe,
  beginAutoColorProbeSession,
  endAutoColorProbeSession,
} from "../adapters/autoColorProbe";
import { computeNextDateFormat, hasMixedDateFormats } from "../core/dateFormatCycle";
import { resolveFormatSettings, FORMAT_SETTINGS_KEY } from "../core/formatSettings";
import { runAutoColor as runAutoColorService } from "../services/autoColor.service";
import { runInsertCagr as runInsertCagrService } from "../services/cagr.service";
import { runCycleCellFormat as runCycleCellFormatService } from "../services/cycleCellFormat.service";
import { runCycleDateFormat as runCycleDateFormatService } from "../services/cycleDateFormat.service";
import { runCycleNumberFormat as runCycleNumberFormatService } from "../services/cycleNumberFormat.service";
import { runCycleTextStyle as runCycleTextStyleService } from "../services/cycleTextStyle.service";
import { runSwitchSign as runSwitchSignService } from "../services/switchSign.service";
import { readTextStyleCycleIndex, writeTextStyleCycleIndex } from "./cycleStateStorage";
import { openTraceDialog } from "./traceDialogLauncher";
import {
  applyFormulaConsistencyAction,
  applySmartFillRightAction,
} from "./workbookActions";

// DIAGNOSTIC (autocolor hang on Ctrl+A): catch anything that escapes finish()
// or fires after a handler returns, so we always leave a trace in the console.
if (typeof window !== "undefined") {
  window.addEventListener("unhandledrejection", (ev) => {
    // eslint-disable-next-line no-console
    console.warn("[xlerate-autocolor-probe] unhandledrejection", { reason: String(ev.reason) });
  });
  window.addEventListener("error", (ev) => {
    // eslint-disable-next-line no-console
    console.warn("[xlerate-autocolor-probe] window.error", { message: ev.message });
  });
  // eslint-disable-next-line no-console
  console.log("[xlerate-autocolor-probe] module-loaded", { t: performance.now() });
}

// Every ribbon handler MUST call event.completed() even on failure, or
// Office leaves the button in a "busy" state.
async function finish(event: Office.AddinCommands.Event, work: () => Promise<void>): Promise<void> {
  try {
    await work();
  } catch (error) {
    // eslint-disable-next-line no-console
    console.error("[XLerate ribbon]", error);
    // eslint-disable-next-line no-console
    console.error("[xlerate-autocolor-probe] 99 finish-caught", {
      error: String(error),
      stack: (error as Error)?.stack,
    });
  } finally {
    event.completed();
  }
}

function readFormatSettings(): ReturnType<typeof resolveFormatSettings> {
  const raw = Office.context.document.settings.get(FORMAT_SETTINGS_KEY);
  return resolveFormatSettings(raw);
}

async function runOpenTracePrecedentsDialog(event: Office.AddinCommands.Event): Promise<void> {
  await finish(event, () => {
    const settings = readFormatSettings();
    return openTraceDialog("precedents", settings.trace);
  });
}

async function runOpenTraceDependentsDialog(event: Office.AddinCommands.Event): Promise<void> {
  await finish(event, () => {
    const settings = readFormatSettings();
    return openTraceDialog("dependents", settings.trace);
  });
}

async function runSwitchSignFromRibbon(event: Office.AddinCommands.Event): Promise<void> {
  await finish(event, () => runSwitchSignService(new ExcelPortLive()));
}

async function runSmartFillRightFromRibbon(event: Office.AddinCommands.Event): Promise<void> {
  await finish(event, async () => {
    await applySmartFillRightAction();
  });
}

async function runFormulaConsistencyFromRibbon(event: Office.AddinCommands.Event): Promise<void> {
  await finish(event, async () => {
    await applyFormulaConsistencyAction();
  });
}

async function runCycleNumberFormatFromRibbon(event: Office.AddinCommands.Event): Promise<void> {
  await finish(event, () => {
    const settings = readFormatSettings();
    return runCycleNumberFormatService(new ExcelPortLive(), settings.numberFormats);
  });
}

async function runCycleDateFormatFromRibbon(event: Office.AddinCommands.Event): Promise<void> {
  await finish(event, async () => {
    const settings = readFormatSettings();
    const port = new ExcelPortLive();
    const before = await port.getSelectionFormatting();
    const beforeFormats = before.map((snap) => ({
      address: `${snap.address.sheet}!R${snap.address.row}C${snap.address.col}`,
      numberFormat: snap.numberFormat,
    }));
    const mixed = hasMixedDateFormats(before.map((snap) => snap.numberFormat));
    const nextFormat =
      before.length > 0
        ? computeNextDateFormat(before[0].numberFormat, mixed, settings.dateFormats)
        : null;

    // eslint-disable-next-line no-console
    console.log("[xlerate-date-debug] before", {
      configuredDateFormats: settings.dateFormats.map((item) => ({
        name: item.name,
        formatCode: item.formatCode,
      })),
      selection: beforeFormats,
      mixedSelection: mixed,
      computedNextFormat: nextFormat,
    });

    await runCycleDateFormatService(port, settings.dateFormats);

    const after = await port.getSelectionFormatting();
    // eslint-disable-next-line no-console
    console.log("[xlerate-date-debug] after", {
      selection: after.map((snap) => ({
        address: `${snap.address.sheet}!R${snap.address.row}C${snap.address.col}`,
        numberFormat: snap.numberFormat,
      })),
    });
  });
}

async function runCycleCellFormatFromRibbon(event: Office.AddinCommands.Event): Promise<void> {
  await finish(event, () => {
    const settings = readFormatSettings();
    return runCycleCellFormatService(new ExcelPortLive(), settings.cellFormats);
  });
}

async function runCycleTextStyleFromRibbon(event: Office.AddinCommands.Event): Promise<void> {
  await finish(event, async () => {
    const settings = readFormatSettings();
    const { index } = await runCycleTextStyleService(
      new ExcelPortLive(),
      readTextStyleCycleIndex(),
      settings.textStyles
    );
    writeTextStyleCycleIndex(index);
  });
}

async function runAutoColorFromRibbon(event: Office.AddinCommands.Event): Promise<void> {
  beginAutoColorProbeSession();
  autoColorProbe("01 handler-enter");
  try {
    await finish(event, async () => {
      autoColorProbe("02 finish-inner-enter");
      const settings = readFormatSettings();
      autoColorProbe("03 settings-read", {
        paletteKeys: Object.keys(settings.autoColorPalette ?? {}),
      });
      autoColorProbe("04 before-runAutoColorService");
      await runAutoColorService(new ExcelPortLive(), settings.autoColorPalette);
      autoColorProbe("05 after-runAutoColorService");
    });
    endAutoColorProbeSession();
  } catch (error) {
    autoColorProbe("99 ribbon-error", {
      message: error instanceof Error ? error.message : String(error),
    });
    throw error;
  } finally {
    autoColorProbe("06 handler-exit");
  }
}

async function runInsertCagrFromRibbon(event: Office.AddinCommands.Event): Promise<void> {
  await finish(event, () => runInsertCagrService(new ExcelPortLive()).then(() => undefined));
}

Office.actions.associate("openTracePrecedentsDialog", runOpenTracePrecedentsDialog);
Office.actions.associate("openTraceDependentsDialog", runOpenTraceDependentsDialog);
Office.actions.associate("runSwitchSign", runSwitchSignFromRibbon);
Office.actions.associate("runSmartFillRight", runSmartFillRightFromRibbon);
Office.actions.associate("runFormulaConsistency", runFormulaConsistencyFromRibbon);
Office.actions.associate("runCycleNumberFormat", runCycleNumberFormatFromRibbon);
Office.actions.associate("runCycleDateFormat", runCycleDateFormatFromRibbon);
Office.actions.associate("runCycleCellFormat", runCycleCellFormatFromRibbon);
Office.actions.associate("runCycleTextStyle", runCycleTextStyleFromRibbon);
Office.actions.associate("runAutoColor", runAutoColorFromRibbon);
Office.actions.associate("runInsertCagr", runInsertCagrFromRibbon);
