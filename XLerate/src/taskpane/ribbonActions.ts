/* global Office */

/**
 * Ribbon ExecuteFunction handlers, registered with Office.actions.associate
 * at module-load time. In shared-runtime mode these run in the same JS
 * context as the taskpane — no separate commands iframe, no IPC, no cold
 * start. Importing this module from taskpane.ts is what triggers the
 * registrations.
 */

import { ExcelPortLive } from "../adapters/excelPortLive";
import { resolveFormatSettings, FORMAT_SETTINGS_KEY } from "../core/formatSettings";
import { runAutoColor as runAutoColorService } from "../services/autoColor.service";
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

// Every ribbon handler MUST call event.completed() even on failure, or
// Office leaves the button in a "busy" state.
async function finish(event: Office.AddinCommands.Event, work: () => Promise<void>): Promise<void> {
  try {
    await work();
  } catch (error) {
    // eslint-disable-next-line no-console
    console.error("[XLerate ribbon]", error);
  } finally {
    event.completed();
  }
}

function readFormatSettings(): ReturnType<typeof resolveFormatSettings> {
  const raw = Office.context.document.settings.get(FORMAT_SETTINGS_KEY);
  return resolveFormatSettings(raw);
}

async function runOpenTracePrecedentsDialog(event: Office.AddinCommands.Event): Promise<void> {
  await finish(event, () => openTraceDialog("precedents"));
}

async function runOpenTraceDependentsDialog(event: Office.AddinCommands.Event): Promise<void> {
  await finish(event, () => openTraceDialog("dependents"));
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
  await finish(event, () => {
    const settings = readFormatSettings();
    return runCycleDateFormatService(new ExcelPortLive(), settings.dateFormats);
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
  await finish(event, () => runAutoColorService(new ExcelPortLive()));
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
