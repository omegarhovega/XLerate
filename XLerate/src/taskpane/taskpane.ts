/* global Excel, Office */

import "./ribbonActions";
import { VALUE_ERROR } from "../core/cagr";
import { FORMAT_SETTINGS_KEY, resolveFormatSettings, type ResolvedFormatSettings } from "../core/formatSettings";
import { runCagrCalculator } from "../services/cagr.service";
import { resetTextStyleCycleIndex } from "./cycleStateStorage";

type CellValue = string | number | boolean | null;

const FORMAT_SETTINGS_EDITOR_ID = "format-settings-json";

function setStatus(message: string): void {
  const target = document.getElementById("status-text");
  if (target) {
    target.textContent = message;
  }
}

function setCagrResult(message: string): void {
  const target = document.getElementById("cagr-result");
  if (target) {
    target.textContent = message;
  }
}

function getFormatSettingsEditor(): HTMLTextAreaElement | null {
  const node = document.getElementById(FORMAT_SETTINGS_EDITOR_ID);
  return node instanceof HTMLTextAreaElement ? node : null;
}

function setFormatSettingsEditorText(value: string): void {
  const editor = getFormatSettingsEditor();
  if (editor) {
    editor.value = value;
  }
}

function getFormatSettingsEditorText(): string | null {
  const editor = getFormatSettingsEditor();
  return editor ? editor.value : null;
}

function stringifyFormatSettings(settings: ResolvedFormatSettings): string {
  return JSON.stringify(settings, null, 2);
}

// Office.context.document.settings.saveAsync breaks the native Excel undo
// chain on Desktop when used in the same click as cell mutations, but it is
// the right persistence mechanism for settings-editor actions that only
// modify workbook settings.
function saveDocumentSettingsAsync(): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

function readResolvedFormatSettings(): ResolvedFormatSettings {
  const raw = Office.context.document.settings.get(FORMAT_SETTINGS_KEY);
  return resolveFormatSettings(raw);
}

async function clearFormatSettingsAndCycleState(): Promise<void> {
  Office.context.document.settings.remove(FORMAT_SETTINGS_KEY);
  await saveDocumentSettingsAsync();
  resetTextStyleCycleIndex();
}

async function writeFormatSettingsAndResetCycleState(settings: ResolvedFormatSettings): Promise<void> {
  Office.context.document.settings.set(FORMAT_SETTINGS_KEY, JSON.stringify(settings));
  await saveDocumentSettingsAsync();
  resetTextStyleCycleIndex();
}

async function runResetFormatSettings(): Promise<void> {
  await clearFormatSettingsAndCycleState();
  setFormatSettingsEditorText(stringifyFormatSettings(resolveFormatSettings(undefined)));
  setStatus("Format settings reset to defaults.");
}

async function runLoadFormatSettingsEditor(): Promise<void> {
  const settings = readResolvedFormatSettings();
  setFormatSettingsEditorText(stringifyFormatSettings(settings));
  setStatus("Loaded saved format settings into editor.");
}

async function runLoadDefaultFormatSettingsEditor(): Promise<void> {
  const defaults = resolveFormatSettings(undefined);
  setFormatSettingsEditorText(stringifyFormatSettings(defaults));
  setStatus("Loaded default format settings into editor.");
}

async function runSaveFormatSettingsFromEditor(): Promise<void> {
  const raw = getFormatSettingsEditorText();
  if (raw === null) {
    setStatus("Format settings editor not found.");
    return;
  }

  const trimmed = raw.trim();
  if (trimmed.length === 0) {
    setStatus("Format settings editor is empty.");
    return;
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(trimmed);
  } catch {
    setStatus("Format settings JSON is invalid.");
    return;
  }

  const resolved = resolveFormatSettings(parsed);
  await writeFormatSettingsAndResetCycleState(resolved);
  setFormatSettingsEditorText(stringifyFormatSettings(resolved));
  setStatus("Format settings saved. Cycle state reset.");
}

async function runCagr(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["values", "rowCount", "columnCount"]);
    await context.sync();

    const values: number[] = [];
    for (let r = 0; r < range.rowCount; r += 1) {
      for (let c = 0; c < range.columnCount; c += 1) {
        const raw = range.values[r][c] as CellValue;
        const parsed = typeof raw === "number" ? raw : Number(raw);
        if (!Number.isFinite(parsed)) {
          setCagrResult(VALUE_ERROR);
          setStatus("CAGR failed: selected range includes non-numeric values.");
          return;
        }
        values.push(parsed);
      }
    }

    const result = runCagrCalculator(values);
    if (result === VALUE_ERROR) {
      setCagrResult(VALUE_ERROR);
      setStatus("CAGR returned #VALUE! based on baseline rules.");
      return;
    }

    const formatted = result.toFixed(10);
    setCagrResult(formatted);
    setStatus("CAGR calculated successfully.");
  });
}

async function guardedRun(action: () => Promise<void>): Promise<void> {
  try {
    await action();
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`Error: ${message}`);
    // eslint-disable-next-line no-console
    console.error(error);
  }
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    return;
  }

  const sideloadMessage = document.getElementById("sideload-msg");
  const appBody = document.getElementById("app-body");
  if (sideloadMessage && appBody) {
    sideloadMessage.style.display = "none";
    appBody.style.display = "block";
  }

  setFormatSettingsEditorText(stringifyFormatSettings(readResolvedFormatSettings()));

  document
    .getElementById("load-format-settings")
    ?.addEventListener("click", () => guardedRun(runLoadFormatSettingsEditor));
  document
    .getElementById("load-default-format-settings")
    ?.addEventListener("click", () => guardedRun(runLoadDefaultFormatSettingsEditor));
  document
    .getElementById("save-format-settings")
    ?.addEventListener("click", () => guardedRun(runSaveFormatSettingsFromEditor));
  document
    .getElementById("run-reset-format-settings")
    ?.addEventListener("click", () => guardedRun(runResetFormatSettings));
  document.getElementById("run-cagr")?.addEventListener("click", () => guardedRun(runCagr));
});
