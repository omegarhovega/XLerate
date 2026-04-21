/* global Office */

import {
  buildDefaultFormatSettings,
  FORMAT_SETTINGS_KEY,
  resolveFormatSettings,
  type ResolvedFormatSettings,
} from "../core/formatSettings";
import { resetTextStyleCycleIndex } from "./cycleStateStorage";

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

export function readWorkbookFormatSettings(): ResolvedFormatSettings {
  const raw = Office.context.document.settings.get(FORMAT_SETTINGS_KEY);
  return resolveFormatSettings(raw);
}

export async function saveWorkbookFormatSettings(settings: ResolvedFormatSettings): Promise<void> {
  const defaults = buildDefaultFormatSettings();
  if (JSON.stringify(settings) === JSON.stringify(defaults)) {
    Office.context.document.settings.remove(FORMAT_SETTINGS_KEY);
  } else {
    Office.context.document.settings.set(FORMAT_SETTINGS_KEY, JSON.stringify(settings));
  }

  await saveDocumentSettingsAsync();
  resetTextStyleCycleIndex();
}

export async function resetWorkbookFormatSettings(): Promise<void> {
  await saveWorkbookFormatSettings(buildDefaultFormatSettings());
}
