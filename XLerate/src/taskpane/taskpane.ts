/* global Office */

import "./ribbonActions";
import { clearAutoColorProbeLog, readAutoColorProbeLog } from "../adapters/autoColorProbe";
import {
  getFormatSettingsValidationError,
  resolveFormatSettings,
  type ResolvedFormatSettings,
} from "../core/formatSettings";
import {
  readWorkbookFormatSettings,
  saveWorkbookFormatSettings,
} from "./formatSettingsStore";
import { initSettingsWorkspace } from "./settingsWorkspace";

const SETTINGS_FILE_NAME = "xlerate-settings.json";
const SETTINGS_FILE_ACCEPT = ".json";

type FilePickerAcceptType = {
  description?: string;
  accept: Record<string, string[]>;
};

type OpenFilePickerOptions = {
  excludeAcceptAllOption?: boolean;
  id?: string;
  multiple?: boolean;
  startIn?: "documents" | "downloads";
  types?: FilePickerAcceptType[];
};

type SaveFilePickerOptions = {
  excludeAcceptAllOption?: boolean;
  id?: string;
  suggestedName?: string;
  startIn?: "documents" | "downloads";
  types?: FilePickerAcceptType[];
};

type FileSystemWritableFileStreamLike = {
  write(data: string): Promise<void>;
  close(): Promise<void>;
};

type FileSystemFileHandleLike = {
  createWritable?: () => Promise<FileSystemWritableFileStreamLike>;
  getFile?: () => Promise<File>;
};

type WindowWithFilePickers = Window & {
  showOpenFilePicker?: (options?: OpenFilePickerOptions) => Promise<FileSystemFileHandleLike[]>;
  showSaveFilePicker?: (options?: SaveFilePickerOptions) => Promise<FileSystemFileHandleLike>;
};

type AutoColorDebugHelpers = {
  clearAutoColorProbes: () => void;
  readAutoColorProbes: () => unknown;
};

type WindowWithAutoColorDebug = WindowWithFilePickers & {
  __xlerateDebug?: AutoColorDebugHelpers;
};

function setStatus(message: string): void {
  const target = document.getElementById("status-text");
  if (target) {
    target.textContent = message;
  }
}

function settingsFilePickerTypes(): FilePickerAcceptType[] {
  return [
    {
      description: "XLerate settings",
      accept: {
        "application/json": [SETTINGS_FILE_ACCEPT],
      },
    },
  ];
}

function promptForSettingsFile(): Promise<File> {
  return new Promise((resolve, reject) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = SETTINGS_FILE_ACCEPT;
    input.style.position = "fixed";
    input.style.left = "-9999px";
    input.addEventListener("change", () => {
      const file = input.files?.[0];
      input.remove();
      if (file) {
        resolve(file);
      } else {
        reject(new Error("No settings file selected."));
      }
    });
    document.body.appendChild(input);
    input.click();
  });
}

async function loadSettingsFromFile(): Promise<ResolvedFormatSettings> {
  const pickerWindow = window as WindowWithFilePickers;
  let file: File | null = null;

  if (typeof pickerWindow.showOpenFilePicker === "function") {
    const [handle] = await pickerWindow.showOpenFilePicker({
      excludeAcceptAllOption: true,
      id: "xlerate-settings",
      multiple: false,
      startIn: "documents",
      types: settingsFilePickerTypes(),
    });
    if (!handle?.getFile) {
      throw new Error("Could not open the selected settings file.");
    }
    file = await handle.getFile();
  } else {
    file = await promptForSettingsFile();
  }

  const text = await file.text();
  let parsed: unknown;
  try {
    parsed = JSON.parse(text);
  } catch {
    throw new Error("The selected settings file is not valid JSON.");
  }

  const resolved = resolveFormatSettings(parsed);
  const validationError = getFormatSettingsValidationError(resolved);
  if (validationError) {
    throw new Error(validationError);
  }

  return resolved;
}

async function exportSettingsToFile(settings: ResolvedFormatSettings): Promise<void> {
  const pickerWindow = window as WindowWithFilePickers;
  const text = JSON.stringify(settings, null, 2);

  if (typeof pickerWindow.showSaveFilePicker === "function") {
    const handle = await pickerWindow.showSaveFilePicker({
      excludeAcceptAllOption: true,
      id: "xlerate-settings",
      suggestedName: SETTINGS_FILE_NAME,
      startIn: "documents",
      types: settingsFilePickerTypes(),
    });
    if (!handle?.createWritable) {
      throw new Error("Could not create the settings export file.");
    }

    const writable = await handle.createWritable();
    await writable.write(text);
    await writable.close();
    return;
  }

  const blob = new Blob([text], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = SETTINGS_FILE_NAME;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 0);
}

function installAutoColorProbeHelpers(): void {
  const debugWindow = window as WindowWithAutoColorDebug;
  debugWindow.__xlerateDebug = {
    clearAutoColorProbes: () => {
      clearAutoColorProbeLog();
    },
    readAutoColorProbes: () => {
      const probes = readAutoColorProbeLog();
      // eslint-disable-next-line no-console
      console.log("[xlerate-debug] readAutoColorProbes", probes);
      return probes;
    },
  };
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

  initSettingsWorkspace({
    initialSettings: readWorkbookFormatSettings(),
    loadSavedSettings: loadSettingsFromFile,
    exportSettings: exportSettingsToFile,
    saveSettings: saveWorkbookFormatSettings,
    onStatus: setStatus,
  });
  installAutoColorProbeHelpers();

  setStatus("Ribbon actions are live on the XLerate tab. Workbook settings are ready.");
});
