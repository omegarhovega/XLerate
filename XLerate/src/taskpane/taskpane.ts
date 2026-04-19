/* global Excel, Office */
import { runAutoColor as runAutoColorService } from "../services/autoColor.service";
import { VALUE_ERROR } from "../core/cagr";
import { runCagrCalculator } from "../services/cagr.service";
import {
  computeNextCellFormat,
  type CellFormatDefinition,
  type SelectionCellFormatState
} from "../core/cellFormatCycle";
import { FORMAT_SETTINGS_KEY, resolveFormatSettings, type ResolvedFormatSettings } from "../core/formatSettings";
import { runCycleNumberFormat as runCycleNumberFormatService } from "../services/cycleNumberFormat.service";
import { runCycleDateFormat as runCycleDateFormatService } from "../services/cycleDateFormat.service";
import { runCycleCellFormat as runCycleCellFormatService } from "../services/cycleCellFormat.service";
import { ExcelPortLive } from "../adapters/excelPortLive";
import { runErrorWrap as runErrorWrapService } from "../services/errorWrap.service";
import { runSwitchSign as runSwitchSignService } from "../services/switchSign.service";
import { runCycleTextStyle as runCycleTextStyleService } from "../services/cycleTextStyle.service";
import {
  parseWorksheetScopedAddress,
  sanitizeTraceDepth,
  type TraceDirection,
} from "../core/traceUtils";
import { buildTrace, type TraceCellInfo, type TraceRow } from "../core/traceBuilder";
import {
  getAllDirectTraceNeighbors,
  loadTraceCellProperties,
  snapshotRangeForTrace,
} from "./traceExcelNeighbors";
import { openTraceDialog } from "./traceDialogLauncher";
import {
  applyFormulaConsistencyAction,
  applySmartFillRightAction,
} from "./workbookActions";
import {
  readTextStyleCycleIndex,
  resetTextStyleCycleIndex,
  writeTextStyleCycleIndex,
} from "./cycleStateStorage";

type CellValue = string | number | boolean | null;
const FORMAT_SETTINGS_EDITOR_ID = "format-settings-json";
const TRACE_MAX_DEPTH_INPUT_ID = "trace-max-depth";
const TRACE_RESULTS_TBODY_ID = "trace-results-body";
const BORDER_SIDE_ITEMS = [
  "EdgeLeft",
  "EdgeTop",
  "EdgeBottom",
  "EdgeRight"
] as const;
type BorderSideItem = (typeof BORDER_SIDE_ITEMS)[number];

// Text-style cycle index lives in src/taskpane/cycleStateStorage.ts
// (window.localStorage under a versioned key) so the ribbon button's
// commands runtime and the taskpane stay in sync. Saving via
// Office.context.document.settings.saveAsync would break the Excel
// undo chain on Desktop; localStorage is session-scoped and doesn't
// touch the workbook. Cell-format / date-format / number-format
// cycles infer position from the current cell and need no index.

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

function getTraceMaxDepthInputValue(): number {
  const input = document.getElementById(TRACE_MAX_DEPTH_INPUT_ID);
  if (!(input instanceof HTMLInputElement)) {
    return sanitizeTraceDepth(undefined);
  }
  return sanitizeTraceDepth(Number(input.value));
}

// Keyboard-navigation state for the trace results table.
// `currentTraceRows` mirrors what's rendered so key handlers can look up
// addresses without re-reading the DOM. `currentTraceFocusIndex` tracks the
// row that arrow keys operate on — null means "no list focus" (post-Esc or
// empty results).
let currentTraceRows: TraceRow[] = [];
let currentTraceFocusIndex: number | null = null;
const TRACE_ROW_FOCUSED_CLASS = "trace-row-focused";
const TRACE_ROW_INDEX_ATTR = "data-trace-index";

function getTraceRowElements(): HTMLTableRowElement[] {
  const body = document.getElementById(TRACE_RESULTS_TBODY_ID);
  if (!(body instanceof HTMLTableSectionElement)) return [];
  return Array.from(body.querySelectorAll<HTMLTableRowElement>(`tr[${TRACE_ROW_INDEX_ATTR}]`));
}

function focusTraceRow(targetIndex: number): void {
  const elements = getTraceRowElements();
  if (elements.length === 0) {
    currentTraceFocusIndex = null;
    return;
  }
  const clamped = Math.max(0, Math.min(elements.length - 1, targetIndex));
  elements.forEach((el, i) => {
    el.classList.toggle(TRACE_ROW_FOCUSED_CLASS, i === clamped);
    el.setAttribute("aria-selected", i === clamped ? "true" : "false");
  });
  currentTraceFocusIndex = clamped;
  const target = elements[clamped];
  // .focus() is silent if the taskpane iframe doesn't currently have
  // document focus (e.g. right after running trace from the Excel grid).
  // That's OK — the .trace-row-focused class still renders the ring so the
  // user can see where Enter will land once they click/Alt+F6 into the pane.
  target.focus();
  target.scrollIntoView({ block: "nearest" });
}

function clearTraceFocus(): void {
  currentTraceFocusIndex = null;
  for (const el of getTraceRowElements()) {
    el.classList.remove(TRACE_ROW_FOCUSED_CLASS);
    el.setAttribute("aria-selected", "false");
  }
}

function handleTraceKeydown(event: KeyboardEvent): void {
  if (currentTraceRows.length === 0) return;
  const focusIndex = currentTraceFocusIndex ?? 0;
  switch (event.key) {
    case "ArrowDown":
      event.preventDefault();
      focusTraceRow(focusIndex + 1);
      break;
    case "ArrowUp":
      event.preventDefault();
      focusTraceRow(focusIndex - 1);
      break;
    case "Home":
      event.preventDefault();
      focusTraceRow(0);
      break;
    case "End":
      event.preventDefault();
      focusTraceRow(currentTraceRows.length - 1);
      break;
    case "Enter": {
      event.preventDefault();
      const row = currentTraceRows[focusIndex];
      if (row) {
        void guardedRun(() => runSelectTraceAddress(row.address));
      }
      break;
    }
    case "Escape":
      event.preventDefault();
      if (document.activeElement instanceof HTMLElement) {
        document.activeElement.blur();
      }
      clearTraceFocus();
      break;
    default:
      // Don't preventDefault — let unrelated keys (Tab, shortcuts) through.
      break;
  }
}

function handleTraceFocusIn(event: FocusEvent): void {
  // Fired when a row (or an element inside it, e.g. the link button) gains
  // focus. Sync our in-memory index with whichever row is now the focus target
  // so subsequent arrow keys continue from there, not from row 0.
  const eventTarget = event.target;
  if (!(eventTarget instanceof Element)) return;
  const row = eventTarget.closest(`tr[${TRACE_ROW_INDEX_ATTR}]`);
  if (!(row instanceof HTMLTableRowElement)) return;
  const raw = row.getAttribute(TRACE_ROW_INDEX_ATTR);
  const idx = raw === null ? NaN : Number(raw);
  if (!Number.isInteger(idx) || idx < 0 || idx >= currentTraceRows.length) return;
  if (currentTraceFocusIndex !== idx) {
    focusTraceRow(idx);
  }
}

function wireTraceListKeyboard(): void {
  const body = document.getElementById(TRACE_RESULTS_TBODY_ID);
  if (!(body instanceof HTMLTableSectionElement)) return;
  body.addEventListener("keydown", handleTraceKeydown);
  body.addEventListener("focusin", handleTraceFocusIn);
}

function renderTraceRows(rows: TraceRow[]): void {
  const body = document.getElementById(TRACE_RESULTS_TBODY_ID);
  if (!(body instanceof HTMLTableSectionElement)) {
    return;
  }

  body.textContent = "";
  currentTraceRows = rows;
  currentTraceFocusIndex = null;

  if (rows.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 4;
    cell.textContent = "No trace results.";
    row.appendChild(cell);
    body.appendChild(row);
    return;
  }

  rows.forEach((item, index) => {
    const row = document.createElement("tr");
    row.className = "trace-row-clickable";
    row.setAttribute("role", "option");
    row.setAttribute("tabindex", "0");
    row.setAttribute(TRACE_ROW_INDEX_ATTR, String(index));
    row.setAttribute("aria-selected", "false");

    const level = document.createElement("td");
    const address = document.createElement("td");
    const value = document.createElement("td");
    const formula = document.createElement("td");
    const addressBtn = document.createElement("button");

    level.textContent = String(item.level);
    addressBtn.type = "button";
    addressBtn.className = "trace-link-btn";
    addressBtn.textContent = item.address;
    // The row itself is the keyboard target; the inner button shouldn't steal
    // tab focus away from the listbox navigation pattern.
    addressBtn.tabIndex = -1;
    address.appendChild(addressBtn);
    value.textContent = item.value;
    formula.textContent = item.formula;

    row.addEventListener("click", () => {
      void guardedRun(() => runSelectTraceAddress(item.address));
    });
    addressBtn.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      void guardedRun(() => runSelectTraceAddress(item.address));
    });

    row.append(level, address, value, formula);
    body.appendChild(row);
  });

  // Auto-select row 0 so the user sees where Enter will land and can arrow
  // from there. Programmatic .focus() is a no-op when the taskpane lacks
  // document focus, but the class-driven ring still renders.
  focusTraceRow(0);
}

function makeFormatMatrix(rowCount: number, columnCount: number, formatCode: string): string[][] {
  const matrix: string[][] = [];
  for (let r = 0; r < rowCount; r += 1) {
    const row: string[] = [];
    for (let c = 0; c < columnCount; c += 1) {
      row.push(formatCode);
    }
    matrix.push(row);
  }
  return matrix;
}

function flattenFormatMatrix(matrix: unknown[][]): string[] {
  const values: string[] = [];
  for (const row of matrix) {
    for (const item of row) {
      values.push(typeof item === "string" ? item : String(item ?? ""));
    }
  }
  return values;
}

function createCellPropertiesMatrix(rowCount: number, columnCount: number): Excel.SettableCellProperties[][] {
  const matrix: Excel.SettableCellProperties[][] = [];
  for (let r = 0; r < rowCount; r += 1) {
    const row: Excel.SettableCellProperties[] = [];
    for (let c = 0; c < columnCount; c += 1) {
      row.push({});
    }
    matrix.push(row);
  }
  return matrix;
}

function toUnderlinedBoolean(value: unknown): boolean | null {
  if (typeof value !== "string") {
    return null;
  }
  return value.toLowerCase() !== "none";
}

function toBorderColor(style: unknown, color: unknown): string | null {
  return typeof style === "string" && style.toLowerCase() !== "none" && typeof color === "string" ? color : null;
}

function readBorderState(border: Excel.RangeBorder): { style: string | null; color: string | null } {
  const style = typeof border.style === "string" ? border.style : null;
  const color = toBorderColor(style, border.color);
  return { style, color };
}

function buildSelectionCellFormatState(
  range: Excel.Range,
  borders: Record<BorderSideItem, Excel.RangeBorder>
): SelectionCellFormatState {
  const left = readBorderState(borders.EdgeLeft);
  const top = readBorderState(borders.EdgeTop);
  const bottom = readBorderState(borders.EdgeBottom);
  const right = readBorderState(borders.EdgeRight);

  return {
    fillPattern: typeof range.format.fill.pattern === "string" ? range.format.fill.pattern : null,
    fillColor: typeof range.format.fill.color === "string" ? range.format.fill.color : null,
    fontColor: typeof range.format.font.color === "string" ? range.format.font.color : null,
    fontBold: typeof range.format.font.bold === "boolean" ? range.format.font.bold : null,
    fontItalic: typeof range.format.font.italic === "boolean" ? range.format.font.italic : null,
    fontUnderline: toUnderlinedBoolean(range.format.font.underline),
    fontStrikethrough:
      typeof range.format.font.strikethrough === "boolean" ? range.format.font.strikethrough : null,
    edgeLeftStyle: left.style,
    edgeTopStyle: top.style,
    edgeBottomStyle: bottom.style,
    edgeRightStyle: right.style,
    edgeLeftColor: left.color,
    edgeTopColor: top.color,
    edgeBottomColor: bottom.color,
    edgeRightColor: right.color
  };
}

function setRangeBorder(range: Excel.Range, side: Excel.BorderIndex, style: string, color: string): void {
  const border = range.format.borders.getItem(side);
  border.style = style as Excel.BorderLineStyle;
  if (style.toLowerCase() !== "none") {
    border.color = color;
  }
}

function applyCellFormatToRange(range: Excel.Range, format: CellFormatDefinition): void {
  range.format.fill.pattern = format.fillPattern;
  range.format.fill.color = format.fillColor;
  range.format.font.color = format.fontColor;
  range.format.font.bold = format.fontBold;
  range.format.font.italic = format.fontItalic;
  range.format.font.underline = format.fontUnderline ? "Single" : "None";
  range.format.font.strikethrough = format.fontStrikethrough;

  setRangeBorder(range, Excel.BorderIndex.edgeLeft, format.borderStyle, format.borderColor);
  setRangeBorder(range, Excel.BorderIndex.edgeTop, format.borderStyle, format.borderColor);
  setRangeBorder(range, Excel.BorderIndex.edgeBottom, format.borderStyle, format.borderColor);
  setRangeBorder(range, Excel.BorderIndex.edgeRight, format.borderStyle, format.borderColor);

  if (range.rowCount * range.columnCount > 1) {
    setRangeBorder(range, Excel.BorderIndex.insideHorizontal, format.borderStyle, format.borderColor);
    setRangeBorder(range, Excel.BorderIndex.insideVertical, format.borderStyle, format.borderColor);
  }
}

// Office.context.document.settings.saveAsync breaks the native Excel undo
// chain on Desktop: any click on the sheet after a sequence of
// Excel.run mutations + saveAsync flushes the undo boundary so Ctrl+Z no
// longer reverses the mutations. We therefore use saveAsync ONLY for
// settings-editor actions that the user explicitly commits (Save / Reset
// Format Settings) and NEVER in handlers that also mutate cells in the
// same click. The Format Settings editor is the only feature that
// persists to document settings.
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


async function runTrace(direction: TraceDirection): Promise<void> {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.12")) {
    setStatus("Trace requires ExcelApi 1.12 or later on this Excel host.");
    return;
  }

  await Excel.run(async (context) => {
    const rootRange = context.workbook.getActiveCell();
    loadTraceCellProperties(rootRange);
    await context.sync();

    const root = snapshotRangeForTrace(rootRange);
    const maxDepth = getTraceMaxDepthInputValue();

    // Batched per-level neighbor lookup: one sync pair per BFS level
    // regardless of breadth. See getAllDirectTraceNeighbors for the
    // round-trip accounting.
    const getAllNeighbors = async (cells: TraceCellInfo[]): Promise<TraceCellInfo[][]> => {
      const ranges = cells.map((info) => {
        const worksheet = context.workbook.worksheets.getItem(info.worksheetName);
        return worksheet.getRangeByIndexes(info.rowIndex, info.columnIndex, 1, 1);
      });
      const neighborLists = await getAllDirectTraceNeighbors(context, ranges, direction);
      return neighborLists.map((list) => list.map(snapshotRangeForTrace));
    };

    const { rows, truncated } = await buildTrace({ root, maxDepth, getAllNeighbors });

    renderTraceRows(rows);
    setStatus(
      `Trace ${direction} complete on ${root.address}: ${rows.length} cells${truncated ? " (truncated)." : "."}`
    );
  });
}

async function runTracePrecedents(): Promise<void> {
  await runTrace("precedents");
}

async function runTraceDependents(): Promise<void> {
  await runTrace("dependents");
}

async function openTraceDialogFromTaskpane(direction: TraceDirection): Promise<void> {
  await openTraceDialog(direction, { maxDepth: getTraceMaxDepthInputValue() });
  setStatus(`Trace ${direction} dialog opened.`);
}

async function runSelectTraceAddress(fullAddress: string): Promise<void> {
  const parsed = parseWorksheetScopedAddress(fullAddress);
  if (!parsed) {
    setStatus(`Unable to parse trace target "${fullAddress}".`);
    return;
  }

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(parsed.worksheetName);
    sheet.load("isNullObject,name");
    await context.sync();

    if (sheet.isNullObject) {
      setStatus(`Trace target sheet "${parsed.worksheetName}" no longer exists.`);
      return;
    }

    const target = sheet.getRanges(parsed.rangeAddress);
    target.select();
    await context.sync();
    setStatus(`Selected ${fullAddress}.`);
  });
}

async function runSwitchSign(): Promise<void> {
  await runSwitchSignService(new ExcelPortLive());
  setStatus("Switch Sign applied.");
}

async function runCycleNumberFormat(): Promise<void> {
  const formatSettings = readResolvedFormatSettings();
  const configuredFormats = formatSettings.numberFormats;
  await runCycleNumberFormatService(new ExcelPortLive(), configuredFormats);
  setStatus("Cycled number format.");
}

async function runCycleDateFormat(): Promise<void> {
  const formatSettings = readResolvedFormatSettings();
  await runCycleDateFormatService(new ExcelPortLive(), formatSettings.dateFormats);
  setStatus("Cycled date format.");
}

async function runCycleCellFormat(): Promise<void> {
  const formatSettings = readResolvedFormatSettings();
  await runCycleCellFormatService(new ExcelPortLive(), formatSettings.cellFormats);
  setStatus("Cycled cell format.");
}

async function runCycleTextStyle(): Promise<void> {
  const formatSettings = readResolvedFormatSettings();
  const { index } = await runCycleTextStyleService(
    new ExcelPortLive(),
    readTextStyleCycleIndex(),
    formatSettings.textStyles,
  );
  writeTextStyleCycleIndex(index);
  setStatus("Cycled text style.");
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

async function runAutoColor(): Promise<void> {
  await runAutoColorService(new ExcelPortLive());
  setStatus("Auto-color applied.");
}

async function runErrorWrap(): Promise<void> {
  const fallbackInput = (document.getElementById("error-value") as HTMLInputElement | null)?.value?.trim() || "NA()";
  await runErrorWrapService(new ExcelPortLive(), fallbackInput);
  setStatus(`Error Wrap applied with fallback "${fallbackInput}".`);
}

async function runSmartFillRight(): Promise<void> {
  const result = await applySmartFillRightAction();
  if (!result.ok) {
    if (result.reason === "no_formula") {
      setStatus(`Smart Fill Right skipped: active cell ${result.address} has no formula.`);
    } else if (result.reason === "merged") {
      setStatus(`Smart Fill Right skipped: active cell ${result.address} is merged.`);
    } else {
      setStatus(
        `Smart Fill Right skipped: no boundary found within 3 rows above ${result.address}.`
      );
    }
    return;
  }
  setStatus(`Smart Fill Right applied through column ${result.boundaryColumn1Based}.`);
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

async function runFormulaConsistency(): Promise<void> {
  const result = await applyFormulaConsistencyAction();
  if (!result.applied) {
    setStatus(`Formula Consistency found no formula cells to mark in ${result.address}.`);
    return;
  }
  setStatus(
    `Formula Consistency applied on ${result.address} (consistent: ${result.consistent}, inconsistent: ${result.inconsistent}). Ctrl+Z to remove.`
  );
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
  wireTraceListKeyboard();

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
    .getElementById("run-trace-precedents")
    ?.addEventListener("click", () => guardedRun(runTracePrecedents));
  document
    .getElementById("run-trace-dependents")
    ?.addEventListener("click", () => guardedRun(runTraceDependents));
  document
    .getElementById("run-trace-dialog-precedents")
    ?.addEventListener("click", () => guardedRun(() => openTraceDialogFromTaskpane("precedents")));
  document
    .getElementById("run-trace-dialog-dependents")
    ?.addEventListener("click", () => guardedRun(() => openTraceDialogFromTaskpane("dependents")));
  document
    .getElementById("run-cycle-number-format")
    ?.addEventListener("click", () => guardedRun(runCycleNumberFormat));
  document
    .getElementById("run-cycle-date-format")
    ?.addEventListener("click", () => guardedRun(runCycleDateFormat));
  document
    .getElementById("run-cycle-cell-format")
    ?.addEventListener("click", () => guardedRun(runCycleCellFormat));
  document
    .getElementById("run-cycle-text-style")
    ?.addEventListener("click", () => guardedRun(runCycleTextStyle));
  document
    .getElementById("run-reset-format-settings")
    ?.addEventListener("click", () => guardedRun(runResetFormatSettings));
  document.getElementById("run-auto-color")?.addEventListener("click", () => guardedRun(runAutoColor));
  document.getElementById("run-switch-sign")?.addEventListener("click", () => guardedRun(runSwitchSign));
  document.getElementById("run-error-wrap")?.addEventListener("click", () => guardedRun(runErrorWrap));
  document.getElementById("run-smart-fill-right")?.addEventListener("click", () => guardedRun(runSmartFillRight));
  document
    .getElementById("run-formula-consistency")
    ?.addEventListener("click", () => guardedRun(runFormulaConsistency));
  document.getElementById("run-cagr")?.addEventListener("click", () => guardedRun(runCagr));
});
