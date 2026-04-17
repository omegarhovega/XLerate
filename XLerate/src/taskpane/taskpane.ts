/* global Excel, Office */
import { runAutoColor as runAutoColorService } from "../services/autoColor.service";
import { VALUE_ERROR } from "../core/cagr";
import { runCagrCalculator } from "../services/cagr.service";
import {
  computeNextCellFormat,
  type CellFormatDefinition,
  type SelectionCellFormatState
} from "../core/cellFormatCycle";
import { computeNextDateFormat, hasMixedDateFormats } from "../core/dateFormatCycle";
import { FORMAT_SETTINGS_KEY, resolveFormatSettings, type ResolvedFormatSettings } from "../core/formatSettings";
import {
  analyzeHorizontalFormulaConsistency,
  type FormulaConsistencyCell,
  type FormulaConsistencyMark
} from "../core/formulaConsistency";
import { computeNextNumberFormat, hasMixedNumberFormats } from "../core/numberFormatCycle";
import { computeSmartFillRight, type SmartFillRow } from "../core/smartFillRight";
import { ExcelPortLive } from "../adapters/excelPortLive";
import { runErrorWrap as runErrorWrapService } from "../services/errorWrap.service";
import { runSwitchSign as runSwitchSignService } from "../services/switchSign.service";
import {
  computeNextTextStyle,
  mapBorderWeight,
  type TextStyleDefinition
} from "../core/textStyleCycle";
import {
  buildTraceCellKey,
  formatTraceFormula,
  formatTraceValue,
  MAX_TRACE_ROWS,
  parseWorksheetScopedAddress,
  sanitizeTraceDepth,
  scalarFromMatrix,
  type TraceDirection
} from "../core/traceUtils";

type CellFormula = string | number | boolean;
type CellValue = string | number | boolean | null;
const CONSISTENT_COLOR = "#00F2DA";
const INCONSISTENT_COLOR = "#FF0000";
const FORMULA_CONSISTENCY_STATE_KEY = "xlerate_formula_consistency_state_v1";
const CELL_FORMAT_CYCLE_STATE_KEY = "xlerate_cell_format_cycle_state_v1";
const TEXT_STYLE_CYCLE_STATE_KEY = "xlerate_text_style_cycle_state_v1";
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

type FormulaConsistencyCellState = {
  rowOffset: number;
  colOffset: number;
  originalColor: string | null;
};

type FormulaConsistencyState = {
  sheetName: string;
  rangeAddress: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  cells: FormulaConsistencyCellState[];
};

type CellFormatCycleState = {
  sheetName: string;
  rangeAddress: string;
  lastIndex: number;
};

type TextStyleCycleState = {
  lastIndex: number;
};

type TraceRow = {
  level: number;
  address: string;
  value: string;
  formula: string;
};

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

function renderTraceRows(rows: TraceRow[]): void {
  const body = document.getElementById(TRACE_RESULTS_TBODY_ID);
  if (!(body instanceof HTMLTableSectionElement)) {
    return;
  }

  body.textContent = "";
  if (rows.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 4;
    cell.textContent = "No trace results.";
    row.appendChild(cell);
    body.appendChild(row);
    return;
  }

  for (const item of rows) {
    const row = document.createElement("tr");
    row.className = "trace-row-clickable";
    const level = document.createElement("td");
    const address = document.createElement("td");
    const value = document.createElement("td");
    const formula = document.createElement("td");
    const addressBtn = document.createElement("button");

    level.textContent = String(item.level);
    addressBtn.type = "button";
    addressBtn.className = "trace-link-btn";
    addressBtn.textContent = item.address;
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
  }
}

function asFormulaCell(cell: CellFormula): string | null {
  return typeof cell === "string" && cell.startsWith("=") ? cell : null;
}

function toFormulaConsistencyRows(formulasR1C1: CellFormula[][]): FormulaConsistencyCell[][] {
  return formulasR1C1.map((row) =>
    row.map((raw) => {
      const formula = asFormulaCell(raw);
      return {
        isFormula: formula !== null,
        formulaR1C1: formula ?? undefined
      };
    })
  );
}

function applyConsistencyColor(cell: Excel.Range, mark: FormulaConsistencyMark): void {
  if (mark === "consistent") {
    cell.format.fill.color = CONSISTENT_COLOR;
  } else if (mark === "inconsistent") {
    cell.format.fill.color = INCONSISTENT_COLOR;
  }
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

function applyTextStyleToRange(range: Excel.Range, style: TextStyleDefinition): void {
  range.format.font.name = style.fontName;
  range.format.font.size = style.fontSize;
  range.format.font.bold = style.bold;
  range.format.font.italic = style.italic;
  range.format.font.underline = style.underline ? "Single" : "None";
  range.format.font.color = style.fontColor;
  range.format.fill.color = style.backColor;

  // Match VBA behavior: clear all borders, then apply configured edges only.
  range.format.borders.getItem(Excel.BorderIndex.edgeLeft).style = Excel.BorderLineStyle.none;
  range.format.borders.getItem(Excel.BorderIndex.edgeTop).style = Excel.BorderLineStyle.none;
  range.format.borders.getItem(Excel.BorderIndex.edgeBottom).style = Excel.BorderLineStyle.none;
  range.format.borders.getItem(Excel.BorderIndex.edgeRight).style = Excel.BorderLineStyle.none;
  range.format.borders.getItem(Excel.BorderIndex.insideHorizontal).style = Excel.BorderLineStyle.none;
  range.format.borders.getItem(Excel.BorderIndex.insideVertical).style = Excel.BorderLineStyle.none;

  if (style.borderStyle === "None") {
    return;
  }

  const weight = mapBorderWeight(style.borderStyle) as Excel.BorderWeight;
  const apply = (side: Excel.BorderIndex, enabled: boolean): void => {
    if (!enabled) {
      return;
    }
    const border = range.format.borders.getItem(side);
    border.style = style.borderStyle as Excel.BorderLineStyle;
    border.weight = weight;
  };

  apply(Excel.BorderIndex.edgeTop, style.borderTop);
  apply(Excel.BorderIndex.edgeBottom, style.borderBottom);
  apply(Excel.BorderIndex.edgeLeft, style.borderLeft);
  apply(Excel.BorderIndex.edgeRight, style.borderRight);
}

function readFormulaConsistencyState(): FormulaConsistencyState | null {
  const raw = Office.context.document.settings.get(FORMULA_CONSISTENCY_STATE_KEY);
  if (typeof raw !== "string" || raw.length === 0) {
    return null;
  }

  try {
    const parsed = JSON.parse(raw) as FormulaConsistencyState;
    if (
      typeof parsed.sheetName !== "string" ||
      typeof parsed.rangeAddress !== "string" ||
      typeof parsed.rowIndex !== "number" ||
      typeof parsed.columnIndex !== "number" ||
      typeof parsed.rowCount !== "number" ||
      typeof parsed.columnCount !== "number" ||
      !Array.isArray(parsed.cells)
    ) {
      return null;
    }
    return parsed;
  } catch {
    return null;
  }
}

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

async function writeFormulaConsistencyState(state: FormulaConsistencyState): Promise<void> {
  Office.context.document.settings.set(FORMULA_CONSISTENCY_STATE_KEY, JSON.stringify(state));
  await saveDocumentSettingsAsync();
}

async function clearFormulaConsistencyState(): Promise<void> {
  Office.context.document.settings.remove(FORMULA_CONSISTENCY_STATE_KEY);
  await saveDocumentSettingsAsync();
}

function readCellFormatCycleState(): CellFormatCycleState | null {
  const raw = Office.context.document.settings.get(CELL_FORMAT_CYCLE_STATE_KEY);
  if (typeof raw !== "string" || raw.length === 0) {
    return null;
  }

  try {
    const parsed = JSON.parse(raw) as CellFormatCycleState;
    if (
      typeof parsed.sheetName !== "string" ||
      typeof parsed.rangeAddress !== "string" ||
      typeof parsed.lastIndex !== "number"
    ) {
      return null;
    }
    return parsed;
  } catch {
    return null;
  }
}

async function writeCellFormatCycleState(state: CellFormatCycleState): Promise<void> {
  Office.context.document.settings.set(CELL_FORMAT_CYCLE_STATE_KEY, JSON.stringify(state));
  await saveDocumentSettingsAsync();
}

function readTextStyleCycleState(): TextStyleCycleState {
  const raw = Office.context.document.settings.get(TEXT_STYLE_CYCLE_STATE_KEY);
  if (typeof raw !== "string" || raw.length === 0) {
    return { lastIndex: -1 };
  }

  try {
    const parsed = JSON.parse(raw) as TextStyleCycleState;
    if (typeof parsed.lastIndex === "number") {
      return parsed;
    }
    return { lastIndex: -1 };
  } catch {
    return { lastIndex: -1 };
  }
}

async function writeTextStyleCycleState(state: TextStyleCycleState): Promise<void> {
  Office.context.document.settings.set(TEXT_STYLE_CYCLE_STATE_KEY, JSON.stringify(state));
  await saveDocumentSettingsAsync();
}

function readResolvedFormatSettings(): ResolvedFormatSettings {
  const raw = Office.context.document.settings.get(FORMAT_SETTINGS_KEY);
  return resolveFormatSettings(raw);
}

async function clearFormatSettingsAndCycleState(): Promise<void> {
  Office.context.document.settings.remove(FORMAT_SETTINGS_KEY);
  Office.context.document.settings.remove(CELL_FORMAT_CYCLE_STATE_KEY);
  Office.context.document.settings.remove(TEXT_STYLE_CYCLE_STATE_KEY);
  await saveDocumentSettingsAsync();
}

async function writeFormatSettingsAndResetCycleState(settings: ResolvedFormatSettings): Promise<void> {
  Office.context.document.settings.set(FORMAT_SETTINGS_KEY, JSON.stringify(settings));
  Office.context.document.settings.remove(CELL_FORMAT_CYCLE_STATE_KEY);
  Office.context.document.settings.remove(TEXT_STYLE_CYCLE_STATE_KEY);
  await saveDocumentSettingsAsync();
}

async function restoreFormulaConsistencyState(
  context: Excel.RequestContext,
  state: FormulaConsistencyState
): Promise<{ restoredCells: number; sheetMissing: boolean }> {
  const worksheet = context.workbook.worksheets.getItemOrNullObject(state.sheetName);
  worksheet.load("isNullObject");
  await context.sync();

  if (worksheet.isNullObject) {
    return { restoredCells: 0, sheetMissing: true };
  }

  const target = worksheet.getRangeByIndexes(
    state.rowIndex,
    state.columnIndex,
    Math.max(1, state.rowCount),
    Math.max(1, state.columnCount)
  );

  for (const entry of state.cells) {
    const cell = target.getCell(entry.rowOffset, entry.colOffset);
    if (!entry.originalColor) {
      cell.format.fill.clear();
    } else {
      cell.format.fill.color = entry.originalColor;
    }
  }

  await context.sync();
  return { restoredCells: state.cells.length, sheetMissing: false };
}

function isItemNotFoundError(error: unknown): boolean {
  if (!error || typeof error !== "object") {
    return false;
  }

  const maybe = error as { code?: unknown; message?: unknown };
  if (maybe.code === "ItemNotFound") {
    return true;
  }
  return typeof maybe.message === "string" && maybe.message.includes("ItemNotFound");
}

async function getDirectTraceNeighbors(
  context: Excel.RequestContext,
  source: Excel.Range,
  direction: TraceDirection
): Promise<Excel.Range[]> {
  const links =
    direction === "precedents" ? source.getDirectPrecedents() : source.getDirectDependents();
  links.areas.load("items");

  try {
    await context.sync();
  } catch (error) {
    if (isItemNotFoundError(error)) {
      return [];
    }
    throw error;
  }

  for (const bySheet of links.areas.items) {
    bySheet.areas.load("items/address,rowIndex,columnIndex,rowCount,columnCount,worksheet/name,values,formulas");
  }
  await context.sync();

  const neighbors: Excel.Range[] = [];
  let expanded = false;

  for (const bySheet of links.areas.items) {
    for (const area of bySheet.areas.items) {
      if (area.rowCount === 1 && area.columnCount === 1) {
        neighbors.push(area);
        continue;
      }

      for (let r = 0; r < area.rowCount; r += 1) {
        for (let c = 0; c < area.columnCount; c += 1) {
          const cell = area.getCell(r, c);
          cell.load(["address", "worksheet/name", "rowIndex", "columnIndex", "values", "formulas"]);
          neighbors.push(cell);
          expanded = true;
        }
      }
    }
  }

  if (expanded) {
    await context.sync();
  }

  return neighbors;
}

function toTraceRow(cell: Excel.Range, level: number): TraceRow {
  const value = formatTraceValue(scalarFromMatrix(cell.values));
  const formula = formatTraceFormula(scalarFromMatrix(cell.formulas));
  return {
    level,
    address: cell.address,
    value,
    formula
  };
}

async function runTrace(direction: TraceDirection): Promise<void> {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.12")) {
    setStatus("Trace requires ExcelApi 1.12 or later on this Excel host.");
    return;
  }

  await Excel.run(async (context) => {
    const root = context.workbook.getActiveCell();
    root.load(["address", "worksheet/name", "rowIndex", "columnIndex", "values", "formulas"]);
    await context.sync();

    const maxDepth = getTraceMaxDepthInputValue();
    const rows: TraceRow[] = [toTraceRow(root, 0)];
    const visited = new Set<string>([buildTraceCellKey(root.worksheet.name, root.rowIndex, root.columnIndex)]);
    const queue: Array<{ level: number; cell: Excel.Range }> = [{ level: 0, cell: root }];
    let truncated = false;

    while (queue.length > 0) {
      const current = queue.shift();
      if (!current) {
        break;
      }

      if (current.level >= maxDepth) {
        continue;
      }

      const neighbors = await getDirectTraceNeighbors(context, current.cell, direction);
      for (const neighbor of neighbors) {
        const key = buildTraceCellKey(neighbor.worksheet.name, neighbor.rowIndex, neighbor.columnIndex);
        if (visited.has(key)) {
          continue;
        }
        visited.add(key);

        rows.push(toTraceRow(neighbor, current.level + 1));
        if (rows.length >= MAX_TRACE_ROWS) {
          truncated = true;
          break;
        }

        queue.push({ level: current.level + 1, cell: neighbor });
      }

      if (truncated) {
        break;
      }
    }

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
  await Excel.run(async (context) => {
    const formatSettings = readResolvedFormatSettings();
    const configuredFormats = formatSettings.numberFormats;
    const range = context.workbook.getSelectedRange();
    range.load(["numberFormat", "rowCount", "columnCount", "address"]);
    await context.sync();

    const selectionFormats = flattenFormatMatrix(range.numberFormat as unknown[][]);
    if (selectionFormats.length === 0) {
      setStatus("Cycle Number Format skipped: empty selection.");
      return;
    }

    const currentFormat = selectionFormats[0];
    const mixedSelection = hasMixedNumberFormats(selectionFormats);
    const nextFormat = computeNextNumberFormat(currentFormat, mixedSelection, configuredFormats);

    range.numberFormat = makeFormatMatrix(range.rowCount, range.columnCount, nextFormat);
    await context.sync();

    const formatName = configuredFormats.find((item) => item.formatCode === nextFormat)?.name ?? "custom format";
    setStatus(`Cycle Number Format applied "${formatName}" on ${range.address}.`);
  });
}

async function runCycleDateFormat(): Promise<void> {
  await Excel.run(async (context) => {
    const formatSettings = readResolvedFormatSettings();
    const configuredFormats = formatSettings.dateFormats;
    const range = context.workbook.getSelectedRange();
    range.load(["numberFormat", "rowCount", "columnCount", "address"]);
    await context.sync();

    const selectionFormats = flattenFormatMatrix(range.numberFormat as unknown[][]);
    if (selectionFormats.length === 0) {
      setStatus("Cycle Date Format skipped: empty selection.");
      return;
    }

    const currentFormat = selectionFormats[0];
    const mixedSelection = hasMixedDateFormats(selectionFormats);
    const nextFormat = computeNextDateFormat(currentFormat, mixedSelection, configuredFormats);

    range.numberFormat = makeFormatMatrix(range.rowCount, range.columnCount, nextFormat);
    await context.sync();

    const formatName = configuredFormats.find((item) => item.formatCode === nextFormat)?.name ?? "custom format";
    setStatus(`Cycle Date Format applied "${formatName}" on ${range.address}.`);
  });
}

async function runCycleCellFormat(): Promise<void> {
  await Excel.run(async (context) => {
    const formatSettings = readResolvedFormatSettings();
    const configuredFormats = formatSettings.cellFormats;
    const range = context.workbook.getSelectedRange();
    const worksheet = range.worksheet;
    range.load(["address", "rowCount", "columnCount"]);
    worksheet.load("name");
    range.format.fill.load(["pattern", "color"]);
    range.format.font.load(["color", "bold", "italic", "underline", "strikethrough"]);
    const borders = {
      EdgeLeft: range.format.borders.getItem("EdgeLeft"),
      EdgeTop: range.format.borders.getItem("EdgeTop"),
      EdgeBottom: range.format.borders.getItem("EdgeBottom"),
      EdgeRight: range.format.borders.getItem("EdgeRight")
    } as const;
    BORDER_SIDE_ITEMS.forEach((side) => borders[side].load(["style", "color"]));
    await context.sync();

    const cycleState = readCellFormatCycleState();
    let nextFormat: CellFormatDefinition;
    let nextIndex: number;

    if (
      cycleState &&
      cycleState.sheetName === worksheet.name &&
      cycleState.rangeAddress === range.address &&
      configuredFormats.length > 0
    ) {
      nextIndex = (cycleState.lastIndex + 1) % configuredFormats.length;
      nextFormat = configuredFormats[nextIndex];
    } else {
      const state = buildSelectionCellFormatState(range, borders);
      nextFormat = computeNextCellFormat(state, configuredFormats);
      nextIndex = configuredFormats.indexOf(nextFormat);
      if (nextIndex < 0) {
        nextIndex = 0;
      }
    }

    applyCellFormatToRange(range, nextFormat);
    await context.sync();
    await writeCellFormatCycleState({
      sheetName: worksheet.name,
      rangeAddress: range.address,
      lastIndex: nextIndex
    });

    setStatus(`Cycle Cell Format applied "${nextFormat.name}" on ${range.address}.`);
  });
}

async function runCycleTextStyle(): Promise<void> {
  await Excel.run(async (context) => {
    const formatSettings = readResolvedFormatSettings();
    const configuredStyles = formatSettings.textStyles;
    const range = context.workbook.getSelectedRange();
    range.load(["address"]);
    await context.sync();

    const current = readTextStyleCycleState();
    const next = computeNextTextStyle(current.lastIndex, configuredStyles);
    applyTextStyleToRange(range, next.style);
    await context.sync();

    await writeTextStyleCycleState({ lastIndex: next.index });
    setStatus(`Cycle Text Style applied "${next.style.name}" on ${range.address}.`);
  });
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

function toSmartFillRows(values: CellValue[][], formulas: CellFormula[][]): SmartFillRow[] {
  return values.map((rowValues, r) =>
    rowValues.map((value, c) => {
      const formula = asFormulaCell(formulas[r][c]);
      const isEmpty = formula === null && (value === null || value === "");
      return {
        isEmpty,
        isMerged: false
      };
    })
  );
}

async function runSmartFillRight(): Promise<void> {
  await Excel.run(async (context) => {
    const workbook = context.workbook;
    const worksheet = workbook.worksheets.getActiveWorksheet();
    const activeCell = workbook.getActiveCell();
    const usedRange = worksheet.getUsedRangeOrNullObject();

    activeCell.load(["rowIndex", "columnIndex", "formulas", "address"]);
    usedRange.load(["isNullObject", "columnIndex", "columnCount"]);
    await context.sync();

    const activeFormula = asFormulaCell(activeCell.formulas[0][0] as CellFormula) ?? "";

    const startRowIndex = Math.max(0, activeCell.rowIndex - 3);
    const sampleRowCount = activeCell.rowIndex - startRowIndex + 1;

    const usedLastColExclusive = usedRange.isNullObject ? activeCell.columnIndex + 1 : usedRange.columnIndex + usedRange.columnCount;
    const sampleColCount = Math.max(1, Math.min(2000, usedLastColExclusive - activeCell.columnIndex));

    const sample = worksheet.getRangeByIndexes(startRowIndex, activeCell.columnIndex, sampleRowCount, sampleColCount);
    sample.load(["values", "formulas"]);
    await context.sync();

    const rows = toSmartFillRows(sample.values as CellValue[][], sample.formulas as CellFormula[][]);
    const result = computeSmartFillRight(rows, {
      row: sampleRowCount,
      col: 1,
      formula: activeFormula,
      isMerged: false
    });

    if (!result.ok) {
      if (result.reason === "active_cell_must_have_formula") {
        setStatus(`Smart Fill Right skipped: active cell ${activeCell.address} has no formula.`);
      } else if (result.reason === "active_cell_is_merged") {
        setStatus(`Smart Fill Right skipped: active cell ${activeCell.address} is merged.`);
      } else {
        setStatus(`Smart Fill Right skipped: no boundary found within 3 rows above ${activeCell.address}.`);
      }
      return;
    }

    const boundaryAbsCol = activeCell.columnIndex + (result.boundaryCol - 1);
    const destination = worksheet.getRangeByIndexes(
      activeCell.rowIndex,
      activeCell.columnIndex,
      1,
      boundaryAbsCol - activeCell.columnIndex + 1
    );

    destination.copyFrom(activeCell, Excel.RangeCopyType.formulas);
    await context.sync();
    setStatus(`Smart Fill Right applied through column ${boundaryAbsCol + 1}.`);
  });
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
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const worksheet = range.worksheet;
    range.load(["formulasR1C1", "rowCount", "columnCount", "address", "rowIndex", "columnIndex"]);
    worksheet.load("name");
    await context.sync();

    const existingState = readFormulaConsistencyState();
    if (existingState) {
      const restoreResult = await restoreFormulaConsistencyState(context, existingState);
      await clearFormulaConsistencyState();

      if (existingState.sheetName === worksheet.name) {
        if (restoreResult.sheetMissing) {
          setStatus("Formula Consistency toggle off: previous sheet no longer exists.");
        } else {
          setStatus(
            `Formula Consistency formatting restored on ${existingState.rangeAddress} (${restoreResult.restoredCells} cells).`
          );
        }
        return;
      }
    }

    const rows = toFormulaConsistencyRows(range.formulasR1C1 as CellFormula[][]);
    const marks = analyzeHorizontalFormulaConsistency(rows);
    const changedCells: Array<{ row: number; col: number; mark: FormulaConsistencyMark; cell: Excel.Range }> = [];

    let consistentCount = 0;
    let inconsistentCount = 0;

    for (let r = 0; r < range.rowCount; r += 1) {
      for (let c = 0; c < range.columnCount; c += 1) {
        const mark = marks[r][c];
        if (mark === "none") {
          continue;
        }

        const cell = range.getCell(r, c);
        cell.format.fill.load("color");
        changedCells.push({ row: r, col: c, mark, cell });

        if (mark === "consistent") {
          consistentCount += 1;
        } else {
          inconsistentCount += 1;
        }
      }
    }

    if (changedCells.length === 0) {
      setStatus(`Formula Consistency found no formula cells to mark in ${range.address}.`);
      return;
    }

    await context.sync();

    const stateCells: FormulaConsistencyCellState[] = [];
    for (const item of changedCells) {
      const originalColor = item.cell.format.fill.color;
      stateCells.push({
        rowOffset: item.row,
        colOffset: item.col,
        originalColor: typeof originalColor === "string" && originalColor.length > 0 ? originalColor : null
      });
      applyConsistencyColor(item.cell, item.mark);
    }

    await context.sync();
    await writeFormulaConsistencyState({
      sheetName: worksheet.name,
      rangeAddress: range.address,
      rowIndex: range.rowIndex,
      columnIndex: range.columnIndex,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      cells: stateCells
    });

    setStatus(
      `Formula Consistency applied on ${range.address} (consistent: ${consistentCount}, inconsistent: ${inconsistentCount}).`
    );
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
    .getElementById("run-trace-precedents")
    ?.addEventListener("click", () => guardedRun(runTracePrecedents));
  document
    .getElementById("run-trace-dependents")
    ?.addEventListener("click", () => guardedRun(runTraceDependents));
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
