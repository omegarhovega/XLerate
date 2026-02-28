/* global Excel, Office */
import { calculateCagr, VALUE_ERROR } from "../core/cagr";
import { wrapFormulaWithError } from "../core/errorWrap";
import {
  analyzeHorizontalFormulaConsistency,
  type FormulaConsistencyCell,
  type FormulaConsistencyMark
} from "../core/formulaConsistency";
import {
  computeNextNumberFormat,
  DEFAULT_NUMBER_FORMATS,
  hasMixedNumberFormats
} from "../core/numberFormatCycle";
import { computeSmartFillRight, type SmartFillRow } from "../core/smartFillRight";
import { switchSignCell } from "../core/switchSign";

type CellFormula = string | number | boolean;
type CellValue = string | number | boolean | null;
const CONSISTENT_COLOR = "#00F2DA";
const INCONSISTENT_COLOR = "#FF0000";
const FORMULA_CONSISTENCY_STATE_KEY = "xlerate_formula_consistency_state_v1";

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

async function runSwitchSign(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["formulas", "values", "rowCount", "columnCount"]);
    await context.sync();

    const updated: CellFormula[][] = [];
    for (let r = 0; r < range.rowCount; r += 1) {
      updated[r] = [];
      for (let c = 0; c < range.columnCount; c += 1) {
        const rawFormula = range.formulas[r][c] as CellFormula;
        const formula = asFormulaCell(rawFormula);
        const value = range.values[r][c] as CellValue;
        const result = switchSignCell({
          isFormula: formula !== null,
          formula: formula ?? undefined,
          value: formula === null ? value : undefined
        });
        updated[r][c] = result.isFormula ? (result.formula as string) : (result.value as CellFormula);
      }
    }

    range.formulas = updated as unknown as string[][];
    await context.sync();
    setStatus("Switch Sign applied.");
  });
}

async function runCycleNumberFormat(): Promise<void> {
  await Excel.run(async (context) => {
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
    const nextFormat = computeNextNumberFormat(currentFormat, mixedSelection, DEFAULT_NUMBER_FORMATS);

    range.numberFormat = makeFormatMatrix(range.rowCount, range.columnCount, nextFormat);
    await context.sync();

    const formatName = DEFAULT_NUMBER_FORMATS.find((item) => item.formatCode === nextFormat)?.name ?? "custom format";
    setStatus(`Cycle Number Format applied "${formatName}" on ${range.address}.`);
  });
}

async function runErrorWrap(): Promise<void> {
  await Excel.run(async (context) => {
    const fallbackInput = (document.getElementById("error-value") as HTMLInputElement | null)?.value?.trim() || "NA()";
    const range = context.workbook.getSelectedRange();
    range.load(["formulas", "rowCount", "columnCount"]);
    await context.sync();

    const updated: CellFormula[][] = [];
    for (let r = 0; r < range.rowCount; r += 1) {
      updated[r] = [];
      for (let c = 0; c < range.columnCount; c += 1) {
        const rawFormula = range.formulas[r][c] as CellFormula;
        const formula = asFormulaCell(rawFormula);
        updated[r][c] = formula ? wrapFormulaWithError(formula, fallbackInput) : rawFormula;
      }
    }

    range.formulas = updated as unknown as string[][];
    await context.sync();
    setStatus(`Error Wrap applied with fallback "${fallbackInput}".`);
  });
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

    const result = calculateCagr(values);
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

  document
    .getElementById("run-cycle-number-format")
    ?.addEventListener("click", () => guardedRun(runCycleNumberFormat));
  document.getElementById("run-switch-sign")?.addEventListener("click", () => guardedRun(runSwitchSign));
  document.getElementById("run-error-wrap")?.addEventListener("click", () => guardedRun(runErrorWrap));
  document.getElementById("run-smart-fill-right")?.addEventListener("click", () => guardedRun(runSmartFillRight));
  document
    .getElementById("run-formula-consistency")
    ?.addEventListener("click", () => guardedRun(runFormulaConsistency));
  document.getElementById("run-cagr")?.addEventListener("click", () => guardedRun(runCagr));
});
