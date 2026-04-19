/* global Excel */

import {
  analyzeHorizontalFormulaConsistency,
  type FormulaConsistencyCell,
  type FormulaConsistencyMark,
} from "../core/formulaConsistency";
import { computeSmartFillRight, type SmartFillRow } from "../core/smartFillRight";

/**
 * Shared Office.js-using action helpers for features that live on more
 * than one surface (taskpane buttons + ribbon ExecuteFunction handlers).
 * Each exported `apply*Action()` returns a structured result; callers
 * wrap their own status reporting (setStatus in the taskpane, logging
 * in the commands runtime) around it. No DOM access here — the commands
 * runtime has no taskpane DOM to touch.
 */

type CellFormula = string | number | boolean;
type CellValue = string | number | boolean | null;

const CONSISTENT_COLOR = "#00F2DA";
const INCONSISTENT_COLOR = "#FF0000";

function asFormulaCell(cell: CellFormula): string | null {
  return typeof cell === "string" && cell.startsWith("=") ? cell : null;
}

function toFormulaConsistencyRows(formulasR1C1: CellFormula[][]): FormulaConsistencyCell[][] {
  return formulasR1C1.map((row) =>
    row.map((raw) => {
      const formula = asFormulaCell(raw);
      return {
        isFormula: formula !== null,
        formulaR1C1: formula ?? undefined,
      };
    })
  );
}

function toSmartFillRows(values: CellValue[][], formulas: CellFormula[][]): SmartFillRow[] {
  return values.map((rowValues, r) =>
    rowValues.map((value, c) => {
      const formula = asFormulaCell(formulas[r][c]);
      const isEmpty = formula === null && (value === null || value === "");
      return { isEmpty, isMerged: false };
    })
  );
}

function applyConsistencyColor(cell: Excel.Range, mark: FormulaConsistencyMark): void {
  const color =
    mark === "consistent"
      ? CONSISTENT_COLOR
      : mark === "inconsistent"
        ? INCONSISTENT_COLOR
        : null;
  if (color === null) return;
  // Office.js requires fill.pattern to be set before color renders
  // reliably on a previously-unfilled cell. See CLAUDE.md gotchas.
  cell.format.fill.pattern = "Solid";
  cell.format.fill.color = color;
}

// ---- Formula Consistency ----

export type FormulaConsistencyResult =
  | { applied: false; address: string }
  | { applied: true; address: string; consistent: number; inconsistent: number };

/**
 * Apply green/red fills to formula cells in the selection that are
 * consistent / inconsistent with their horizontal neighbors. Single Excel
 * undo step; Ctrl+Z reverses. No persistent state (see CLAUDE.md —
 * saveAsync breaks the undo chain).
 */
export async function applyFormulaConsistencyAction(): Promise<FormulaConsistencyResult> {
  let out: FormulaConsistencyResult = { applied: false, address: "" };
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["formulasR1C1", "rowCount", "columnCount", "address"]);
    await context.sync();

    const rows = toFormulaConsistencyRows(range.formulasR1C1 as CellFormula[][]);
    const marks = analyzeHorizontalFormulaConsistency(rows);

    let consistentCount = 0;
    let inconsistentCount = 0;
    let applied = false;

    for (let r = 0; r < range.rowCount; r += 1) {
      for (let c = 0; c < range.columnCount; c += 1) {
        const mark = marks[r][c];
        if (mark === "none") continue;
        applyConsistencyColor(range.getCell(r, c), mark);
        applied = true;
        if (mark === "consistent") consistentCount += 1;
        else inconsistentCount += 1;
      }
    }

    if (!applied) {
      out = { applied: false, address: range.address };
      return;
    }

    await context.sync();
    out = {
      applied: true,
      address: range.address,
      consistent: consistentCount,
      inconsistent: inconsistentCount,
    };
  });
  return out;
}

// ---- Smart Fill Right ----

export type SmartFillRightResult =
  | { ok: false; reason: "no_formula" | "merged" | "no_boundary"; address: string }
  | { ok: true; boundaryColumn1Based: number };

/**
 * Fill the active cell's formula to the right up to the boundary found
 * in the 3 rows above (or the used range's right edge, whichever is
 * closer). Skipped with a structured reason when the active cell has no
 * formula, is merged, or no boundary exists.
 */
export async function applySmartFillRightAction(): Promise<SmartFillRightResult> {
  let out: SmartFillRightResult = {
    ok: false,
    reason: "no_boundary",
    address: "",
  };
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
    const usedLastColExclusive = usedRange.isNullObject
      ? activeCell.columnIndex + 1
      : usedRange.columnIndex + usedRange.columnCount;
    const sampleColCount = Math.max(
      1,
      Math.min(2000, usedLastColExclusive - activeCell.columnIndex)
    );

    const sample = worksheet.getRangeByIndexes(
      startRowIndex,
      activeCell.columnIndex,
      sampleRowCount,
      sampleColCount
    );
    sample.load(["values", "formulas"]);
    await context.sync();

    const sampleRows = toSmartFillRows(
      sample.values as CellValue[][],
      sample.formulas as CellFormula[][]
    );
    const result = computeSmartFillRight(sampleRows, {
      row: sampleRowCount,
      col: 1,
      formula: activeFormula,
      isMerged: false,
    });

    if (!result.ok) {
      const reason: SmartFillRightResult extends { ok: false; reason: infer R }
        ? R
        : never =
        result.reason === "active_cell_must_have_formula"
          ? "no_formula"
          : result.reason === "active_cell_is_merged"
            ? "merged"
            : "no_boundary";
      out = { ok: false, reason, address: activeCell.address };
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
    out = { ok: true, boundaryColumn1Based: boundaryAbsCol + 1 };
  });
  return out;
}
