import { CellMutation, ExcelPort } from "../adapters/excelPort";
import {
  buildCagrInsertFormula,
  CAGR_RESULT_NUMBER_FORMAT,
  findContiguousLeftCagrSeries,
  toA1Address,
} from "../core/cagrInsert";

export type InsertCagrResult =
  | { ok: false; reason: "no_series"; destination: string }
  | {
      ok: true;
      destination: string;
      sourceRange: string;
      insertedFormula: string;
      periodCount: number;
    };

/**
 * Implements spec §3.13. Discovers the contiguous numeric series immediately
 * left of the active cell, inserts a CAGR worksheet formula into the active
 * cell, and applies a percent number format in the same undo step.
 */
export async function runInsertCagr(port: ExcelPort): Promise<InsertCagrResult> {
  const snapshot = await port.getActiveCellLeftRowSnapshot();
  const destination = toA1Address(snapshot.activeCell.row, snapshot.activeCell.col);
  const series = findContiguousLeftCagrSeries(
    snapshot.leftCells.map((cell) => ({
      row: cell.address.row,
      col: cell.address.col,
      value: cell.value,
    }))
  );

  if (!series) {
    return {
      ok: false,
      reason: "no_series",
      destination,
    };
  }

  const formula = buildCagrInsertFormula(series);
  const mutations: CellMutation[] = [
    {
      address: snapshot.activeCell,
      kind: "formula",
      formula,
    },
    {
      address: snapshot.activeCell,
      kind: "numberFormat",
      format: CAGR_RESULT_NUMBER_FORMAT,
    },
  ];

  await port.applyMutations(mutations);

  return {
    ok: true,
    destination,
    sourceRange: `${toA1Address(series.start.row, series.start.col)}:${toA1Address(series.end.row, series.end.col)}`,
    insertedFormula: formula,
    periodCount: series.periodCount,
  };
}
