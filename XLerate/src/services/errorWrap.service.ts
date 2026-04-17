import { CellMutation, ExcelPort } from "../adapters/excelPort";
import { wrapFormulaWithError } from "../core/errorWrap";

/**
 * Implements spec §3.11. Wraps every formula cell in the selection with
 * IFERROR. Existing wrappers are not detected — wrapping can nest, per spec.
 * Non-formula cells are left unchanged.
 */
export async function runErrorWrap(port: ExcelPort, errorValue = "NA()"): Promise<void> {
  const cells = await port.getSelectionCells();
  const mutations: CellMutation[] = [];

  for (const cell of cells) {
    if (!cell.isFormula) continue;
    const wrapped = wrapFormulaWithError(cell.formula, errorValue);
    if (wrapped !== cell.formula) {
      mutations.push({
        address: cell.address,
        kind: cell.isArrayFormula ? "arrayFormula" : "formula",
        formula: wrapped
      });
    }
  }

  await port.applyMutations(mutations);
}
