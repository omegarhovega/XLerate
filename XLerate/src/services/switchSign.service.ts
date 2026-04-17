import { CellMutation, ExcelPort } from "../adapters/excelPort";
import { switchSignCell } from "../core/switchSign";

/**
 * Implements spec §3.3 Switch Sign by composing the existing pure core
 * `switchSignCell` with an ExcelPort. Produces a single batch of mutations
 * so that the action is a single Excel undo step.
 */
export async function runSwitchSign(port: ExcelPort): Promise<void> {
  const cells = await port.getSelectionCells();
  const mutations: CellMutation[] = [];

  for (const cell of cells) {
    const before = {
      isFormula: cell.isFormula,
      isArrayFormula: cell.isArrayFormula,
      formula: cell.formula,
      value: cell.value,
    };
    const after = switchSignCell(before);

    if (after.isFormula) {
      const newFormula = after.formula ?? "";
      if (newFormula !== cell.formula) {
        mutations.push({
          address: cell.address,
          kind: after.isArrayFormula ? "arrayFormula" : "formula",
          formula: newFormula,
        });
      }
      continue;
    }

    if (after.value !== cell.value) {
      mutations.push({
        address: cell.address,
        kind: "value",
        value: after.value,
      });
    }
  }

  await port.applyMutations(mutations);
}
