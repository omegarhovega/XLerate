import { CellMutation, ExcelPort } from "../adapters/excelPort";
import {
  AutoColorPalette,
  classifyAutoColorCell,
  DEFAULT_AUTO_COLOR_PALETTE,
} from "../core/autoColor";

/**
 * Implements spec §3.12. Classifies each selected cell and applies the
 * palette's color to the cell's font. Blank cells and cells classified as
 * "none" are left unchanged.
 */
export async function runAutoColor(
  port: ExcelPort,
  palette: AutoColorPalette = DEFAULT_AUTO_COLOR_PALETTE
): Promise<void> {
  const cells = await port.getSelectionCells();
  const formats = await port.getSelectionFormatting();
  const mutations: CellMutation[] = [];

  for (let i = 0; i < cells.length; i++) {
    const cell = cells[i];
    const fmt = formats[i];
    const category = classifyAutoColorCell({
      formula: cell.isFormula ? cell.formula : null,
      value: cell.isFormula ? undefined : cell.value,
      numberFormat: fmt?.numberFormat ?? null,
      hasHyperlink: fmt?.hasHyperlink ?? false,
    });

    if (category === "none") continue;
    mutations.push({
      address: cell.address,
      kind: "fontColor",
      color: palette[category],
    });
  }

  await port.applyMutations(mutations);
}
