import { CellMutation, ExcelPort } from "../adapters/excelPort";
import { autoColorProbe } from "../adapters/autoColorProbe";
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
  autoColorProbe("10 service-enter");
  const cells = await port.getSelectionAutoColorCells();
  autoColorProbe("11 getSelectionAutoColorCells-returned", {
    count: cells.length,
  });
  const mutations: CellMutation[] = [];
  const categoryCounts = new Map<string, number>();

  for (const cell of cells) {
    const category = classifyAutoColorCell({
      formula: cell.isFormula ? cell.formula : null,
      value: cell.isFormula ? undefined : cell.value,
      numberFormat: cell.numberFormat ?? null,
      hasHyperlink: cell.hasHyperlink ?? false,
    });

    if (category === "none") continue;
    categoryCounts.set(category, (categoryCounts.get(category) ?? 0) + 1);
    mutations.push({
      address: cell.address,
      kind: "fontColor",
      color: palette[category],
    });
  }

  autoColorProbe("12 classification-complete", {
    mutationCount: mutations.length,
    categoryCounts: Object.fromEntries(categoryCounts.entries()),
  });
  await port.applyMutations(mutations);
  autoColorProbe("13 applyMutations-returned", {
    mutationCount: mutations.length,
  });
}
