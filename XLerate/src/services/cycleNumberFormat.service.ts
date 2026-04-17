import { CellMutation, ExcelPort } from "../adapters/excelPort";
import {
  computeNextNumberFormat,
  DEFAULT_NUMBER_FORMATS,
  hasMixedNumberFormats,
  NumberFormatDefinition
} from "../core/numberFormatCycle";

/**
 * Implements spec §3.7. Reads the current number format from the selection,
 * asks the core which preset comes next, and applies it to every selected
 * cell in one batch.
 */
export async function runCycleNumberFormat(
  port: ExcelPort,
  configuredFormats: NumberFormatDefinition[] = DEFAULT_NUMBER_FORMATS
): Promise<void> {
  const snaps = await port.getSelectionFormatting();
  if (snaps.length === 0) return;

  const formats = snaps.map((s) => s.numberFormat);
  const mixed = hasMixedNumberFormats(formats);
  const nextFormat = computeNextNumberFormat(formats[0], mixed, configuredFormats);

  const mutations: CellMutation[] = snaps.map((s) => ({
    address: s.address,
    kind: "numberFormat",
    format: nextFormat
  }));
  await port.applyMutations(mutations);
}
