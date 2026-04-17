import { CellMutation, ExcelPort } from "../adapters/excelPort";
import {
  computeNextDateFormat,
  DateFormatDefinition,
  DEFAULT_DATE_FORMATS,
  hasMixedDateFormats,
} from "../core/dateFormatCycle";

/**
 * Implements spec §3.9. Same shape as runCycleNumberFormat but using the
 * date-format preset list.
 */
export async function runCycleDateFormat(
  port: ExcelPort,
  configuredFormats: DateFormatDefinition[] = DEFAULT_DATE_FORMATS,
): Promise<void> {
  const snaps = await port.getSelectionFormatting();
  if (snaps.length === 0) return;

  const formats = snaps.map((s) => s.numberFormat);
  const mixed = hasMixedDateFormats(formats);
  const nextFormat = computeNextDateFormat(formats[0], mixed, configuredFormats);

  const mutations: CellMutation[] = snaps.map((s) => ({
    address: s.address,
    kind: "numberFormat",
    format: nextFormat,
  }));
  await port.applyMutations(mutations);
}
