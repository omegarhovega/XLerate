import { ExcelPort } from "../adapters/excelPort";

/**
 * Implements spec §3.6. Removes all fill color from every cell on the named
 * sheet. The confirmation dialog is surfaced in the taskpane; by the time
 * this service runs, the user has already confirmed.
 */
export async function runClearConsistencyMarks(port: ExcelPort, sheetName: string): Promise<void> {
  await port.clearSheetFill(sheetName);
}
