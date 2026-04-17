import { CellAddress, CellMutation, ExcelPort } from "../adapters/excelPort";

/**
 * One cell's pre-mark fill color, captured when the consistency check first
 * applied its green/red fill. `originalColor === null` means the cell had no
 * fill before the check ran.
 */
export type ConsistencyMarkRestore = {
  address: CellAddress;
  originalColor: string | null;
};

/**
 * Implements spec §3.6. Restores each marked cell to its pre-check fill color.
 * The caller (taskpane) owns reading the stored consistency state from
 * document settings and turning it into the `restores` list, as well as
 * deleting the state key after this succeeds.
 *
 * Design note: an earlier implementation called `port.clearSheetFill(sheetName)`
 * to bulk-wipe fills on the active sheet. That was architecturally wrong — it
 * destroyed the user's own fill formatting alongside consistency marks. The
 * per-cell restore pattern here mirrors the pre-migration VBA / TS behavior
 * and is strictly surgical: only cells the check previously modified are touched.
 */
export async function runClearConsistencyMarks(
  port: ExcelPort,
  restores: ConsistencyMarkRestore[]
): Promise<void> {
  if (restores.length === 0) return;
  const mutations: CellMutation[] = restores.map((entry) => ({
    address: entry.address,
    kind: "formatBundle",
    format:
      entry.originalColor === null || entry.originalColor === ""
        ? { fill: { pattern: "None" } }
        : { fill: { pattern: "Solid", color: entry.originalColor } },
  }));
  await port.applyMutations(mutations);
}
