/**
 * Session-only storage for the text-style cycle index.
 *
 * Shared runtime means ribbon handlers and the taskpane live in the same
 * module graph, so a plain module variable is enough to keep the current
 * session in sync without persisting anything across workbook reopen.
 *
 * `-1` means "before the first preset" so the next cycle applies preset 0.
 * The Format Settings Save / Reset flows call `resetTextStyleCycleIndex()`
 * to restart the cycle from the top of the configured list.
 */

let textStyleCycleIndex = -1;

export function readTextStyleCycleIndex(): number {
  return textStyleCycleIndex;
}

export function writeTextStyleCycleIndex(index: number): void {
  textStyleCycleIndex = Number.isInteger(index) ? index : -1;
}

export function resetTextStyleCycleIndex(): void {
  textStyleCycleIndex = -1;
}
