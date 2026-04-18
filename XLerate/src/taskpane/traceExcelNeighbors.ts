/* global Excel */

import type { TraceCellInfo } from "../core/traceBuilder";
import { parseWorksheetScopedAddress, type TraceDirection } from "../core/traceUtils";

/**
 * Office.js-bound helpers that back the pure `buildTrace` BFS. Shared
 * between the taskpane (live trace panel) and the trace dialog page
 * (Phase B). Kept in `src/taskpane/` because they use the Office.js
 * globals; the dependency-cruiser rule confines `office-js` *imports* to
 * `src/adapters/` but globals are unrestricted, matching the pattern the
 * taskpane already uses.
 *
 * Every function here assumes the Office.js properties listed in
 * `TRACE_CELL_LOAD_FIELDS` have been loaded on the relevant `Excel.Range`
 * before it is used.
 */

export const TRACE_CELL_LOAD_FIELDS = [
  "address",
  "worksheet/name",
  "rowIndex",
  "columnIndex",
  "values",
  "formulas",
] as const;

export function loadTraceCellProperties(range: Excel.Range): void {
  range.load([...TRACE_CELL_LOAD_FIELDS]);
}

/**
 * Snapshot a loaded Excel.Range into the plain TraceCellInfo the core
 * builder expects. Caller must have loaded `TRACE_CELL_LOAD_FIELDS` and
 * `context.sync()`ed.
 */
export function snapshotRangeForTrace(cell: Excel.Range): TraceCellInfo {
  return {
    worksheetName: cell.worksheet.name,
    rowIndex: cell.rowIndex,
    columnIndex: cell.columnIndex,
    address: cell.address,
    value: cell.values,
    formula: cell.formulas,
  };
}

function isItemNotFoundError(error: unknown): boolean {
  if (!error || typeof error !== "object") return false;
  const maybe = error as { code?: unknown; message?: unknown };
  if (maybe.code === "ItemNotFound") return true;
  return typeof maybe.message === "string" && maybe.message.includes("ItemNotFound");
}

/**
 * Fetch the direct precedents or dependents of `source`. Expands
 * multi-cell precedent areas into individual cells so the BFS can
 * traverse them one at a time. Returns empty when Office.js reports no
 * links (or when the underlying link enumeration throws ItemNotFound,
 * which Excel uses for "no links found" on some hosts).
 */
export async function getDirectTraceNeighbors(
  context: Excel.RequestContext,
  source: Excel.Range,
  direction: TraceDirection
): Promise<Excel.Range[]> {
  const links =
    direction === "precedents" ? source.getDirectPrecedents() : source.getDirectDependents();
  links.areas.load("items");

  try {
    await context.sync();
  } catch (error) {
    if (isItemNotFoundError(error)) return [];
    throw error;
  }

  for (const bySheet of links.areas.items) {
    bySheet.areas.load(
      "items/address,rowIndex,columnIndex,rowCount,columnCount,worksheet/name,values,formulas"
    );
  }
  await context.sync();

  const neighbors: Excel.Range[] = [];
  let expanded = false;

  for (const bySheet of links.areas.items) {
    for (const area of bySheet.areas.items) {
      if (area.rowCount === 1 && area.columnCount === 1) {
        neighbors.push(area);
        continue;
      }
      for (let r = 0; r < area.rowCount; r += 1) {
        for (let c = 0; c < area.columnCount; c += 1) {
          const cell = area.getCell(r, c);
          loadTraceCellProperties(cell);
          neighbors.push(cell);
          expanded = true;
        }
      }
    }
  }

  if (expanded) await context.sync();
  return neighbors;
}

/**
 * Resolve the starting cell for a trace. If `address` parses to a valid
 * worksheet-scoped reference that exists in the workbook, that range is
 * returned (loaded with the standard trace fields). Otherwise falls back
 * to `context.workbook.getActiveCell()`. Callers must `await context.sync()`
 * before inspecting properties.
 *
 * Returning null means the fallback itself failed, which is unusual —
 * getActiveCell typically always resolves — but callers should handle it.
 */
export async function resolveTraceStartCell(
  context: Excel.RequestContext,
  address: string | null
): Promise<Excel.Range | null> {
  if (address) {
    const parsed = parseWorksheetScopedAddress(address);
    if (parsed) {
      const worksheet = context.workbook.worksheets.getItemOrNullObject(parsed.worksheetName);
      worksheet.load("isNullObject");
      await context.sync();
      if (!worksheet.isNullObject) {
        try {
          const range = worksheet.getRange(parsed.rangeAddress);
          loadTraceCellProperties(range);
          await context.sync();
          return range;
        } catch {
          // fall through to active-cell fallback
        }
      }
    }
  }

  const active = context.workbook.getActiveCell();
  loadTraceCellProperties(active);
  await context.sync();
  return active;
}
