import {
  MAX_TRACE_ROWS,
  buildTraceCellKey,
  formatTraceFormula,
  formatTraceValue,
  scalarFromMatrix,
} from "./traceUtils";

/**
 * One row in the trace result list, independent of how the cells were
 * discovered. Shared between the taskpane table and the Phase B trace
 * dialog.
 */
export type TraceRow = {
  level: number;
  address: string;
  value: string;
  formula: string;
};

/**
 * Plain-data snapshot of a traced cell. The `Excel.Range`-bound callers
 * (taskpane / trace dialog) convert live Range instances to these before
 * passing them to the BFS. Identified by `(worksheetName, rowIndex,
 * columnIndex)` — the triple is the cycle-prevention key.
 *
 * `value` / `formula` are the raw payloads as seen via Office.js (matrix or
 * scalar); the builder normalizes via `scalarFromMatrix` + the formatters.
 */
export type TraceCellInfo = {
  worksheetName: string;
  rowIndex: number;
  columnIndex: number;
  /** Display address, e.g. `"Sheet1!B5"`. */
  address: string;
  value: unknown;
  formula: unknown;
};

export type TraceBuilderInput = {
  root: TraceCellInfo;
  /** Hard cap on BFS levels. Root is level 0. */
  maxDepth: number;
  /** Hard cap on total rows, including the root. Defaults to `MAX_TRACE_ROWS`. */
  maxRows?: number;
  /** Returns the direct precedents (or dependents) of the given cell. */
  getNeighbors: (cell: TraceCellInfo) => Promise<TraceCellInfo[]>;
};

export type TraceBuilderResult = {
  rows: TraceRow[];
  /** True if BFS stopped early because `maxRows` was hit. */
  truncated: boolean;
};

export function toTraceRow(cell: TraceCellInfo, level: number): TraceRow {
  return {
    level,
    address: cell.address,
    value: formatTraceValue(scalarFromMatrix(cell.value)),
    formula: formatTraceFormula(scalarFromMatrix(cell.formula)),
  };
}

/**
 * Pure BFS over the trace graph. The Office.js-specific neighbor lookup is
 * passed in via `getNeighbors` so this function is testable without any
 * Excel runtime. Cycle prevention uses `(worksheetName, rowIndex,
 * columnIndex)`; depth / row caps match the pre-refactor behavior of
 * `runTrace` in `taskpane.ts`.
 */
export async function buildTrace(input: TraceBuilderInput): Promise<TraceBuilderResult> {
  const { root, maxDepth, getNeighbors } = input;
  const maxRows = input.maxRows ?? MAX_TRACE_ROWS;

  const rows: TraceRow[] = [toTraceRow(root, 0)];
  const visited = new Set<string>([
    buildTraceCellKey(root.worksheetName, root.rowIndex, root.columnIndex),
  ]);
  const queue: Array<{ level: number; cell: TraceCellInfo }> = [{ level: 0, cell: root }];
  let truncated = false;

  while (queue.length > 0) {
    const current = queue.shift();
    if (!current) break;
    if (current.level >= maxDepth) continue;

    const neighbors = await getNeighbors(current.cell);
    for (const neighbor of neighbors) {
      const key = buildTraceCellKey(neighbor.worksheetName, neighbor.rowIndex, neighbor.columnIndex);
      if (visited.has(key)) continue;
      visited.add(key);

      rows.push(toTraceRow(neighbor, current.level + 1));
      if (rows.length >= maxRows) {
        truncated = true;
        break;
      }
      queue.push({ level: current.level + 1, cell: neighbor });
    }

    if (truncated) break;
  }

  return { rows, truncated };
}
