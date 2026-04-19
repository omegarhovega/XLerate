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
  /**
   * Returns neighbors for every cell in `cells`, batched. `result[i]`
   * is the neighbor list of `cells[i]`. Batching is the whole point of
   * this signature: the live implementation does one `context.sync()`
   * pair per BFS level instead of one per cell, turning O(N) Office.js
   * round-trips into O(depth). Test doubles can ignore batching and
   * just return per-cell lists; correctness is identical.
   */
  getAllNeighbors: (cells: TraceCellInfo[]) => Promise<TraceCellInfo[][]>;
  /**
   * Optional progressive-loading callback. Fires once per BFS level,
   * starting with the root alone (level 0). Consumer receives the
   * full cumulative `rows` plus `isFinal` — when `isFinal` is true
   * this is the last callback and no further emissions will occur.
   * The builder `await`s each callback, so the consumer can drive
   * IPC (e.g. `dialog.messageChild`) and yield a frame before the
   * next BFS level starts. Skipping the callback (not passing
   * `onProgress`) preserves the pre-existing all-at-once behavior.
   */
  onProgress?: (progress: TraceBuilderProgress) => void | Promise<void>;
};

export type TraceBuilderProgress = {
  /** Cumulative rows discovered so far. Safe to retain; the builder passes a copy. */
  rows: TraceRow[];
  /** Deepest BFS level included in `rows`. Root is 0. */
  level: number;
  /** True when this is the last `onProgress` call; no further emissions follow. */
  isFinal: boolean;
  /** True when the row cap was hit and BFS stopped early. */
  truncated: boolean;
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
 * Pure BFS over the trace graph, processing one BFS level per round-trip
 * through `getAllNeighbors`. Visitation order and truncation semantics
 * match the pre-batching behavior of the prior signature — the only
 * observable change is that all cells at depth N are passed to the
 * neighbor callback together, so a live Office.js caller can batch
 * their loads. Cycle prevention uses `(worksheetName, rowIndex,
 * columnIndex)`; row cap wins over depth cap when both apply.
 */
export async function buildTrace(input: TraceBuilderInput): Promise<TraceBuilderResult> {
  const { root, maxDepth, getAllNeighbors, onProgress } = input;
  const maxRows = input.maxRows ?? MAX_TRACE_ROWS;

  const rows: TraceRow[] = [toTraceRow(root, 0)];
  const visited = new Set<string>([
    buildTraceCellKey(root.worksheetName, root.rowIndex, root.columnIndex),
  ]);
  let truncated = false;

  // Emit the root immediately so a progressive-loading consumer can
  // render "level 0" before any neighbor lookups run. For maxDepth=0
  // this is also the final emit.
  if (onProgress) {
    await onProgress({
      rows: [...rows],
      level: 0,
      isFinal: maxDepth === 0,
      truncated: false,
    });
  }

  // At each iteration, `currentLevelCells` holds all cells at depth
  // `level`, in the order they were first discovered. We ask for all
  // their neighbors in one call, then promote novel neighbors to the
  // next level.
  let currentLevelCells: TraceCellInfo[] = [root];
  for (let level = 0; level < maxDepth && currentLevelCells.length > 0 && !truncated; level += 1) {
    const neighborLists = await getAllNeighbors(currentLevelCells);
    const nextLevelCells: TraceCellInfo[] = [];

    for (let i = 0; i < neighborLists.length && !truncated; i += 1) {
      const neighbors = neighborLists[i] ?? [];
      for (const neighbor of neighbors) {
        const key = buildTraceCellKey(
          neighbor.worksheetName,
          neighbor.rowIndex,
          neighbor.columnIndex
        );
        if (visited.has(key)) continue;
        visited.add(key);

        rows.push(toTraceRow(neighbor, level + 1));
        if (rows.length >= maxRows) {
          truncated = true;
          break;
        }
        nextLevelCells.push(neighbor);
      }
    }

    currentLevelCells = nextLevelCells;

    if (onProgress) {
      const willBeFinal =
        truncated || currentLevelCells.length === 0 || level + 1 >= maxDepth;
      await onProgress({
        rows: [...rows],
        level: level + 1,
        isFinal: willBeFinal,
        truncated,
      });
    }
  }

  return { rows, truncated };
}
