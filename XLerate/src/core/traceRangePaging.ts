export type TraceRangePageBlock = {
  rowOffset: number;
  columnOffset: number;
  rowCount: number;
  columnCount: number;
  cellCount: number;
};

/**
 * Convert a flat row-major page into at most three rectangular reads:
 * a partial first row, a block of complete rows, and a partial last row.
 */
export function planTraceRangePageBlocks(
  sourceRowCount: number,
  sourceColumnCount: number,
  start: number,
  count: number
): TraceRangePageBlock[] {
  if (
    !Number.isInteger(sourceRowCount) ||
    !Number.isInteger(sourceColumnCount) ||
    sourceRowCount <= 0 ||
    sourceColumnCount <= 0 ||
    !Number.isInteger(start) ||
    !Number.isInteger(count) ||
    start < 0 ||
    count <= 0
  ) {
    return [];
  }

  const totalCells = sourceRowCount * sourceColumnCount;
  if (start >= totalCells) return [];

  let flatIndex = start;
  let remaining = Math.min(count, totalCells - start);
  const blocks: TraceRangePageBlock[] = [];
  const startingColumn = flatIndex % sourceColumnCount;

  if (startingColumn > 0) {
    const firstCount = Math.min(remaining, sourceColumnCount - startingColumn);
    blocks.push({
      rowOffset: Math.floor(flatIndex / sourceColumnCount),
      columnOffset: startingColumn,
      rowCount: 1,
      columnCount: firstCount,
      cellCount: firstCount,
    });
    flatIndex += firstCount;
    remaining -= firstCount;
  }

  const completeRows = Math.floor(remaining / sourceColumnCount);
  if (completeRows > 0) {
    const fullCellCount = completeRows * sourceColumnCount;
    blocks.push({
      rowOffset: Math.floor(flatIndex / sourceColumnCount),
      columnOffset: 0,
      rowCount: completeRows,
      columnCount: sourceColumnCount,
      cellCount: fullCellCount,
    });
    flatIndex += fullCellCount;
    remaining -= fullCellCount;
  }

  if (remaining > 0) {
    blocks.push({
      rowOffset: Math.floor(flatIndex / sourceColumnCount),
      columnOffset: flatIndex % sourceColumnCount,
      rowCount: 1,
      columnCount: remaining,
      cellCount: remaining,
    });
  }

  return blocks;
}
