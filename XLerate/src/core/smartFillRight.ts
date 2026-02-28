export type SmartFillCell = {
  isEmpty: boolean;
  isMerged: boolean;
};

export type SmartFillRow = SmartFillCell[];

export type SmartFillActiveCell = {
  row: number;
  col: number;
  formula: string;
  isMerged: boolean;
};

export type SmartFillResult =
  | { ok: true; boundaryCol: number }
  | {
      ok: false;
      reason: "active_cell_must_have_formula" | "active_cell_is_merged" | "no_boundary_found";
    };

function getCell(row: SmartFillRow, col1: number): SmartFillCell | undefined {
  return row[col1 - 1];
}

function hasMergedInContiguousBlock(row: SmartFillRow, startCol: number): boolean {
  let col = startCol;
  let current = getCell(row, col);

  while (current && !current.isEmpty) {
    if (current.isMerged) {
      return true;
    }
    col += 1;
    current = getCell(row, col);
  }

  return false;
}

function findLastCellInRow(row: SmartFillRow, startCol: number): number {
  let current = getCell(row, startCol);
  if (!current || current.isEmpty) {
    return 0;
  }

  let col = startCol;
  let next = getCell(row, col + 1);
  while (next && !next.isEmpty) {
    col += 1;
    next = getCell(row, col + 1);
  }

  return col;
}

export function findSmartFillBoundary(rows: SmartFillRow[], startRow: number, startCol: number): number {
  const maxRowsUp = 3;
  let rowsChecked = 0;
  let currentRow = startRow - 1;

  while (rowsChecked < maxRowsUp && currentRow > 0) {
    const row = rows[currentRow - 1];
    if (row && !hasMergedInContiguousBlock(row, startCol)) {
      const boundary = findLastCellInRow(row, startCol);
      if (boundary > 0) {
        return boundary;
      }
    }

    currentRow -= 1;
    rowsChecked += 1;
  }

  return 0;
}

export function computeSmartFillRight(rows: SmartFillRow[], active: SmartFillActiveCell): SmartFillResult {
  if (active.formula.length === 0 || !active.formula.startsWith("=")) {
    return { ok: false, reason: "active_cell_must_have_formula" };
  }

  if (active.isMerged) {
    return { ok: false, reason: "active_cell_is_merged" };
  }

  const boundaryCol = findSmartFillBoundary(rows, active.row, active.col);
  if (boundaryCol === 0) {
    return { ok: false, reason: "no_boundary_found" };
  }

  return { ok: true, boundaryCol };
}
