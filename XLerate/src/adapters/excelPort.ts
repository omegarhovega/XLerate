/**
 * Boundary between domain code and Excel. Services depend on this interface,
 * never on Office.js directly. The live implementation is the only place
 * that calls Excel.run; the fake implementation drives tests.
 */

export type CellAddress = {
  sheet: string;
  /** Zero-based absolute row index on the sheet. */
  row: number;
  /** Zero-based absolute column index on the sheet. */
  col: number;
};

export type CellSnapshot = {
  address: CellAddress;
  isFormula: boolean;
  isArrayFormula: boolean;
  /** Formula text including the leading '='. For array formulas, includes surrounding braces. Empty string when not a formula. */
  formula: string;
  /** Current value when the cell is not a formula. Undefined when it is. */
  value: unknown;
};

export type CellMutation =
  | { address: CellAddress; kind: "value"; value: unknown }
  | { address: CellAddress; kind: "formula"; formula: string }
  | { address: CellAddress; kind: "arrayFormula"; formula: string };

export interface ExcelPort {
  /** Read the current user selection as individual cell snapshots, flattened in row-major order. */
  getSelectionCells(): Promise<CellSnapshot[]>;

  /**
   * Apply a batch of mutations atomically (one native Excel undo step).
   * Empty mutation arrays are a no-op. Ordering within the batch is
   * observationally irrelevant — mutations target distinct cells.
   */
  applyMutations(mutations: CellMutation[]): Promise<void>;
}
