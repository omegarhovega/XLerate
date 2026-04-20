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

export type FontMutation = {
  name?: string;
  size?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
};

export type FillMutation = {
  pattern?: "Solid" | "None";
  color?: string;
};

export type BorderEdgeMutation = {
  style: "None" | "Continuous" | "Dash" | "Dot" | "Double";
  color?: string;
  weight?: "Thin" | "Medium" | "Thick";
};

export type BordersMutation = {
  /** If true, clear all edges before applying the supplied ones. */
  clearAll?: boolean;
  left?: BorderEdgeMutation;
  top?: BorderEdgeMutation;
  bottom?: BorderEdgeMutation;
  right?: BorderEdgeMutation;
  /** Only meaningful for multi-cell ranges; ignored by the fake for single cells. */
  insideHorizontal?: BorderEdgeMutation;
  insideVertical?: BorderEdgeMutation;
};

export type CellFormatMutation = {
  numberFormat?: string;
  font?: FontMutation;
  fill?: FillMutation;
  borders?: BordersMutation;
};

export type CellFormattingSnapshot = {
  address: CellAddress;
  numberFormat: string;
  hasHyperlink: boolean;
  fillPattern: string | null;
  fillColor: string | null;
  fontName: string | null;
  fontSize: number | null;
  fontColor: string | null;
  fontBold: boolean | null;
  fontItalic: boolean | null;
  fontUnderline: boolean | null;
  fontStrikethrough: boolean | null;
  edgeLeftStyle: string | null;
  edgeTopStyle: string | null;
  edgeBottomStyle: string | null;
  edgeRightStyle: string | null;
  edgeLeftColor: string | null;
  edgeTopColor: string | null;
  edgeBottomColor: string | null;
  edgeRightColor: string | null;
};

export type ActiveCellLeftRowSnapshot = {
  activeCell: CellAddress;
  leftCells: Array<{
    address: CellAddress;
    value: unknown;
  }>;
};

export type AutoColorCellSnapshot = {
  address: CellAddress;
  isFormula: boolean;
  formula: string;
  value: unknown;
  numberFormat: string;
  hasHyperlink: boolean;
};

export type CellMutation =
  | { address: CellAddress; kind: "value"; value: unknown }
  | { address: CellAddress; kind: "formula"; formula: string }
  | { address: CellAddress; kind: "arrayFormula"; formula: string }
  | { address: CellAddress; kind: "numberFormat"; format: string }
  | { address: CellAddress; kind: "fontColor"; color: string }
  | { address: CellAddress; kind: "formatBundle"; format: CellFormatMutation };

export interface ExcelPort {
  /** Read the current user selection as individual cell snapshots, flattened in row-major order. */
  getSelectionCells(): Promise<CellSnapshot[]>;

  /**
   * Apply a batch of mutations atomically (one native Excel undo step).
   * Empty mutation arrays are a no-op. Ordering within the batch is
   * observationally irrelevant — mutations target distinct cells.
   */
  applyMutations(mutations: CellMutation[]): Promise<void>;

  /**
   * Read formatting for each cell in the current selection. Separate from
   * getSelectionCells because loading all formatting axes is expensive; features
   * that only need value/formula should not pay that cost.
   */
  getSelectionFormatting(): Promise<CellFormattingSnapshot[]>;

  /**
   * Read the active cell plus every cell to its left on the same row up to the
   * used-range boundary, ordered left-to-right. Used by the in-sheet CAGR
   * action to discover the contiguous numeric series adjacent to the
   * destination cell.
   */
  getActiveCellLeftRowSnapshot(): Promise<ActiveCellLeftRowSnapshot>;

  /**
   * Read the non-empty cells in the current selection with just the metadata
   * needed by Auto-color. This intentionally skips blank cells so large
   * selections do not degenerate into full-sheet scans.
   */
  getSelectionAutoColorCells(): Promise<AutoColorCellSnapshot[]>;

  /**
   * Apply one uniform format bundle to the current selection. This is used by
   * formatting cycles where every selected cell receives the same style and is
   * substantially faster than issuing one mutation per cell.
   */
  applySelectionFormatBundle(format: CellFormatMutation): Promise<void>;

  /**
   * Remove all fill color from every cell on the named sheet. Used by
   * Clear Consistency Marks (spec §3.6). Single Excel undo step.
   */
  clearSheetFill(sheetName: string): Promise<void>;
}
