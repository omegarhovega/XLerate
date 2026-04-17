import { CellAddress, CellMutation, CellSnapshot, ExcelPort } from "./excelPort";

type FakeCellState =
  | { kind: "empty" }
  | { kind: "value"; value: unknown }
  | { kind: "formula"; formula: string; isArray: boolean };

function key(a: CellAddress): string {
  return `${a.sheet}!${a.row},${a.col}`;
}

export class ExcelPortFake implements ExcelPort {
  private cells = new Map<string, FakeCellState>();
  private selection: CellAddress[] = [];

  /** Seed a cell with a value. Test helper only. */
  setCellValue(address: CellAddress, value: unknown): void {
    this.cells.set(key(address), { kind: "value", value });
  }

  /** Seed a cell with a formula. Test helper only. */
  setCellFormula(address: CellAddress, formula: string, isArray = false): void {
    this.cells.set(key(address), { kind: "formula", formula, isArray });
  }

  /** Set which cells are "selected". Test helper only. */
  setSelection(addresses: CellAddress[]): void {
    this.selection = [...addresses];
  }

  async getSelectionCells(): Promise<CellSnapshot[]> {
    return this.selection.map((address) => {
      const state = this.cells.get(key(address)) ?? { kind: "empty" as const };
      if (state.kind === "formula") {
        return {
          address,
          isFormula: true,
          isArrayFormula: state.isArray,
          formula: state.formula,
          value: undefined
        };
      }
      if (state.kind === "value") {
        return {
          address,
          isFormula: false,
          isArrayFormula: false,
          formula: "",
          value: state.value
        };
      }
      return {
        address,
        isFormula: false,
        isArrayFormula: false,
        formula: "",
        value: null
      };
    });
  }

  async applyMutations(mutations: CellMutation[]): Promise<void> {
    for (const m of mutations) {
      if (m.kind === "value") {
        this.cells.set(key(m.address), { kind: "value", value: m.value });
      } else if (m.kind === "formula") {
        this.cells.set(key(m.address), { kind: "formula", formula: m.formula, isArray: false });
      } else {
        this.cells.set(key(m.address), { kind: "formula", formula: m.formula, isArray: true });
      }
    }
  }
}
