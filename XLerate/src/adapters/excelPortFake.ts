import {
  ActiveCellLeftRowSnapshot,
  AutoColorCellSnapshot,
  CellAddress,
  CellFormatMutation,
  CellFormattingSnapshot,
  CellMutation,
  CellSnapshot,
  ExcelPort,
} from "./excelPort";

type FakeCellState =
  | { kind: "empty" }
  | { kind: "value"; value: unknown }
  | { kind: "formula"; formula: string; isArray: boolean };

type FakeCellFormatting = Omit<CellFormattingSnapshot, "address">;

function key(a: CellAddress): string {
  return `${a.sheet}!${a.row},${a.col}`;
}

function defaultFormatting(): FakeCellFormatting {
  return {
    numberFormat: "General",
    hasHyperlink: false,
    fillPattern: null,
    fillColor: null,
    fontName: null,
    fontSize: null,
    fontColor: null,
    fontBold: null,
    fontItalic: null,
    fontUnderline: null,
    fontStrikethrough: null,
    edgeLeftStyle: null,
    edgeTopStyle: null,
    edgeBottomStyle: null,
    edgeRightStyle: null,
    edgeLeftColor: null,
    edgeTopColor: null,
    edgeBottomColor: null,
    edgeRightColor: null,
  };
}

export class ExcelPortFake implements ExcelPort {
  private cells = new Map<string, FakeCellState>();
  private formatting = new Map<string, FakeCellFormatting>();
  private selection: CellAddress[] = [];

  setCellValue(address: CellAddress, value: unknown): void {
    this.cells.set(key(address), { kind: "value", value });
  }

  setCellFormula(address: CellAddress, formula: string, isArray = false): void {
    this.cells.set(key(address), { kind: "formula", formula, isArray });
  }

  setSelection(addresses: CellAddress[]): void {
    this.selection = [...addresses];
  }

  /** Test helper: patch formatting fields on a cell. Unspecified fields keep their defaults. */
  setCellFormatting(address: CellAddress, patch: Partial<FakeCellFormatting>): void {
    const existing = this.formatting.get(key(address)) ?? defaultFormatting();
    this.formatting.set(key(address), { ...existing, ...patch });
  }

  setCellHyperlink(address: CellAddress, hasHyperlink: boolean): void {
    this.setCellFormatting(address, { hasHyperlink });
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
          value: undefined,
        };
      }
      if (state.kind === "value") {
        return {
          address,
          isFormula: false,
          isArrayFormula: false,
          formula: "",
          value: state.value,
        };
      }
      return {
        address,
        isFormula: false,
        isArrayFormula: false,
        formula: "",
        value: null,
      };
    });
  }

  async getSelectionFormatting(): Promise<CellFormattingSnapshot[]> {
    return this.selection.map((address) => {
      const fmt = this.formatting.get(key(address)) ?? defaultFormatting();
      return { address, ...fmt };
    });
  }

  async getActiveCellLeftRowSnapshot(): Promise<ActiveCellLeftRowSnapshot> {
    const activeCell = this.selection[0] ?? { sheet: "Sheet1", row: 0, col: 0 };
    const leftCells = Array.from({ length: activeCell.col }, (_, index) => {
      const address = { sheet: activeCell.sheet, row: activeCell.row, col: index };
      const state = this.cells.get(key(address)) ?? { kind: "empty" as const };
      return {
        address,
        value: state.kind === "value" ? state.value : null,
      };
    });

    return {
      activeCell,
      leftCells,
    };
  }

  async getSelectionAutoColorCells(): Promise<AutoColorCellSnapshot[]> {
    return this.selection
      .map((address) => {
        const state = this.cells.get(key(address)) ?? { kind: "empty" as const };
        const fmt = this.formatting.get(key(address)) ?? defaultFormatting();
        const isFormula = state.kind === "formula";
        const value = state.kind === "value" ? state.value : null;
        const hasMeaningfulValue =
          isFormula || (value !== null && value !== undefined && value !== "");

        if (!hasMeaningfulValue && !fmt.hasHyperlink) {
          return null;
        }

        return {
          address,
          isFormula,
          formula: isFormula ? state.formula : "",
          value: isFormula ? undefined : value,
          numberFormat: fmt.numberFormat,
          hasHyperlink: fmt.hasHyperlink,
        };
      })
      .filter((cell): cell is AutoColorCellSnapshot => cell !== null);
  }

  async applySelectionFormatBundle(format: CellFormatMutation): Promise<void> {
    await this.applyMutations(
      this.selection.map((address) => ({
        address,
        kind: "formatBundle" as const,
        format,
      }))
    );
  }

  async applyMutations(mutations: CellMutation[]): Promise<void> {
    for (const m of mutations) {
      if (m.kind === "value") {
        this.cells.set(key(m.address), { kind: "value", value: m.value });
      } else if (m.kind === "formula") {
        this.cells.set(key(m.address), { kind: "formula", formula: m.formula, isArray: false });
      } else if (m.kind === "arrayFormula") {
        this.cells.set(key(m.address), { kind: "formula", formula: m.formula, isArray: true });
      } else if (m.kind === "numberFormat") {
        this.setCellFormatting(m.address, { numberFormat: m.format });
      } else if (m.kind === "fontColor") {
        this.setCellFormatting(m.address, { fontColor: m.color });
      } else if (m.kind === "formatBundle") {
        const fmt = m.format;
        const patch: Partial<FakeCellFormatting> = {};
        if (fmt.numberFormat !== undefined) {
          patch.numberFormat = fmt.numberFormat;
        }
        if (fmt.fill) {
          // Mirror Excel.RangeFill.clear(): pattern "None" clears BOTH pattern
          // and color, since the live adapter invokes fill.clear() for this
          // case. Keeping the fake aligned prevents contract tests from
          // disagreeing with live behavior.
          if (fmt.fill.pattern === "None") {
            patch.fillPattern = null;
            patch.fillColor = null;
          } else {
            if (fmt.fill.pattern !== undefined) {
              patch.fillPattern = fmt.fill.pattern;
            }
            if (fmt.fill.color !== undefined) {
              patch.fillColor = fmt.fill.color;
            }
          }
        }
        if (fmt.font) {
          if (fmt.font.name !== undefined) patch.fontName = fmt.font.name;
          if (fmt.font.size !== undefined) patch.fontSize = fmt.font.size;
          if (fmt.font.color !== undefined) patch.fontColor = fmt.font.color;
          if (fmt.font.bold !== undefined) patch.fontBold = fmt.font.bold;
          if (fmt.font.italic !== undefined) patch.fontItalic = fmt.font.italic;
          if (fmt.font.underline !== undefined) patch.fontUnderline = fmt.font.underline;
          if (fmt.font.strikethrough !== undefined)
            patch.fontStrikethrough = fmt.font.strikethrough;
        }
        if (fmt.borders) {
          if (fmt.borders.clearAll) {
            patch.edgeLeftStyle = null;
            patch.edgeTopStyle = null;
            patch.edgeBottomStyle = null;
            patch.edgeRightStyle = null;
            patch.edgeLeftColor = null;
            patch.edgeTopColor = null;
            patch.edgeBottomColor = null;
            patch.edgeRightColor = null;
          }
          if (fmt.borders.left) {
            patch.edgeLeftStyle = fmt.borders.left.style;
            if (fmt.borders.left.color !== undefined) patch.edgeLeftColor = fmt.borders.left.color;
          }
          if (fmt.borders.top) {
            patch.edgeTopStyle = fmt.borders.top.style;
            if (fmt.borders.top.color !== undefined) patch.edgeTopColor = fmt.borders.top.color;
          }
          if (fmt.borders.bottom) {
            patch.edgeBottomStyle = fmt.borders.bottom.style;
            if (fmt.borders.bottom.color !== undefined)
              patch.edgeBottomColor = fmt.borders.bottom.color;
          }
          if (fmt.borders.right) {
            patch.edgeRightStyle = fmt.borders.right.style;
            if (fmt.borders.right.color !== undefined)
              patch.edgeRightColor = fmt.borders.right.color;
          }
        }
        this.setCellFormatting(m.address, patch);
      }
    }
  }

  async clearSheetFill(sheetName: string): Promise<void> {
    for (const [k, fmt] of this.formatting) {
      if (k.startsWith(`${sheetName}!`)) {
        this.formatting.set(k, { ...fmt, fillPattern: null, fillColor: null });
      }
    }
  }
}
