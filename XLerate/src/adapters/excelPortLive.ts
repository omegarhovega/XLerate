import { CellAddress, CellMutation, CellSnapshot, ExcelPort } from "./excelPort";

/**
 * The only place in the codebase that imports Office.js from domain code.
 * Every other module depends on the ExcelPort interface, never on Excel.run
 * directly. This keeps the surface area of "things that only work in real
 * Excel" to a single file.
 *
 * Known limitation: array formula mutation (kind: "arrayFormula") is not yet
 * implemented for the live adapter. The pure core and the fake adapter handle
 * array formulas correctly; wiring this through Office.js requires
 * Range.setArrayFormula, which will be addressed in Phase 2.
 */
export class ExcelPortLive implements ExcelPort {
  async getSelectionCells(): Promise<CellSnapshot[]> {
    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["rowCount", "columnCount", "rowIndex", "columnIndex", "values", "formulas"]);
      const worksheet = range.worksheet;
      worksheet.load("name");
      await context.sync();

      const snapshots: CellSnapshot[] = [];
      for (let r = 0; r < range.rowCount; r++) {
        for (let c = 0; c < range.columnCount; c++) {
          const formula = range.formulas[r][c];
          const value = range.values[r][c];
          const address: CellAddress = {
            sheet: worksheet.name,
            row: range.rowIndex + r,
            col: range.columnIndex + c
          };
          const formulaText = typeof formula === "string" ? formula : "";
          const isFormula = formulaText.startsWith("=") || formulaText.startsWith("{=");
          const isArrayFormula = formulaText.startsWith("{=");
          snapshots.push({
            address,
            isFormula,
            isArrayFormula,
            formula: isFormula ? formulaText : "",
            value: isFormula ? undefined : value
          });
        }
      }
      return snapshots;
    });
  }

  async applyMutations(mutations: CellMutation[]): Promise<void> {
    if (mutations.length === 0) return;

    const arrayFormulaMutations = mutations.filter((m) => m.kind === "arrayFormula");
    if (arrayFormulaMutations.length > 0) {
      throw new Error(
        "ExcelPortLive: array formula mutation is not yet supported. This will be addressed in Phase 2."
      );
    }

    await Excel.run(async (context) => {
      const sheetCache = new Map<string, Excel.Worksheet>();
      const sheetFor = (name: string): Excel.Worksheet => {
        let s = sheetCache.get(name);
        if (!s) {
          s = context.workbook.worksheets.getItem(name);
          sheetCache.set(name, s);
        }
        return s;
      };

      for (const m of mutations) {
        const cell = sheetFor(m.address.sheet).getRangeByIndexes(m.address.row, m.address.col, 1, 1);
        if (m.kind === "value") {
          cell.values = [[m.value as string | number | boolean | null]];
        } else if (m.kind === "formula") {
          cell.formulas = [[m.formula]];
        }
      }
      await context.sync();
    });
  }
}
