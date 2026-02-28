export type SelectionCell = {
  isFormula: boolean;
  formula?: string;
  value?: unknown;
};

export function wrapFormulaWithError(formula: string, errorValue = "NA()"): string {
  const innerFormula = formula.startsWith("=") ? formula.slice(1) : formula;
  return `=IFERROR(${innerFormula}, ${errorValue})`;
}

export function wrapSelectionFormulas(values: SelectionCell[], errorValue = "NA()"): SelectionCell[] {
  return values.map((cell) => {
    if (cell.isFormula && typeof cell.formula === "string") {
      return {
        isFormula: true,
        formula: wrapFormulaWithError(cell.formula, errorValue)
      };
    }

    if ("value" in cell) {
      return { isFormula: false, value: cell.value };
    }

    if ("formula" in cell) {
      return { isFormula: false, formula: cell.formula };
    }

    return { isFormula: false };
  });
}
