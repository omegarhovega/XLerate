export type FormulaConsistencyCell = {
  isFormula: boolean;
  formulaR1C1?: string;
};

export type FormulaConsistencyMark = "consistent" | "inconsistent" | "none";

export function collectAdjacentEqualFormulas(rows: FormulaConsistencyCell[][]): Set<string> {
  const consistentFormulas = new Set<string>();

  for (const row of rows) {
    for (let c = 0; c < row.length - 1; c += 1) {
      const current = row[c];
      const right = row[c + 1];

      if (!current?.isFormula || !right?.isFormula) {
        continue;
      }

      const currentFormula = current.formulaR1C1 ?? "";
      const rightFormula = right.formulaR1C1 ?? "";
      if (currentFormula.length > 0 && currentFormula === rightFormula) {
        consistentFormulas.add(currentFormula);
      }
    }
  }

  return consistentFormulas;
}

export function analyzeHorizontalFormulaConsistency(rows: FormulaConsistencyCell[][]): FormulaConsistencyMark[][] {
  const consistentFormulas = collectAdjacentEqualFormulas(rows);

  return rows.map((row) =>
    row.map((cell, c) => {
      if (!cell?.isFormula) {
        return "none";
      }

      const currentFormula = cell.formulaR1C1 ?? "";
      const right = c + 1 < row.length ? row[c + 1] : undefined;

      if (right?.isFormula) {
        const rightFormula = right.formulaR1C1 ?? "";
        return currentFormula.length > 0 && currentFormula === rightFormula ? "consistent" : "inconsistent";
      }

      return currentFormula.length > 0 && consistentFormulas.has(currentFormula) ? "consistent" : "none";
    })
  );
}
