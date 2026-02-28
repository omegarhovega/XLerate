export type CellInput = {
  isFormula: boolean;
  isArrayFormula?: boolean;
  formula?: string;
  value?: unknown;
};

export function switchSignCell(input: CellInput): CellInput {
  if (input.isFormula) {
    const isArrayFormula = Boolean(input.isArrayFormula);
    let formulaText = input.formula ?? "";

    if (isArrayFormula && formulaText.startsWith("{") && formulaText.endsWith("}")) {
      formulaText = formulaText.slice(1, -1);
    }

    if (formulaText.startsWith("=")) {
      formulaText = `=-(${formulaText.slice(1)})`;
    } else {
      formulaText = `-(${formulaText})`;
    }

    if (isArrayFormula) {
      return {
        isFormula: true,
        isArrayFormula: true,
        formula: `{${formulaText}}`
      };
    }

    return {
      isFormula: true,
      formula: formulaText
    };
  }

  const value = input.value;
  const isNumericValue = typeof value === "number" && Number.isFinite(value);

  if (isNumericValue) {
    const negated = -(value as number);
    return {
      isFormula: false,
      value: Object.is(negated, -0) ? 0 : negated
    };
  }

  return {
    isFormula: false,
    value
  };
}
