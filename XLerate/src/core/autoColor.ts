export type AutoColorCategory =
  | "input"
  | "formula"
  | "worksheetLink"
  | "workbookLink"
  | "external"
  | "hyperlink"
  | "partialInput"
  | "none";

export type AutoColorCell = {
  formula?: string | null;
  value?: unknown;
  numberFormat?: string | null;
  hasHyperlink?: boolean;
};

export type AutoColorPalette = Record<Exclude<AutoColorCategory, "none">, string>;

const CELL_REFERENCE_REGEX = /[$]?[A-Za-z]+[$]?[0-9]+|R[0-9]*C[0-9]*/;
const EXTERNAL_WORKBOOK_REFERENCE_REGEX = /\[[^\]]+\]/;
const SHEET_REFERENCE_REGEX = /(\[[^\]]+\])?'?[^!]+!'?/g;
const A1_REFERENCE_REGEX = /[$]?[A-Za-z]+[$]?[0-9]+/g;
const R1C1_REFERENCE_REGEX = /R[0-9]*C[0-9]*/g;

const COMMON_FUNCTIONS = ["SUM", "AVERAGE", "COUNT", "LEFT", "RIGHT", "MID", "ROUND"];

export const DEFAULT_AUTO_COLOR_PALETTE: AutoColorPalette = {
  input: "#0000FF",
  formula: "#000000",
  worksheetLink: "#008000",
  workbookLink: "#CC99FF",
  external: "#00B0F0",
  hyperlink: "#FF8000",
  partialInput: "#800000",
};

function normalizeFormula(formula: string): string {
  const trimmed = formula.trim();
  if (trimmed.startsWith("{=") && trimmed.endsWith("}")) {
    return trimmed.slice(2, -1);
  }
  if (trimmed.startsWith("=")) {
    return trimmed.slice(1);
  }
  return trimmed;
}

function isFormulaCell(formula: string | null | undefined): boolean {
  if (typeof formula !== "string") {
    return false;
  }
  const trimmed = formula.trim();
  return trimmed.startsWith("=") || (trimmed.startsWith("{=") && trimmed.endsWith("}"));
}

function isDateLikeFormat(numberFormat: string | null | undefined): boolean {
  if (!numberFormat) {
    return false;
  }

  const cleaned = numberFormat
    .toLowerCase()
    .replace(/"[^"]*"/g, "")
    .replace(/\\./g, "");

  return /(\[\$-f800\]|yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|h|s|am\/pm)/i.test(cleaned);
}

function isDateLikeValue(value: unknown, numberFormat: string | null | undefined): boolean {
  if (value instanceof Date) {
    return true;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    return isDateLikeFormat(numberFormat);
  }

  return false;
}

export function isOnlyNumbersAndOperators(formula: string): boolean {
  const normalized = normalizeFormula(formula);
  return /^[-+*/\d\s.,()]*$/.test(normalized);
}

function hasCellReference(formula: string): boolean {
  return CELL_REFERENCE_REGEX.test(normalizeFormula(formula));
}

export function isWorkbookLinkFormula(formula: string): boolean {
  return formula.includes("!") && !EXTERNAL_WORKBOOK_REFERENCE_REGEX.test(formula);
}

export function isWorksheetLinkFormula(formula: string): boolean {
  return (
    hasCellReference(formula) &&
    !formula.includes("!") &&
    !EXTERNAL_WORKBOOK_REFERENCE_REGEX.test(formula)
  );
}

export function isExternalReferenceFormula(formula: string): boolean {
  const upper = formula.toUpperCase();
  return (
    EXTERNAL_WORKBOOK_REFERENCE_REGEX.test(formula) ||
    upper.includes("WEBSERVICE") ||
    upper.includes("ODBC") ||
    upper.includes("SQL")
  );
}

export function isInputCell(cell: AutoColorCell): boolean {
  const value = cell.value;
  const formula = cell.formula ?? null;

  if (value === null || value === undefined || value === "") {
    return false;
  }

  if (typeof value === "string") {
    return false;
  }

  if (isDateLikeValue(value, cell.numberFormat ?? null)) {
    return false;
  }

  if (isFormulaCell(formula)) {
    if (!hasCellReference(formula!)) {
      return true;
    }
    return isOnlyNumbersAndOperators(formula!);
  }

  return true;
}

export function isPartialInputFormula(cell: AutoColorCell): boolean {
  const formula = cell.formula ?? null;
  if (!isFormulaCell(formula)) {
    return false;
  }

  if (typeof cell.value === "string") {
    return false;
  }

  if (isDateLikeValue(cell.value, cell.numberFormat ?? null)) {
    return false;
  }

  if (isOnlyNumbersAndOperators(formula!)) {
    return false;
  }

  // A formula with no cell references cannot be a partial input — it is a
  // pure formula even if it contains numeric literals (e.g. =IF(TRUE,1,0)).
  if (!hasCellReference(formula!)) {
    return false;
  }

  let candidate = normalizeFormula(formula!);
  candidate = candidate.replace(SHEET_REFERENCE_REGEX, "SHEET_REF!");

  for (const func of COMMON_FUNCTIONS) {
    candidate = candidate.replace(new RegExp(func, "g"), "");
  }

  candidate = candidate.replace(/[$%]/g, "");
  candidate = candidate.replace(A1_REFERENCE_REGEX, "");
  candidate = candidate.replace(R1C1_REFERENCE_REGEX, "");

  return /[0-9]+/.test(candidate);
}

export function classifyAutoColorCell(cell: AutoColorCell): AutoColorCategory {
  const formula = cell.formula ?? null;
  if (isFormulaCell(formula)) {
    if (isPartialInputFormula(cell)) {
      return "partialInput";
    }
    if (isWorkbookLinkFormula(formula!)) {
      return "workbookLink";
    }
    if (isWorksheetLinkFormula(formula!)) {
      return "worksheetLink";
    }
    if (isExternalReferenceFormula(formula!)) {
      return "external";
    }
    if (isInputCell(cell)) {
      return "input";
    }
    return "formula";
  }

  if (cell.hasHyperlink) {
    return "hyperlink";
  }

  if (isInputCell(cell)) {
    return "input";
  }

  return "none";
}

export function classifyAutoColorGrid(cells: AutoColorCell[][]): AutoColorCategory[][] {
  return cells.map((row) => row.map((cell) => classifyAutoColorCell(cell)));
}
