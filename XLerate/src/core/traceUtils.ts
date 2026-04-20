export type TraceDirection = "precedents" | "dependents";
export type ParsedTraceAddress = {
  worksheetName: string;
  rangeAddress: string;
};

export const DEFAULT_TRACE_MAX_DEPTH = 10;
export const MAX_TRACE_MAX_DEPTH = 20;
export const DEFAULT_TRACE_SAFETY_LIMIT = 500;
export const MAX_TRACE_SAFETY_LIMIT = 5000;
export const MAX_TRACE_ROWS = DEFAULT_TRACE_SAFETY_LIMIT;

export function sanitizeTraceDepth(raw: unknown): number {
  if (typeof raw !== "number" || !Number.isFinite(raw)) {
    return DEFAULT_TRACE_MAX_DEPTH;
  }

  const normalized = Math.trunc(raw);
  if (normalized < 1) {
    return 1;
  }

  if (normalized > MAX_TRACE_MAX_DEPTH) {
    return MAX_TRACE_MAX_DEPTH;
  }

  return normalized;
}

export function sanitizeTraceSafetyLimit(raw: unknown): number {
  if (typeof raw !== "number" || !Number.isFinite(raw)) {
    return DEFAULT_TRACE_SAFETY_LIMIT;
  }

  const normalized = Math.trunc(raw);
  if (normalized < 1) {
    return 1;
  }

  if (normalized > MAX_TRACE_SAFETY_LIMIT) {
    return MAX_TRACE_SAFETY_LIMIT;
  }

  return normalized;
}

export function scalarFromMatrix(value: unknown): unknown {
  if (!Array.isArray(value) || value.length === 0) {
    return value;
  }

  const row = value[0];
  if (!Array.isArray(row) || row.length === 0) {
    return value;
  }

  return row[0];
}

export function formatTraceValue(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "string") {
    return value;
  }

  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }

  if (typeof value === "object") {
    const maybeError = value as { error?: unknown };
    if (typeof maybeError.error === "string") {
      return maybeError.error;
    }

    try {
      return JSON.stringify(value);
    } catch {
      return String(value);
    }
  }

  return String(value);
}

export function formatTraceFormula(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "string") {
    return value.startsWith("=") ? value : "";
  }

  return "";
}

export function buildTraceCellKey(
  worksheetName: string,
  rowIndex: number,
  columnIndex: number
): string {
  return `${worksheetName}!R${rowIndex}C${columnIndex}`;
}

export function parseWorksheetScopedAddress(address: string): ParsedTraceAddress | null {
  const trimmed = address.trim();
  const bang = trimmed.lastIndexOf("!");
  if (bang <= 0 || bang >= trimmed.length - 1) {
    return null;
  }

  let worksheetName = trimmed.slice(0, bang).trim();
  const rangeAddress = trimmed.slice(bang + 1).trim();
  if (!worksheetName || !rangeAddress) {
    return null;
  }

  // Excel escapes single quotes in worksheet names by doubling them.
  if (worksheetName.startsWith("'") && worksheetName.endsWith("'") && worksheetName.length >= 2) {
    worksheetName = worksheetName.slice(1, -1).replace(/''/g, "'");
  }

  // If workbook prefix is present, strip it: [Book1]Sheet1 -> Sheet1.
  if (worksheetName.startsWith("[")) {
    const close = worksheetName.indexOf("]");
    if (close > 0 && close < worksheetName.length - 1) {
      worksheetName = worksheetName.slice(close + 1);
    }
  }

  if (!worksheetName) {
    return null;
  }

  return { worksheetName, rangeAddress };
}
