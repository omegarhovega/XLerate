import type { TraceTarget } from "./traceGraph";

export type FormulaTextSegment = {
  kind: "text";
  text: string;
};

export type FormulaReferenceSegment = {
  kind: "reference";
  text: string;
  target: TraceTarget;
};

export type FormulaSegment = FormulaTextSegment | FormulaReferenceSegment;

const A1_REFERENCE_AT_START =
  /^(?:(?<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(?<range>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})?)/;

function isIdentifierCharacter(value: string | undefined): boolean {
  return value !== undefined && /[A-Za-z0-9_.]/.test(value);
}

function columnNumber(letters: string): number {
  let result = 0;
  for (const character of letters.toUpperCase()) {
    result = result * 26 + (character.charCodeAt(0) - 64);
  }
  return result;
}

function isValidCellAddress(address: string): boolean {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$/.exec(address);
  if (!match) return false;
  return columnNumber(match[1]) <= 16_384 && Number(match[2]) <= 1_048_576;
}

function isValidRangeAddress(address: string): boolean {
  return address.split(":").every(isValidCellAddress);
}

function unquoteWorksheetName(token: string): string {
  if (token.startsWith("'") && token.endsWith("'")) {
    return token.slice(1, -1).replace(/''/g, "'");
  }
  return token;
}

function quoteWorksheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

function pushText(segments: FormulaSegment[], text: string): void {
  if (text.length === 0) return;
  const previous = segments[segments.length - 1];
  if (previous?.kind === "text") previous.text += text;
  else segments.push({ kind: "text", text });
}

function stringLiteralEnd(formula: string, start: number): number {
  let index = start + 1;
  while (index < formula.length) {
    if (formula[index] !== '"') {
      index += 1;
      continue;
    }
    if (formula[index + 1] === '"') {
      index += 2;
      continue;
    }
    return index + 1;
  }
  return formula.length;
}

export function tokenizeFormulaReferences(
  formula: string,
  contextWorksheetName: string
): FormulaSegment[] {
  if (formula.length === 0) return [];

  const segments: FormulaSegment[] = [];
  let index = 0;
  while (index < formula.length) {
    if (formula[index] === '"') {
      const end = stringLiteralEnd(formula, index);
      pushText(segments, formula.slice(index, end));
      index = end;
      continue;
    }

    const candidate = A1_REFERENCE_AT_START.exec(formula.slice(index));
    if (!candidate?.groups) {
      pushText(segments, formula[index]);
      index += 1;
      continue;
    }

    const token = candidate[0];
    const sheetToken = candidate.groups.sheet;
    const rangeAddress = candidate.groups.range;
    const previous = index > 0 ? formula[index - 1] : undefined;
    const next = formula[index + token.length];
    const invalidBoundary =
      isIdentifierCharacter(previous) ||
      isIdentifierCharacter(next) ||
      previous === "[" ||
      previous === "]" ||
      (previous === ":" && sheetToken !== undefined) ||
      next === "[" ||
      next === "(";
    const externalReference = sheetToken?.includes("[") || sheetToken?.includes("]");

    if (invalidBoundary || externalReference || !isValidRangeAddress(rangeAddress)) {
      pushText(segments, formula[index]);
      index += 1;
      continue;
    }

    const worksheetName = sheetToken ? unquoteWorksheetName(sheetToken) : contextWorksheetName;
    segments.push({
      kind: "reference",
      text: token,
      target: {
        address: `${quoteWorksheetName(worksheetName)}!${rangeAddress}`,
      },
    });
    index += token.length;
  }

  return segments;
}
