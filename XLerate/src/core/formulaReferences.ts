import type { TraceTableSection, TraceTableTarget, TraceTarget } from "./traceGraph";

export type FormulaTextSegment = {
  kind: "text";
  text: string;
};

export type FormulaReferenceSegment = {
  kind: "reference";
  referenceKind?: "namedRange" | "table";
  text: string;
  target: TraceTarget;
};

export type FormulaExternalReferenceSegment = {
  kind: "external";
  text: string;
  workbookName: string;
  worksheetName: string;
  rangeAddress: string;
  reason: string;
};

export type FormulaSegment =
  | FormulaTextSegment
  | FormulaReferenceSegment
  | FormulaExternalReferenceSegment;

const CELL_PATTERN = "\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]{0,6}";
const RANGE_PATTERN = `${CELL_PATTERN}(?::${CELL_PATTERN})?`;
const SHEET_PATTERN = "'(?:[^']|'')+'|[A-Za-z_\\\\][A-Za-z0-9_.]*";
const NAME_PATTERN = "[A-Za-z_\\\\][A-Za-z0-9_.]*";
const A1_REFERENCE_AT_START = new RegExp(
  `^(?:(?<sheet>${SHEET_PATTERN})!)?(?<range>${RANGE_PATTERN})`
);
const NAMED_REFERENCE_AT_START = new RegExp(
  `^(?:(?<sheet>${SHEET_PATTERN})!)?(?<name>${NAME_PATTERN})`
);
const IDENTIFIER_AT_START = new RegExp(`^(?<name>${NAME_PATTERN})`);
const UNQUOTED_EXTERNAL_AT_START = new RegExp(
  `^\\[(?<workbook>[^\\]]+)\\](?<sheet>[A-Za-z_][A-Za-z0-9_.]*)!(?<range>${RANGE_PATTERN})`
);
const QUOTED_EXTERNAL_AT_START = new RegExp(
  `^'(?<path>(?:[^']|'')*?)\\[(?<workbook>[^\\]]+)\\](?<sheet>(?:[^']|'')+)'!(?<range>${RANGE_PATTERN})`
);
const EXTERNAL_REASON =
  "External workbook navigation is not available in the Excel JavaScript API.";
const RESERVED_NAMES = new Set(["TRUE", "FALSE"]);

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

function bracketExpressionEnd(formula: string, start: number): number {
  let depth = 0;
  for (let index = start; index < formula.length; index += 1) {
    if (formula[index] === "[") depth += 1;
    else if (formula[index] === "]") {
      depth -= 1;
      if (depth === 0) return index + 1;
    }
  }
  return start;
}

function sectionFromToken(token: string): TraceTableSection | null {
  switch (token.trim().toLowerCase()) {
    case "#all":
      return "all";
    case "#data":
      return "data";
    case "#headers":
      return "headers";
    case "#totals":
      return "totals";
    default:
      return null;
  }
}

function buildTableTarget(tableName: string, bracketText: string): TraceTableTarget | null {
  const inner = bracketText.slice(1, -1).trim();
  const sectionOnly = sectionFromToken(inner);
  if (sectionOnly) return { kind: "table", tableName, section: sectionOnly };

  if (!inner.includes("[") && !inner.includes("]") && !inner.startsWith("@")) {
    return {
      kind: "table",
      tableName,
      section: "data",
      columnStart: inner,
    };
  }

  const sectionAndColumns = /^\[#([^\]]+)\]\s*,\s*\[([^\]]+)\](?:\s*:\s*\[([^\]]+)\])?$/.exec(
    inner
  );
  if (sectionAndColumns) {
    const section = sectionFromToken(`#${sectionAndColumns[1]}`);
    if (!section) return null;
    return {
      kind: "table",
      tableName,
      section,
      columnStart: sectionAndColumns[2].trim(),
      columnEnd: sectionAndColumns[3]?.trim(),
    };
  }

  const columns = /^\[([^\]]+)\](?:\s*:\s*\[([^\]]+)\])?$/.exec(inner);
  if (!columns) return null;
  return {
    kind: "table",
    tableName,
    section: "data",
    columnStart: columns[1].trim(),
    columnEnd: columns[2]?.trim(),
  };
}

function tryExternalReference(
  formula: string,
  index: number
): FormulaExternalReferenceSegment | null {
  const remaining = formula.slice(index);
  const match =
    QUOTED_EXTERNAL_AT_START.exec(remaining) ?? UNQUOTED_EXTERNAL_AT_START.exec(remaining);
  if (!match?.groups || !isValidRangeAddress(match.groups.range)) return null;
  const next = formula[index + match[0].length];
  if (isIdentifierCharacter(next)) return null;
  return {
    kind: "external",
    text: match[0],
    workbookName: match.groups.workbook,
    worksheetName: match.groups.sheet.replace(/''/g, "'"),
    rangeAddress: match.groups.range,
    reason: EXTERNAL_REASON,
  };
}

function isExcelWorkbookName(value: string): boolean {
  return /\.(xlsx|xlsm|xlsb|xls)$/i.test(value);
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

    const external = tryExternalReference(formula, index);
    if (external) {
      segments.push(external);
      index += external.text.length;
      continue;
    }

    const a1Candidate = A1_REFERENCE_AT_START.exec(formula.slice(index));
    if (a1Candidate?.groups) {
      const token = a1Candidate[0];
      const sheetToken = a1Candidate.groups.sheet;
      const rangeAddress = a1Candidate.groups.range;
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
      if (!invalidBoundary && isValidRangeAddress(rangeAddress)) {
        const worksheetName = sheetToken ? unquoteWorksheetName(sheetToken) : contextWorksheetName;
        segments.push({
          kind: "reference",
          text: token,
          target: {
            address: `${quoteWorksheetName(worksheetName)}!${rangeAddress}`,
          },
        });
        index += token.length;
        continue;
      }
    }

    const identifier = IDENTIFIER_AT_START.exec(formula.slice(index));
    if (identifier?.groups && formula[index + identifier[0].length] === "[") {
      const bracketStart = index + identifier[0].length;
      const end = bracketExpressionEnd(formula, bracketStart);
      if (end > bracketStart) {
        const token = formula.slice(index, end);
        const previous = index > 0 ? formula[index - 1] : undefined;
        const target = buildTableTarget(identifier.groups.name, formula.slice(bracketStart, end));
        if (!isIdentifierCharacter(previous) && target) {
          segments.push({
            kind: "reference",
            referenceKind: "table",
            text: token,
            target,
          });
        } else {
          pushText(segments, token);
        }
        index = end;
        continue;
      }
    }

    if (formula[index] === "[") {
      const end = bracketExpressionEnd(formula, index);
      if (end > index) {
        pushText(segments, formula.slice(index, end));
        index = end;
        continue;
      }
    }

    const namedCandidate = NAMED_REFERENCE_AT_START.exec(formula.slice(index));
    if (namedCandidate?.groups) {
      const token = namedCandidate[0];
      const sheetToken = namedCandidate.groups.sheet;
      const name = namedCandidate.groups.name;
      const previous = index > 0 ? formula[index - 1] : undefined;
      const next = formula[index + token.length];
      const worksheetName = sheetToken ? unquoteWorksheetName(sheetToken) : undefined;
      const validName =
        !isIdentifierCharacter(previous) &&
        !isIdentifierCharacter(next) &&
        next !== "(" &&
        next !== "[" &&
        !RESERVED_NAMES.has(name.toUpperCase()) &&
        !(worksheetName && isExcelWorkbookName(worksheetName));
      if (validName) {
        segments.push({
          kind: "reference",
          referenceKind: "namedRange",
          text: token,
          target: {
            kind: "namedRange",
            name,
            formulaWorksheetName: contextWorksheetName,
            ...(worksheetName ? { worksheetName } : {}),
          },
        });
        index += token.length;
        continue;
      }
    }

    pushText(segments, formula[index]);
    index += 1;
  }

  return segments;
}
