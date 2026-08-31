import {
  buildTraceRangeKey,
  type TraceCellData,
  type TraceNodeData,
  type TraceRangeData,
} from "./traceGraph";
import type { TraceLoadingPolicy } from "./tracePolicy";
import { shouldSummarizeTraceRange } from "./tracePolicy";
import { buildTraceCellKey, formatTraceFormula, formatTraceValue } from "./traceUtils";

export type TraceAreaSnapshot = {
  worksheetName: string;
  address: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  values?: unknown[][];
  formulas?: unknown[][];
};

export type TraceAreaMaterializationResult = {
  nodes: TraceNodeData[];
  truncated: boolean;
  materializedCellCount: number;
};

export type TraceCellMaterializationResult = {
  nodes: TraceCellData[];
  truncated: boolean;
  materializedCellCount: number;
};

function columnIndexToLetters(columnIndex: number): string {
  let remainder = columnIndex + 1;
  let letters = "";
  while (remainder > 0) {
    const zeroBased = (remainder - 1) % 26;
    letters = String.fromCharCode(65 + zeroBased) + letters;
    remainder = Math.floor((remainder - 1) / 26);
  }
  return letters;
}

function quoteWorksheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) {
    return name;
  }
  return `'${name.replace(/'/g, "''")}'`;
}

function buildCellAddress(worksheetName: string, rowIndex: number, columnIndex: number): string {
  return `${quoteWorksheetName(worksheetName)}!${columnIndexToLetters(columnIndex)}${rowIndex + 1}`;
}

function toRangeData(area: TraceAreaSnapshot): TraceRangeData {
  return {
    kind: "range",
    key: buildTraceRangeKey(
      area.worksheetName,
      area.rowIndex,
      area.columnIndex,
      area.rowCount,
      area.columnCount
    ),
    worksheetName: area.worksheetName,
    rowIndex: area.rowIndex,
    columnIndex: area.columnIndex,
    rowCount: area.rowCount,
    columnCount: area.columnCount,
    cellCount: area.rowCount * area.columnCount,
    address: area.address,
  };
}

export function materializeTraceArea(
  area: TraceAreaSnapshot,
  policy: TraceLoadingPolicy,
  remainingNodes: number
): TraceAreaMaterializationResult {
  const cellCount = area.rowCount * area.columnCount;
  if (remainingNodes <= 0) {
    return { nodes: [], truncated: cellCount > 0, materializedCellCount: 0 };
  }

  if (shouldSummarizeTraceRange(cellCount, policy) || !area.values || !area.formulas) {
    return {
      nodes: [toRangeData(area)],
      truncated: false,
      materializedCellCount: 0,
    };
  }

  return materializeTraceCells(area, remainingNodes);
}

export function materializeTraceCells(
  area: TraceAreaSnapshot,
  remainingNodes: number
): TraceCellMaterializationResult {
  const cellCount = area.rowCount * area.columnCount;
  if (!area.values || !area.formulas || remainingNodes <= 0) {
    return {
      nodes: [],
      truncated: cellCount > 0,
      materializedCellCount: 0,
    };
  }
  const nodes: TraceCellData[] = [];
  const limit = Math.min(cellCount, remainingNodes);
  for (let index = 0; index < limit; index += 1) {
    const rowOffset = Math.floor(index / area.columnCount);
    const columnOffset = index % area.columnCount;
    const rowIndex = area.rowIndex + rowOffset;
    const columnIndex = area.columnIndex + columnOffset;
    const formula = area.formulas[rowOffset]?.[columnOffset];
    const value = area.values[rowOffset]?.[columnOffset];
    nodes.push({
      kind: "cell",
      key: buildTraceCellKey(area.worksheetName, rowIndex, columnIndex),
      worksheetName: area.worksheetName,
      rowIndex,
      columnIndex,
      address: buildCellAddress(area.worksheetName, rowIndex, columnIndex),
      value: formatTraceValue(value),
      formula: formatTraceFormula(formula),
    });
  }

  return {
    nodes,
    truncated: limit < cellCount,
    materializedCellCount: nodes.length,
  };
}
