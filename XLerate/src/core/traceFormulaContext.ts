import type { FormulaSegment } from "./formulaReferences";
import type { TraceDialogState } from "./traceDialogState";
import type { TraceNode, TraceTarget } from "./traceGraph";
import type { TraceDirection } from "./traceUtils";
import { parseWorksheetScopedAddress } from "./traceUtils";

type TraceBounds = {
  worksheetName: string;
  rowStart: number;
  rowEnd: number;
  columnStart: number;
  columnEnd: number;
};

export type TraceFormulaContext = {
  formulaNode: TraceNode;
  relationNode: TraceNode | null;
  breadcrumb: TraceNode[];
};

function columnIndex(letters: string): number {
  let result = 0;
  for (const character of letters.toUpperCase()) {
    result = result * 26 + character.charCodeAt(0) - 64;
  }
  return result - 1;
}

function parseCellAddress(address: string): { rowIndex: number; columnIndex: number } | null {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$/.exec(address.trim());
  if (!match) return null;
  return {
    rowIndex: Number(match[2]) - 1,
    columnIndex: columnIndex(match[1]),
  };
}

function parseRangeBounds(address: string): TraceBounds | null {
  const parsed = parseWorksheetScopedAddress(address);
  if (!parsed) return null;
  const parts = parsed.rangeAddress.split(":");
  if (parts.length < 1 || parts.length > 2) return null;
  const start = parseCellAddress(parts[0]);
  const end = parseCellAddress(parts[1] ?? parts[0]);
  if (!start || !end) return null;
  return {
    worksheetName: parsed.worksheetName,
    rowStart: Math.min(start.rowIndex, end.rowIndex),
    rowEnd: Math.max(start.rowIndex, end.rowIndex),
    columnStart: Math.min(start.columnIndex, end.columnIndex),
    columnEnd: Math.max(start.columnIndex, end.columnIndex),
  };
}

function nodeBounds(node: TraceNode): TraceBounds | null {
  if (node.kind === "cell") {
    return {
      worksheetName: node.worksheetName,
      rowStart: node.rowIndex,
      rowEnd: node.rowIndex,
      columnStart: node.columnIndex,
      columnEnd: node.columnIndex,
    };
  }
  if (node.kind === "range") {
    return {
      worksheetName: node.worksheetName,
      rowStart: node.rowIndex,
      rowEnd: node.rowIndex + node.rowCount - 1,
      columnStart: node.columnIndex,
      columnEnd: node.columnIndex + node.columnCount - 1,
    };
  }
  return parseRangeBounds(node.address);
}

function targetBounds(target: TraceTarget): TraceBounds | null {
  return "address" in target ? parseRangeBounds(target.address) : null;
}

function boundsOverlap(left: TraceBounds, right: TraceBounds): boolean {
  return (
    left.worksheetName.localeCompare(right.worksheetName, undefined, { sensitivity: "accent" }) ===
      0 &&
    left.rowStart <= right.rowEnd &&
    left.rowEnd >= right.rowStart &&
    left.columnStart <= right.columnEnd &&
    left.columnEnd >= right.columnStart
  );
}

function hasFormula(node: TraceNode): boolean {
  return node.kind !== "range" && node.formula.length > 0;
}

export function buildTraceBreadcrumb(state: TraceDialogState, nodeId: string): TraceNode[] {
  const reversed: TraceNode[] = [];
  const seen = new Set<string>();
  let current = state.nodesById.get(nodeId);
  while (current && !seen.has(current.id)) {
    reversed.push(current);
    seen.add(current.id);
    current = current.parentId ? state.nodesById.get(current.parentId) : undefined;
  }
  return reversed.reverse();
}

function nearestNonRangeAncestor(path: TraceNode[]): TraceNode | null {
  for (let index = path.length - 2; index >= 0; index -= 1) {
    if (path[index].kind !== "range") return path[index];
  }
  return null;
}

export function buildTraceFormulaContext(
  state: TraceDialogState,
  nodeId: string,
  direction: TraceDirection
): TraceFormulaContext | null {
  const breadcrumb = buildTraceBreadcrumb(state, nodeId);
  const focused = breadcrumb[breadcrumb.length - 1];
  if (!focused) return null;

  let formulaNode = focused;
  let relationNode: TraceNode | null = null;
  if (direction === "precedents" && breadcrumb.length > 1) {
    relationNode = focused;
    for (let index = breadcrumb.length - 2; index >= 0; index -= 1) {
      if (hasFormula(breadcrumb[index])) {
        formulaNode = breadcrumb[index];
        break;
      }
    }
  } else if (direction === "dependents" && breadcrumb.length > 1) {
    relationNode = nearestNonRangeAncestor(breadcrumb);
  }

  return {
    formulaNode,
    relationNode,
    breadcrumb,
  };
}

export function formulaTargetMatchesNode(target: TraceTarget, node: TraceNode): boolean {
  const formulaBounds = targetBounds(target);
  const selectedBounds = nodeBounds(node);
  return (
    formulaBounds !== null &&
    selectedBounds !== null &&
    boundsOverlap(formulaBounds, selectedBounds)
  );
}

export function getHighlightedFormulaSegmentIndexes(
  segments: FormulaSegment[],
  relationNode: TraceNode | null
): Set<number> {
  if (!relationNode) return new Set();
  const referenceIndexes: number[] = [];
  const matchingIndexes: number[] = [];
  segments.forEach((segment, index) => {
    if (segment.kind !== "reference") return;
    referenceIndexes.push(index);
    if (formulaTargetMatchesNode(segment.target, relationNode)) matchingIndexes.push(index);
  });
  if (matchingIndexes.length > 0) return new Set(matchingIndexes);
  return referenceIndexes.length === 1 ? new Set(referenceIndexes) : new Set();
}
