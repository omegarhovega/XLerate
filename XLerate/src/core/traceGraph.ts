export type TraceNodeLoadState = "unloaded" | "loading" | "loaded" | "leaf" | "error";

export type TraceCellData = {
  kind: "cell";
  /** Stable workbook-session identity, normally worksheet + row + column. */
  key: string;
  worksheetName: string;
  rowIndex: number;
  columnIndex: number;
  address: string;
  value: string;
  formula: string;
};

export type TraceRangeData = {
  kind: "range";
  /** Stable identity for the rectangular block. */
  key: string;
  worksheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  cellCount: number;
  address: string;
};

export type TraceNodeData = TraceCellData | TraceRangeData;

type TraceNodePlacement = {
  /** Tree-instance identity. References to the same cell get distinct IDs. */
  id: string;
  parentId: string | null;
  level: number;
  loadState: TraceNodeLoadState;
};

export type TraceCellNode = TraceCellData & TraceNodePlacement;
export type TraceRangeNode = TraceRangeData & TraceNodePlacement;

export type TraceReferenceNode = TraceNodePlacement & {
  kind: "reference";
  key: string;
  targetId: string;
  worksheetName: string;
  address: string;
  value: string;
  formula: string;
};

export type TraceNode = TraceCellNode | TraceRangeNode | TraceReferenceNode;

export type TraceAddressTarget = {
  kind?: "address";
  address: string;
};

export type TraceNamedRangeTarget = {
  kind: "namedRange";
  name: string;
  /** Worksheet containing the formula, used for Excel's local-name precedence. */
  formulaWorksheetName: string;
  /** Present only when the formula explicitly qualifies the name as Sheet!Name. */
  worksheetName?: string;
};

export type TraceTableSection = "all" | "data" | "headers" | "totals";

export type TraceTableTarget = {
  kind: "table";
  tableName: string;
  section: TraceTableSection;
  columnStart?: string;
  columnEnd?: string;
};

export type TraceTarget = TraceAddressTarget | TraceNamedRangeTarget | TraceTableTarget;

export type DirectNeighborResult = {
  nodes: TraceNodeData[];
  truncated: boolean;
  materializedCellCount: number;
};

export type TraceRangePageResult = {
  nodes: TraceCellData[];
  page: number;
  hasMore: boolean;
  truncated: boolean;
};

export function buildTraceRangeKey(
  worksheetName: string,
  rowIndex: number,
  columnIndex: number,
  rowCount: number,
  columnCount: number
): string {
  return `${worksheetName}!R${rowIndex}C${columnIndex}:${rowCount}x${columnCount}`;
}

export function isTraceNodeExpandable(node: TraceNode): boolean {
  return node.kind !== "reference" && node.loadState !== "leaf";
}
