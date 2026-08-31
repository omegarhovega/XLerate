import type {
  DirectNeighborResult,
  TraceCellData,
  TraceRangeData,
  TraceRangePageResult,
  TraceTableSection,
  TraceTableTarget,
  TraceTarget,
} from "../core/traceGraph";
import {
  materializeTraceArea,
  materializeTraceCells,
  type TraceAreaSnapshot,
} from "../core/traceAreaMaterializer";
import { getRangePageBounds } from "../core/tracePolicy";
import { planTraceRangePageBlocks } from "../core/traceRangePaging";
import {
  buildTraceCellKey,
  formatTraceFormula,
  formatTraceValue,
  parseWorksheetScopedAddress,
  scalarFromMatrix,
  type TraceDirection,
} from "../core/traceUtils";
import type { TraceLoadRequest, TracePort } from "./tracePort";

const RANGE_COLLECTION_LOAD_SPEC =
  "items/address,rowIndex,columnIndex,rowCount,columnCount,worksheet/name";

export type TracePortMetric = {
  operation: "root" | "neighbors" | "rangePage" | "select";
  syncCount: number;
  returnedAreaCount?: number;
  returnedNodeCount?: number;
  materializedCellCount?: number;
};

export type TracePortMetricSink = (metric: TracePortMetric) => void;

function isItemNotFoundError(error: unknown): boolean {
  if (!error || typeof error !== "object") return false;
  const candidate = error as { code?: unknown; message?: unknown };
  if (candidate.code === "ItemNotFound") return true;
  return typeof candidate.message === "string" && candidate.message.includes("ItemNotFound");
}

function snapshotCell(
  worksheetName: string,
  rowIndex: number,
  columnIndex: number,
  address: string,
  values: unknown,
  formulas: unknown
): TraceCellData {
  return {
    kind: "cell",
    key: buildTraceCellKey(worksheetName, rowIndex, columnIndex),
    worksheetName,
    rowIndex,
    columnIndex,
    address,
    value: formatTraceValue(scalarFromMatrix(values)),
    formula: formatTraceFormula(scalarFromMatrix(formulas)),
  };
}

function getTableRangeForSection(table: Excel.Table, section: TraceTableSection): Excel.Range {
  if (section === "all") return table.getRange();
  if (section === "headers") return table.getHeaderRowRange();
  if (section === "totals") return table.getTotalRowRange();
  return table.getDataBodyRange();
}

function getTableColumnRangeForSection(
  column: Excel.TableColumn,
  section: TraceTableSection
): Excel.Range {
  if (section === "all") return column.getRange();
  if (section === "headers") return column.getHeaderRowRange();
  if (section === "totals") return column.getTotalRowRange();
  return column.getDataBodyRange();
}

function getStructuredTableRange(
  context: Excel.RequestContext,
  target: TraceTableTarget
): Excel.Range {
  const table = context.workbook.tables.getItem(target.tableName);
  if (!target.columnStart) return getTableRangeForSection(table, target.section);

  const start = getTableColumnRangeForSection(
    table.columns.getItem(target.columnStart),
    target.section
  );
  if (!target.columnEnd) return start;
  const end = getTableColumnRangeForSection(
    table.columns.getItem(target.columnEnd),
    target.section
  );
  return start.getBoundingRect(end);
}

export class TracePortLive implements TracePort {
  constructor(private readonly onMetric: TracePortMetricSink = () => undefined) {}

  async getRoot(): Promise<TraceCellData> {
    let syncCount = 0;
    const result = await Excel.run(async (context) => {
      const cell = context.workbook.getActiveCell();
      cell.load(["address", "worksheet/name", "rowIndex", "columnIndex", "values", "formulas"]);
      await context.sync();
      syncCount += 1;
      return snapshotCell(
        cell.worksheet.name,
        cell.rowIndex,
        cell.columnIndex,
        cell.address,
        cell.values,
        cell.formulas
      );
    });
    this.onMetric({ operation: "root", syncCount, returnedNodeCount: 1 });
    return result;
  }

  async getDirectNeighbors(
    node: TraceCellData,
    direction: TraceDirection,
    request: TraceLoadRequest
  ): Promise<DirectNeighborResult> {
    let syncCount = 0;
    const result = await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(node.worksheetName);
      const source = worksheet.getRangeByIndexes(node.rowIndex, node.columnIndex, 1, 1);
      const links =
        direction === "precedents" ? source.getDirectPrecedents() : source.getDirectDependents();
      const ranges = links.ranges;
      ranges.load(RANGE_COLLECTION_LOAD_SPEC);

      try {
        await context.sync();
        syncCount += 1;
      } catch (error) {
        syncCount += 1;
        if (isItemNotFoundError(error)) {
          return {
            nodes: [],
            truncated: false,
            materializedCellCount: 0,
            areaCount: 0,
          };
        }
        throw error;
      }

      let remainingNodes = request.remainingNodes;
      const smallAreas: Excel.Range[] = [];
      for (const area of ranges.items) {
        const cellCount = area.rowCount * area.columnCount;
        if (cellCount <= request.policy.rangeSummaryThreshold && remainingNodes > 0) {
          area.load(["values", "formulas"]);
          smallAreas.push(area);
        }
        remainingNodes -=
          cellCount > request.policy.rangeSummaryThreshold
            ? 1
            : Math.min(cellCount, Math.max(remainingNodes, 0));
        remainingNodes = Math.max(0, remainingNodes);
      }

      if (smallAreas.length > 0) {
        await context.sync();
        syncCount += 1;
      }

      const nodes: DirectNeighborResult["nodes"] = [];
      let truncated = false;
      let materializedCellCount = 0;
      remainingNodes = request.remainingNodes;
      for (const area of ranges.items) {
        if (remainingNodes <= 0) {
          truncated = true;
          break;
        }
        const snapshot: TraceAreaSnapshot = {
          worksheetName: area.worksheet.name,
          address: area.address,
          rowIndex: area.rowIndex,
          columnIndex: area.columnIndex,
          rowCount: area.rowCount,
          columnCount: area.columnCount,
          values:
            area.rowCount * area.columnCount <= request.policy.rangeSummaryThreshold
              ? (area.values as unknown[][])
              : undefined,
          formulas:
            area.rowCount * area.columnCount <= request.policy.rangeSummaryThreshold
              ? (area.formulas as unknown[][])
              : undefined,
        };
        const materialized = materializeTraceArea(snapshot, request.policy, remainingNodes);
        nodes.push(...materialized.nodes);
        remainingNodes -= materialized.nodes.length;
        materializedCellCount += materialized.materializedCellCount;
        truncated = truncated || materialized.truncated;
      }

      return {
        nodes,
        truncated,
        materializedCellCount,
        areaCount: ranges.items.length,
      };
    });

    const measured = result as DirectNeighborResult & { areaCount?: number };
    this.onMetric({
      operation: "neighbors",
      syncCount,
      returnedAreaCount: measured.areaCount,
      returnedNodeCount: measured.nodes.length,
      materializedCellCount: measured.materializedCellCount,
    });
    return {
      nodes: measured.nodes,
      truncated: measured.truncated,
      materializedCellCount: measured.materializedCellCount,
    };
  }

  async getRangePage(
    node: TraceRangeData,
    page: number,
    request: TraceLoadRequest
  ): Promise<TraceRangePageResult> {
    const bounds = getRangePageBounds(node.cellCount, page, request.policy);
    const count = Math.min(bounds.count, request.remainingNodes);
    if (count === 0) {
      return {
        nodes: [],
        page,
        hasMore: bounds.hasMore,
        truncated: bounds.count > 0,
      };
    }

    const blocks = planTraceRangePageBlocks(node.rowCount, node.columnCount, bounds.start, count);

    let syncCount = 0;
    const result = await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(node.worksheetName);
      const ranges = blocks.map((block) => {
        const range = worksheet.getRangeByIndexes(
          node.rowIndex + block.rowOffset,
          node.columnIndex + block.columnOffset,
          block.rowCount,
          block.columnCount
        );
        range.load([
          "address",
          "rowIndex",
          "columnIndex",
          "rowCount",
          "columnCount",
          "values",
          "formulas",
        ]);
        return range;
      });
      if (ranges.length === 0) {
        return [];
      }
      await context.sync();
      syncCount += 1;

      const nodes: TraceCellData[] = [];
      let remainingNodes = count;
      for (const range of ranges) {
        const materialized = materializeTraceCells(
          {
            worksheetName: node.worksheetName,
            address: range.address,
            rowIndex: range.rowIndex,
            columnIndex: range.columnIndex,
            rowCount: range.rowCount,
            columnCount: range.columnCount,
            values: range.values as unknown[][],
            formulas: range.formulas as unknown[][],
          },
          remainingNodes
        );
        nodes.push(...materialized.nodes);
        remainingNodes -= materialized.nodes.length;
      }
      return nodes;
    });

    this.onMetric({
      operation: "rangePage",
      syncCount,
      returnedAreaCount: blocks.length,
      returnedNodeCount: result.length,
      materializedCellCount: result.length,
    });
    return {
      nodes: result,
      page,
      hasMore: bounds.hasMore || count < bounds.count,
      truncated: count < bounds.count,
    };
  }

  async select(target: TraceTarget): Promise<void> {
    let syncCount = 0;
    await Excel.run(async (context) => {
      if ("address" in target) {
        const parsed = parseWorksheetScopedAddress(target.address);
        if (!parsed) return;
        const worksheet = context.workbook.worksheets.getItem(parsed.worksheetName);
        worksheet.getRanges(parsed.rangeAddress).select();
      } else if (target.kind === "namedRange") {
        if (target.worksheetName) {
          const worksheet = context.workbook.worksheets.getItem(target.worksheetName);
          worksheet.names.getItem(target.name).getRange().select();
        } else {
          const formulaWorksheet = context.workbook.worksheets.getItem(target.formulaWorksheetName);
          const localName = formulaWorksheet.names.getItemOrNullObject(target.name);
          const workbookName = context.workbook.names.getItemOrNullObject(target.name);
          localName.load("name");
          workbookName.load("name");
          await context.sync();
          syncCount += 1;

          const resolvedName = localName.isNullObject ? workbookName : localName;
          if (resolvedName.isNullObject) {
            throw new Error(`Named range "${target.name}" was not found.`);
          }
          resolvedName.getRange().select();
        }
      } else {
        getStructuredTableRange(context, target).select();
      }
      await context.sync();
      syncCount += 1;
    });
    this.onMetric({ operation: "select", syncCount });
  }
}
