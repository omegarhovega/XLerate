/* global Excel, Office */

import { buildTrace, type TraceCellInfo, type TraceRow } from "../core/traceBuilder";
import { sanitizeTraceDepth, type TraceDirection } from "../core/traceUtils";
import {
  getDirectTraceNeighbors,
  loadTraceCellProperties,
  resolveTraceStartCell,
  snapshotRangeForTrace,
} from "./traceExcelNeighbors";

// Phase B.3: the dialog computes its own trace and renders the list.
// Keyboard focus-nav (B.4) and messageParent protocol (B.5) land next.
// This file has its own Excel.run context, independent of the taskpane's.

const BODY_ID = "trace-dialog-body";
const STATUS_ID = "trace-dialog-status";
const TITLE_ID = "trace-dialog-title";

function setDialogStatus(message: string): void {
  const el = document.getElementById(STATUS_ID);
  if (el) el.textContent = message;
}

function setDialogTitle(message: string): void {
  const el = document.getElementById(TITLE_ID);
  if (el) el.textContent = message;
}

type DialogParams = {
  direction: TraceDirection;
  address: string | null;
  maxDepth: number;
};

function parseDialogParams(): DialogParams {
  const params = new URLSearchParams(window.location.search);
  const rawDirection = params.get("direction");
  const direction: TraceDirection = rawDirection === "dependents" ? "dependents" : "precedents";
  const address = params.get("address");
  const depthRaw = params.get("maxDepth");
  const maxDepth = sanitizeTraceDepth(depthRaw === null ? undefined : Number(depthRaw));
  return { direction, address: address && address.length > 0 ? address : null, maxDepth };
}

function renderDialogTraceRows(rows: TraceRow[]): void {
  const body = document.getElementById(BODY_ID);
  if (!(body instanceof HTMLTableSectionElement)) return;

  body.textContent = "";

  if (rows.length === 0) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 4;
    td.textContent = "No trace results.";
    tr.appendChild(td);
    body.appendChild(tr);
    return;
  }

  rows.forEach((item, index) => {
    const tr = document.createElement("tr");
    tr.className = "trace-row-clickable";
    tr.setAttribute("role", "option");
    tr.setAttribute("tabindex", "0");
    tr.setAttribute("data-trace-index", String(index));
    tr.setAttribute("aria-selected", "false");

    const level = document.createElement("td");
    const address = document.createElement("td");
    const value = document.createElement("td");
    const formula = document.createElement("td");

    level.textContent = String(item.level);
    address.textContent = item.address;
    value.textContent = item.value;
    formula.textContent = item.formula;

    tr.append(level, address, value, formula);
    body.appendChild(tr);
  });
}

async function runDialogTrace(params: DialogParams): Promise<void> {
  setDialogTitle(`Trace ${params.direction}`);
  setDialogStatus(params.address ? `Tracing ${params.direction} from ${params.address}…` : `Tracing ${params.direction}…`);

  try {
    await Excel.run(async (context) => {
      const rootRange = await resolveTraceStartCell(context, params.address);
      if (!rootRange) {
        setDialogStatus("Could not resolve a starting cell for the trace.");
        return;
      }

      const root = snapshotRangeForTrace(rootRange);

      const getNeighbors = async (info: TraceCellInfo): Promise<TraceCellInfo[]> => {
        const worksheet = context.workbook.worksheets.getItem(info.worksheetName);
        const cell = worksheet.getRangeByIndexes(info.rowIndex, info.columnIndex, 1, 1);
        // resolveTraceStartCell/getDirectTraceNeighbors do their own loads; the
        // ephemeral cell above only needs to exist to call getDirectPrecedents/Dependents.
        loadTraceCellProperties(cell);
        const neighbors = await getDirectTraceNeighbors(context, cell, params.direction);
        return neighbors.map(snapshotRangeForTrace);
      };

      const { rows, truncated } = await buildTrace({
        root,
        maxDepth: params.maxDepth,
        getNeighbors,
      });

      renderDialogTraceRows(rows);
      setDialogStatus(
        `Trace ${params.direction} on ${root.address}: ${rows.length} cell${rows.length === 1 ? "" : "s"}${truncated ? " (truncated)" : ""}.`
      );
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setDialogStatus(`Trace failed: ${message}`);
  }
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setDialogStatus("Trace dialog requires Excel.");
    return;
  }
  const params = parseDialogParams();
  void runDialogTrace(params);
});
