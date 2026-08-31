/// <reference types="office-js" />

import { afterEach, describe, expect, it, vi } from "vitest";
import { TracePortLive } from "../src/adapters/tracePortLive";

type FakeNamedItem = {
  isNullObject: boolean;
  load: ReturnType<typeof vi.fn>;
  getRange: ReturnType<typeof vi.fn>;
  select: ReturnType<typeof vi.fn>;
};

function namedItem(isNullObject: boolean): FakeNamedItem {
  const select = vi.fn();
  return {
    isNullObject,
    load: vi.fn(),
    getRange: vi.fn(() => ({ select })),
    select,
  };
}

function setupExcel(localName: FakeNamedItem, workbookName: FakeNamedItem) {
  const getWorksheetName = vi.fn(() => localName);
  const getLocalNameOrNull = vi.fn(() => localName);
  const getWorkbookNameOrNull = vi.fn(() => workbookName);
  const getWorksheet = vi.fn(() => ({
    names: {
      getItem: getWorksheetName,
      getItemOrNullObject: getLocalNameOrNull,
    },
  }));
  const sync = vi.fn(async () => undefined);
  const context = {
    workbook: {
      worksheets: { getItem: getWorksheet },
      names: { getItemOrNullObject: getWorkbookNameOrNull },
    },
    sync,
  };
  vi.stubGlobal("Excel", {
    run: vi.fn(async (callback: (request: Excel.RequestContext) => Promise<void>) =>
      callback(context as unknown as Excel.RequestContext)
    ),
  });
  return {
    getWorksheet,
    getWorksheetName,
    getLocalNameOrNull,
    getWorkbookNameOrNull,
    sync,
  };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("TracePortLive named-range navigation", () => {
  it("resolves an unqualified Summary formula name to the workbook-scoped named item", async () => {
    const localName = namedItem(true);
    const workbookName = namedItem(false);
    const harness = setupExcel(localName, workbookName);
    const metrics = vi.fn();

    await new TracePortLive(metrics).select({
      kind: "namedRange",
      name: "Annual_volume_growth",
      formulaWorksheetName: "Summary",
    });

    expect(harness.getWorksheet).toHaveBeenCalledWith("Summary");
    expect(harness.getLocalNameOrNull).toHaveBeenCalledWith("Annual_volume_growth");
    expect(harness.getWorkbookNameOrNull).toHaveBeenCalledWith("Annual_volume_growth");
    expect(localName.getRange).not.toHaveBeenCalled();
    expect(workbookName.getRange).toHaveBeenCalledOnce();
    expect(workbookName.select).toHaveBeenCalledOnce();
    expect(harness.sync).toHaveBeenCalledTimes(2);
    expect(metrics).toHaveBeenCalledWith({ operation: "select", syncCount: 2 });
  });

  it("gives a same-sheet local name precedence over a workbook name", async () => {
    const localName = namedItem(false);
    const workbookName = namedItem(false);
    setupExcel(localName, workbookName);

    await new TracePortLive().select({
      kind: "namedRange",
      name: "Annual_volume_growth",
      formulaWorksheetName: "Summary",
    });

    expect(localName.getRange).toHaveBeenCalledOnce();
    expect(localName.select).toHaveBeenCalledOnce();
    expect(workbookName.getRange).not.toHaveBeenCalled();
  });

  it("uses the named-item collection of an explicitly qualified worksheet", async () => {
    const worksheetName = namedItem(false);
    const workbookName = namedItem(false);
    const harness = setupExcel(worksheetName, workbookName);

    await new TracePortLive().select({
      kind: "namedRange",
      name: "Annual_volume_growth",
      formulaWorksheetName: "Summary",
      worksheetName: "Global_Drivers",
    });

    expect(harness.getWorksheet).toHaveBeenCalledWith("Global_Drivers");
    expect(harness.getWorksheetName).toHaveBeenCalledWith("Annual_volume_growth");
    expect(harness.getLocalNameOrNull).not.toHaveBeenCalled();
    expect(harness.getWorkbookNameOrNull).not.toHaveBeenCalled();
    expect(worksheetName.select).toHaveBeenCalledOnce();
    expect(harness.sync).toHaveBeenCalledOnce();
  });
});
