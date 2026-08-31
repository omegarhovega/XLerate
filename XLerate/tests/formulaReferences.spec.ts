import { describe, expect, it } from "vitest";
import { tokenizeFormulaReferences } from "../src/core/formulaReferences";

describe("formula reference tokenization", () => {
  it("tokenizes local cells, absolute references, and ranges", () => {
    expect(tokenizeFormulaReferences("=A1+$B$2+C$3:$D4", "Model")).toEqual([
      { kind: "text", text: "=" },
      { kind: "reference", text: "A1", target: { address: "Model!A1" } },
      { kind: "text", text: "+" },
      { kind: "reference", text: "$B$2", target: { address: "Model!$B$2" } },
      { kind: "text", text: "+" },
      { kind: "reference", text: "C$3:$D4", target: { address: "Model!C$3:$D4" } },
    ]);
  });

  it("resolves quoted and unquoted worksheet references", () => {
    expect(
      tokenizeFormulaReferences("=SUM('North America'!B5:B12)+Sheet_2!$C$7", "Summary")
    ).toEqual([
      { kind: "text", text: "=SUM(" },
      {
        kind: "reference",
        text: "'North America'!B5:B12",
        target: { address: "'North America'!B5:B12" },
      },
      { kind: "text", text: ")+" },
      {
        kind: "reference",
        text: "Sheet_2!$C$7",
        target: { address: "Sheet_2!$C$7" },
      },
    ]);
  });

  it("unescapes and requotes worksheet names containing apostrophes", () => {
    expect(tokenizeFormulaReferences("='O''Brien'!$A$1", "Summary")).toEqual([
      { kind: "text", text: "=" },
      {
        kind: "reference",
        text: "'O''Brien'!$A$1",
        target: { address: "'O''Brien'!$A$1" },
      },
    ]);
  });

  it("does not link reference-looking text inside Excel string literals", () => {
    const segments = tokenizeFormulaReferences(
      '=IF(A1="B2 and Sheet2!C3",LOG10(C3),1E10)',
      "Model"
    );
    expect(segments.filter((segment) => segment.kind === "reference")).toEqual([
      { kind: "reference", text: "A1", target: { address: "Model!A1" } },
      { kind: "reference", text: "C3", target: { address: "Model!C3" } },
    ]);
  });

  it("handles escaped quotes inside string literals", () => {
    const segments = tokenizeFormulaReferences('="Text says ""A1"""&B2', "Model");
    expect(segments.filter((segment) => segment.kind === "reference")).toEqual([
      { kind: "reference", text: "B2", target: { address: "Model!B2" } },
    ]);
  });

  it("does not misclassify functions or scientific notation and preserves name candidates", () => {
    const segments = tokenizeFormulaReferences(
      "=Revenue_A1+LOG10(B2)+1E10+XFE1+A1048577+XFD1048576",
      "Model"
    );
    expect(segments.filter((segment) => segment.kind === "reference")).toEqual([
      {
        kind: "reference",
        referenceKind: "namedRange",
        text: "Revenue_A1",
        target: {
          kind: "namedRange",
          name: "Revenue_A1",
          formulaWorksheetName: "Model",
        },
      },
      { kind: "reference", text: "B2", target: { address: "Model!B2" } },
      {
        kind: "reference",
        referenceKind: "namedRange",
        text: "XFE1",
        target: {
          kind: "namedRange",
          name: "XFE1",
          formulaWorksheetName: "Model",
        },
      },
      {
        kind: "reference",
        referenceKind: "namedRange",
        text: "A1048577",
        target: {
          kind: "namedRange",
          name: "A1048577",
          formulaWorksheetName: "Model",
        },
      },
      {
        kind: "reference",
        text: "XFD1048576",
        target: { address: "Model!XFD1048576" },
      },
    ]);
  });

  it("keeps spill syntax outside the clickable A1 token", () => {
    expect(tokenizeFormulaReferences("=A1#", "Model")).toEqual([
      { kind: "text", text: "=" },
      { kind: "reference", text: "A1", target: { address: "Model!A1" } },
      { kind: "text", text: "#" },
    ]);
  });

  it("returns one text segment when no supported references exist", () => {
    expect(tokenizeFormulaReferences("=SUM(1,2,3)", "Model")).toEqual([
      { kind: "text", text: "=SUM(1,2,3)" },
    ]);
    expect(tokenizeFormulaReferences("", "Model")).toEqual([]);
  });

  it("tokenizes workbook- and worksheet-scoped named ranges", () => {
    expect(
      tokenizeFormulaReferences("=Revenue_Growth*Inputs!Local_Assumption+TRUE", "Model")
    ).toEqual([
      { kind: "text", text: "=" },
      {
        kind: "reference",
        referenceKind: "namedRange",
        text: "Revenue_Growth",
        target: {
          kind: "namedRange",
          name: "Revenue_Growth",
          formulaWorksheetName: "Model",
        },
      },
      { kind: "text", text: "*" },
      {
        kind: "reference",
        referenceKind: "namedRange",
        text: "Inputs!Local_Assumption",
        target: {
          kind: "namedRange",
          name: "Local_Assumption",
          formulaWorksheetName: "Model",
          worksheetName: "Inputs",
        },
      },
      { kind: "text", text: "+TRUE" },
    ]);
  });

  it("does not convert an unqualified workbook name into a Summary-scoped name", () => {
    expect(tokenizeFormulaReferences("=Annual_volume_growth*100", "Summary")).toEqual([
      { kind: "text", text: "=" },
      {
        kind: "reference",
        referenceKind: "namedRange",
        text: "Annual_volume_growth",
        target: {
          kind: "namedRange",
          name: "Annual_volume_growth",
          formulaWorksheetName: "Summary",
        },
      },
      { kind: "text", text: "*100" },
    ]);
  });

  it("tokenizes fully qualified table columns, spans, and sections", () => {
    expect(
      tokenizeFormulaReferences(
        "=RevenueTable[Revenue]+RevenueTable[[Price]:[Volume]]+RevenueTable[#All]+RevenueTable[[#Data],[Margin]]",
        "Model"
      )
    ).toEqual([
      { kind: "text", text: "=" },
      {
        kind: "reference",
        referenceKind: "table",
        text: "RevenueTable[Revenue]",
        target: {
          kind: "table",
          tableName: "RevenueTable",
          section: "data",
          columnStart: "Revenue",
        },
      },
      { kind: "text", text: "+" },
      {
        kind: "reference",
        referenceKind: "table",
        text: "RevenueTable[[Price]:[Volume]]",
        target: {
          kind: "table",
          tableName: "RevenueTable",
          section: "data",
          columnStart: "Price",
          columnEnd: "Volume",
        },
      },
      { kind: "text", text: "+" },
      {
        kind: "reference",
        referenceKind: "table",
        text: "RevenueTable[#All]",
        target: { kind: "table", tableName: "RevenueTable", section: "all" },
      },
      { kind: "text", text: "+" },
      {
        kind: "reference",
        referenceKind: "table",
        text: "RevenueTable[[#Data],[Margin]]",
        target: {
          kind: "table",
          tableName: "RevenueTable",
          section: "data",
          columnStart: "Margin",
        },
      },
    ]);
  });

  it("leaves current-row and unsupported structured references as text", () => {
    expect(
      tokenizeFormulaReferences("=[@Revenue]+RevenueTable[[#This Row],[Revenue]]", "Model")
    ).toEqual([{ kind: "text", text: "=[@Revenue]+RevenueTable[[#This Row],[Revenue]]" }]);
  });

  it("detects open and closed external workbook cell/range references", () => {
    const segments = tokenizeFormulaReferences(
      "=SUM([Budget.xlsx]Annual!C10:C25)+'C:\\Reports\\[Ops.xlsx]North America'!$A$1",
      "Model"
    );
    expect(segments.filter((segment) => segment.kind === "external")).toEqual([
      {
        kind: "external",
        text: "[Budget.xlsx]Annual!C10:C25",
        workbookName: "Budget.xlsx",
        worksheetName: "Annual",
        rangeAddress: "C10:C25",
        reason: "External workbook navigation is not available in the Excel JavaScript API.",
      },
      {
        kind: "external",
        text: "'C:\\Reports\\[Ops.xlsx]North America'!$A$1",
        workbookName: "Ops.xlsx",
        worksheetName: "North America",
        rangeAddress: "$A$1",
        reason: "External workbook navigation is not available in the Excel JavaScript API.",
      },
    ]);
  });
});
