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

  it("does not misclassify names, functions, scientific notation, or invalid cells", () => {
    const segments = tokenizeFormulaReferences(
      "=Revenue_A1+LOG10(B2)+1E10+XFE1+A1048577+XFD1048576",
      "Model"
    );
    expect(segments.filter((segment) => segment.kind === "reference")).toEqual([
      { kind: "reference", text: "B2", target: { address: "Model!B2" } },
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
});
