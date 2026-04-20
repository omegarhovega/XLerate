import { describe, expect, it } from "vitest";
import {
  classifyAutoColorCell,
  classifyAutoColorGrid,
  isExternalReferenceFormula,
  isOnlyNumbersAndOperators,
  isWorkbookLinkFormula,
  isWorksheetLinkFormula,
} from "../src/core/autoColor";

describe("auto color baseline parity", () => {
  it("detects formula helper categories", () => {
    expect(isWorksheetLinkFormula("=A1")).toBe(true);
    expect(isWorkbookLinkFormula("=Sheet2!A1")).toBe(true);
    expect(isExternalReferenceFormula("='[Book1.xlsx]Sheet1'!A1")).toBe(true);
    expect(isExternalReferenceFormula('=WEBSERVICE("https://x")')).toBe(true);
    expect(isOnlyNumbersAndOperators("=1+2*(3-4)")).toBe(true);
  });

  it("classifies blank cells as none", () => {
    expect(classifyAutoColorCell({ value: "" })).toBe("none");
    expect(classifyAutoColorCell({ value: null })).toBe("none");
  });

  it("classifies numeric constants as input", () => {
    expect(classifyAutoColorCell({ value: 10, numberFormat: "General" })).toBe("input");
  });

  it("does not classify text constants as input", () => {
    expect(classifyAutoColorCell({ value: "hello", numberFormat: "General" })).toBe("none");
  });

  it("classifies explicit hyperlinks as hyperlink for non-formula cells", () => {
    expect(
      classifyAutoColorCell({
        value: "Open",
        numberFormat: "General",
        hasHyperlink: true,
      })
    ).toBe("hyperlink");
  });

  it("classifies partial input formulas before link checks", () => {
    expect(
      classifyAutoColorCell({
        formula: "=A1+10",
        value: 42,
        numberFormat: "General",
      })
    ).toBe("partialInput");
  });

  it("classifies same-sheet link formulas", () => {
    expect(
      classifyAutoColorCell({
        formula: "=A1",
        value: 123,
        numberFormat: "General",
      })
    ).toBe("worksheetLink");
  });

  it("classifies wrapped array-formula syntax the same as regular formulas", () => {
    expect(
      classifyAutoColorCell({
        formula: "{=A1}",
        value: 123,
        numberFormat: "General",
      })
    ).toBe("worksheetLink");

    expect(
      classifyAutoColorCell({
        formula: "{=IF(TRUE,1,0)}",
        value: 1,
        numberFormat: "General",
      })
    ).toBe("input");
  });

  it("classifies workbook link formulas", () => {
    expect(
      classifyAutoColorCell({
        formula: "=Sheet2!A1",
        value: 123,
        numberFormat: "General",
      })
    ).toBe("workbookLink");
  });

  it("classifies external formulas", () => {
    expect(
      classifyAutoColorCell({
        formula: "='[Book1.xlsx]Sheet1'!A1",
        value: 123,
        numberFormat: "General",
      })
    ).toBe("external");

    expect(
      classifyAutoColorCell({
        formula: '=WEBSERVICE("https://example.com")',
        value: 123,
        numberFormat: "General",
      })
    ).toBe("external");
  });

  it("classifies formula with no references as input", () => {
    expect(
      classifyAutoColorCell({
        formula: "=1+2",
        value: 3,
        numberFormat: "General",
      })
    ).toBe("input");
  });

  it("classifies same-sheet referenced formulas as worksheet links", () => {
    expect(
      classifyAutoColorCell({
        formula: "=SUM(A1:A3)",
        value: 10,
        numberFormat: "General",
      })
    ).toBe("worksheetLink");
  });

  it("does not classify date-like numeric values as input", () => {
    expect(
      classifyAutoColorCell({
        value: 45292,
        numberFormat: "dd-mmm-yy",
      })
    ).toBe("none");
  });

  it("classifies a grid of mixed cells", () => {
    const grid = classifyAutoColorGrid([
      [
        { formula: "=A1+5", value: 10, numberFormat: "General" },
        { formula: "=Sheet2!A1", value: 9, numberFormat: "General" },
        { value: 100, numberFormat: "General" },
      ],
      [
        { formula: "=SUM(A1:A3)", value: 50, numberFormat: "General" },
        { value: "x", hasHyperlink: true, numberFormat: "General" },
        { value: "", numberFormat: "General" },
      ],
    ]);

    expect(grid).toEqual([
      ["partialInput", "workbookLink", "input"],
      ["worksheetLink", "hyperlink", "none"],
    ]);
  });
});
