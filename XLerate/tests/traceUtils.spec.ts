import { describe, expect, it } from "vitest";
import {
  buildTraceCellKey,
  DEFAULT_TRACE_MAX_DEPTH,
  formatTraceFormula,
  formatTraceValue,
  MAX_TRACE_MAX_DEPTH,
  parseWorksheetScopedAddress,
  sanitizeTraceDepth,
  scalarFromMatrix
} from "../src/core/traceUtils";

describe("trace utils", () => {
  it("sanitizes depth to allowed bounds", () => {
    expect(sanitizeTraceDepth(undefined)).toBe(DEFAULT_TRACE_MAX_DEPTH);
    expect(sanitizeTraceDepth(NaN)).toBe(DEFAULT_TRACE_MAX_DEPTH);
    expect(sanitizeTraceDepth(0)).toBe(1);
    expect(sanitizeTraceDepth(1.9)).toBe(1);
    expect(sanitizeTraceDepth(MAX_TRACE_MAX_DEPTH + 10)).toBe(MAX_TRACE_MAX_DEPTH);
  });

  it("extracts scalar values from 2D matrices", () => {
    expect(scalarFromMatrix([[123]])).toBe(123);
    expect(scalarFromMatrix([["abc"]])).toBe("abc");
    expect(scalarFromMatrix("raw")).toBe("raw");
  });

  it("formats value and formula payloads", () => {
    expect(formatTraceValue(null)).toBe("");
    expect(formatTraceValue(42)).toBe("42");
    expect(formatTraceValue({ error: "#DIV/0!" })).toBe("#DIV/0!");
    expect(formatTraceFormula("=A1+B1")).toBe("=A1+B1");
    expect(formatTraceFormula("plain-text")).toBe("");
    expect(formatTraceFormula(42)).toBe("");
  });

  it("builds stable trace keys by worksheet and position", () => {
    expect(buildTraceCellKey("Sheet1", 0, 0)).toBe("Sheet1!R0C0");
    expect(buildTraceCellKey("Inputs", 7, 3)).toBe("Inputs!R7C3");
  });

  it("parses worksheet-scoped addresses including quoted and workbook-prefixed names", () => {
    expect(parseWorksheetScopedAddress("Sheet1!A1")).toEqual({ worksheetName: "Sheet1", rangeAddress: "A1" });
    expect(parseWorksheetScopedAddress("'My Sheet'!B2:C3")).toEqual({
      worksheetName: "My Sheet",
      rangeAddress: "B2:C3"
    });
    expect(parseWorksheetScopedAddress("'O''Brien'!D4")).toEqual({
      worksheetName: "O'Brien",
      rangeAddress: "D4"
    });
    expect(parseWorksheetScopedAddress("[Book1]SheetX!A1")).toEqual({
      worksheetName: "SheetX",
      rangeAddress: "A1"
    });
    expect(parseWorksheetScopedAddress("A1")).toBeNull();
  });
});
