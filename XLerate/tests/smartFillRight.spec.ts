import { describe, expect, it } from "vitest";
import { computeSmartFillRight, findSmartFillBoundary, type SmartFillCell, type SmartFillRow } from "../src/core/smartFillRight";

function cell(isEmpty: boolean, isMerged = false): SmartFillCell {
  return { isEmpty, isMerged };
}

function rowFromPattern(pattern: string): SmartFillRow {
  // Pattern chars: . = empty, x = non-empty, m = non-empty merged.
  return pattern.split("").map((ch) => {
    if (ch === ".") return cell(true, false);
    if (ch === "m") return cell(false, true);
    return cell(false, false);
  });
}

describe("findSmartFillBoundary parity", () => {
  it("uses nearest valid row above", () => {
    const rows: SmartFillRow[] = [
      rowFromPattern("........"), // row 1
      rowFromPattern("........"), // row 2
      rowFromPattern(".xxx...."), // row 3: boundary 4 from startCol=2
      rowFromPattern(".xxxxx.."), // row 4: boundary 6
      rowFromPattern(".x......") // row 5 active row
    ];

    expect(findSmartFillBoundary(rows, 5, 2)).toBe(6);
  });

  it("skips rows with merged cells in the contiguous block", () => {
    const rows: SmartFillRow[] = [
      rowFromPattern("........"), // row 1
      rowFromPattern(".xxx...."), // row 2: boundary 4
      rowFromPattern(".xmx...."), // row 3: has merge in contiguous block
      rowFromPattern(".x......"), // row 4 active row
      rowFromPattern("........")
    ];

    expect(findSmartFillBoundary(rows, 4, 2)).toBe(4);
  });

  it("returns 0 when no valid boundary is found in 3 rows above", () => {
    const rows: SmartFillRow[] = [
      rowFromPattern("........"), // row 1
      rowFromPattern("........"), // row 2
      rowFromPattern(".m......"), // row 3 merged at start block
      rowFromPattern("........"), // row 4 active row
      rowFromPattern("........")
    ];

    expect(findSmartFillBoundary(rows, 4, 2)).toBe(0);
  });
});

describe("computeSmartFillRight parity", () => {
  it("rejects non-formula active cells", () => {
    const rows: SmartFillRow[] = [rowFromPattern("..."), rowFromPattern("..."), rowFromPattern("...")];
    expect(computeSmartFillRight(rows, { row: 3, col: 2, formula: "123", isMerged: false })).toEqual({
      ok: false,
      reason: "active_cell_must_have_formula"
    });
  });

  it("rejects merged active cells", () => {
    const rows: SmartFillRow[] = [rowFromPattern("..."), rowFromPattern(".xx"), rowFromPattern(".x.")];
    expect(computeSmartFillRight(rows, { row: 3, col: 2, formula: "=A1", isMerged: true })).toEqual({
      ok: false,
      reason: "active_cell_is_merged"
    });
  });

  it("returns boundary when valid", () => {
    const rows: SmartFillRow[] = [
      rowFromPattern("........"),
      rowFromPattern("........"),
      rowFromPattern(".xxxx..."), // row 3 boundary 5
      rowFromPattern(".x......") // row 4 active
    ];

    expect(computeSmartFillRight(rows, { row: 4, col: 2, formula: "=B4+C4", isMerged: false })).toEqual({
      ok: true,
      boundaryCol: 5
    });
  });

  it("returns no-boundary when search fails", () => {
    const rows: SmartFillRow[] = [
      rowFromPattern("........"),
      rowFromPattern("........"),
      rowFromPattern("........"),
      rowFromPattern(".x......") // row 4 active
    ];

    expect(computeSmartFillRight(rows, { row: 4, col: 2, formula: "=B4+C4", isMerged: false })).toEqual({
      ok: false,
      reason: "no_boundary_found"
    });
  });
});
