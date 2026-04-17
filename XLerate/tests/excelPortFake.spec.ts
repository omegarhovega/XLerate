import { describe, expect, it } from "vitest";
import { ExcelPortFake } from "../src/adapters/excelPortFake";
import { CellAddress } from "../src/adapters/excelPort";

const addr = (row: number, col: number, sheet = "Sheet1"): CellAddress => ({ sheet, row, col });

describe("ExcelPortFake", () => {
  it("returns empty selection when none set", async () => {
    const port = new ExcelPortFake();
    expect(await port.getSelectionCells()).toEqual([]);
  });

  it("returns value cells in selection", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 0), 42);
    port.setSelection([addr(0, 0)]);
    const snapshots = await port.getSelectionCells();
    expect(snapshots).toEqual([
      {
        address: addr(0, 0),
        isFormula: false,
        isArrayFormula: false,
        formula: "",
        value: 42,
      },
    ]);
  });

  it("returns formula cells in selection", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "=A1+B1");
    port.setSelection([addr(0, 0)]);
    expect(await port.getSelectionCells()).toEqual([
      {
        address: addr(0, 0),
        isFormula: true,
        isArrayFormula: false,
        formula: "=A1+B1",
        value: undefined,
      },
    ]);
  });

  it("returns array formulas with isArrayFormula=true", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "{=A1:A3*2}", true);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionCells();
    expect(snap.isArrayFormula).toBe(true);
    expect(snap.formula).toBe("{=A1:A3*2}");
  });

  it("applies value mutations", async () => {
    const port = new ExcelPortFake();
    await port.applyMutations([{ address: addr(0, 0), kind: "value", value: 7 }]);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionCells();
    expect(snap.value).toBe(7);
    expect(snap.isFormula).toBe(false);
  });

  it("applies formula mutations", async () => {
    const port = new ExcelPortFake();
    await port.applyMutations([{ address: addr(0, 0), kind: "formula", formula: "=SUM(A1:A3)" }]);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionCells();
    expect(snap.isFormula).toBe(true);
    expect(snap.formula).toBe("=SUM(A1:A3)");
  });

  it("applies array formula mutations with isArray flag", async () => {
    const port = new ExcelPortFake();
    await port.applyMutations([
      { address: addr(0, 0), kind: "arrayFormula", formula: "{=-(A1:A3*2)}" },
    ]);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionCells();
    expect(snap.isArrayFormula).toBe(true);
    expect(snap.formula).toBe("{=-(A1:A3*2)}");
  });

  it("treats missing cells as empty (null value, not a formula)", async () => {
    const port = new ExcelPortFake();
    port.setSelection([addr(5, 5)]);
    expect(await port.getSelectionCells()).toEqual([
      {
        address: addr(5, 5),
        isFormula: false,
        isArrayFormula: false,
        formula: "",
        value: null,
      },
    ]);
  });

  it("handles multi-cell selection in row-major order", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 0), 1);
    port.setCellValue(addr(0, 1), 2);
    port.setCellValue(addr(1, 0), 3);
    port.setSelection([addr(0, 0), addr(0, 1), addr(1, 0)]);
    const snaps = await port.getSelectionCells();
    expect(snaps.map((s) => s.value)).toEqual([1, 2, 3]);
  });
});
