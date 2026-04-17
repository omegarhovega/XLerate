import { describe, expect, it } from "vitest";
import { ExcelPortFake } from "../src/adapters/excelPortFake";
import { runSwitchSign } from "../src/services/switchSign.service";
import { CellAddress } from "../src/adapters/excelPort";

const addr = (row: number, col: number, sheet = "Sheet1"): CellAddress => ({ sheet, row, col });

describe("Switch Sign contract (spec §3.3)", () => {
  it("negates a positive numeric value", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 0), 10);
    port.setSelection([addr(0, 0)]);
    await runSwitchSign(port);
    port.setSelection([addr(0, 0)]);
    expect((await port.getSelectionCells())[0].value).toBe(-10);
  });

  it("negates a negative numeric value", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 0), -42);
    port.setSelection([addr(0, 0)]);
    await runSwitchSign(port);
    expect((await port.getSelectionCells())[0].value).toBe(42);
  });

  it("leaves zero unchanged", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 0), 0);
    port.setSelection([addr(0, 0)]);
    await runSwitchSign(port);
    expect((await port.getSelectionCells())[0].value).toBe(0);
  });

  it("wraps a standard formula", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "=A1+B1");
    port.setSelection([addr(0, 0)]);
    await runSwitchSign(port);
    const [snap] = await port.getSelectionCells();
    expect(snap.isFormula).toBe(true);
    expect(snap.formula).toBe("=-(A1+B1)");
  });

  it("wraps a formula that is itself negative (does not simplify)", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "=-A1");
    port.setSelection([addr(0, 0)]);
    await runSwitchSign(port);
    expect((await port.getSelectionCells())[0].formula).toBe("=-(-A1)");
  });

  it("wraps a function formula", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "=SUM(A1:A3)");
    port.setSelection([addr(0, 0)]);
    await runSwitchSign(port);
    expect((await port.getSelectionCells())[0].formula).toBe("=-(SUM(A1:A3))");
  });

  it("wraps an array formula preserving braces", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "{=A1:A3*2}", true);
    port.setSelection([addr(0, 0)]);
    await runSwitchSign(port);
    const [snap] = await port.getSelectionCells();
    expect(snap.isArrayFormula).toBe(true);
    expect(snap.formula).toBe("{=-(A1:A3*2)}");
  });

  it("leaves text values unchanged", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 0), "hello");
    port.setSelection([addr(0, 0)]);
    await runSwitchSign(port);
    expect((await port.getSelectionCells())[0].value).toBe("hello");
  });

  it("leaves blank cells unchanged", async () => {
    const port = new ExcelPortFake();
    port.setSelection([addr(0, 0)]);
    await runSwitchSign(port);
    const [snap] = await port.getSelectionCells();
    expect(snap.value).toBe(null);
    expect(snap.isFormula).toBe(false);
  });

  it("processes multi-cell selections mixing numbers, formulas, and text as one batch", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 0), 5);
    port.setCellFormula(addr(0, 1), "=SUM(B2:B5)");
    port.setCellValue(addr(0, 2), "skip");
    port.setSelection([addr(0, 0), addr(0, 1), addr(0, 2)]);
    await runSwitchSign(port);
    port.setSelection([addr(0, 0), addr(0, 1), addr(0, 2)]);
    const snaps = await port.getSelectionCells();
    expect(snaps[0].value).toBe(-5);
    expect(snaps[1].formula).toBe("=-(SUM(B2:B5))");
    expect(snaps[2].value).toBe("skip");
  });
});
