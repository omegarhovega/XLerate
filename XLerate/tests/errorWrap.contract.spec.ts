import { describe, expect, it } from "vitest";
import { ExcelPortFake } from "../src/adapters/excelPortFake";
import { runErrorWrap } from "../src/services/errorWrap.service";
import { CellAddress } from "../src/adapters/excelPort";

const addr = (row: number, col: number, sheet = "Sheet1"): CellAddress => ({ sheet, row, col });

describe("Error Wrap contract (spec §3.11)", () => {
  it("wraps a formula with default NA() fallback", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "=A1/B1");
    port.setSelection([addr(0, 0)]);
    await runErrorWrap(port);
    port.setSelection([addr(0, 0)]);
    expect((await port.getSelectionCells())[0].formula).toBe("=IFERROR(A1/B1, NA())");
  });

  it("wraps a formula with an explicit fallback value", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "=SUM(A1:A3)");
    port.setSelection([addr(0, 0)]);
    await runErrorWrap(port, "0");
    port.setSelection([addr(0, 0)]);
    expect((await port.getSelectionCells())[0].formula).toBe("=IFERROR(SUM(A1:A3), 0)");
  });

  it("nests existing IFERROR wrappers rather than deduplicating", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "=IFERROR(A1/B1,0)");
    port.setSelection([addr(0, 0)]);
    await runErrorWrap(port);
    port.setSelection([addr(0, 0)]);
    expect((await port.getSelectionCells())[0].formula).toBe("=IFERROR(IFERROR(A1/B1,0), NA())");
  });

  it("leaves numeric constants unchanged", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 0), 100);
    port.setSelection([addr(0, 0)]);
    await runErrorWrap(port);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionCells();
    expect(snap.isFormula).toBe(false);
    expect(snap.value).toBe(100);
  });

  it("leaves text constants unchanged", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 0), "x");
    port.setSelection([addr(0, 0)]);
    await runErrorWrap(port);
    port.setSelection([addr(0, 0)]);
    expect((await port.getSelectionCells())[0].value).toBe("x");
  });

  it("leaves blank cells unchanged", async () => {
    const port = new ExcelPortFake();
    port.setSelection([addr(0, 0)]);
    await runErrorWrap(port);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionCells();
    expect(snap.isFormula).toBe(false);
    expect(snap.value).toBe(null);
  });

  it("processes mixed selections", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "=A1+B1");
    port.setCellValue(addr(0, 1), 42);
    port.setCellFormula(addr(0, 2), "=SUM(A:A)");
    port.setSelection([addr(0, 0), addr(0, 1), addr(0, 2)]);
    await runErrorWrap(port, "0");
    port.setSelection([addr(0, 0), addr(0, 1), addr(0, 2)]);
    const snaps = await port.getSelectionCells();
    expect(snaps[0].formula).toBe("=IFERROR(A1+B1, 0)");
    expect(snaps[1].value).toBe(42);
    expect(snaps[2].formula).toBe("=IFERROR(SUM(A:A), 0)");
  });

  it("wraps an array formula and preserves braces", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "{=A1:A3*2}", true);
    port.setSelection([addr(0, 0)]);
    await runErrorWrap(port, "0");
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionCells();
    expect(snap.isArrayFormula).toBe(true);
    expect(snap.formula).toBe("{=IFERROR(A1:A3*2, 0)}");
  });
});
