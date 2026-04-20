import { describe, expect, it } from "vitest";
import { ExcelPortFake } from "../src/adapters/excelPortFake";
import { CellAddress } from "../src/adapters/excelPort";
import { runInsertCagr } from "../src/services/cagr.service";

const addr = (row: number, col: number, sheet = "Sheet1"): CellAddress => ({ sheet, row, col });

describe("Insert CAGR contract (spec §3.13)", () => {
  it("inserts a CAGR formula from the contiguous numeric series to the left", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(4, 1), 100);
    port.setCellValue(addr(4, 2), 110);
    port.setCellValue(addr(4, 3), 121);
    port.setSelection([addr(4, 4)]);

    const result = await runInsertCagr(port);

    expect(result).toEqual({
      ok: true,
      destination: "E5",
      sourceRange: "B5:D5",
      insertedFormula: "=POWER(D5/B5,1/2)-1",
      periodCount: 2,
    });

    port.setSelection([addr(4, 4)]);
    const [destination] = await port.getSelectionCells();
    const [formatting] = await port.getSelectionFormatting();
    expect(destination.formula).toBe("=POWER(D5/B5,1/2)-1");
    expect(formatting.numberFormat).toBe("0.0%");
  });

  it("stops at the first non-numeric boundary to the left", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 0), 100);
    port.setCellValue(addr(0, 1), 110);
    port.setCellValue(addr(0, 2), "blocked");
    port.setCellValue(addr(0, 3), 121);
    port.setSelection([addr(0, 4)]);

    const result = await runInsertCagr(port);

    expect(result).toEqual({
      ok: false,
      reason: "no_series",
      destination: "E1",
    });
  });

  it("does not insert when fewer than two numeric cells are adjacent to the destination", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 2), 121);
    port.setSelection([addr(0, 3)]);

    const result = await runInsertCagr(port);

    expect(result).toEqual({
      ok: false,
      reason: "no_series",
      destination: "D1",
    });

    port.setSelection([addr(0, 3)]);
    const [destination] = await port.getSelectionCells();
    expect(destination.isFormula).toBe(false);
    expect(destination.value).toBe(null);
  });
});
