import { describe, expect, it } from "vitest";
import { ExcelPortFake } from "../src/adapters/excelPortFake";
import { runCycleNumberFormat } from "../src/services/cycleNumberFormat.service";
import { CellAddress } from "../src/adapters/excelPort";

const addr = (row: number, col: number, sheet = "Sheet1"): CellAddress => ({ sheet, row, col });

const FORMATS = [
  { name: "A", formatCode: "#,##0" },
  { name: "B", formatCode: "#,##0.00" },
  { name: "C", formatCode: "0.0%" }
];

describe("Cycle Number Format contract (spec §3.7)", () => {
  it("applies the next format when the current matches a preset", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), { numberFormat: "#,##0" });
    port.setCellValue(addr(0, 0), 1000);
    port.setSelection([addr(0, 0)]);
    await runCycleNumberFormat(port, FORMATS);
    port.setSelection([addr(0, 0)]);
    expect((await port.getSelectionFormatting())[0].numberFormat).toBe("#,##0.00");
  });

  it("wraps to the first format when at the end", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), { numberFormat: "0.0%" });
    port.setSelection([addr(0, 0)]);
    await runCycleNumberFormat(port, FORMATS);
    port.setSelection([addr(0, 0)]);
    expect((await port.getSelectionFormatting())[0].numberFormat).toBe("#,##0");
  });

  it("applies the first format on mixed selection", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), { numberFormat: "#,##0" });
    port.setCellFormatting(addr(0, 1), { numberFormat: "General" });
    port.setSelection([addr(0, 0), addr(0, 1)]);
    await runCycleNumberFormat(port, FORMATS);
    port.setSelection([addr(0, 0), addr(0, 1)]);
    const snaps = await port.getSelectionFormatting();
    expect(snaps[0].numberFormat).toBe("#,##0");
    expect(snaps[1].numberFormat).toBe("#,##0");
  });

  it("applies the first format when the current format is unknown", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), { numberFormat: "[$-409]mmm-yy" });
    port.setSelection([addr(0, 0)]);
    await runCycleNumberFormat(port, FORMATS);
    port.setSelection([addr(0, 0)]);
    expect((await port.getSelectionFormatting())[0].numberFormat).toBe("#,##0");
  });

  it("applies to every cell in the selection in one batch", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), { numberFormat: "#,##0" });
    port.setCellFormatting(addr(1, 0), { numberFormat: "#,##0" });
    port.setCellFormatting(addr(2, 0), { numberFormat: "#,##0" });
    port.setSelection([addr(0, 0), addr(1, 0), addr(2, 0)]);
    await runCycleNumberFormat(port, FORMATS);
    port.setSelection([addr(0, 0), addr(1, 0), addr(2, 0)]);
    const snaps = await port.getSelectionFormatting();
    expect(snaps.every((s) => s.numberFormat === "#,##0.00")).toBe(true);
  });
});
