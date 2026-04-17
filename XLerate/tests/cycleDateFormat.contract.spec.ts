import { describe, expect, it } from "vitest";
import { ExcelPortFake } from "../src/adapters/excelPortFake";
import { runCycleDateFormat } from "../src/services/cycleDateFormat.service";
import { CellAddress } from "../src/adapters/excelPort";

const addr = (row: number, col: number, sheet = "Sheet1"): CellAddress => ({ sheet, row, col });

const FORMATS = [
  { name: "A", formatCode: "yyyy" },
  { name: "B", formatCode: "mmm-yyyy" },
  { name: "C", formatCode: "dd-mmm-yy" }
];

describe("Cycle Date Format contract (spec §3.9)", () => {
  it("applies the next format when current matches a preset", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), { numberFormat: "yyyy" });
    port.setSelection([addr(0, 0)]);
    await runCycleDateFormat(port, FORMATS);
    port.setSelection([addr(0, 0)]);
    expect((await port.getSelectionFormatting())[0].numberFormat).toBe("mmm-yyyy");
  });

  it("wraps to the first format when at the end", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), { numberFormat: "dd-mmm-yy" });
    port.setSelection([addr(0, 0)]);
    await runCycleDateFormat(port, FORMATS);
    port.setSelection([addr(0, 0)]);
    expect((await port.getSelectionFormatting())[0].numberFormat).toBe("yyyy");
  });

  it("applies the first format on mixed selection", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), { numberFormat: "yyyy" });
    port.setCellFormatting(addr(0, 1), { numberFormat: "General" });
    port.setSelection([addr(0, 0), addr(0, 1)]);
    await runCycleDateFormat(port, FORMATS);
    port.setSelection([addr(0, 0), addr(0, 1)]);
    const snaps = await port.getSelectionFormatting();
    expect(snaps[0].numberFormat).toBe("yyyy");
    expect(snaps[1].numberFormat).toBe("yyyy");
  });

  it("applies the first format when current is unknown", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), { numberFormat: "#,##0" });
    port.setSelection([addr(0, 0)]);
    await runCycleDateFormat(port, FORMATS);
    port.setSelection([addr(0, 0)]);
    expect((await port.getSelectionFormatting())[0].numberFormat).toBe("yyyy");
  });
});
