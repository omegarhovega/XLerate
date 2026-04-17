import { describe, expect, it } from "vitest";
import { ExcelPortFake } from "../src/adapters/excelPortFake";
import { runCycleCellFormat } from "../src/services/cycleCellFormat.service";
import { CellAddress } from "../src/adapters/excelPort";
import { DEFAULT_CELL_FORMATS } from "../src/core/cellFormatCycle";

const addr = (row: number, col: number, sheet = "Sheet1"): CellAddress => ({ sheet, row, col });

describe("Cycle Cell Format contract (spec §3.8)", () => {
  it("applies the first preset (Normal) when no preset matches", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), {
      fillColor: "#ABCDEF",
      fontColor: "#000000",
    });
    port.setSelection([addr(0, 0)]);
    await runCycleCellFormat(port);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fillColor).toBe("#FFFFFF");
    expect(snap.fontColor).toBe("#000000");
  });

  it("applies the second preset (Inputs) when current matches Normal", async () => {
    const port = new ExcelPortFake();
    const normal = DEFAULT_CELL_FORMATS[0];
    port.setCellFormatting(addr(0, 0), {
      fillPattern: normal.fillPattern,
      fillColor: normal.fillColor,
      fontColor: normal.fontColor,
      fontBold: normal.fontBold,
      fontItalic: normal.fontItalic,
      fontUnderline: normal.fontUnderline,
      fontStrikethrough: normal.fontStrikethrough,
    });
    port.setSelection([addr(0, 0)]);
    await runCycleCellFormat(port);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fillColor).toBe("#FFFFCC");
    expect(snap.fontColor).toBe("#0000FF");
  });

  it("wraps from last preset (Important) back to the first (Normal)", async () => {
    const port = new ExcelPortFake();
    const important = DEFAULT_CELL_FORMATS[DEFAULT_CELL_FORMATS.length - 1];
    port.setCellFormatting(addr(0, 0), {
      fillPattern: important.fillPattern,
      fillColor: important.fillColor,
      fontColor: important.fontColor,
      fontBold: important.fontBold,
      fontItalic: important.fontItalic,
      fontUnderline: important.fontUnderline,
      fontStrikethrough: important.fontStrikethrough,
    });
    port.setSelection([addr(0, 0)]);
    await runCycleCellFormat(port);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fillColor).toBe("#FFFFFF");
  });

  it("applies to every cell in the selection", async () => {
    const port = new ExcelPortFake();
    port.setSelection([addr(0, 0), addr(1, 0)]);
    await runCycleCellFormat(port);
    port.setSelection([addr(0, 0), addr(1, 0)]);
    const snaps = await port.getSelectionFormatting();
    expect(snaps[0].fillColor).toBe("#FFFFFF");
    expect(snaps[1].fillColor).toBe("#FFFFFF");
  });
});
