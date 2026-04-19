import { describe, expect, it } from "vitest";
import { ExcelPortFake } from "../src/adapters/excelPortFake";
import { runAutoColor } from "../src/services/autoColor.service";
import { CellAddress } from "../src/adapters/excelPort";
import { DEFAULT_AUTO_COLOR_PALETTE } from "../src/core/autoColor";

const addr = (row: number, col: number, sheet = "Sheet1"): CellAddress => ({ sheet, row, col });

describe("Auto-color contract (spec §3.12)", () => {
  it("colors a numeric input blue", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 0), 100);
    port.setSelection([addr(0, 0)]);
    await runAutoColor(port);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fontColor).toBe(DEFAULT_AUTO_COLOR_PALETTE.input);
  });

  it("colors a same-sheet formula green", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "=A1");
    port.setSelection([addr(0, 0)]);
    await runAutoColor(port);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fontColor).toBe(DEFAULT_AUTO_COLOR_PALETTE.worksheetLink);
  });

  it("colors a workbook-link formula purple", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "=Inputs!A1");
    port.setSelection([addr(0, 0)]);
    await runAutoColor(port);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fontColor).toBe(DEFAULT_AUTO_COLOR_PALETTE.workbookLink);
  });

  it("colors an external workbook-link formula cyan", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "='[Model.xlsx]Sheet1'!A1");
    port.setSelection([addr(0, 0)]);
    await runAutoColor(port);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fontColor).toBe(DEFAULT_AUTO_COLOR_PALETTE.external);
  });

  it("colors a hyperlink cell orange", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 0), "click here");
    port.setCellHyperlink(addr(0, 0), true);
    port.setSelection([addr(0, 0)]);
    await runAutoColor(port);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fontColor).toBe(DEFAULT_AUTO_COLOR_PALETTE.hyperlink);
  });

  it("colors a partialInput formula dark red", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "=100+A1");
    port.setSelection([addr(0, 0)]);
    await runAutoColor(port);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fontColor).toBe(DEFAULT_AUTO_COLOR_PALETTE.partialInput);
  });

  it("colors a plain formula black", async () => {
    const port = new ExcelPortFake();
    port.setCellFormula(addr(0, 0), "=IF(TRUE,1,0)");
    port.setSelection([addr(0, 0)]);
    await runAutoColor(port);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fontColor).toBe(DEFAULT_AUTO_COLOR_PALETTE.formula);
  });

  it("does not change blank cells", async () => {
    const port = new ExcelPortFake();
    port.setSelection([addr(0, 0)]);
    await runAutoColor(port);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fontColor).toBe(null);
  });

  it("does not change text values", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 0), "hello");
    port.setSelection([addr(0, 0)]);
    await runAutoColor(port);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fontColor).toBe(null);
  });

  it("applies a custom palette when provided", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 0), 100);
    port.setSelection([addr(0, 0)]);
    const custom = { ...DEFAULT_AUTO_COLOR_PALETTE, input: "#123456" };
    await runAutoColor(port, custom);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fontColor).toBe("#123456");
  });
});
