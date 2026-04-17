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

  it("returns empty formatting snapshots when no selection", async () => {
    const port = new ExcelPortFake();
    expect(await port.getSelectionFormatting()).toEqual([]);
  });

  it("reports default formatting for a cell with a value only", async () => {
    const port = new ExcelPortFake();
    port.setCellValue(addr(0, 0), 100);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.address).toEqual(addr(0, 0));
    expect(snap.numberFormat).toBe("General");
    expect(snap.hasHyperlink).toBe(false);
    expect(snap.fontColor).toBe(null);
    expect(snap.fillColor).toBe(null);
  });

  it("lets tests seed formatting via setCellFormatting", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), {
      numberFormat: "#,##0.00",
      fillColor: "#FFFF00",
      fontColor: "#0000FF",
      fontBold: true,
    });
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.numberFormat).toBe("#,##0.00");
    expect(snap.fillColor).toBe("#FFFF00");
    expect(snap.fontColor).toBe("#0000FF");
    expect(snap.fontBold).toBe(true);
  });

  it("lets tests seed hyperlink presence", async () => {
    const port = new ExcelPortFake();
    port.setCellHyperlink(addr(0, 0), true);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.hasHyperlink).toBe(true);
  });

  it("applies numberFormat mutations", async () => {
    const port = new ExcelPortFake();
    await port.applyMutations([{ address: addr(0, 0), kind: "numberFormat", format: "0.0%" }]);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.numberFormat).toBe("0.0%");
  });

  it("applies fontColor mutations", async () => {
    const port = new ExcelPortFake();
    await port.applyMutations([{ address: addr(0, 0), kind: "fontColor", color: "#00FF00" }]);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fontColor).toBe("#00FF00");
  });

  it("applies formatBundle mutations with fill, font, and borders", async () => {
    const port = new ExcelPortFake();
    await port.applyMutations([
      {
        address: addr(0, 0),
        kind: "formatBundle",
        format: {
          fill: { pattern: "Solid", color: "#FFFFCC" },
          font: { color: "#0000FF", bold: true },
          borders: {
            clearAll: true,
            top: { style: "Continuous", color: "#808080" },
            bottom: { style: "Continuous", color: "#808080" },
          },
        },
      },
    ]);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fillPattern).toBe("Solid");
    expect(snap.fillColor).toBe("#FFFFCC");
    expect(snap.fontColor).toBe("#0000FF");
    expect(snap.fontBold).toBe(true);
    expect(snap.edgeTopStyle).toBe("Continuous");
    expect(snap.edgeTopColor).toBe("#808080");
    expect(snap.edgeBottomStyle).toBe("Continuous");
    expect(snap.edgeLeftStyle).toBe(null);
    expect(snap.edgeRightStyle).toBe(null);
  });

  it("clearAll in borders wipes existing edges before applying new ones", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), {
      edgeLeftStyle: "Continuous",
      edgeTopStyle: "Continuous",
      edgeBottomStyle: "Continuous",
      edgeRightStyle: "Continuous",
    });
    await port.applyMutations([
      {
        address: addr(0, 0),
        kind: "formatBundle",
        format: {
          borders: { clearAll: true, top: { style: "Double" } },
        },
      },
    ]);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.edgeTopStyle).toBe("Double");
    expect(snap.edgeLeftStyle).toBe(null);
    expect(snap.edgeBottomStyle).toBe(null);
    expect(snap.edgeRightStyle).toBe(null);
  });

  it("clearSheetFill removes fills on the named sheet only", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0, "Sheet1"), { fillColor: "#FFFF00" });
    port.setCellFormatting(addr(0, 0, "Sheet2"), { fillColor: "#FF0000" });
    await port.clearSheetFill("Sheet1");

    port.setSelection([addr(0, 0, "Sheet1")]);
    expect((await port.getSelectionFormatting())[0].fillColor).toBe(null);

    port.setSelection([addr(0, 0, "Sheet2")]);
    expect((await port.getSelectionFormatting())[0].fillColor).toBe("#FF0000");
  });
});
