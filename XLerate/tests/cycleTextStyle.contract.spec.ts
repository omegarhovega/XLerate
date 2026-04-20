import { describe, expect, it } from "vitest";
import { ExcelPortFake } from "../src/adapters/excelPortFake";
import { runCycleTextStyle } from "../src/services/cycleTextStyle.service";
import { CellAddress } from "../src/adapters/excelPort";
import { DEFAULT_TEXT_STYLES } from "../src/core/textStyleCycle";

const addr = (row: number, col: number, sheet = "Sheet1"): CellAddress => ({ sheet, row, col });

describe("Cycle Text Style contract (spec §3.10)", () => {
  it("first invocation (index -1) applies the first style (Heading) and returns index 0", async () => {
    const port = new ExcelPortFake();
    port.setSelection([addr(0, 0)]);
    const { index } = await runCycleTextStyle(port, -1);
    expect(index).toBe(0);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fontName).toBe("Calibri");
    expect(snap.fontSize).toBe(14);
    expect(snap.fontBold).toBe(true);
    expect(snap.fillColor).toBe("#F0F0F0");
  });

  it("second invocation applies Subheading and returns index 1", async () => {
    const port = new ExcelPortFake();
    port.setSelection([addr(0, 0)]);
    const { index } = await runCycleTextStyle(port, 0);
    expect(index).toBe(1);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fontSize).toBe(12);
    expect(snap.fontColor).toBe("#595959");
  });

  it("wraps from last (Normal) back to the first (Heading)", async () => {
    const port = new ExcelPortFake();
    port.setSelection([addr(0, 0)]);
    const lastIndex = DEFAULT_TEXT_STYLES.length - 1;
    const { index } = await runCycleTextStyle(port, lastIndex);
    expect(index).toBe(0);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fontSize).toBe(14);
  });

  it("clears all borders before applying the new style's borders", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), {
      edgeLeftStyle: "Continuous",
      edgeRightStyle: "Continuous",
    });
    port.setSelection([addr(0, 0)]);
    await runCycleTextStyle(port, -1);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.edgeTopStyle).toBe("Continuous");
    expect(snap.edgeBottomStyle).toBe("Continuous");
    expect(snap.edgeLeftStyle).toBe(null);
    expect(snap.edgeRightStyle).toBe(null);
  });

  it("applies to every cell in the selection", async () => {
    const port = new ExcelPortFake();
    port.setSelection([addr(0, 0), addr(1, 0)]);
    await runCycleTextStyle(port, -1);
    port.setSelection([addr(0, 0), addr(1, 0)]);
    const snaps = await port.getSelectionFormatting();
    expect(snaps.every((s) => s.fontSize === 14)).toBe(true);
  });

  it("supports styles with no fill", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), {
      fillPattern: "Solid",
      fillColor: "#FFFF00",
    });
    port.setSelection([addr(0, 0)]);
    await runCycleTextStyle(port, -1, [
      {
        name: "Plain",
        fontName: "Calibri",
        fontSize: 11,
        bold: false,
        italic: false,
        underline: false,
        fontColor: "#000000",
        fillPattern: "None",
        backColor: "#FFFFFF",
        borderStyle: "None",
        borderTop: false,
        borderBottom: false,
        borderLeft: false,
        borderRight: false,
      },
    ]);

    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fillPattern).toBe(null);
    expect(snap.fillColor).toBe(null);
  });
});
