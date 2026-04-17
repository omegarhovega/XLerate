import { describe, expect, it } from "vitest";
import { ExcelPortFake } from "../src/adapters/excelPortFake";
import { runClearConsistencyMarks } from "../src/services/clearConsistencyMarks.service";
import { CellAddress } from "../src/adapters/excelPort";

const addr = (row: number, col: number, sheet = "Sheet1"): CellAddress => ({ sheet, row, col });

describe("Clear Consistency Marks contract (spec §3.6)", () => {
  it("clears fill color on every cell of the named sheet", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0, "Sheet1"), { fillColor: "#00FF00" });
    port.setCellFormatting(addr(1, 0, "Sheet1"), { fillColor: "#FF0000" });
    await runClearConsistencyMarks(port, "Sheet1");

    port.setSelection([addr(0, 0, "Sheet1"), addr(1, 0, "Sheet1")]);
    const snaps = await port.getSelectionFormatting();
    expect(snaps[0].fillColor).toBe(null);
    expect(snaps[1].fillColor).toBe(null);
  });

  it("leaves other sheets' fills intact", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0, "Sheet1"), { fillColor: "#00FF00" });
    port.setCellFormatting(addr(0, 0, "Sheet2"), { fillColor: "#FF0000" });
    await runClearConsistencyMarks(port, "Sheet1");

    port.setSelection([addr(0, 0, "Sheet2")]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fillColor).toBe("#FF0000");
  });

  it("is a no-op when the sheet has no fills", async () => {
    const port = new ExcelPortFake();
    await runClearConsistencyMarks(port, "Sheet1");
    port.setSelection([addr(0, 0, "Sheet1")]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fillColor).toBe(null);
  });
});
