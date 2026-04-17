import { describe, expect, it } from "vitest";
import { ExcelPortFake } from "../src/adapters/excelPortFake";
import {
  ConsistencyMarkRestore,
  runClearConsistencyMarks,
} from "../src/services/clearConsistencyMarks.service";
import { CellAddress } from "../src/adapters/excelPort";

const addr = (row: number, col: number, sheet = "Sheet1"): CellAddress => ({ sheet, row, col });

describe("Clear Consistency Marks contract (spec §3.6)", () => {
  it("restores each cell to its original fill color", async () => {
    const port = new ExcelPortFake();
    // Simulate: cells currently marked green/red by the consistency check
    port.setCellFormatting(addr(0, 0), { fillPattern: "Solid", fillColor: "#00FF00" });
    port.setCellFormatting(addr(1, 0), { fillPattern: "Solid", fillColor: "#FF0000" });

    const restores: ConsistencyMarkRestore[] = [
      { address: addr(0, 0), originalColor: "#FFFF00" }, // was yellow before marking
      { address: addr(1, 0), originalColor: null }, // had no fill before marking
    ];
    await runClearConsistencyMarks(port, restores);

    port.setSelection([addr(0, 0), addr(1, 0)]);
    const snaps = await port.getSelectionFormatting();
    expect(snaps[0].fillColor).toBe("#FFFF00");
    expect(snaps[0].fillPattern).toBe("Solid");
    expect(snaps[1].fillColor).toBe(null);
  });

  it("leaves cells not in the restore list untouched", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), { fillPattern: "Solid", fillColor: "#00FF00" });
    // This cell was never marked and is not in restores
    port.setCellFormatting(addr(5, 5), { fillPattern: "Solid", fillColor: "#ABCDEF" });

    const restores: ConsistencyMarkRestore[] = [{ address: addr(0, 0), originalColor: null }];
    await runClearConsistencyMarks(port, restores);

    port.setSelection([addr(5, 5)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fillColor).toBe("#ABCDEF");
  });

  it("only touches cells on the specified addresses (including sheet scope)", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0, "Sheet1"), { fillPattern: "Solid", fillColor: "#00FF00" });
    port.setCellFormatting(addr(0, 0, "Sheet2"), { fillPattern: "Solid", fillColor: "#00FF00" });

    const restores: ConsistencyMarkRestore[] = [
      { address: addr(0, 0, "Sheet1"), originalColor: null },
    ];
    await runClearConsistencyMarks(port, restores);

    port.setSelection([addr(0, 0, "Sheet2")]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fillColor).toBe("#00FF00");
  });

  it("is a no-op when the restore list is empty", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), { fillPattern: "Solid", fillColor: "#FFFF00" });
    await runClearConsistencyMarks(port, []);
    port.setSelection([addr(0, 0)]);
    const [snap] = await port.getSelectionFormatting();
    expect(snap.fillColor).toBe("#FFFF00");
  });
});
