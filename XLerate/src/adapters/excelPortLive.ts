import {
  ActiveCellLeftRowSnapshot,
  AutoColorCellSnapshot,
  BorderEdgeMutation,
  BordersMutation,
  CellAddress,
  CellFormatMutation,
  CellFormattingSnapshot,
  CellMutation,
  CellSnapshot,
  ExcelPort,
  FillMutation,
  FontMutation,
} from "./excelPort";
import { autoColorProbe } from "./autoColorProbe";

/**
 * The only place in the codebase that imports Office.js from domain code.
 * Services depend on the ExcelPort interface; the taskpane constructs this
 * concrete implementation once per feature invocation.
 *
 * Array formula mutations are still deferred (throw). Regular formulas,
 * values, number format, font color, formatBundle, and clearSheetFill are
 * all implemented.
 */
export class ExcelPortLive implements ExcelPort {
  private static readonly MAX_AUTO_COLOR_CELL_COUNT = 50000;
  private static readonly AUTO_COLOR_READ_CHUNK_CELL_BUDGET = 5000;
  private static readonly FONT_COLOR_WRITE_CHUNK_SIZE = 1000;

  async getSelectionCells(): Promise<CellSnapshot[]> {
    return Excel.run(async (context) => {
      const areas = await getSelectedAreas(context);

      for (const area of areas) {
        area.load(["rowCount", "columnCount", "rowIndex", "columnIndex", "values", "formulas"]);
        area.worksheet.load("name");
      }
      await context.sync();

      const formulaArrayCells: Excel.Range[][][] = [];
      for (const area of areas) {
        const areaCells: Excel.Range[][] = [];
        for (let r = 0; r < area.rowCount; r++) {
          const rowCells: Excel.Range[] = [];
          for (let c = 0; c < area.columnCount; c++) {
            const cell = area.getCell(r, c);
            // eslint-disable-next-line office-addins/no-navigational-load
            cell.load("formulaArray");
            rowCells.push(cell);
          }
          areaCells.push(rowCells);
        }
        formulaArrayCells.push(areaCells);
      }
      await context.sync();

      const snapshots: CellSnapshot[] = [];
      for (let areaIndex = 0; areaIndex < areas.length; areaIndex++) {
        const area = areas[areaIndex];
        const worksheet = area.worksheet;
        const areaCells = formulaArrayCells[areaIndex];

        for (let r = 0; r < area.rowCount; r++) {
          for (let c = 0; c < area.columnCount; c++) {
            const formula = area.formulas[r][c];
            const value = area.values[r][c];
            const address: CellAddress = {
              sheet: worksheet.name,
              row: area.rowIndex + r,
              col: area.columnIndex + c,
            };
            const formulaText = typeof formula === "string" ? formula : "";
            const formulaArrayText = areaCells[r][c].formulaArray;
            const isArrayFormula =
              typeof formulaArrayText === "string" && formulaArrayText.length > 0;
            const arrayFormula = isArrayFormula
              ? formatArrayFormulaForSnapshot(formulaArrayText)
              : "";
            const isFormula = isArrayFormula || formulaText.startsWith("=");

            snapshots.push({
              address,
              isFormula,
              isArrayFormula,
              formula: isArrayFormula ? arrayFormula : isFormula ? formulaText : "",
              value: isFormula ? undefined : value,
            });
          }
        }
      }
      return snapshots;
    });
  }

  async getSelectionFormatting(): Promise<CellFormattingSnapshot[]> {
    return Excel.run(async (context) => {
      const areas = await getSelectedAreas(context);
      const cellProperties = areas.map((area) => {
        area.load(["rowCount", "columnCount", "rowIndex", "columnIndex", "numberFormat"]);
        area.worksheet.load("name");
        return area.getCellProperties({
          hyperlink: true,
          format: {
            fill: {
              color: true,
              pattern: true,
            },
            font: {
              name: true,
              size: true,
              color: true,
              bold: true,
              italic: true,
              underline: true,
              strikethrough: true,
            },
            borders: {
              color: true,
              style: true,
            },
          },
        });
      });
      await context.sync();

      const snapshots: CellFormattingSnapshot[] = [];
      for (let areaIndex = 0; areaIndex < areas.length; areaIndex++) {
        const area = areas[areaIndex];
        const worksheet = area.worksheet;
        const props = cellProperties[areaIndex].value;

        for (let r = 0; r < area.rowCount; r++) {
          for (let c = 0; c < area.columnCount; c++) {
            const address: CellAddress = {
              sheet: worksheet.name,
              row: area.rowIndex + r,
              col: area.columnIndex + c,
            };
            const numberFormatCell =
              Array.isArray(area.numberFormat) && Array.isArray(area.numberFormat[r])
                ? String(area.numberFormat[r][c] ?? "General")
                : "General";
            const cellProp = props[r][c];
            const format = cellProp.format;
            const borders = format?.borders;

            snapshots.push({
              address,
              numberFormat: numberFormatCell,
              hasHyperlink: hasCellHyperlink(cellProp.hyperlink),
              fillPattern: format?.fill?.pattern ?? null,
              fillColor: format?.fill?.color ?? null,
              fontName: format?.font?.name ?? null,
              fontSize: typeof format?.font?.size === "number" ? format.font.size : null,
              fontColor: format?.font?.color ?? null,
              fontBold: typeof format?.font?.bold === "boolean" ? format.font.bold : null,
              fontItalic: typeof format?.font?.italic === "boolean" ? format.font.italic : null,
              fontUnderline:
                typeof format?.font?.underline === "string"
                  ? format.font.underline !== "None"
                  : null,
              fontStrikethrough:
                typeof format?.font?.strikethrough === "boolean" ? format.font.strikethrough : null,
              edgeLeftStyle: borders?.left?.style ?? null,
              edgeTopStyle: borders?.top?.style ?? null,
              edgeBottomStyle: borders?.bottom?.style ?? null,
              edgeRightStyle: borders?.right?.style ?? null,
              edgeLeftColor: borders?.left?.color ?? null,
              edgeTopColor: borders?.top?.color ?? null,
              edgeBottomColor: borders?.bottom?.color ?? null,
              edgeRightColor: borders?.right?.color ?? null,
            });
          }
        }
      }
      return snapshots;
    });
  }

  async getActiveCellLeftRowSnapshot(): Promise<ActiveCellLeftRowSnapshot> {
    return Excel.run(async (context) => {
      const workbook = context.workbook;
      const activeCell = workbook.getActiveCell();
      const worksheet = workbook.worksheets.getActiveWorksheet();
      const usedRange = worksheet.getUsedRangeOrNullObject(true);

      activeCell.load(["rowIndex", "columnIndex"]);
      activeCell.worksheet.load("name");
      // eslint-disable-next-line office-addins/no-navigational-load
      usedRange.load(["isNullObject", "columnIndex"]);
      await context.sync();

      const snapshot: ActiveCellLeftRowSnapshot = {
        activeCell: {
          sheet: activeCell.worksheet.name,
          row: activeCell.rowIndex,
          col: activeCell.columnIndex,
        },
        leftCells: [],
      };

      const startCol = usedRange.isNullObject
        ? 0
        : Math.max(0, Math.min(activeCell.columnIndex, usedRange.columnIndex));
      const cellCount = Math.max(0, activeCell.columnIndex - startCol);

      if (cellCount === 0) {
        return snapshot;
      }

      const rowRange = activeCell.worksheet.getRangeByIndexes(
        activeCell.rowIndex,
        startCol,
        1,
        cellCount
      );
      rowRange.load("values");
      await context.sync();

      snapshot.leftCells = Array.from({ length: cellCount }, (_, index) => ({
        address: {
          sheet: activeCell.worksheet.name,
          row: activeCell.rowIndex,
          col: startCol + index,
        },
        value: rowRange.values[0][index],
      }));

      return snapshot;
    });
  }

  async getSelectionAutoColorCells(): Promise<AutoColorCellSnapshot[]> {
    return Excel.run(async (context) => {
      autoColorProbe("20 getSelectionAutoColorCells.enter");
      const snapshots: AutoColorCellSnapshot[] = [];
      const snapshotByKey = new Map<string, AutoColorCellSnapshot>();
      const worksheetCache = new Map<string, Excel.Worksheet>();
      autoColorProbe("21 before-resolve-target-ranges");
      const targetRanges = await resolveAutoColorTargetRanges(context, worksheetCache);
      autoColorProbe("22 after-resolve-target-ranges", {
        count: targetRanges.length,
        sample: targetRanges.slice(0, 10).map((target) => ({
          sheet: target.worksheetName,
          rowIndex: target.rowIndex,
          columnIndex: target.columnIndex,
          rowCount: target.rowCount,
          columnCount: target.columnCount,
        })),
      });
      let totalTargetCellCount = 0;

      for (let targetIndex = 0; targetIndex < targetRanges.length; targetIndex += 1) {
        const target = targetRanges[targetIndex];
        const boundedRowCount = target.rowCount;
        const boundedColumnCount = target.columnCount;
        const chunkRowCount = Math.max(
          1,
          Math.floor(ExcelPortLive.AUTO_COLOR_READ_CHUNK_CELL_BUDGET / boundedColumnCount)
        );
        autoColorProbe("23 target-range-enter", {
          targetIndex,
          sheet: target.worksheetName,
          rowIndex: target.rowIndex,
          columnIndex: target.columnIndex,
          rowCount: boundedRowCount,
          columnCount: boundedColumnCount,
          chunkRowCount,
        });

        for (let offset = 0; offset < boundedRowCount; offset += chunkRowCount) {
          const rowsInChunk = Math.min(chunkRowCount, boundedRowCount - offset);
          const chunkRange = target.worksheet.getRangeByIndexes(
            target.rowIndex + offset,
            target.columnIndex,
            rowsInChunk,
            boundedColumnCount
          );

          const formulaCells = chunkRange.getSpecialCellsOrNullObject(
            Excel.SpecialCellType.formulas
          );
          const constantCells = chunkRange.getSpecialCellsOrNullObject(
            Excel.SpecialCellType.constants
          );
          const textConstantCells = chunkRange.getSpecialCellsOrNullObject(
            Excel.SpecialCellType.constants,
            Excel.SpecialCellValueType.text
          );

          // eslint-disable-next-line office-addins/no-navigational-load
          formulaCells.load(["isNullObject", "cellCount"]);
          // eslint-disable-next-line office-addins/no-navigational-load
          constantCells.load(["isNullObject", "cellCount"]);
          // eslint-disable-next-line office-addins/no-navigational-load
          textConstantCells.load("isNullObject");
          autoColorProbe("24 before-sync-special-cells", {
            targetIndex,
            offset,
            rowsInChunk,
          });
          await context.sync();
          autoColorProbe("25 after-sync-special-cells", {
            targetIndex,
            offset,
            rowsInChunk,
            formulaCellCount: formulaCells.isNullObject ? 0 : formulaCells.cellCount,
            constantCellCount: constantCells.isNullObject ? 0 : constantCells.cellCount,
            textConstantsExist: !textConstantCells.isNullObject,
          });

          totalTargetCellCount += formulaCells.isNullObject ? 0 : formulaCells.cellCount;
          totalTargetCellCount += constantCells.isNullObject ? 0 : constantCells.cellCount;

          if (totalTargetCellCount > ExcelPortLive.MAX_AUTO_COLOR_CELL_COUNT) {
            autoColorProbe("26 abort-too-many-target-cells", {
              totalTargetCellCount,
              limit: ExcelPortLive.MAX_AUTO_COLOR_CELL_COUNT,
            });
            throw new Error(
              `Auto-color supports up to ${ExcelPortLive.MAX_AUTO_COLOR_CELL_COUNT.toLocaleString()} non-empty cells at a time. Narrow the selection and try again.`
            );
          }

          if (!textConstantCells.isNullObject) {
            textConstantCells.areas.load("items");
          }
          if (!formulaCells.isNullObject) {
            formulaCells.areas.load("items");
          }
          if (!constantCells.isNullObject) {
            constantCells.areas.load("items");
          }
          autoColorProbe("27 before-sync-area-items", {
            targetIndex,
            offset,
          });
          await context.sync();
          autoColorProbe("28 after-sync-area-items", {
            targetIndex,
            offset,
            formulaAreaCount: formulaCells.isNullObject ? 0 : formulaCells.areas.items.length,
            constantAreaCount: constantCells.isNullObject ? 0 : constantCells.areas.items.length,
            textAreaCount: textConstantCells.isNullObject
              ? 0
              : textConstantCells.areas.items.length,
          });

          const formulaAreas = formulaCells.isNullObject ? [] : formulaCells.areas.items;
          const constantAreas = constantCells.isNullObject ? [] : constantCells.areas.items;
          const textAreas = textConstantCells.isNullObject ? [] : textConstantCells.areas.items;

          const formulaPayloads = formulaAreas.map((range) => {
            range.load([
              "rowCount",
              "columnCount",
              "rowIndex",
              "columnIndex",
              "values",
              "formulas",
              "numberFormat",
            ]);
            range.worksheet.load("name");
            return range;
          });

          const constantPayloads = constantAreas.map((range) => {
            range.load([
              "rowCount",
              "columnCount",
              "rowIndex",
              "columnIndex",
              "values",
              "formulas",
              "numberFormat",
            ]);
            range.worksheet.load("name");
            return range;
          });

          const textHyperlinkPayloads = textAreas.map((range) => {
            range.load(["rowCount", "columnCount", "rowIndex", "columnIndex"]);
            range.worksheet.load("name");
            return range.getCellProperties({
              hyperlink: true,
            });
          });

          autoColorProbe("29 before-sync-payload-load", {
            targetIndex,
            offset,
            formulaAreas: formulaAreas.length,
            constantAreas: constantAreas.length,
            textAreas: textAreas.length,
          });
          await context.sync();
          autoColorProbe("30 after-sync-payload-load", {
            targetIndex,
            offset,
          });

          for (const area of formulaPayloads) {
            pushAutoColorSnapshots(area, false, snapshots, snapshotByKey);
          }

          for (const area of constantPayloads) {
            pushAutoColorSnapshots(area, true, snapshots, snapshotByKey);
          }

          for (let textIndex = 0; textIndex < textAreas.length; textIndex += 1) {
            const area = textAreas[textIndex];
            const props = textHyperlinkPayloads[textIndex].value;
            markHyperlinks(area, props, snapshotByKey);
          }
          autoColorProbe("31 chunk-complete", {
            targetIndex,
            offset,
            snapshotCount: snapshots.length,
          });
        }
      }

      autoColorProbe("32 getSelectionAutoColorCells.return", {
        snapshotCount: snapshots.length,
      });
      return snapshots;
    });
  }

  async applySelectionFormatBundle(format: CellFormatMutation): Promise<void> {
    await Excel.run(async (context) => {
      const areas = await getSelectedAreas(context);
      for (const area of areas) {
        applyFormatBundle(area, format);
      }
      await context.sync();
    });
  }

  async applyMutations(mutations: CellMutation[]): Promise<void> {
    if (mutations.length === 0) return;

    await Excel.run(async (context) => {
      autoColorProbe("40 applyMutations.enter", {
        mutationCount: mutations.length,
      });
      const sheetCache = new Map<string, Excel.Worksheet>();
      const sheetFor = (name: string): Excel.Worksheet => {
        let s = sheetCache.get(name);
        if (!s) {
          s = context.workbook.worksheets.getItem(name);
          sheetCache.set(name, s);
        }
        return s;
      };

      const fontColorGroups = new Map<string, CellAddress[]>();

      for (const m of mutations) {
        if (m.kind === "fontColor") {
          const groupKey = `${m.address.sheet}|${m.color}`;
          const group = fontColorGroups.get(groupKey) ?? [];
          group.push(m.address);
          fontColorGroups.set(groupKey, group);
          continue;
        }

        const cell = sheetFor(m.address.sheet).getRangeByIndexes(
          m.address.row,
          m.address.col,
          1,
          1
        );
        if (m.kind === "value") {
          cell.values = [[m.value as string | number | boolean | null]];
        } else if (m.kind === "formula") {
          cell.formulas = [[m.formula]];
        } else if (m.kind === "arrayFormula") {
          cell.formulaArray = normalizeFormulaArrayForWrite(m.formula);
        } else if (m.kind === "numberFormat") {
          cell.numberFormat = [[m.format]];
        } else if (m.kind === "formatBundle") {
          applyFormatBundle(cell, m.format);
        }
      }

      for (const [groupKey, addresses] of fontColorGroups) {
        const [sheetName, color] = groupKey.split("|");
        const worksheet = sheetFor(sheetName);
        for (const chunk of chunkArray(addresses, ExcelPortLive.FONT_COLOR_WRITE_CHUNK_SIZE)) {
          const refs = chunk.map((address) => toWorksheetScopedAddress(address)).join(",");
          worksheet.getRanges(refs).format.font.color = color;
        }
      }

      autoColorProbe("41 before-sync-applyMutations", {
        groupCount: fontColorGroups.size,
        groupedTargetCount: Array.from(fontColorGroups.values()).reduce(
          (sum, addresses) => sum + addresses.length,
          0
        ),
      });
      await context.sync();
      autoColorProbe("42 after-sync-applyMutations");
    });
  }

  async clearSheetFill(sheetName: string): Promise<void> {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const used = sheet.getUsedRange(true);
      used.load(["rowCount", "columnCount"]);
      await context.sync();
      if (used.rowCount === 0 || used.columnCount === 0) return;
      used.format.fill.clear();
      await context.sync();
    });
  }
}

async function getSelectedAreas(context: Excel.RequestContext): Promise<Excel.Range[]> {
  const selectedRanges = context.workbook.getSelectedRanges();
  const areas = selectedRanges.areas;
  areas.load("items");
  await context.sync();
  return areas.items;
}

function hasCellHyperlink(hyperlink: Excel.RangeHyperlink | undefined): boolean {
  if (!hyperlink) return false;
  return Boolean(
    hyperlink.address ||
    hyperlink.documentReference ||
    hyperlink.screenTip ||
    hyperlink.textToDisplay
  );
}

function cellKey(address: CellAddress): string {
  return `${address.sheet}!${address.row},${address.col}`;
}

function toWorksheetScopedAddress(address: CellAddress): string {
  return `${columnIndexToLetters(address.col)}${address.row + 1}`;
}

function columnIndexToLetters(columnIndex: number): string {
  let remainder = columnIndex + 1;
  let letters = "";

  while (remainder > 0) {
    const zeroBased = (remainder - 1) % 26;
    letters = String.fromCharCode(65 + zeroBased) + letters;
    remainder = Math.floor((remainder - 1) / 26);
  }

  return letters;
}

function chunkArray<T>(items: T[], chunkSize: number): T[][] {
  const chunks: T[][] = [];
  for (let index = 0; index < items.length; index += chunkSize) {
    chunks.push(items.slice(index, index + chunkSize));
  }
  return chunks;
}

type AutoColorTargetRange = {
  worksheet: Excel.Worksheet;
  worksheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
};

async function resolveAutoColorTargetRanges(
  context: Excel.RequestContext,
  worksheetCache: Map<string, Excel.Worksheet>
): Promise<AutoColorTargetRange[]> {
  autoColorProbe("50 resolveTargetRanges.enter");
  const selectedRange = context.workbook.getSelectedRange();
  selectedRange.load(["rowCount", "columnCount"]);
  autoColorProbe("51 before-sync-selection-shape");
  await context.sync();
  autoColorProbe("52 after-sync-selection-shape", {
    selectedRangeRowCount: selectedRange.rowCount,
    selectedRangeColumnCount: selectedRange.columnCount,
  });

  const isWholeSheetSelection =
    selectedRange.rowCount === 1048576 && selectedRange.columnCount === 16384;
  autoColorProbe("53 whole-sheet-check", {
    isWholeSheetSelection,
  });

  if (isWholeSheetSelection) {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = worksheet.getUsedRangeOrNullObject(true);
    usedRange.load(["isNullObject", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    autoColorProbe("54 before-sync-whole-sheet-used-range");
    await context.sync();
    autoColorProbe("55 after-sync-whole-sheet-used-range", {
      isNullObject: usedRange.isNullObject,
      rowIndex: usedRange.isNullObject ? null : usedRange.rowIndex,
      columnIndex: usedRange.isNullObject ? null : usedRange.columnIndex,
      rowCount: usedRange.isNullObject ? 0 : usedRange.rowCount,
      columnCount: usedRange.isNullObject ? 0 : usedRange.columnCount,
    });

    if (usedRange.isNullObject || usedRange.rowCount === 0 || usedRange.columnCount === 0) {
      autoColorProbe("56 resolveTargetRanges.return-empty-whole-sheet");
      return [];
    }

    const targetRanges = [
      {
        worksheet,
        worksheetName: worksheet.name,
        rowIndex: usedRange.rowIndex,
        columnIndex: usedRange.columnIndex,
        rowCount: usedRange.rowCount,
        columnCount: usedRange.columnCount,
      },
    ];
    autoColorProbe("57 resolveTargetRanges.return-whole-sheet", {
      count: targetRanges.length,
    });
    return targetRanges;
  }

  const selectionAreas = await getSelectedAreas(context);

  for (const area of selectionAreas) {
    area.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    area.worksheet.load("name");
  }
  autoColorProbe("58 before-sync-selection-areas");
  await context.sync();
  autoColorProbe("59 after-sync-selection-areas", {
    selectionAreaCount: selectionAreas.length,
  });

  const usedRanges = selectionAreas.map((selectionArea) => {
    const usedRange = selectionArea.worksheet.getUsedRangeOrNullObject(true);
    usedRange.load(["isNullObject", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    return usedRange;
  });
  autoColorProbe("60 before-sync-area-used-ranges");
  await context.sync();
  autoColorProbe("61 after-sync-area-used-ranges", {
    count: usedRanges.length,
  });

  const targetRanges: AutoColorTargetRange[] = [];
  for (let areaIndex = 0; areaIndex < selectionAreas.length; areaIndex += 1) {
    const selectionArea = selectionAreas[areaIndex];
    const usedRange = usedRanges[areaIndex];

    if (usedRange.isNullObject) {
      continue;
    }

    const rowStart = Math.max(selectionArea.rowIndex, usedRange.rowIndex);
    const columnStart = Math.max(selectionArea.columnIndex, usedRange.columnIndex);
    const rowEnd = Math.min(
      selectionArea.rowIndex + selectionArea.rowCount,
      usedRange.rowIndex + usedRange.rowCount
    );
    const columnEnd = Math.min(
      selectionArea.columnIndex + selectionArea.columnCount,
      usedRange.columnIndex + usedRange.columnCount
    );

    if (rowStart >= rowEnd || columnStart >= columnEnd) {
      continue;
    }

    const worksheetName = selectionArea.worksheet.name;
    let worksheet = worksheetCache.get(worksheetName);
    if (!worksheet) {
      worksheet = context.workbook.worksheets.getItem(worksheetName);
      worksheetCache.set(worksheetName, worksheet);
    }

    targetRanges.push({
      worksheet,
      worksheetName,
      rowIndex: rowStart,
      columnIndex: columnStart,
      rowCount: rowEnd - rowStart,
      columnCount: columnEnd - columnStart,
    });
  }

  autoColorProbe("62 resolveTargetRanges.return", {
    count: targetRanges.length,
  });
  return targetRanges;
}

function pushAutoColorSnapshots(
  area: Excel.Range,
  includeConstants: boolean,
  snapshots: AutoColorCellSnapshot[],
  snapshotByKey: Map<string, AutoColorCellSnapshot>
): void {
  for (let r = 0; r < area.rowCount; r += 1) {
    for (let c = 0; c < area.columnCount; c += 1) {
      const formula = area.formulas[r][c];
      const value = area.values[r][c];
      const formulaText = typeof formula === "string" ? formula : "";
      const isFormula = formulaText.startsWith("=");
      const hasMeaningfulValue =
        isFormula || (value !== null && value !== undefined && value !== "");

      if (!hasMeaningfulValue) {
        continue;
      }

      if (!isFormula && !includeConstants) {
        continue;
      }

      const address: CellAddress = {
        sheet: area.worksheet.name,
        row: area.rowIndex + r,
        col: area.columnIndex + c,
      };
      const snapshot: AutoColorCellSnapshot = {
        address,
        isFormula,
        formula: isFormula ? formulaText : "",
        value: isFormula ? undefined : value,
        numberFormat:
          Array.isArray(area.numberFormat) && Array.isArray(area.numberFormat[r])
            ? String(area.numberFormat[r][c] ?? "General")
            : "General",
        hasHyperlink: false,
      };
      snapshots.push(snapshot);
      snapshotByKey.set(cellKey(address), snapshot);
    }
  }
}

function markHyperlinks(
  area: Excel.Range,
  props: Excel.CellProperties[][],
  snapshotByKey: Map<string, AutoColorCellSnapshot>
): void {
  for (let r = 0; r < area.rowCount; r += 1) {
    for (let c = 0; c < area.columnCount; c += 1) {
      const address: CellAddress = {
        sheet: area.worksheet.name,
        row: area.rowIndex + r,
        col: area.columnIndex + c,
      };
      const snapshot = snapshotByKey.get(cellKey(address));
      if (snapshot && hasCellHyperlink(props[r][c].hyperlink)) {
        snapshot.hasHyperlink = true;
      }
    }
  }
}

function formatArrayFormulaForSnapshot(formulaArray: string): string {
  const trimmed = formulaArray.trim();
  if (trimmed.startsWith("{") && trimmed.endsWith("}")) {
    return trimmed;
  }
  if (trimmed.startsWith("=")) {
    return `{${trimmed}}`;
  }
  return `{=${trimmed}}`;
}

function normalizeFormulaArrayForWrite(formula: string): string {
  const trimmed = formula.trim();
  if (trimmed.startsWith("{") && trimmed.endsWith("}")) {
    return trimmed.slice(1, -1);
  }
  if (trimmed.startsWith("=")) {
    return trimmed;
  }
  return `=${trimmed}`;
}

function applyFontMutation(cell: Excel.Range, font: FontMutation): void {
  if (font.name !== undefined) cell.format.font.name = font.name;
  if (font.size !== undefined) cell.format.font.size = font.size;
  if (font.color !== undefined) cell.format.font.color = font.color;
  if (font.bold !== undefined) cell.format.font.bold = font.bold;
  if (font.italic !== undefined) cell.format.font.italic = font.italic;
  if (font.underline !== undefined) cell.format.font.underline = font.underline ? "Single" : "None";
  if (font.strikethrough !== undefined) cell.format.font.strikethrough = font.strikethrough;
}

function applyFillMutation(cell: Excel.Range, fill: FillMutation): void {
  if (fill.pattern === "None") {
    cell.format.fill.clear();
    return;
  }
  // Order matters: Office.js requires pattern to be set for a color to render
  // as a solid fill on a previously-unfilled cell. Setting color alone does
  // not work reliably. Always set pattern first (when provided), then color.
  if (fill.pattern !== undefined) cell.format.fill.pattern = fill.pattern;
  if (fill.color !== undefined) cell.format.fill.color = fill.color;
}

function applyBorderEdge(cell: Excel.Range, edge: Excel.BorderIndex, m: BorderEdgeMutation): void {
  const border = cell.format.borders.getItem(edge);
  border.style = m.style;
  // Office.js quirk: assigning border.color on a "None"-style border silently
  // upgrades the style to "Continuous". Only set color when the style is
  // visible (not "None"). Mirrors the pre-migration setRangeBorder helper in
  // the old taskpane.ts which had the same guard.
  if (m.style !== "None" && m.color !== undefined) {
    border.color = m.color;
  }
  if (m.style !== "None" && m.weight !== undefined) {
    border.weight = m.weight as Excel.BorderWeight;
  }
}

function applyBordersMutation(cell: Excel.Range, borders: BordersMutation): void {
  if (borders.clearAll) {
    const edges: Excel.BorderIndex[] = [
      Excel.BorderIndex.edgeLeft,
      Excel.BorderIndex.edgeTop,
      Excel.BorderIndex.edgeBottom,
      Excel.BorderIndex.edgeRight,
      Excel.BorderIndex.insideHorizontal,
      Excel.BorderIndex.insideVertical,
    ];
    for (const e of edges) cell.format.borders.getItem(e).style = "None";
  }
  if (borders.left) applyBorderEdge(cell, Excel.BorderIndex.edgeLeft, borders.left);
  if (borders.top) applyBorderEdge(cell, Excel.BorderIndex.edgeTop, borders.top);
  if (borders.bottom) applyBorderEdge(cell, Excel.BorderIndex.edgeBottom, borders.bottom);
  if (borders.right) applyBorderEdge(cell, Excel.BorderIndex.edgeRight, borders.right);
  if (borders.insideHorizontal)
    applyBorderEdge(cell, Excel.BorderIndex.insideHorizontal, borders.insideHorizontal);
  if (borders.insideVertical)
    applyBorderEdge(cell, Excel.BorderIndex.insideVertical, borders.insideVertical);
}

function applyFormatBundle(cell: Excel.Range, format: CellFormatMutation): void {
  if (format.numberFormat !== undefined) cell.numberFormat = [[format.numberFormat]];
  if (format.font) applyFontMutation(cell, format.font);
  if (format.fill) applyFillMutation(cell, format.fill);
  if (format.borders) applyBordersMutation(cell, format.borders);
}
