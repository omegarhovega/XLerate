import {
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

  async applyMutations(mutations: CellMutation[]): Promise<void> {
    if (mutations.length === 0) return;

    await Excel.run(async (context) => {
      const sheetCache = new Map<string, Excel.Worksheet>();
      const sheetFor = (name: string): Excel.Worksheet => {
        let s = sheetCache.get(name);
        if (!s) {
          s = context.workbook.worksheets.getItem(name);
          sheetCache.set(name, s);
        }
        return s;
      };

      for (const m of mutations) {
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
        } else if (m.kind === "fontColor") {
          cell.format.font.color = m.color;
        } else if (m.kind === "formatBundle") {
          applyFormatBundle(cell, m.format);
        }
      }
      await context.sync();
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
