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
      const range = context.workbook.getSelectedRange();
      range.load(["rowCount", "columnCount", "rowIndex", "columnIndex", "values", "formulas"]);
      const worksheet = range.worksheet;
      worksheet.load("name");
      await context.sync();

      const snapshots: CellSnapshot[] = [];
      for (let r = 0; r < range.rowCount; r++) {
        for (let c = 0; c < range.columnCount; c++) {
          const formula = range.formulas[r][c];
          const value = range.values[r][c];
          const address: CellAddress = {
            sheet: worksheet.name,
            row: range.rowIndex + r,
            col: range.columnIndex + c,
          };
          const formulaText = typeof formula === "string" ? formula : "";
          const isFormula = formulaText.startsWith("=") || formulaText.startsWith("{=");
          const isArrayFormula = formulaText.startsWith("{=");
          snapshots.push({
            address,
            isFormula,
            isArrayFormula,
            formula: isFormula ? formulaText : "",
            value: isFormula ? undefined : value,
          });
        }
      }
      return snapshots;
    });
  }

  async getSelectionFormatting(): Promise<CellFormattingSnapshot[]> {
    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load([
        "rowCount",
        "columnCount",
        "rowIndex",
        "columnIndex",
        "numberFormat",
        "hyperlink",
      ]);
      const format = range.format;
      format.load(["fill/color", "fill/pattern"]);
      format.font.load(["name", "size", "color", "bold", "italic", "underline", "strikethrough"]);
      const borders = format.borders;
      const edges: Excel.BorderIndex[] = [
        Excel.BorderIndex.edgeLeft,
        Excel.BorderIndex.edgeTop,
        Excel.BorderIndex.edgeBottom,
        Excel.BorderIndex.edgeRight,
      ];
      const edgeItems = edges.map((e) => borders.getItem(e));
      edgeItems.forEach((b) => b.load(["style", "color"]));
      const worksheet = range.worksheet;
      worksheet.load("name");
      await context.sync();

      const [edgeLeft, edgeTop, edgeBottom, edgeRight] = edgeItems;

      const snapshots: CellFormattingSnapshot[] = [];
      for (let r = 0; r < range.rowCount; r++) {
        for (let c = 0; c < range.columnCount; c++) {
          const address: CellAddress = {
            sheet: worksheet.name,
            row: range.rowIndex + r,
            col: range.columnIndex + c,
          };
          const numberFormatCell =
            Array.isArray(range.numberFormat) && Array.isArray(range.numberFormat[r])
              ? String(range.numberFormat[r][c] ?? "General")
              : "General";
          const hyperlinkRaw = (range as unknown as { hyperlink?: unknown }).hyperlink;
          const hyperlinkCell =
            Array.isArray(hyperlinkRaw) && Array.isArray((hyperlinkRaw as unknown[])[r])
              ? Boolean(((hyperlinkRaw as unknown[])[r] as unknown[])[c])
              : false;
          snapshots.push({
            address,
            numberFormat: numberFormatCell,
            hasHyperlink: hyperlinkCell,
            fillPattern: format.fill.pattern ?? null,
            fillColor: format.fill.color ?? null,
            fontName: format.font.name ?? null,
            fontSize: typeof format.font.size === "number" ? format.font.size : null,
            fontColor: format.font.color ?? null,
            fontBold: typeof format.font.bold === "boolean" ? format.font.bold : null,
            fontItalic: typeof format.font.italic === "boolean" ? format.font.italic : null,
            fontUnderline:
              typeof format.font.underline === "string" ? format.font.underline !== "None" : null,
            fontStrikethrough:
              typeof format.font.strikethrough === "boolean" ? format.font.strikethrough : null,
            edgeLeftStyle: edgeLeft.style ?? null,
            edgeTopStyle: edgeTop.style ?? null,
            edgeBottomStyle: edgeBottom.style ?? null,
            edgeRightStyle: edgeRight.style ?? null,
            edgeLeftColor: edgeLeft.color ?? null,
            edgeTopColor: edgeTop.color ?? null,
            edgeBottomColor: edgeBottom.color ?? null,
            edgeRightColor: edgeRight.color ?? null,
          });
        }
      }
      return snapshots;
    });
  }

  async applyMutations(mutations: CellMutation[]): Promise<void> {
    if (mutations.length === 0) return;

    const arrayFormulaMutations = mutations.filter((m) => m.kind === "arrayFormula");
    if (arrayFormulaMutations.length > 0) {
      throw new Error(
        "ExcelPortLive: array formula mutation is not yet supported. This will be addressed in a later phase."
      );
    }

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
  // Order matters: Office.js requires pattern to be set for a color to render as a
  // solid fill on a previously-unfilled cell. Setting color alone does not work
  // reliably. Always set pattern first (when provided), then color.
  if (fill.pattern !== undefined) cell.format.fill.pattern = fill.pattern;
  if (fill.color !== undefined) cell.format.fill.color = fill.color;
}

function applyBorderEdge(cell: Excel.Range, edge: Excel.BorderIndex, m: BorderEdgeMutation): void {
  const border = cell.format.borders.getItem(edge);
  border.style = m.style;
  if (m.color !== undefined) border.color = m.color;
  if (m.weight !== undefined) border.weight = m.weight as Excel.BorderWeight;
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
