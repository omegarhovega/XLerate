import { describe, expect, it } from "vitest";
import { DEFAULT_CELL_FORMATS } from "../src/core/cellFormatCycle";
import { DEFAULT_DATE_FORMATS } from "../src/core/dateFormatCycle";
import { resolveFormatSettings } from "../src/core/formatSettings";
import { DEFAULT_NUMBER_FORMATS } from "../src/core/numberFormatCycle";
import { DEFAULT_TEXT_STYLES } from "../src/core/textStyleCycle";

describe("format settings store", () => {
  it("falls back to defaults when settings are missing or invalid", () => {
    const missing = resolveFormatSettings(undefined);
    expect(missing.numberFormats).toEqual(DEFAULT_NUMBER_FORMATS);
    expect(missing.dateFormats).toEqual(DEFAULT_DATE_FORMATS);
    expect(missing.cellFormats).toEqual(DEFAULT_CELL_FORMATS);
    expect(missing.textStyles).toEqual(DEFAULT_TEXT_STYLES);

    const invalidJson = resolveFormatSettings("{invalid json");
    expect(invalidJson.numberFormats).toEqual(DEFAULT_NUMBER_FORMATS);
  });

  it("accepts valid custom settings from JSON string", () => {
    const customRaw = JSON.stringify({
      numberFormats: [{ name: "NF-A", formatCode: "0.0" }],
      dateFormats: [{ name: "DF-A", formatCode: "yyyy-mm" }],
      cellFormats: [
        {
          name: "CF-A",
          fillPattern: "Solid",
          fillColor: "#FFFFFF",
          borderStyle: "None",
          borderColor: "#000000",
          fontColor: "#000000",
          fontBold: false,
          fontItalic: false,
          fontUnderline: false,
          fontStrikethrough: false
        }
      ],
      textStyles: [
        {
          name: "TS-A",
          fontName: "Calibri",
          fontSize: 11,
          bold: false,
          italic: false,
          underline: false,
          fontColor: "#000000",
          backColor: "#FFFFFF",
          borderStyle: "None",
          borderTop: false,
          borderBottom: false,
          borderLeft: false,
          borderRight: false
        }
      ]
    });

    const resolved = resolveFormatSettings(customRaw);
    expect(resolved.numberFormats).toEqual([{ name: "NF-A", formatCode: "0.0" }]);
    expect(resolved.dateFormats).toEqual([{ name: "DF-A", formatCode: "yyyy-mm" }]);
    expect(resolved.cellFormats).toHaveLength(1);
    expect(resolved.textStyles).toHaveLength(1);
  });

  it("uses defaults for invalid or empty lists while keeping valid lists", () => {
    const resolved = resolveFormatSettings({
      numberFormats: [{ name: "NF-A", formatCode: "0.0" }, { name: "", formatCode: "" }],
      dateFormats: [],
      cellFormats: [{ name: "bad", fillPattern: "Solid" }],
      textStyles: [{ name: "bad", fontName: "Calibri" }]
    });

    expect(resolved.numberFormats).toEqual([{ name: "NF-A", formatCode: "0.0" }]);
    expect(resolved.dateFormats).toEqual(DEFAULT_DATE_FORMATS);
    expect(resolved.cellFormats).toEqual(DEFAULT_CELL_FORMATS);
    expect(resolved.textStyles).toEqual(DEFAULT_TEXT_STYLES);
  });

  it("returns cloned arrays so defaults are not mutated by callers", () => {
    const resolved = resolveFormatSettings(undefined);
    resolved.numberFormats[0].name = "MUTATED";
    expect(DEFAULT_NUMBER_FORMATS[0].name).not.toBe("MUTATED");
  });
});
