import { describe, expect, it } from "vitest";
import { DEFAULT_AUTO_COLOR_PALETTE } from "../src/core/autoColor";
import { DEFAULT_CELL_FORMATS } from "../src/core/cellFormatCycle";
import { DEFAULT_DATE_FORMATS } from "../src/core/dateFormatCycle";
import {
  buildDefaultFormatSettings,
  cloneResolvedFormatSettings,
  getFormatSettingsValidationError,
  resolveFormatSettings,
} from "../src/core/formatSettings";
import { DEFAULT_NUMBER_FORMATS } from "../src/core/numberFormatCycle";
import { DEFAULT_TEXT_STYLES } from "../src/core/textStyleCycle";
import { DEFAULT_TRACE_MAX_DEPTH, DEFAULT_TRACE_SAFETY_LIMIT } from "../src/core/traceUtils";

describe("format settings store", () => {
  it("falls back to defaults when settings are missing or invalid", () => {
    const missing = resolveFormatSettings(undefined);
    expect(missing.numberFormats).toEqual(DEFAULT_NUMBER_FORMATS);
    expect(missing.dateFormats).toEqual(DEFAULT_DATE_FORMATS);
    expect(missing.cellFormats).toEqual(DEFAULT_CELL_FORMATS);
    expect(missing.textStyles).toEqual(DEFAULT_TEXT_STYLES);
    expect(missing.autoColorPalette).toEqual(DEFAULT_AUTO_COLOR_PALETTE);
    expect(missing.trace).toEqual({
      maxDepth: DEFAULT_TRACE_MAX_DEPTH,
      safetyLimit: DEFAULT_TRACE_SAFETY_LIMIT,
    });

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
      autoColorPalette: {
        input: "#112233",
        formula: "#223344",
        worksheetLink: "#334455",
        workbookLink: "#445566",
        external: "#556677",
        hyperlink: "#667788",
        partialInput: "#778899",
      },
      trace: {
        maxDepth: 12,
        safetyLimit: 900,
      },
      textStyles: [
        {
          name: "TS-A",
          fontName: "Calibri",
          fontSize: 11,
          bold: false,
          italic: false,
          underline: false,
          fontColor: "#000000",
          fillPattern: "Solid",
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
    expect(resolved.autoColorPalette.input).toBe("#112233");
    expect(resolved.trace).toEqual({ maxDepth: 12, safetyLimit: 900 });
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
    expect(resolved.autoColorPalette).toEqual(DEFAULT_AUTO_COLOR_PALETTE);
    expect(resolved.trace).toEqual({
      maxDepth: DEFAULT_TRACE_MAX_DEPTH,
      safetyLimit: DEFAULT_TRACE_SAFETY_LIMIT,
    });
  });

  it("keeps older text-style settings compatible by defaulting missing fillPattern to Solid", () => {
    const resolved = resolveFormatSettings({
      textStyles: [
        {
          name: "Legacy",
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
          borderRight: false,
        },
      ],
    });

    expect(resolved.textStyles).toEqual([
      {
        name: "Legacy",
        fontName: "Calibri",
        fontSize: 11,
        bold: false,
        italic: false,
        underline: false,
        fontColor: "#000000",
        fillPattern: "Solid",
        backColor: "#FFFFFF",
        borderStyle: "None",
        borderTop: false,
        borderBottom: false,
        borderLeft: false,
        borderRight: false,
      },
    ]);
  });

  it("returns cloned structures so defaults are not mutated by callers", () => {
    const resolved = resolveFormatSettings(undefined);
    resolved.numberFormats[0].name = "MUTATED";
    resolved.autoColorPalette.input = "#ABCDEF";
    resolved.trace.maxDepth = 99;
    expect(DEFAULT_NUMBER_FORMATS[0].name).not.toBe("MUTATED");
    expect(DEFAULT_AUTO_COLOR_PALETTE.input).not.toBe("#ABCDEF");
    expect(DEFAULT_TRACE_MAX_DEPTH).not.toBe(99);
  });

  it("builds and clones default settings bundles", () => {
    const defaults = buildDefaultFormatSettings();
    const clone = cloneResolvedFormatSettings(defaults);

    clone.textStyles[0].name = "Changed";
    clone.autoColorPalette.formula = "#111111";
    clone.trace.safetyLimit = 1234;

    expect(defaults.textStyles[0].name).not.toBe("Changed");
    expect(defaults.autoColorPalette.formula).not.toBe("#111111");
    expect(defaults.trace.safetyLimit).toBe(DEFAULT_TRACE_SAFETY_LIMIT);
  });

  it("surfaces save-time validation errors for incomplete settings", () => {
    const invalid = buildDefaultFormatSettings();
    invalid.numberFormats[0].name = "";

    expect(getFormatSettingsValidationError(invalid)).toBe(
      "Every number format needs a name and an Excel format code."
    );
    expect(getFormatSettingsValidationError(buildDefaultFormatSettings())).toBeNull();
  });
});
