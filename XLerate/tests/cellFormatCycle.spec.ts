import { describe, expect, it } from "vitest";
import {
  type CellFormatDefinition,
  computeNextCellFormat,
  DEFAULT_CELL_FORMATS,
  doesSelectionMatchCellFormat,
  type SelectionCellFormatState
} from "../src/core/cellFormatCycle";

function stateFromFormat(preset: CellFormatDefinition): SelectionCellFormatState {
  return {
    fillPattern: preset.fillPattern,
    fillColor: preset.fillColor,
    fontColor: preset.fontColor,
    fontBold: preset.fontBold,
    fontItalic: preset.fontItalic,
    fontUnderline: preset.fontUnderline,
    fontStrikethrough: preset.fontStrikethrough,
    edgeLeftStyle: preset.borderStyle,
    edgeTopStyle: preset.borderStyle,
    edgeBottomStyle: preset.borderStyle,
    edgeRightStyle: preset.borderStyle,
    edgeLeftColor: preset.borderColor,
    edgeTopColor: preset.borderColor,
    edgeBottomColor: preset.borderColor,
    edgeRightColor: preset.borderColor
  };
}

function stateFromPreset(index: number): SelectionCellFormatState {
  return stateFromFormat(DEFAULT_CELL_FORMATS[index]);
}

describe("cycle cell format baseline parity", () => {
  it("matches a preset when fill/font/border properties align", () => {
    expect(doesSelectionMatchCellFormat(stateFromPreset(0), DEFAULT_CELL_FORMATS[0])).toBe(true);
    expect(doesSelectionMatchCellFormat(stateFromPreset(2), DEFAULT_CELL_FORMATS[2])).toBe(true);
  });

  it("requires no borders for presets with borderStyle None", () => {
    const normal = stateFromPreset(0);
    normal.edgeLeftStyle = "Continuous";
    expect(doesSelectionMatchCellFormat(normal, DEFAULT_CELL_FORMATS[0])).toBe(false);
  });

  it("accepts Excel-reported default style variants for Normal", () => {
    const normalLike: SelectionCellFormatState = {
      ...stateFromPreset(0),
      fillPattern: "None",
      fontBold: null,
      fontItalic: null,
      fontUnderline: null,
      fontStrikethrough: null,
      edgeLeftStyle: null,
      edgeTopStyle: null,
      edgeBottomStyle: null,
      edgeRightStyle: null
    };

    expect(doesSelectionMatchCellFormat(normalLike, DEFAULT_CELL_FORMATS[0])).toBe(true);
    expect(computeNextCellFormat(normalLike).name).toBe("Inputs");
  });

  it("cycles to next preset when current one matches", () => {
    const current = stateFromPreset(1); // Inputs
    expect(computeNextCellFormat(current).name).toBe("Good");
  });

  it("wraps to first preset from last", () => {
    const current = stateFromPreset(DEFAULT_CELL_FORMATS.length - 1);
    expect(computeNextCellFormat(current).name).toBe(DEFAULT_CELL_FORMATS[0].name);
  });

  it("falls back to first preset when no match exists", () => {
    const unknown: SelectionCellFormatState = {
      ...stateFromPreset(0),
      fillColor: "#ABCDEF"
    };

    expect(computeNextCellFormat(unknown).name).toBe(DEFAULT_CELL_FORMATS[0].name);
  });

  it("uses provided custom configured presets", () => {
    const customFormats: CellFormatDefinition[] = [
      {
        name: "Custom-1",
        fillPattern: "Solid",
        fillColor: "#FFFFFF",
        borderStyle: "None",
        borderColor: "#000000",
        fontColor: "#000000",
        fontBold: false,
        fontItalic: false,
        fontUnderline: false,
        fontStrikethrough: false
      },
      {
        name: "Custom-2",
        fillPattern: "Solid",
        fillColor: "#FFF2CC",
        borderStyle: "Continuous",
        borderColor: "#808080",
        fontColor: "#000000",
        fontBold: true,
        fontItalic: false,
        fontUnderline: false,
        fontStrikethrough: false
      }
    ];

    const current = stateFromFormat(customFormats[0]);
    expect(computeNextCellFormat(current, customFormats).name).toBe("Custom-2");
    expect(computeNextCellFormat(stateFromFormat(customFormats[1]), customFormats).name).toBe("Custom-1");
  });
});
