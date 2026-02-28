export type CellFormatDefinition = {
  name: string;
  fillPattern: "Solid" | "None";
  fillColor: string;
  borderStyle: "None" | "Continuous";
  borderColor: string;
  fontColor: string;
  fontBold: boolean;
  fontItalic: boolean;
  fontUnderline: boolean;
  fontStrikethrough: boolean;
};

export type SelectionCellFormatState = {
  fillPattern: string | null;
  fillColor: string | null;
  fontColor: string | null;
  fontBold: boolean | null;
  fontItalic: boolean | null;
  fontUnderline: boolean | null;
  fontStrikethrough: boolean | null;
  edgeLeftStyle: string | null;
  edgeTopStyle: string | null;
  edgeBottomStyle: string | null;
  edgeRightStyle: string | null;
  edgeLeftColor: string | null;
  edgeTopColor: string | null;
  edgeBottomColor: string | null;
  edgeRightColor: string | null;
};

export const DEFAULT_CELL_FORMATS: CellFormatDefinition[] = [
  {
    name: "Normal",
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
    name: "Inputs",
    fillPattern: "Solid",
    fillColor: "#FFFFCC",
    borderStyle: "Continuous",
    borderColor: "#808080",
    fontColor: "#0000FF",
    fontBold: false,
    fontItalic: false,
    fontUnderline: false,
    fontStrikethrough: false
  },
  {
    name: "Good",
    fillPattern: "Solid",
    fillColor: "#C6EFCE",
    borderStyle: "Continuous",
    borderColor: "#808080",
    fontColor: "#006100",
    fontBold: false,
    fontItalic: false,
    fontUnderline: false,
    fontStrikethrough: false
  },
  {
    name: "Bad",
    fillPattern: "Solid",
    fillColor: "#FFC7CE",
    borderStyle: "Continuous",
    borderColor: "#808080",
    fontColor: "#9C0006",
    fontBold: false,
    fontItalic: false,
    fontUnderline: false,
    fontStrikethrough: false
  },
  {
    name: "Important",
    fillPattern: "Solid",
    fillColor: "#FFFF00",
    borderStyle: "None",
    borderColor: "#000000",
    fontColor: "#000000",
    fontBold: false,
    fontItalic: false,
    fontUnderline: false,
    fontStrikethrough: false
  }
];

function normalizeColor(value: string | null): string | null {
  if (!value) {
    return null;
  }

  const trimmed = value.trim();
  if (trimmed.length === 0) {
    return null;
  }

  const withHash = trimmed.startsWith("#") ? trimmed : `#${trimmed}`;
  return withHash.toUpperCase();
}

function normalizeStyle(value: string | null): string | null {
  if (!value) {
    return null;
  }
  return value.trim().toLowerCase();
}

function normalizeBool(value: boolean | null): boolean {
  return value === true;
}

function doesFillPatternMatch(state: SelectionCellFormatState, expected: CellFormatDefinition): boolean {
  const actualPattern = normalizeStyle(state.fillPattern);
  const expectedPattern = normalizeStyle(expected.fillPattern);
  if (actualPattern === expectedPattern) {
    return true;
  }

  // Excel may report default white fills as "None" even when explicitly set to solid white.
  return expectedPattern === "solid" && normalizeColor(expected.fillColor) === "#FFFFFF" && actualPattern === "none";
}

function doesBorderMatch(state: SelectionCellFormatState, expected: CellFormatDefinition): boolean {
  const styles = [
    normalizeStyle(state.edgeLeftStyle),
    normalizeStyle(state.edgeTopStyle),
    normalizeStyle(state.edgeBottomStyle),
    normalizeStyle(state.edgeRightStyle)
  ];

  const expectedStyle = normalizeStyle(expected.borderStyle);
  if (expectedStyle === "none") {
    return styles.every((style) => style === "none" || style === null);
  }

  const colors = [
    normalizeColor(state.edgeLeftColor),
    normalizeColor(state.edgeTopColor),
    normalizeColor(state.edgeBottomColor),
    normalizeColor(state.edgeRightColor)
  ];
  const expectedColor = normalizeColor(expected.borderColor);

  if (!styles.every((style) => style === expectedStyle)) {
    return false;
  }

  // Border color can be omitted by Excel for some uniform ranges; ignore missing values.
  return colors.every((color) => color === null || color === expectedColor);
}

export function doesSelectionMatchCellFormat(state: SelectionCellFormatState, expected: CellFormatDefinition): boolean {
  const fillPatternMatches = doesFillPatternMatch(state, expected);
  const fillColorMatches = normalizeColor(state.fillColor) === normalizeColor(expected.fillColor);
  const fontColorMatches = normalizeColor(state.fontColor) === normalizeColor(expected.fontColor);

  if (!fillPatternMatches || !fillColorMatches || !fontColorMatches) {
    return false;
  }

  if (
    normalizeBool(state.fontBold) !== expected.fontBold ||
    normalizeBool(state.fontItalic) !== expected.fontItalic ||
    normalizeBool(state.fontUnderline) !== expected.fontUnderline ||
    normalizeBool(state.fontStrikethrough) !== expected.fontStrikethrough
  ) {
    return false;
  }

  return doesBorderMatch(state, expected);
}

export function computeNextCellFormat(
  state: SelectionCellFormatState,
  formats: CellFormatDefinition[] = DEFAULT_CELL_FORMATS
): CellFormatDefinition {
  if (formats.length === 0) {
    throw new Error("formats must contain at least one preset");
  }

  for (let i = 0; i < formats.length; i += 1) {
    if (doesSelectionMatchCellFormat(state, formats[i])) {
      const nextIndex = i < formats.length - 1 ? i + 1 : 0;
      return formats[nextIndex];
    }
  }

  return formats[0];
}
