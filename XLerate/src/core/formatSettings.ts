import { type CellFormatDefinition, DEFAULT_CELL_FORMATS } from "./cellFormatCycle";
import { type DateFormatDefinition, DEFAULT_DATE_FORMATS } from "./dateFormatCycle";
import { type NumberFormatDefinition, DEFAULT_NUMBER_FORMATS } from "./numberFormatCycle";
import { type TextStyleDefinition, DEFAULT_TEXT_STYLES } from "./textStyleCycle";

export const FORMAT_SETTINGS_KEY = "xlerate_format_settings_v1";

type PersistedFormatSettings = Partial<{
  numberFormats: NumberFormatDefinition[];
  dateFormats: DateFormatDefinition[];
  cellFormats: CellFormatDefinition[];
  textStyles: TextStyleDefinition[];
}>;

export type ResolvedFormatSettings = {
  numberFormats: NumberFormatDefinition[];
  dateFormats: DateFormatDefinition[];
  cellFormats: CellFormatDefinition[];
  textStyles: TextStyleDefinition[];
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

function cloneList<T>(items: T[]): T[] {
  return items.map((item) => ({ ...item }));
}

function resolveList<T>(value: unknown, fallback: T[], isItem: (item: unknown) => item is T): T[] {
  if (!Array.isArray(value)) {
    return cloneList(fallback);
  }

  const valid = value.filter(isItem);
  if (valid.length === 0) {
    return cloneList(fallback);
  }

  return cloneList(valid);
}

function asPersistedFormatSettings(raw: unknown): PersistedFormatSettings | null {
  if (typeof raw === "string") {
    if (raw.trim().length === 0) {
      return null;
    }

    try {
      const parsed = JSON.parse(raw);
      if (!isRecord(parsed)) {
        return null;
      }
      return parsed as PersistedFormatSettings;
    } catch {
      return null;
    }
  }

  if (!isRecord(raw)) {
    return null;
  }

  return raw as PersistedFormatSettings;
}

function isNumberFormatDefinition(value: unknown): value is NumberFormatDefinition {
  if (!isRecord(value)) {
    return false;
  }
  return isNonEmptyString(value.name) && isNonEmptyString(value.formatCode);
}

function isDateFormatDefinition(value: unknown): value is DateFormatDefinition {
  if (!isRecord(value)) {
    return false;
  }
  return isNonEmptyString(value.name) && isNonEmptyString(value.formatCode);
}

function isCellFormatDefinition(value: unknown): value is CellFormatDefinition {
  if (!isRecord(value)) {
    return false;
  }

  if (
    !isNonEmptyString(value.name) ||
    (value.fillPattern !== "Solid" && value.fillPattern !== "None") ||
    !isNonEmptyString(value.fillColor) ||
    (value.borderStyle !== "None" && value.borderStyle !== "Continuous") ||
    !isNonEmptyString(value.borderColor) ||
    !isNonEmptyString(value.fontColor)
  ) {
    return false;
  }

  return (
    typeof value.fontBold === "boolean" &&
    typeof value.fontItalic === "boolean" &&
    typeof value.fontUnderline === "boolean" &&
    typeof value.fontStrikethrough === "boolean"
  );
}

function isTextStyleDefinition(value: unknown): value is TextStyleDefinition {
  if (!isRecord(value)) {
    return false;
  }

  if (
    !isNonEmptyString(value.name) ||
    !isNonEmptyString(value.fontName) ||
    typeof value.fontSize !== "number" ||
    !Number.isFinite(value.fontSize) ||
    value.fontSize <= 0 ||
    !isNonEmptyString(value.fontColor) ||
    !isNonEmptyString(value.backColor)
  ) {
    return false;
  }

  const validBorderStyle =
    value.borderStyle === "None" ||
    value.borderStyle === "Continuous" ||
    value.borderStyle === "Double" ||
    value.borderStyle === "Dash" ||
    value.borderStyle === "Dot";

  if (!validBorderStyle) {
    return false;
  }

  return (
    typeof value.bold === "boolean" &&
    typeof value.italic === "boolean" &&
    typeof value.underline === "boolean" &&
    typeof value.borderTop === "boolean" &&
    typeof value.borderBottom === "boolean" &&
    typeof value.borderLeft === "boolean" &&
    typeof value.borderRight === "boolean"
  );
}

export function resolveFormatSettings(raw: unknown): ResolvedFormatSettings {
  const parsed = asPersistedFormatSettings(raw);
  if (!parsed) {
    return {
      numberFormats: cloneList(DEFAULT_NUMBER_FORMATS),
      dateFormats: cloneList(DEFAULT_DATE_FORMATS),
      cellFormats: cloneList(DEFAULT_CELL_FORMATS),
      textStyles: cloneList(DEFAULT_TEXT_STYLES)
    };
  }

  return {
    numberFormats: resolveList(parsed.numberFormats, DEFAULT_NUMBER_FORMATS, isNumberFormatDefinition),
    dateFormats: resolveList(parsed.dateFormats, DEFAULT_DATE_FORMATS, isDateFormatDefinition),
    cellFormats: resolveList(parsed.cellFormats, DEFAULT_CELL_FORMATS, isCellFormatDefinition),
    textStyles: resolveList(parsed.textStyles, DEFAULT_TEXT_STYLES, isTextStyleDefinition)
  };
}
