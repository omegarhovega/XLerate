import { type AutoColorPalette, DEFAULT_AUTO_COLOR_PALETTE } from "./autoColor";
import { type CellFormatDefinition, DEFAULT_CELL_FORMATS } from "./cellFormatCycle";
import { type DateFormatDefinition, DEFAULT_DATE_FORMATS } from "./dateFormatCycle";
import { type NumberFormatDefinition, DEFAULT_NUMBER_FORMATS } from "./numberFormatCycle";
import { type TextStyleDefinition, DEFAULT_TEXT_STYLES } from "./textStyleCycle";
import {
  DEFAULT_TRACE_MAX_DEPTH,
  DEFAULT_TRACE_SAFETY_LIMIT,
  sanitizeTraceDepth,
  sanitizeTraceSafetyLimit,
} from "./traceUtils";

export const FORMAT_SETTINGS_KEY = "xlerate_format_settings_v1";

const AUTO_COLOR_KEYS: Array<keyof AutoColorPalette> = [
  "input",
  "formula",
  "worksheetLink",
  "workbookLink",
  "external",
  "hyperlink",
  "partialInput",
];

type PersistedFormatSettings = Partial<{
  numberFormats: NumberFormatDefinition[];
  dateFormats: DateFormatDefinition[];
  cellFormats: CellFormatDefinition[];
  textStyles: TextStyleDefinition[];
  autoColorPalette: Partial<AutoColorPalette>;
  trace: Partial<TraceSettings>;
}>;

export type TraceSettings = {
  maxDepth: number;
  safetyLimit: number;
};

export type ResolvedFormatSettings = {
  numberFormats: NumberFormatDefinition[];
  dateFormats: DateFormatDefinition[];
  cellFormats: CellFormatDefinition[];
  textStyles: TextStyleDefinition[];
  autoColorPalette: AutoColorPalette;
  trace: TraceSettings;
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

function clonePalette(palette: AutoColorPalette): AutoColorPalette {
  return { ...palette };
}

function cloneTraceSettings(settings: TraceSettings): TraceSettings {
  return { ...settings };
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

function resolveAutoColorPalette(value: unknown): AutoColorPalette {
  const palette = clonePalette(DEFAULT_AUTO_COLOR_PALETTE);
  if (!isRecord(value)) {
    return palette;
  }

  for (const key of AUTO_COLOR_KEYS) {
    const candidate = value[key];
    if (isNonEmptyString(candidate)) {
      palette[key] = candidate.trim();
    }
  }

  return palette;
}

function resolveTraceSettings(value: unknown): TraceSettings {
  if (!isRecord(value)) {
    return {
      maxDepth: DEFAULT_TRACE_MAX_DEPTH,
      safetyLimit: DEFAULT_TRACE_SAFETY_LIMIT,
    };
  }

  return {
    maxDepth: sanitizeTraceDepth(value.maxDepth),
    safetyLimit: sanitizeTraceSafetyLimit(value.safetyLimit),
  };
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
    (value.fillPattern !== undefined && value.fillPattern !== "Solid" && value.fillPattern !== "None") ||
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

function normalizeTextStyleDefinition(value: unknown): TextStyleDefinition | null {
  if (!isTextStyleDefinition(value)) {
    return null;
  }

  return {
    ...value,
    fillPattern: value.fillPattern === "None" ? "None" : "Solid",
  };
}

function resolveTextStyles(value: unknown): TextStyleDefinition[] {
  if (!Array.isArray(value)) {
    return cloneList(DEFAULT_TEXT_STYLES);
  }

  const valid = value
    .map((item) => normalizeTextStyleDefinition(item))
    .filter((item): item is TextStyleDefinition => item !== null);

  if (valid.length === 0) {
    return cloneList(DEFAULT_TEXT_STYLES);
  }

  return cloneList(valid);
}

export function resolveFormatSettings(raw: unknown): ResolvedFormatSettings {
  const parsed = asPersistedFormatSettings(raw);
  if (!parsed) {
    return buildDefaultFormatSettings();
  }

  return {
    numberFormats: resolveList(parsed.numberFormats, DEFAULT_NUMBER_FORMATS, isNumberFormatDefinition),
    dateFormats: resolveList(parsed.dateFormats, DEFAULT_DATE_FORMATS, isDateFormatDefinition),
    cellFormats: resolveList(parsed.cellFormats, DEFAULT_CELL_FORMATS, isCellFormatDefinition),
    textStyles: resolveTextStyles(parsed.textStyles),
    autoColorPalette: resolveAutoColorPalette(parsed.autoColorPalette),
    trace: resolveTraceSettings(parsed.trace),
  };
}

export function buildDefaultFormatSettings(): ResolvedFormatSettings {
  return {
    numberFormats: cloneList(DEFAULT_NUMBER_FORMATS),
    dateFormats: cloneList(DEFAULT_DATE_FORMATS),
    cellFormats: cloneList(DEFAULT_CELL_FORMATS),
    textStyles: cloneList(DEFAULT_TEXT_STYLES),
    autoColorPalette: clonePalette(DEFAULT_AUTO_COLOR_PALETTE),
    trace: {
      maxDepth: DEFAULT_TRACE_MAX_DEPTH,
      safetyLimit: DEFAULT_TRACE_SAFETY_LIMIT,
    },
  };
}

export function cloneResolvedFormatSettings(settings: ResolvedFormatSettings): ResolvedFormatSettings {
  return {
    numberFormats: cloneList(settings.numberFormats),
    dateFormats: cloneList(settings.dateFormats),
    cellFormats: cloneList(settings.cellFormats),
    textStyles: cloneList(settings.textStyles),
    autoColorPalette: clonePalette(settings.autoColorPalette),
    trace: cloneTraceSettings(settings.trace),
  };
}

export function getFormatSettingsValidationError(settings: ResolvedFormatSettings): string | null {
  if (settings.numberFormats.length === 0) {
    return "Add at least one number format before saving.";
  }
  if (!settings.numberFormats.every((item) => isNumberFormatDefinition(item))) {
    return "Every number format needs a name and an Excel format code.";
  }

  if (settings.dateFormats.length === 0) {
    return "Add at least one date format before saving.";
  }
  if (!settings.dateFormats.every((item) => isDateFormatDefinition(item))) {
    return "Every date format needs a name and an Excel format code.";
  }

  if (settings.cellFormats.length === 0) {
    return "Add at least one cell format before saving.";
  }
  if (!settings.cellFormats.every((item) => isCellFormatDefinition(item))) {
    return "Every cell format needs a name plus valid fill, font, and border settings.";
  }

  if (settings.textStyles.length === 0) {
    return "Add at least one text style before saving.";
  }
  if (!settings.textStyles.every((item) => isTextStyleDefinition(item))) {
    return "Every text style needs a name, font settings, and valid border choices.";
  }

  for (const key of AUTO_COLOR_KEYS) {
    if (!isNonEmptyString(settings.autoColorPalette[key])) {
      return "Every auto-color category needs a color.";
    }
  }

  if (
    !Number.isFinite(settings.trace.maxDepth) ||
    settings.trace.maxDepth < 1 ||
    !Number.isFinite(settings.trace.safetyLimit) ||
    settings.trace.safetyLimit < 1
  ) {
    return "Trace settings need positive numeric values.";
  }

  return null;
}
