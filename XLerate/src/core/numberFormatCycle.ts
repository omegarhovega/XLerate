export type NumberFormatDefinition = {
  name: string;
  formatCode: string;
};

export const DEFAULT_NUMBER_FORMATS: NumberFormatDefinition[] = [
  {
    name: "Comma 0 Dec Lg Align",
    formatCode: "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"
  },
  {
    name: "Comma 1 Dec Lg Align",
    formatCode: "_(* #,##0.0_);_(* (#,##0.0);_(* \"-\"_);_(@_)"
  },
  {
    name: "Comma 2 Dec Lg Align",
    formatCode: "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"_);_(@_)"
  }
];

export function hasMixedNumberFormats(formats: string[]): boolean {
  if (formats.length <= 1) {
    return false;
  }

  const first = formats[0];
  return formats.some((value) => value !== first);
}

export function computeNextNumberFormat(
  currentFormat: string,
  hasMixedSelection: boolean,
  configuredFormats: NumberFormatDefinition[] = DEFAULT_NUMBER_FORMATS
): string {
  if (configuredFormats.length === 0) {
    throw new Error("configuredFormats must contain at least one format");
  }

  if (hasMixedSelection) {
    return configuredFormats[0].formatCode;
  }

  const index = configuredFormats.findIndex((item) => item.formatCode === currentFormat);
  if (index < 0) {
    return configuredFormats[0].formatCode;
  }

  const nextIndex = index < configuredFormats.length - 1 ? index + 1 : 0;
  return configuredFormats[nextIndex].formatCode;
}
