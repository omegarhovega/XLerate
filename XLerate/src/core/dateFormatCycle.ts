export type DateFormatDefinition = {
  name: string;
  formatCode: string;
};

export const DEFAULT_DATE_FORMATS: DateFormatDefinition[] = [
  {
    name: "Year Only",
    formatCode: "yyyy"
  },
  {
    name: "Month Year",
    formatCode: "mmm-yyyy"
  },
  {
    name: "Full Date",
    formatCode: "dd-mmm-yy"
  }
];

export function hasMixedDateFormats(formats: string[]): boolean {
  if (formats.length <= 1) {
    return false;
  }

  const first = formats[0];
  return formats.some((value) => value !== first);
}

export function computeNextDateFormat(
  currentFormat: string,
  hasMixedSelection: boolean,
  configuredFormats: DateFormatDefinition[] = DEFAULT_DATE_FORMATS
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
