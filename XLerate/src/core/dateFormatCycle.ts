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

function normalizeDateFormatCode(formatCode: string): string {
  const compact = formatCode
    .trim()
    .toLowerCase()
    .replace(/\[\$-[^\]]+\]/g, "")
    .replace(/\\/g, "")
    .replace(/\s+/g, " ");

  let result = "";
  for (let i = 0; i < compact.length; ) {
    if (compact[i] !== "d") {
      result += compact[i];
      i += 1;
      continue;
    }

    let j = i;
    while (j < compact.length && compact[j] === "d") {
      j += 1;
    }

    const count = j - i;
    result += count === 2 ? "d" : "d".repeat(count);
    i = j;
  }

  return result;
}

function matchesDateFormat(actual: string, expected: string): boolean {
  return normalizeDateFormatCode(actual) === normalizeDateFormatCode(expected);
}

export function hasMixedDateFormats(formats: string[]): boolean {
  if (formats.length <= 1) {
    return false;
  }

  const first = formats[0];
  return formats.some((value) => !matchesDateFormat(value, first));
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

  const index = configuredFormats.findIndex((item) =>
    matchesDateFormat(currentFormat, item.formatCode)
  );
  if (index < 0) {
    return configuredFormats[0].formatCode;
  }

  const nextIndex = index < configuredFormats.length - 1 ? index + 1 : 0;
  return configuredFormats[nextIndex].formatCode;
}
