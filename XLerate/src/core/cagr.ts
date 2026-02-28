export const VALUE_ERROR = "#VALUE!" as const;

export type CagrResult = number | typeof VALUE_ERROR;

export function calculateCagr(values: number[]): CagrResult {
  try {
    const firstValue = values[0];
    const lastValue = values[values.length - 1];
    const periodCount = values.length - 1;

    if (firstValue <= 0 || lastValue <= 0) {
      return VALUE_ERROR;
    }

    if (periodCount <= 0) {
      return VALUE_ERROR;
    }

    return (lastValue / firstValue) ** (1 / periodCount) - 1;
  } catch {
    return VALUE_ERROR;
  }
}
