export type CagrSeriesCell = {
  row: number;
  col: number;
  value: unknown;
};

export type CagrSeries = {
  start: { row: number; col: number };
  end: { row: number; col: number };
  values: number[];
  periodCount: number;
};

export const CAGR_RESULT_NUMBER_FORMAT = "0.0%";

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

export function findContiguousLeftCagrSeries(cells: CagrSeriesCell[]): CagrSeries | null {
  const contiguous: CagrSeriesCell[] = [];

  for (let index = cells.length - 1; index >= 0; index -= 1) {
    const cell = cells[index];
    if (!isFiniteNumber(cell.value)) {
      break;
    }
    contiguous.push(cell);
  }

  if (contiguous.length < 2) {
    return null;
  }

  contiguous.reverse();
  const values = contiguous.map((cell) => cell.value as number);
  const start = contiguous[0];
  const end = contiguous[contiguous.length - 1];

  return {
    start: { row: start.row, col: start.col },
    end: { row: end.row, col: end.col },
    values,
    periodCount: values.length - 1,
  };
}

export function toA1Address(row: number, col: number): string {
  let remainder = col + 1;
  let letters = "";

  while (remainder > 0) {
    const zeroBased = (remainder - 1) % 26;
    letters = String.fromCharCode(65 + zeroBased) + letters;
    remainder = Math.floor((remainder - 1) / 26);
  }

  return `${letters}${row + 1}`;
}

export function buildCagrInsertFormula(series: CagrSeries): string {
  const startAddress = toA1Address(series.start.row, series.start.col);
  const endAddress = toA1Address(series.end.row, series.end.col);
  return `=POWER(${endAddress}/${startAddress},1/${series.periodCount})-1`;
}
