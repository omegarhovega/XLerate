# Golden Baseline: CAGR

Source logic: `src/modules/ModCAGR.bas`

## Contract

1. Input is an ordered range of numeric values.
2. `firstValue = values[0]`, `lastValue = values[last]`.
3. `periodCount = values.length - 1`.
4. Result: `(lastValue / firstValue)^(1 / periodCount) - 1`.
5. Return `#VALUE!` behavior when:
   - first value <= 0
   - last value <= 0
   - periodCount <= 0
   - any conversion/runtime error

## Baseline Cases

1. `[100, 121]` -> `0.21`
2. `[100, 110, 121]` -> `0.1`
3. `[0, 121]` -> `#VALUE!`
4. `[100]` -> `#VALUE!`
5. `[-100, 121]` -> `#VALUE!`
