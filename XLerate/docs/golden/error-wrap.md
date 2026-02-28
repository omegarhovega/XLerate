# Golden Baseline: Error Wrap

Source logic: `src/modules/ModErrorWrap.bas`

## Contract

1. For each formula cell in selection, wrap with `=IFERROR(<existing_formula_without_equal>, <errorValue>)`.
2. Non-formula cells are skipped.
3. Default `errorValue` is `NA()` when missing.
4. Existing wrappers are not deduplicated; wrapping can nest.

## Baseline Cases

1. Formula `=A1/B1`, error `NA()` -> `=IFERROR(A1/B1, NA())`
2. Formula `=SUM(A1:A3)`, error `0` -> `=IFERROR(SUM(A1:A3), 0)`
3. Formula `=IFERROR(A1/B1,0)`, error `NA()` -> `=IFERROR(IFERROR(A1/B1,0), NA())`
4. Constant `100` unchanged
5. Text `"x"` unchanged
6. Blank unchanged
