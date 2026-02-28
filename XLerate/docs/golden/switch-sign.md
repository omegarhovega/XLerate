# Golden Baseline: Switch Sign

Source logic: `src/modules/ModSwitchSign.bas`

## Contract

1. For numeric values, negate value in-place.
2. For formulas, wrap expression with `=-(` and `)`.
3. For array formulas, remove outer braces before wrapping, then restore as array formula.
4. Blank/non-numeric constants remain unchanged.

## Baseline Cases

1. Value `10` -> `-10`
2. Value `-42` -> `42`
3. Value `0` -> `0`
4. Formula `=A1+B1` -> `=-(A1+B1)`
5. Formula `=-A1` -> `=-(-A1)`
6. Formula `=SUM(A1:A3)` -> `=-(SUM(A1:A3))`
7. Array formula `{=A1:A3*2}` -> `{=-(A1:A3*2)}`
8. Text `"hello"` unchanged
9. Blank unchanged
