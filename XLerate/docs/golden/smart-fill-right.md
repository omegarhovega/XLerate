# Golden Baseline: Smart Fill Right

Source logic: `src/modules/ModSmartFillRight.bas`

## Contract

1. Active cell must contain a formula starting with `=`.
2. Active cell cannot be merged.
3. Boundary search checks up to 3 rows above active row, nearest row first.
4. In candidate row, scan from active column rightward through the first contiguous non-empty block.
5. If any merged cell exists in that contiguous block, row is rejected.
6. Boundary is the last non-empty cell in the contiguous block.
7. First valid boundary found is used.
8. If no boundary found, operation stops.

## Baseline Cases

1. Active formula at row 5 col 2. Row 4 has non-empty cells at cols 2-6, no merges -> boundary 6.
2. Row 4 has merge in cols 2-6 block, row 3 valid cols 2-4 -> boundary 4 (fallback to next row up).
3. Rows 4/3/2 invalid or empty at start col -> no boundary.
4. Active cell with non-formula value -> rejected.
5. Active cell merged -> rejected.
