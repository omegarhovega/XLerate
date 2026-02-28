# Golden Baseline: Formula Consistency (Horizontal)

Source logic: `src/modules/FormulaConsistency.bas`

## Contract

1. Analyze only cells that contain formulas.
2. Horizontal check compares each formula cell to its immediate right neighbor.
3. If right neighbor is a formula:
   - mark cell as `consistent` when `FormulaR1C1` matches
   - mark cell as `inconsistent` when it differs
4. If right neighbor is not a formula:
   - mark cell as `consistent` only when this formula appears in any adjacent-equal pair in the used range.
   - otherwise no mark.
5. Last-column formula cells are marked `consistent` only when they belong to a formula that appears in any adjacent-equal pair.
6. Non-formula cells are never marked.

## Notes

The VBA algorithm can mark an isolated formula cell as `consistent` when the same formula has an adjacent match elsewhere in the range.
