# Golden Baseline: Cycle Number Format

Source logic: `src/modules/ModNumberFormat.bas`

## Contract

1. Cycle through a configured list of number formats.
2. Use the first selected cell's `NumberFormat` as the current format.
3. If selected cells contain mixed formats, use the first configured format.
4. If current format is found in the list, apply the next one (wrap at end).
5. If current format is not found, apply the first configured format.

## Baseline Cases

1. Current format is first item -> apply second item.
2. Current format is last item -> wrap to first item.
3. Mixed selection formats -> apply first item.
4. Unknown current format -> apply first item.
