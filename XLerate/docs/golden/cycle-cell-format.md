# Golden Baseline: Cycle Cell Format

Source logic: `src/modules/ModCellFormat.bas`

## Contract

1. Cycle through configured cell-format presets.
2. Determine current preset by matching selection formatting:
   - fill pattern and fill color
   - font color + font style flags
   - edge border style/color (left, top, bottom, right)
3. If a preset matches, apply the next preset (wrap at end).
4. If no preset matches, apply the first preset.
5. Applying a preset sets:
   - fill pattern + fill color
   - font color + bold/italic/underline/strikethrough
   - edge borders and, for multi-cell selections, inside borders.

## Default Presets

1. Normal
2. Inputs
3. Good
4. Bad
5. Important
