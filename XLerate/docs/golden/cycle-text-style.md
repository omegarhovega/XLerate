# Golden Baseline: Cycle Text Style

Source logic: `src/modules/ModTextStyle.bas`

## Contract

1. Keep an index of the current text style.
2. Each action increments index and wraps (`(index + 1) mod count`).
3. Apply selected style to current selection:
   - font name, size, bold, italic, underline, font color
   - fill color
   - clear all borders first
   - apply selected edge borders only
4. Border weight is derived from border style:
   - `Continuous` -> `Medium`
   - `Double` -> `Thick`
   - `Dash` or `Dot` -> `Thin`
   - default -> `Thin`

## Default Styles

1. Heading
2. Subheading
3. Sum
4. Normal (Excel default: Calibri 11, no emphasis, no borders, white fill)
