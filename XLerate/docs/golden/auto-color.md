# Golden Baseline: Auto-color Numbers

Source logic: `src/modules/AutoColorModule.bas`

## Contract

1. Classify each non-empty selected cell and set font color by category.
2. For formula cells, apply precedence in this order:
   - `partialInput`
   - `workbookLink`
   - `worksheetLink`
   - `external`
   - `input`
   - `formula`
3. For non-formula cells:
   - `hyperlink` first
   - otherwise `input` for numeric/non-date values
4. Blank cells are not changed.

## Default Palette

1. `input`: `#0000FF`
2. `formula`: `#000000`
3. `worksheetLink`: `#008000`
4. `workbookLink`: `#CC99FF`
5. `external`: `#00B0F0`
6. `hyperlink`: `#FF8000`
7. `partialInput`: `#800000`
