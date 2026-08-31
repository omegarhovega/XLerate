# Golden Baseline: Trace Precedents / Dependents

Source logic: `src/modules/TraceUtils.bas`, `src/modules/RibbonCallbacks.bas`

## Contract

1. Trace starts from the active cell (root at level 0).
2. User can run either precedents or dependents mode.
3. Initial load contains only the root and its direct relationships.
4. Deeper relationships load when the user expands a node and are cached for the dialog session.
5. Circular references / repeated cells appear as non-expandable reference nodes.
6. Rectangular areas above 50 cells appear as summarized ranges and load in pages of 100.
7. Results show level, address, value, and formula for each discovered cell.
8. Supported A1-style references, workbook/worksheet named ranges, and fully
   qualified Excel Table references inside displayed formulas are clickable.
   Unqualified names follow Excel resolution rules: a local name on the formula
   sheet takes precedence, otherwise the workbook-scoped name is used. A
   `Sheet!Name` token resolves only through that worksheet's named items.
   String literals remain text; external workbook cells/ranges are identified
   but non-clickable. Current-row table, 3D, and external-name references remain text.
9. A compact pinned context panel shows the relevant complete formula and its
   root-to-node breadcrumb while the trace tree remains the navigation control.
   Supported A1 cells/ranges are highlighted in place without additional Excel
   reads during navigation.
10. Selecting a result row jumps to and selects the referenced worksheet/range.
11. Enter closes and keeps the current selection; Escape restores the original
    trace-root cell before closing.
12. If no direct references exist, output includes only the root cell.
