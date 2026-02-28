# Golden Baseline: Trace Precedents / Dependents

Source logic: `src/modules/TraceUtils.bas`, `src/modules/RibbonCallbacks.bas`

## Contract

1. Trace starts from the active cell (root at level 0).
2. User can run either precedents or dependents mode.
3. Traversal is recursive up to configurable max depth.
4. Circular references / repeated cells are skipped via visited-cell tracking.
5. Results show level, address, value, and formula for each discovered cell.
6. If no direct references exist, output includes only the root cell.
7. Selecting a result row jumps to and selects the referenced worksheet/range.
