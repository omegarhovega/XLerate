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
8. Supported A1-style references inside displayed formulas are clickable; string
   literals, names, structured references, 3D references, and external links remain text.
9. If no direct references exist, output includes only the root cell.
10. Selecting a result row jumps to and selects the referenced worksheet/range.
