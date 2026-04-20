# XLerate Open Issues

This file tracks the highest-signal known gaps that are still open on top of
the current branch. It should describe repo reality, not ideal future state.

## Build and integration

- `manifest.xml` still under-declares the true host/runtime floor. Shared
  runtime behavior in practice needs newer Desktop builds than the manifest
  currently communicates to tooling.

## Product behavior

- The spec-defined keyboard shortcuts are not fully represented in the
  manifest/runtime yet.

## Platform limitations

- Trace-dialog focus return now uses `Workbook.focus()` on Desktop
  (`ExcelApiDesktop 1.1`). Equivalent focus restoration is still unresolved on
  non-Desktop hosts where that API is unavailable.

## Recently fixed

- Stale trace compute can no longer stream into a newly-opened dialog instance.

- Live Excel now supports array-formula writes for the current mutation paths.

- Live hyperlink detection now reads the actual `RangeHyperlink` shape rather
  than treating it as a matrix.

- Live selection reads now support multi-area selections.

- Auto-color link bucket classification now matches the functional spec.

- Error Wrap now preserves braces correctly for array formulas.

- Trace defaults now match the functional spec baseline (depth `10`, safety
  limit `500`).

- Text-style cycle state is back to session-only module state instead of
  `localStorage` persistence.

- The task pane no longer duplicates ribbon-first worksheet action buttons.

- The task pane settings surface is now a form-based editor with reorderable
  lists for number/date/cell/text presets plus the auto-color palette and
  trace defaults.

- CAGR is now exposed as an in-sheet ribbon action that inserts a worksheet
  formula into the selected destination cell from the contiguous numeric series
  immediately to its left.

- Production builds now default to the GitHub Pages host instead of the
  Contoso placeholder URL.

- `ci:all` now includes full lint, manifest validation, and the production
  build artifacts that matter for release readiness, while keeping lint scoped
  to the harness-checked paths by design.
