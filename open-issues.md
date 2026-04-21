# XLerate Open Issues

This file tracks the highest-signal known gaps that are still open on top of
the current branch. It should describe repo reality, not ideal future state.

## Release readiness

- No ship blockers are currently tracked in this file.

## Platform notes

- Trace-dialog focus return uses `Workbook.focus()` when the host supports
  it. Excel Desktop gets the best post-close keyboard handoff; Excel on the
  web remains usable but does not have the same host-level focus restore.

## Recently fixed

- `manifest.xml` now declares the desktop/runtime floor in the base manifest,
  so unsupported hosts are filtered out before installation.

- The checked-in production manifest now points at the GitHub Pages host, while
  `manifest.dev.xml` remains the local sideload manifest for development.

- GitHub Actions now deploy `XLerate/dist/` to GitHub Pages on pushes to
  `master` / `main`, including the hosted production manifest and shortcut
  manifest.

- The spec-defined keyboard shortcuts are now represented through
  `ExtendedOverrides` + `shortcuts.json`, including the reset-formats action.

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
