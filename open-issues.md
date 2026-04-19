# XLerate Open Issues

This file tracks the highest-signal known gaps that are still open on top of
the current branch. It should describe repo reality, not ideal future state.

## Build and integration

- Production builds still rewrite manifest URLs to `https://www.contoso.com/`.
  `webpack.config.js` needs a real production host or environment-driven
  configuration before `dist/manifest.xml` is releaseable.

- `npm run ci:all` is not yet the full integration gate described in older
  docs and review notes. It currently runs `typecheck:core`,
  `typecheck:harness`, `lint:harness`, `arch:check`, `test:core`, and
  `build:dev`; it does not include full lint, manifest validation, or a
  production build.

- `manifest.xml` still under-declares the true host/runtime floor. Shared
  runtime behavior in practice needs newer Desktop builds than the manifest
  currently communicates to tooling.

## Product behavior

- Text-style cycle state is still persisted via `localStorage`, so it survives
  workbook close/reopen on the same origin instead of resetting to session-only
  state.

- The spec-defined keyboard shortcuts are not fully represented in the
  manifest/runtime yet.

- The taskpane trace surface is still the older flat/eager rendering path,
  while the dialog uses the newer tree/progressive interaction model.

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
