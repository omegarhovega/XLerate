# XLerate

XLerate is an Excel add-in for financial modeling workflows. This repository contains:

- `src/` — the legacy VBA implementation kept for reference/parity work
- `XLerate/` — the current Office.js + TypeScript add-in

All active product development happens in `XLerate/`.

## Current product shape

- Worksheet actions are **ribbon-first**: the XLerate ribbon tab is the primary trigger surface and actions can also be added to Excel's Quick Access Toolbar.
- The task pane is reserved for **settings** and the **CAGR calculator** so worksheet actions do not have duplicate trigger paths.
- The trace dialog uses Desktop `Workbook.focus()` when available so keyboard events return to the workbook after the dialog closes.

See [open-issues.md](./open-issues.md) for the remaining known gaps.

## Requirements

Recommended / intended environment:

- Excel Desktop on Windows
- Modern Excel Desktop builds with shared runtime support
- Best trace focus behavior on Desktop hosts that support `ExcelApiDesktop 1.1` (`Workbook.focus()`)

Notes:

- The add-in is hosted as a web app and loaded by manifest URL; it is **not** a standalone `.xlam` macro add-in.
- Excel on the web and mobile are not the primary target for v1.
- Some Desktop-only behavior, especially trace focus return, depends on APIs that are not available on every host.

## Distribution

Production hosting is designed around a static HTTPS site. The default production build target is:

- `https://omegarhovega.github.io/XLerate/`

That means a practical distribution model is:

1. Host the built add-in web assets on GitHub Pages.
2. Distribute the Office add-in manifest that points to that hosted site.
3. Users install the manifest by sideloading it or through a trusted add-in catalog/admin deployment path.

The web assets must be reachable over HTTPS for normal Office add-in use.

## Development

From the repo root:

1. `cd XLerate`
2. `npm install`
3. `npm run dev-server`
4. In another terminal: `npm start`
5. Open Excel Desktop and sideload the add-in

## Validation

The repo’s main automated gate is:

- `npm run ci:all`

It currently runs:

- `typecheck:core`
- `typecheck:harness`
- `lint:harness`
- `arch:check`
- `test:core`
- manifest `validate`
- production `build`

For anything touching `src/adapters/excelPortLive.ts` or taskpane handlers, also run the manual checklist in [sideload-checklist.md](./sideload-checklist.md).

## Contributing

- Read the repo guidance in [CLAUDE.md](./CLAUDE.md)
- Treat the functional spec in `specs/2026-04-17-xlerate-functional-spec.md` as the behavior source of truth
- Add or update tests before changing behavior where practical
- Keep ribbon/taskpane behavior aligned so we do not reintroduce duplicated entry points
