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

- Excel on Windows, Mac, or the web
- Modern Excel builds that satisfy the add-in manifest requirements
- Best overall behavior on builds with shared runtime, custom keyboard shortcut, and `Workbook.focus()` support

Notes:

- The add-in is hosted as a web app and loaded by manifest URL; it is **not** a standalone `.xlam` macro add-in.
- The shipped manifest is intentionally scoped to modern shared-runtime Excel hosts.
- Excel Desktop on Windows and Mac are the primary target for v1.
- Excel on the web is supported, but trace-dialog keyboard focus return is a desktop-only enhancement.
- Excel mobile is out of scope for v1.

## Install

The release manifest is:

- [`XLerate/manifest.xml`](./XLerate/manifest.xml)

After GitHub Pages is live, the same manifest is also hosted at:

- `https://omegarhovega.github.io/XLerate/manifest.xml`

Users install XLerate by downloading the manifest and sideloading it in a supported Excel host.

## Distribution

Production hosting is designed around a static HTTPS site. The default production build target is:

- `https://omegarhovega.github.io/XLerate/`

That means the release flow is:

1. Push to `master` or `main`.
2. GitHub Actions deploys `XLerate/dist/` to GitHub Pages.
3. Users install the release manifest that points at that hosted site.

The web assets must be reachable over HTTPS for normal Office add-in use.

## Development

From the repo root:

1. `cd XLerate`
2. `npm install`
3. `npm run dev-server`
4. In another terminal: `npm start`
5. Open Excel Desktop and sideload the add-in

`npm start` uses `manifest.dev.xml` for local sideloading. `manifest.xml` is the production manifest.

## Validation

The repo’s main automated gate is:

- `npm run ci:all`

It currently runs:

- `typecheck:core`
- `typecheck:harness`
- `lint:harness`
- `arch:check`
- `test:core`
- production and local manifest `validate`
- production `build`

For anything touching `src/adapters/excelPortLive.ts` or taskpane handlers, also run the manual checklist in [sideload-checklist.md](./sideload-checklist.md).

## Contributing

- Read the repo guidance in [CLAUDE.md](./CLAUDE.md)
- Treat the functional spec in `specs/2026-04-17-xlerate-functional-spec.md` as the behavior source of truth
- Add or update tests before changing behavior where practical
- Keep ribbon/taskpane behavior aligned so we do not reintroduce duplicated entry points
