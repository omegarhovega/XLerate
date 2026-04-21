# XLerate


![XLerate](https://github.com/omegarhovega/XLerate/blob/master/XLerate/docs/images/XLerate.png)

XLerate is an Excel add-in for financial modeling work. It adds a dedicated **XLerate** ribbon tab with tools for tracing formulas, checking horizontal consistency, cycling through modeling formats, speeding up fill actions, and managing workbook-specific XLerate settings.

If you used the older XLerate VBA-based add-in, this repository now ships the **new Office.js-based version** . The older VBA implementation is still present in the repo for reference but should be treated as legacy.

## What XLerate Does

XLerate is built for repetitive spreadsheet tasks that come up constantly in financial models, operating models, review files, and audit passes.

Core features include:

- **Trace Precedents / Trace Dependents**: open a keyboard-friendly dialog to inspect formula relationships across sheets and multiple levels deep.
- **Switch Sign**: flip the sign of numeric cells or formulas across a selection.
- **Smart Fill Right**: extend a formula rightward to the detected boundary.
- **Horizontal Check**: mark formulas as consistent or inconsistent relative to neighbors.
- **Cycle Number / Date / Cell / Text formats**: move through preset formatting styles quickly.
- **Auto-color**: apply semantic colors to values and formulas.
- **Insert CAGR**: insert a CAGR formula based on the numeric series immediately to the left.
- **Settings workspace**: manage XLerate presets and workbook-specific defaults in the side pane.

Users can edit and save their own presets for:

- number formats
- date formats
- cell formats
- text styles
- auto-color behavior
- trace defaults

## How the Main Functions Work

### Trace Precedents / Trace Dependents

- Starts from the active cell.
- Opens a keyboard-friendly dialog for navigating relationships.
- Can follow formula links across worksheets.
- Can inspect multiple dependency levels, not just direct neighbors.
- Best used on formula cells where you want to understand what feeds a number or what depends on it.

### Switch Sign

- Works on the current selection.
- Numeric constants are negated.
- Formulas are wrapped so the resulting value flips sign.
- Text cells are left unchanged.

### Smart Fill Right

- Works from the **active cell**.
- The active cell must contain a formula.
- XLerate looks up to **three rows above** the active row to find a usable reference row.
- It fills rightward only when it can detect a valid boundary from that data above.
- If there is no suitable row above, or if the active cell is not a formula, the action will not fill.

In practice, Smart Fill Right works best when:

- the row above (or one of the next two rows above) contains a clear contiguous block of cells that defines the intended fill width
- the active cell is already the leftmost formula you want to extend
- the area above is not broken by merges in the boundary region

### Horizontal Check

- Works on formulas in the selected range.
- Compares formulas horizontally to identify copy-across inconsistencies.
- Marks cells as visually consistent or inconsistent relative to their neighbors.
- Useful for finding broken formulas in model rows.

### Cycle Number / Date / Cell / Text Formats

- Works on the current selection.
- Applies the next preset in the configured list.
- If the current cell formatting does not match a known preset, XLerate applies the first preset in that category.
- Presets can be customized in the side pane settings workspace.

### Auto-color

- Works on the current selection.
- Applies semantic colors based on what a cell represents.
- Distinguishes between inputs, formulas, links, hyperlinks, and related categories.
- Best used on populated ranges rather than blank areas.

### Insert CAGR

- Works from the **active cell**, which becomes the destination for the CAGR formula.
- XLerate looks immediately to the **left** and scans the same row for a contiguous numeric series.
- It uses that series to build the CAGR formula automatically.
- At least **two numeric cells** are required for a meaningful CAGR.
- If there is no valid contiguous numeric series directly to the left, nothing is inserted.

In practice, Insert CAGR works best when:

- the destination cell is placed immediately to the right of a row of historical values
- the values to the left are contiguous and numeric
- text cells or blanks do not interrupt the series

### Settings Side Pane

- The side pane is for configuring XLerate behavior rather than duplicating the worksheet actions.
- You can show it when you want to edit presets or defaults, then hide it again while working in the sheet.
- It lets you manage workbook-specific settings for format cycles, auto-color presets, and trace defaults.

## Requirements

Recommended environment:

- Excel Desktop on **Windows** or **Mac**
- A modern Microsoft 365 Excel build

Notes:

- XLerate is now an **Office add-in**, not an `.xlam` macro add-in.
- Excel on the web may load, but Excel Desktop is the primary target.
- Excel mobile is not supported.

## Install

### Get the Manifest

Download the manifest from one of these locations:

- [manifest.xml](https://omegarhovega.github.io/XLerate/manifest.xml)
- [XLerate/manifest.xml](./XLerate/manifest.xml)

## How to Install in Excel

For Excel Desktop, the reliable installation path is via a **shared-folder add-in catalog**.

### Shared-folder install flow

1. Download `manifest.xml`.
2. Put the manifest into a folder and share with yourself, or into a shared/network folder that will act as the add-in catalog. For this see [Create a network shared folder catalog for task pane and content add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
3. In Excel, open **Trust Center** settings and add that shared folder as a trusted add-in catalog.
4. Restart Excel if required.
5. Go to **Home** → **Add-ins**.
6. Open **+ More Add-ins**.
7. Open the **Shared Folder** section.
8. Select XLerate from the shared-folder catalog.
9. The **XLerate** tab will open in the ribbon.

### Important note

The manifest is not installed by simply placing it in an arbitrary folder on your machine. Excel must be told to trust and use that folder as an add-in catalog through Trust Center settings.

If Excel asks which action should own a shortcut such as `Ctrl+Shift+1`, choose **XLerate** for the XLerate actions you want.

If you need to reset those choices later, use Excel’s search box and run:

- `Reset Office Add-ins shortcut preferences`

## First Use

After installation:

1. Open a workbook.
2. Click the **XLerate** ribbon tab.
3. Try **Cycle Number**, **Smart Fill Right**, or **Trace Precedents**.
4. Open **Show Task Pane** to access workbook settings.

Most worksheet actions are ribbon-first. The side pane is mainly for settings and supporting UI; it can be shown when needed and hidden again when you want to work with more sheet space.

## Keyboard Shortcuts

XLerate supports these shortcuts:

- `Ctrl+Shift+1`: Cycle Number Formats
- `Ctrl+Shift+2`: Cycle Cell Formats
- `Ctrl+Shift+3`: Cycle Date Formats
- `Ctrl+Shift+4`: Cycle Text Styles
- `Ctrl+Shift+R`: Smart Fill Right

On first use, Excel may show a conflict dialog if one of these shortcuts is already associated with a built-in Excel feature. Choose the XLerate action if you want XLerate to take precedence on that machine.

For frequent use, it is also a good idea to add your most-used XLerate commands to Excel’s **Quick Access Toolbar**. That gives you fast keyboard access through Excel’s own QAT workflow in addition to the default XLerate shortcuts.

## Usage Notes

- Formatting and worksheet actions are designed to work on the current selection.
- Trace tools operate on the active cell and can follow relationships across worksheets through multiple dependency levels.
- Many actions are intended to be a single Excel undo step.
- Workbook settings are stored as XLerate settings for that workbook, so different workbooks can use different presets.

## Troubleshooting

### The XLerate tab does not appear

- Make sure the add-in manifest was successfully installed.
- Close and reopen Excel.
- Remove and reinstall the manifest if needed.

### A shortcut runs the wrong Excel command

- Choose the XLerate action in the conflict dialog when Excel prompts you.
- If Excel already saved the wrong preference, run:
  - `Reset Office Add-ins shortcut preferences`

### The add-in does not load after an update

- Close Excel fully and reopen it.
- Remove and reinstall the add-in manifest if Excel is still using a stale registration.

## Repository Layout

This repository contains:

- `src/` — the legacy VBA implementation kept for reference/parity work
- `XLerate/` — the current Office.js + TypeScript add-in

All active product development happens in `XLerate/`.

## Development

From the repo root:

1. `cd XLerate`
2. `npm install`
3. `npm run dev-server`
4. In another terminal: `npm start`
5. Open Excel Desktop and sideload the development manifest

Notes:

- `manifest.dev.xml` is for local sideload development.
- `manifest.xml` is the production manifest intended for hosted distribution.

## Validation

Primary automated gate:

- `npm run ci:all`

That runs:

- `typecheck:core`
- `typecheck:harness`
- `lint:harness`
- `arch:check`
- `test:core`
- production and local manifest validation
- production build

For changes that touch live Excel runtime boundaries or manifest/ribbon wiring, also use:

- [sideload-checklist.md](./sideload-checklist.md)
