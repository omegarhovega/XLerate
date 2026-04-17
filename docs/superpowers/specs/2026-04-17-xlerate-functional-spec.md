# XLerate Functional Specification

**Status:** Target specification — defines "migration complete" for the VBA → Office JS / TypeScript port.
**Date:** 2026-04-17
**Scope:** Full product behavior. Functional only — no technical implementation details.
**Authority:** Where the current TypeScript implementation disagrees with this spec, the implementation is the one that changes.

---

## 1. Introduction

### 1.1 Product

XLerate is an Excel add-in for financial modelers. It accelerates the repetitive tasks in a modeling workflow: auditing formula references, cycling between standard number and cell formats, checking that horizontal formulas are consistent, coloring inputs vs formulas vs links, wrapping formulas with error handlers, filling formulas to a detected boundary on the right, flipping signs, and computing compound annual growth rates.

### 1.2 Target user

Financial modelers and investment bankers. Power users building or reviewing models hundreds of times a day. Keyboard-first: the mouse is for ribbon discovery; the hotkey is for the hundredth time today. They value speed, undoability, and the certainty that a workbook emailed to a colleague arrives with all customizations intact.

### 1.3 Platforms

- Excel Desktop on Windows.
- Excel Desktop on Mac.

Excel on the web and Excel Mobile are out of scope for v1.

### 1.4 Design principles

- **Keyboard parity.** Every frequent action has a hotkey. Mouse is never required.
- **Workbook trust.** Customizations travel inside the file when possible; user defaults fill gaps.
- **No surprises.** Destructive actions confirm before acting. Every mutation is a single native Excel undo step.
- **Power-user speed.** Actions complete in under a perceptible beat on realistic models. No eager full-graph or full-sheet traversals on the fast path — see §5.4.

### 1.5 Relationship to the legacy VBA implementation

The legacy VBA add-in defines the authoritative behavior for features already shipped there. This document restates that behavior in functional terms, expands it with decisions made during the migration, and adds new items where VBA parity requires it (see §5). Once a feature is specified here, this document — not the VBA source — is the source of truth.

---

## 2. Surface map

### 2.1 Ribbon

A custom ribbon tab named **XLerate** is present on the Excel ribbon whenever the add-in is loaded. The tab has four groups:

- **Formulas** — Trace Precedents, Trace Dependents, Switch Sign, Smart Fill Right.
- **Auditing** — Horizontal Consistency, Clear Consistency Marks.
- **Formatting** — Format (split button revealing: Cycle Number Formats, Cycle Cell Formats, Cycle Date Formats, Cycle Text Styles), Error Wrap, Auto-color Numbers.
- **Settings** — Settings (opens the task pane editor).

### 2.2 Keyboard shortcuts

XLerate registers six shortcuts and intentionally overrides Excel's built-in bindings where they clash. Bankers use the cycle features far more than the built-in static formats.

| Shortcut | Action |
|---|---|
| Ctrl + Shift + 1 | Cycle Number Formats |
| Ctrl + Shift + 2 | Cycle Cell Formats |
| Ctrl + Shift + 3 | Cycle Date Formats |
| Ctrl + Shift + 4 | Cycle Text Styles |
| Ctrl + Shift + R | Smart Fill Right |
| Ctrl + Shift + 0 | Reset Formats |

If a platform (e.g. Excel Desktop for Mac) cannot fully override a given Excel default, the XLerate binding must still be available; the implementation may surface a platform-specific note but must not silently drop the shortcut.

### 2.3 Task pane

The task pane houses:

- The **Settings editor** (tabbed, one tab per format category plus Auto-color and Trace).
- The **Trace results** panel (tree view, described in §3.1).
- The **CAGR calculator** (one-off calculation without writing to a cell).
- **Clear Consistency Marks** also lives here for quick access.

### 2.4 Worksheet function

`=CAGR(startValue, endValue, years)` is registered as a custom worksheet function available in any cell of any workbook while the add-in is loaded.

---

## 3. Feature specifications

Each feature is defined by: **Purpose** (one sentence), **Trigger** (how the user invokes it), **Contract** (bulleted functional rules), **Examples** (worked scenarios in plain English), and **Edge cases**.

### 3.1 Trace Precedents

**Purpose.** Show the user which cells the active cell depends on, across sheets, to any depth.

**Trigger.** Ribbon: Formulas → Trace Precedents. Task pane: Trace button in Precedents mode.

**Contract.**
- Trace starts at the active cell. The active cell is the tree root at level 0.
- The initial result shows **only the active cell plus its direct (level-1) precedents**. No deeper levels are loaded at open.
- Each result node displays: level, cell address (with workbook and sheet context), current value, and formula.
- Each node is expandable on demand. Expanding a node loads its direct precedents.
- Cycle safety: a cell that has already appeared in the tree is shown as a reference but not re-expanded.
- Selecting a result row navigates Excel to that cell (switching sheets or workbooks as needed) and selects the range.
- If the active cell has no precedents, only the root is shown.

**Examples.**
- Active cell `Sheet1!B7 = A7 + C7`. Open Trace Precedents → tree shows root `Sheet1!B7` with two child nodes `Sheet1!A7` and `Sheet1!C7`. No further nodes visible. Clicking `Sheet1!A7` expands to show its precedents, if any.
- Active cell `Summary!D5 = Inputs!Revenue + Inputs!Costs`. Open Trace Precedents → root plus two children on `Inputs`. Clicking a child jumps the active Excel selection to the referenced cell on the `Inputs` sheet.
- Active cell `=100` (a constant, no references). Open Trace Precedents → tree shows only the root cell.

**Edge cases.**
- No active cell / no worksheet active → button is disabled.
- Circular reference (A→B→A) → B appears as child of A; expanding B shows A as a non-expandable reference.
- External workbook references → if the source workbook is not open, the node shows address and formula but navigation is disabled with a tooltip "External workbook not open."
- Very wide fan-out (e.g. `=SUM(A1:A10000)`) → precedents of that cell are represented as a summarized range node, not ten thousand individual cells.

### 3.2 Trace Dependents

**Purpose.** Show the user which cells depend on the active cell, across sheets, to any depth.

**Trigger.** Ribbon: Formulas → Trace Dependents. Task pane: Trace button in Dependents mode.

**Contract.** Identical in shape to §3.1 but traversing forward: find cells whose formulas reference the active cell, then recurse.
- Initial view: root + direct dependents only.
- Deeper levels load on expansion.
- Cycle safety, navigation, and summarization behaviors are the same as Trace Precedents.

**Examples.**
- Active cell `Inputs!A1`. Open Trace Dependents → tree shows root plus each formula cell that references `Inputs!A1` directly. Expand any of them to see their own dependents.
- Active cell is unreferenced → tree shows root alone.

**Edge cases.** As §3.1, substituting "dependents" for "precedents." In particular, dependents discovery limited to open workbooks; closed workbooks cannot be scanned.

### 3.3 Switch Sign

**Purpose.** Negate the selected cell's value or formula result without retyping.

**Trigger.** Ribbon: Formulas → Switch Sign. Task pane: Switch Sign button.

**Contract.**
- For a numeric constant, replace the value with its negation.
- For a formula, wrap the existing expression as `=-(<original expression without leading =>)`.
- For an array formula, remove the outer array braces, wrap the inner expression with `=-( … )`, and re-enter as an array formula.
- Blank cells and non-numeric text constants are unchanged.
- Operates on every cell in the selection independently.

**Examples.**
- Cell contains `10` → becomes `-10`.
- Cell contains `-42` → becomes `42`.
- Cell contains `=A1+B1` → becomes `=-(A1+B1)`.
- Cell contains `=-A1` → becomes `=-(-A1)` (the operation nests rather than simplifies).
- Array formula `{=A1:A3*2}` → becomes `{=-(A1:A3*2)}`.

**Edge cases.**
- Cell contains `0` → unchanged.
- Cell contains `"hello"` → unchanged.
- Mixed selection: numeric cells flip, formula cells wrap, text cells are skipped — all in one undo step.

### 3.4 Smart Fill Right

**Purpose.** Copy the active cell's formula rightwards, stopping at a boundary detected from the surrounding data.

**Trigger.** Ribbon: Formulas → Smart Fill Right. Hotkey: Ctrl + Shift + R. Task pane: Smart Fill Right button.

**Contract.**
- The active cell must contain a formula (leading `=`). The active cell must not be merged.
- To find the fill boundary, scan up to three rows above the active row, nearest first.
- In each candidate row, scan from the active column rightwards. The boundary is the last non-empty cell of the first contiguous non-empty block.
- A candidate row is rejected if any merged cell exists within that contiguous block.
- The first accepted candidate row wins; its boundary becomes the fill end column.
- Fill the active cell's formula into the range from the active cell to the detected boundary, inclusive.
- If no valid candidate row is found within the three-row search window, the action stops silently with an explanatory message in the task pane.

**Examples.**
- Active formula at row 5, column B. Row 4 has values in B through F with no merges → fill B5 to F5.
- Row 4 has a merge across B–F; row 3 has values in B–D with no merges → fill to D5.
- Rows 4, 3, and 2 are all empty or merged at the start column → no fill, task pane message "Smart Fill Right: no reference row found in the 3 rows above."

**Edge cases.**
- Active cell has no formula → action rejected with task pane message.
- Active cell is merged → action rejected with task pane message.
- Boundary is in the same column as the active cell → no-op; formula is already there.

### 3.5 Horizontal Formula Consistency

**Purpose.** Highlight cells whose formula differs from their horizontal neighbors, so the modeler can spot broken copy-across.

**Trigger.** Ribbon: Auditing → Horizontal Consistency. Task pane: Check Formula Consistency button.

**Contract.**
- Analyze only cells that contain formulas. Non-formula cells are never marked.
- For each formula cell, compare its R1C1-equivalent formula to its immediate right neighbor:
  - If the right neighbor is a formula with the same R1C1 expression → mark **consistent**.
  - If the right neighbor is a formula with a different R1C1 expression → mark **inconsistent**.
  - If the right neighbor is not a formula → the cell is marked consistent only if the same formula appears in at least one adjacent-equal pair elsewhere in the used range; otherwise no mark.
- Cells in the last column of the used range: marked consistent only if their formula matches some adjacent-equal pair in the used range.
- Marks are shown as a cell fill color: green for **consistent**, red for **inconsistent**. No other formatting on the cell (font color, borders, value) is changed by the check.
- Marks are persisted alongside the workbook so they survive close, reopen, and transfer to another machine.
- Marks are removed only by Clear Consistency Marks (§3.6), not by re-running the check. Re-running replaces marks at the affected cells but does not clear cells that fell out of scope.

**Examples.**
- Row 10 contains `=B$9*C10`, `=C$9*D10`, `=D$9*E10` across columns B–D. All three have the same R1C1 pattern → all three marked consistent.
- A single broken cell `=C$9*Z10` in the middle of that row → it is marked inconsistent; its neighbors remain consistent relative to each other.
- Isolated formula `=SUM(A1:A5)` in the corner of a sheet; another cell elsewhere in the used range has the same formula as part of a consistent pair → the isolated cell is marked consistent.

**Edge cases.**
- Empty selection or empty sheet → no-op.
- Sheet is protected → no-op with task pane message "Sheet is protected; unprotect to apply consistency marks."
- Workbook moved between machines → marks persist because they are stored with the workbook.

### 3.6 Clear Consistency Marks

**Purpose.** Remove all formula consistency marks from the active sheet.

**Trigger.** Ribbon: Auditing → Clear Consistency Marks. Task pane: Clear Consistency Marks button.

**Contract.**
- Confirms before acting: "Clear all formula consistency marks on this sheet?" with Clear / Cancel.
- On confirmation, removes marks from every cell on the active sheet and clears the persisted mark state for that sheet.
- Does not affect other sheets in the workbook.
- Produces a single undo step.

**Examples.**
- Two sheets carry marks. User runs Clear Consistency Marks on Sheet2 → Sheet2 marks gone, Sheet1 untouched.

**Edge cases.**
- No marks present → confirmation still shown; on confirm, action is a no-op (does not produce an undo step).
- Sheet is protected → action is rejected with task pane message.

### 3.7 Cycle Number Format

**Purpose.** Cycle the selected cells through the user's ordered list of number format presets.

**Trigger.** Ribbon: Formatting → Format → Cycle Number Formats. Hotkey: Ctrl + Shift + 1. Task pane: Cycle Number Format button.

**Contract.**
- Operates on the current selection.
- Determine the "current" format from the first selected cell's number format.
- If the selection contains mixed number formats, treat the current format as unknown.
- If the current format is present in the configured list, apply the next one. Wrap to the first when reaching the end.
- If the current format is not present (or unknown due to mixed selection), apply the first format in the list.
- The applied format takes effect on every cell in the selection.

**Examples.**
- Preset list: `#,##0`, `#,##0.00`, `$#,##0`, `0.0%`. Current cell format is `#,##0` → becomes `#,##0.00`.
- Current cell format is `0.0%` (the last entry) → becomes `#,##0`.
- Selection has two cells, one is `#,##0` and one is `General` → applies `#,##0` (the first preset) to both.
- Current cell format is `[$-409]mmm-yy` (unknown) → applies `#,##0` (the first preset).

**Edge cases.**
- Empty selection → action disabled.
- Preset list is empty or invalid in storage → fall back to built-in defaults (§4.3) and apply the first default.

### 3.8 Cycle Cell Format

**Purpose.** Cycle the selected cells through presets of combined fill, font, and border styling.

**Trigger.** Ribbon: Formatting → Format → Cycle Cell Formats. Hotkey: Ctrl + Shift + 2. Task pane: Cycle Cell Format button.

**Contract.**
- Operates on the current selection.
- A preset includes: fill pattern + fill color, font color, bold / italic / underline / strikethrough flags, and edge border style + color.
- Match the "current" preset by comparing the first selected cell's formatting to each preset on all of the above axes.
- If a preset matches, apply the next preset (wrap at end).
- If no preset matches, apply the first preset.
- Applying a preset sets: fill pattern + color, font color, emphasis flags, edge borders and — for multi-cell selections — inside borders.

**Default presets.** Normal, Inputs, Good, Bad, Important. (Users may edit or reorder via Settings.)

**Examples.**
- Selection matches the **Inputs** preset → cycles to **Good**.
- Selection matches the last preset (**Important**) → cycles to **Normal**.
- Selection has no fill and default borders (no preset match) → applies **Normal**.

**Edge cases.**
- Empty selection → action disabled.
- Mixed formatting across the selection → treated as no-match; applies the first preset to the entire selection.

### 3.9 Cycle Date Format

**Purpose.** Cycle the selected cells through the user's ordered list of date format presets.

**Trigger.** Ribbon: Formatting → Format → Cycle Date Formats. Hotkey: Ctrl + Shift + 3. Task pane: Cycle Date Format button.

**Contract.** Same shape as Cycle Number Format (§3.7), using the date-format preset list.

**Examples.**
- Preset list: `dd-mmm-yy`, `dd/mm/yyyy`, `mmm yyyy`, `yyyy-mm-dd`. Current format is `dd/mm/yyyy` → becomes `mmm yyyy`.
- Mixed selection of `dd-mmm-yy` and `mmm yyyy` → applies `dd-mmm-yy` (first preset).

**Edge cases.** Empty selection disabled; empty/invalid preset list falls back to built-in defaults.

### 3.10 Cycle Text Style

**Purpose.** Cycle the selected cells through presets of text styling (font, size, emphasis, fill, borders).

**Trigger.** Ribbon: Formatting → Format → Cycle Text Styles. Hotkey: Ctrl + Shift + 4. Task pane: Cycle Text Style button.

**Contract.**
- Keep a single cycle index across the workbook session. Each invocation increments the index modulo the preset count and applies the resulting style to the current selection.
- Applying a preset sets: font name, font size, bold / italic / underline, font color, fill color; clears all borders and then applies the preset's edge borders.
- Border weight is derived from border style: Continuous → Medium, Double → Thick, Dash or Dot → Thin, otherwise Thin.

**Default presets.** Heading, Subheading, Sum, Normal. Normal restores Excel default — Calibri 11, no emphasis, no borders, white fill.

**Examples.**
- First invocation on a plain cell → applies **Heading**.
- Next invocation → applies **Subheading**, regardless of what the cell currently looks like.
- Four invocations complete one loop back to **Heading**.

**Edge cases.**
- Empty selection → action disabled.
- Closing and reopening the workbook does not preserve the cycle index; the next invocation starts from the first preset.

### 3.11 Error Wrap

**Purpose.** Wrap the selected formula cells with `IFERROR(...)` using a configurable fallback value.

**Trigger.** Ribbon: Formatting → Error Wrap. Task pane: Error Wrap button (with input field for fallback value).

**Contract.**
- For each formula cell in the selection, replace its formula with `=IFERROR(<original_expression_without_leading_=>, <fallbackValue>)`.
- Non-formula cells are skipped.
- Default fallback value is `NA()` when not configured.
- Existing `IFERROR` wrappers are not detected or deduplicated — wrapping can nest.
- The configured fallback value is persisted (scope: workbook-first, user default, built-in).

**Examples.**
- `=A1/B1` with fallback `NA()` → `=IFERROR(A1/B1, NA())`.
- `=SUM(A1:A3)` with fallback `0` → `=IFERROR(SUM(A1:A3), 0)`.
- `=IFERROR(A1/B1,0)` with fallback `NA()` → `=IFERROR(IFERROR(A1/B1,0), NA())` (wrapping nests).
- Constant `100`, text `"x"`, blank → unchanged.

**Edge cases.**
- Empty selection → action disabled.
- Fallback value syntactically invalid as a formula fragment → task pane surfaces a validation error and action is rejected before any cell is changed.

### 3.12 Auto-color Numbers

**Purpose.** Color-code the font of each non-empty cell in the selection based on its semantic type (input, formula, link to another sheet or workbook, hyperlink, etc.).

**Trigger.** Ribbon: Formatting → Auto-color Numbers. Task pane: Auto-color Numbers button.

**Contract.**
- Operates on each non-empty cell in the selection.
- Classify each cell into one of seven categories, then apply the corresponding palette color to the cell's font:
  1. `input` — numeric or non-date value, not a formula.
  2. `formula` — formula that does not reference any other cell.
  3. `worksheetLink` — formula that references only cells on the same sheet.
  4. `workbookLink` — formula that references other sheets in the same workbook.
  5. `external` — formula that references cells in other workbooks.
  6. `hyperlink` — cell has a hyperlink attached.
  7. `partialInput` — formula that is primarily a constant combined with a small reference (e.g. `=100 + A1`).
- For formula cells, classification precedence is: `partialInput` → `workbookLink` → `worksheetLink` → `external` → `input` → `formula`.
- For non-formula cells: `hyperlink` first; otherwise `input` for numeric / non-date values.
- Blank cells are not changed.
- The palette is configurable per category (see Settings, §4).

**Default palette.**
- `input`: `#0000FF` (blue)
- `formula`: `#000000` (black)
- `worksheetLink`: `#008000` (green)
- `workbookLink`: `#CC99FF` (purple)
- `external`: `#00B0F0` (cyan)
- `hyperlink`: `#FF8000` (orange)
- `partialInput`: `#800000` (dark red)

**Examples.**
- Cell `=100` → `formula` (black).
- Cell `=A1+B1` on the same sheet → `worksheetLink` (green).
- Cell `=Summary!A1` → `workbookLink` (purple).
- Cell `='[Model.xlsx]Sheet1'!A1` → `external` (cyan).
- Cell with a typed number and no formula → `input` (blue).
- Cell `=100+A1` → `partialInput` (dark red; constants mixed with small references take precedence).

**Edge cases.**
- Selection entirely blank → no-op.
- Very large selections (tens of thousands of cells) must meet the performance budget in §5.4; chunk or use native range APIs as needed.

### 3.13 CAGR

**Purpose.** Compute the compound annual growth rate between two values over a number of periods.

**Trigger.**
- As a worksheet function: `=CAGR(startValue, endValue, years)` typed into any cell.
- As a task pane calculator: the user enters start, end, and years into three fields and the result displays in the task pane (no cell mutation).

**Contract.**
- `CAGR(startValue, endValue, years) = (endValue / startValue)^(1 / years) − 1`.
- `years` may be inferred from a range input: when called as `=CAGR(range)` with a range of values, `startValue = range[0]`, `endValue = range[last]`, `years = range.length − 1`.
- Returns `#VALUE!` when:
  - `startValue ≤ 0`
  - `endValue ≤ 0`
  - `years ≤ 0`
  - any argument fails numeric conversion.

**Examples.**
- `CAGR(100, 121, 2)` → `0.1` (10% per year).
- `CAGR(100, 121)` from range `[100, 121]` → `0.21` (one period).
- `CAGR(100, 110, 121)` from range `[100, 110, 121]` → `0.1`.
- `CAGR(0, 121, 2)` → `#VALUE!`.
- Range of one value `[100]` → `#VALUE!`.
- `CAGR(-100, 121, 2)` → `#VALUE!`.

**Edge cases.**
- Task pane calculator applies the same contract; invalid inputs are surfaced in the task pane as "Invalid input" rather than written as `#VALUE!`.
- Large values that overflow native numeric representation → `#VALUE!` with a task pane tooltip in calculator mode.

### 3.14 Settings editor

**Purpose.** Give the user a visual editor for all configurable presets: number formats, date formats, cell formats, text styles, auto-color palette, error-wrap fallback, and trace parameters.

**Trigger.** Ribbon: Settings → Settings. Task pane: Settings section (always accessible in the task pane).

**Contract.**
- The editor is tabbed. One tab per format category plus tabs for Auto-color, Error Wrap, and Trace.
- Each preset-list tab supports: add a preset, edit a preset's name and format code, reorder presets, delete a preset, restore category defaults.
- The Auto-color tab exposes a color picker per category (seven pickers).
- The Error Wrap tab exposes a single text field for the fallback value, validated as a formula fragment.
- The Trace tab exposes max depth (integer, default 10) and a safety limit (integer, default 500) beyond which a trace expansion stops and shows a warning.
- Two save actions are always visible: **Save to this workbook** and **Save as my default**. The first writes workbook-scoped settings; the second writes user-scoped settings (§4.1).
- Changes in the editor are preview-only until the user clicks a save action. "Revert" discards unsaved edits.

**Examples.**
- User renames `#,##0.00` to `Two decimals`, moves it to the top of the list, saves to workbook → every subsequent Cycle Number Format invocation on this workbook uses the new list starting with `Two decimals`.
- User sets the `input` auto-color to `#000080`, saves as default → every workbook opened by this user without its own Auto-color override uses `#000080` for inputs.

**Edge cases.**
- Invalid format code → save action disabled with inline error ("Format code not recognized").
- Empty preset list after edits → save disabled with inline error ("At least one preset is required").
- Task pane closed with unsaved edits → prompt: "Discard unsaved settings changes?"

### 3.15 Reset Formats

**Purpose.** Restore format presets (number, date, cell, text style, auto-color palette, error-wrap default) to built-in defaults.

**Trigger.** Ribbon: Settings → via Settings editor "Reset to defaults" action. Hotkey: Ctrl + Shift + 0. Task pane: Reset Formats button.

**Contract.**
- Confirms before acting with a clear warning: "This will replace your current XLerate format settings with defaults. Continue?" with Reset / Cancel.
- On confirm, restores defaults at the **current resolved scope**:
  - If the active workbook has workbook-scoped settings, those are cleared (so user defaults or built-ins take over).
  - Otherwise, user-scoped settings are restored to built-ins.
- Cycle state indices (which preset is "current") are cleared so the next cycle invocation starts at the first preset.
- Does not affect Auto-color output already applied to cells, Error Wrap changes already written, or Consistency marks.

**Examples.**
- Workbook has its own number-format list. User presses Ctrl + Shift + 0, confirms → workbook-scoped settings cleared; subsequent cycles use user defaults (or built-ins if user has none).
- Workbook has no overrides. User runs Reset → user-scoped settings cleared; workbook now uses built-in defaults.

**Edge cases.**
- User cancels the confirmation → no state change.
- No custom settings present at any scope → confirmation still shown; on confirm, action is a no-op.

---

## 4. Settings and persistence

### 4.1 Resolution order

For any configurable setting (preset lists, palettes, trace parameters, error-wrap default), the effective value is resolved in order:

1. **Workbook-scoped.** Stored inside the `.xlsx` file so that it travels on email and upload.
2. **User-scoped.** Per-machine defaults the user has customized.
3. **Built-in.** Shipped with XLerate.

The first scope that has a value for the requested setting wins. Scopes do not merge inside a single preset list: if the workbook has any number-format presets at all, its list is used in full; the user-scoped list is not appended.

### 4.2 Categories persisted

| Category | Scope | Notes |
|---|---|---|
| Number format presets | Workbook > User > Built-in | Ordered list |
| Date format presets | Workbook > User > Built-in | Ordered list |
| Cell format presets | Workbook > User > Built-in | Ordered list |
| Text style presets | Workbook > User > Built-in | Ordered list |
| Auto-color palette (7 categories) | Workbook > User > Built-in | Per-category color |
| Error Wrap fallback value | Workbook > User > Built-in | Single string |
| Trace max depth, safety limit | Workbook > User > Built-in | Integers |
| Formula consistency marks | Workbook only | Survives close/reopen |
| Cycle state (current index) | Session only | Reset on workbook reopen |

### 4.3 Built-in defaults

- Number formats: `#,##0`, `#,##0.00`, `0.0%`, `$#,##0`.
- Date formats: `dd-mmm-yy`, `dd/mm/yyyy`, `mmm yyyy`, `yyyy-mm-dd`.
- Cell format presets: Normal, Inputs, Good, Bad, Important.
- Text style presets: Heading, Subheading, Sum, Normal.
- Auto-color palette: per §3.12.
- Error Wrap fallback: `NA()`.
- Trace max depth: `10`. Trace safety limit: `500`.

### 4.4 User-visible persistence behaviors

- The Settings editor always shows which scope the currently-displayed values came from ("From this workbook" / "From your defaults" / "Built-in defaults").
- Save actions are explicit: **Save to this workbook** vs **Save as my default**.
- When the user opens a colleague's workbook that has workbook-scoped overrides, those overrides apply to that workbook's session without modifying the user's own defaults.
- When the user opens a workbook with no overrides, their own user-scoped defaults apply; built-ins apply only if the user has also never customized.

---

## 5. Cross-cutting behaviors

### 5.1 Undo semantics

- Every XLerate action that mutates the workbook produces exactly one native Excel undo step. A single Ctrl+Z reverts the entire action.
- Read-only actions (Trace Precedents / Dependents, CAGR calculator) produce no undo entry.
- Multi-cell operations (e.g. Cycle Number Format on 1000 cells) are a single undo step regardless of cell count.

### 5.2 Confirmations

The following actions confirm before executing. No other action confirms — hotkeys must feel instant.

- Reset Formats (§3.15).
- Clear Consistency Marks (§3.6).
- Settings editor: discarding unsaved edits (§3.14).

### 5.3 Error messaging

- Failures surface in the task pane as a plain-English message. Never an exception dump. Never a silent failure.
- Examples of required copy style:
  - "Smart Fill Right needs a reference row within 3 rows above the active cell."
  - "Sheet is protected; unprotect to apply consistency marks."
  - "Fallback value is not a valid formula fragment."
- An action that fails partway leaves the workbook unchanged. If that is not achievable, the action must roll back to the pre-action state before surfacing the error.

### 5.4 Performance requirements

Performance is a first-class requirement, not a polish item. The following budgets are product requirements and take precedence over implementation convenience.

- **Trace Precedents / Dependents** — opening is imperceptible. The initial view loads **only level 1** (direct precedents/dependents of the active cell). Each deeper level loads on demand within **200 ms** of the expansion request. Eager full-graph traversal across all sheets on open is explicitly forbidden. Where platform APIs provide native direct-precedent / direct-dependent discovery, they are preferred over hand-rolled parsing.
- **Format cycles** (number / cell / date / text) — hotkey or button press to applied change: **under 200 ms** on a selection of up to 10,000 cells.
- **Smart Fill Right** — **under 500 ms** on a destination range of up to 10,000 cells.
- **Auto-color Numbers** — **under 1 second** on a used range of up to 50,000 cells.
- **Horizontal Formula Consistency** — **under 1 second** on a used range of up to 50,000 cells.

**General rule.** No feature blocks the UI for more than one animation frame (~16 ms) on its fast path. Any operation that might exceed its budget above must either (a) use native Excel range APIs where available, (b) run progressively in chunks that yield to the UI, or (c) display a progress indicator and remain cancellable. Eager full-graph or full-sheet traversals on the opening / fast path are banned by default.

### 5.5 Keyboard parity

Every ribbon action must be reachable by keyboard. Either via a hotkey (§2.2) or via the standard Alt-key ribbon navigation. No feature requires a pointing device.

### 5.6 Selection semantics

- Unless otherwise specified, a feature operates on the current Excel selection.
- Multi-area selections (Ctrl-click) are supported for every per-cell feature: Switch Sign, Cycle Number / Date / Cell / Text, Error Wrap, Auto-color Numbers.
- Features that require a single active cell (Trace, Smart Fill Right, CAGR calculator) ignore everything except the active cell and proceed.

### 5.7 Unavailable contexts

- Features that require a selection are disabled (button greyed, hotkey no-op, task pane button disabled) when no worksheet is active or when the selection is empty.
- Features that mutate the workbook are disabled when the active sheet is protected; the task pane explains which sheet is protected and suggests unprotecting.

---

## 6. Out of scope

The following are explicitly not part of v1:

- Excel on the web.
- Excel Mobile.
- Internationalization / localization — English UI copy only.
- Accessibility beyond keyboard parity — full WCAG compliance (screen-reader labels, contrast token system) deferred to a later spec.
- Telemetry, analytics, usage tracking.
- External integrations — no API calls, no cloud sync of settings, no external data sources.
- Cross-device settings sync — user-scoped defaults (§4.1) are per-machine. A user signed in on two different computers maintains two independent sets of defaults in v1.
- Net-new features beyond VBA parity — any capability not present in the VBA add-in (other than the explicit additions listed in this document: Clear Consistency Marks, CAGR as worksheet function, hybrid settings resolution, progressive Trace) is deferred.

---

## Appendix A — Relationship to implementation

This spec describes what the product does, not how it is built. In particular:

- The spec does not name Office.js APIs, TypeScript modules, storage formats, or UI frameworks.
- The spec does not describe the harness, build system, test infrastructure, or CI pipeline. Those are separate documents.
- When implementation questions arise that change user-visible behavior, the spec is updated first; the implementation then follows.

Behaviors documented here that are known to be absent or divergent in the current TypeScript implementation are tracked in the migration plan (see the Phase 1 plan document). The spec does not enumerate them here because the spec describes the destination, not the route.
