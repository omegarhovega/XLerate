# XLerate Sideload Verification Checklist

The harness (`npm run ci:all`) verifies logic against `ExcelPortFake`. It
cannot verify Office.js API behavior in live Excel. This checklist is the
manual protocol that completes verification.

**Run this checklist whenever you change:**
- `src/adapters/excelPortLive.ts` (any extension to the live boundary)
- any handler in `src/taskpane/taskpane.ts`
- `manifest.xml`, `taskpane.html`, or ribbon wiring

**Setup (once per session):**

1. `cd XLerate`
2. `npm run dev-server` in one terminal
3. `npm start` (or manual sideload) in a second terminal
4. Open a blank Excel workbook

Mark each item with ✅ (passes) or ❌ (failing — STOP and fix before merging).

---

## Core cross-cutting invariants

Verify these once per session. They apply to every feature.

- [ ] **Single undo step.** Every mutating feature reverts with one Ctrl+Z.
- [ ] **Status bar updates.** The task pane shows a status message on success.
- [ ] **Error visibility.** On intentional error (e.g. Smart Fill Right with
      no reference row), an explanatory message appears — never a silent
      failure and never an unhandled exception dump.

---

## Feature checklist

### Ribbon tab (spec §2.1, Option B realized)

The XLerate tab holds twelve buttons across four groups. Spec §2.1's
split-button Format menu and the Error Wrap button (needs a text
input) are deferred — the current layout has individual cycle
buttons and no Error Wrap on the ribbon.

**Shared runtime note:** ribbon buttons should feel as fast as taskpane
buttons (no 500 ms–1 s delay). If you see a delay, the manifest's
`SharedRuntime` requirement may not be taking effect — verify the
taskpane loads automatically at Excel startup (you should NOT have to
click Show Task Pane before ribbon buttons work). If needed, clear
`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef` while Excel is closed and
re-sideload; Excel caches manifests aggressively.

- [ ] Open Excel. The ribbon shows an **XLerate** tab alongside Home,
      Insert, etc. Click it.
- [ ] **Formulas** group contains five buttons in this order:
      **Trace Precedents**, **Trace Dependents**, **Switch Sign**,
      **Smart Fill Right**, **Insert CAGR**.
- [ ] **Auditing** group contains one button: **Horizontal Check**.
- [ ] **Formatting** group contains five buttons:
      **Cycle Number**, **Cycle Date**, **Cycle Cell**,
      **Cycle Text Style**, **Auto-color**.
- [ ] **Settings** group contains one button: **Show Task Pane**.
- [ ] Hover each button — a supertip appears with the button's label
      as the title and an explanatory sentence below.
- [ ] No buttons appear on the **Home** tab (single-entry migration —
      the previous Home-tab CommandsGroup is removed).

Functional spot-check per group (full feature verification is
covered in the per-feature sections below; here we just confirm the
ribbon wiring reaches the handler):

- [ ] Click **Switch Sign** → selection's numeric/formula cells flip
      sign in one undo step. No taskpane status is shown (ribbon
      handlers intentionally don't update the taskpane DOM — they
      only call services); behavior is the visible cell change.
- [ ] Click **Smart Fill Right** on a valid active-cell formula →
      selection fills right. On an invalid cell (no formula, merged,
      no boundary) the ribbon button silently no-ops — no error
      popup, no status line. This is expected; the taskpane button
      gives the structured error message if you want one.
- [ ] Click **Horizontal Check** on a row of formulas → green/red
      marks appear; Ctrl+Z removes them.
- [ ] Click **Cycle Number** on a cell with a value → its number
      format advances. Click again → advances. Click **Cycle Cell**
      instead → advances the cell-format preset.
- [ ] Click **Cycle Text Style** on a plain cell → the next text-style
      preset applies immediately, with no separate taskpane trigger path.
- [ ] Click **Auto-color** on a column mixing numbers and formulas →
      each cell gets its category color per spec §3.12.
- [ ] Click **Insert CAGR** on a destination cell with a numeric
      series immediately to its left → the cell receives a CAGR
      worksheet formula and a percent number format.
- [ ] Click **Show Task Pane** → the XLerate task pane opens on the
      right (or re-activates if already open).

If any ribbon button appears to do nothing on click, the first
suspect is the `ribbonActions.ts` handler either missing its
`Office.actions.associate` registration or failing silently before
`event.completed()`. Open DevTools on the **taskpane** (right-click
taskpane → Inspect) — shared runtime means ribbon handlers run in the
taskpane iframe, so their `[XLerate ribbon]` error logs from the
`finish()` wrapper surface in the same console.

### Switch Sign (spec §3.3)

- [ ] Single numeric cell `10` → `-10`.
- [ ] Single formula `=A1+B1` → `=-(A1+B1)`.
- [ ] Text cell `"hello"` → unchanged.
- [ ] Multi-cell selection mixing numbers, formulas, text — all flip in one
      undo step; text cells unchanged.
- [ ] Ctrl+Z reverts the whole batch in one press.

### Error Wrap (spec §3.11)

- [ ] `=A1/0` with default fallback → `=IFERROR(A1/0, NA())`.
- [ ] Change fallback input to `0` and re-run → `=IFERROR(A1/0, 0)`.
- [ ] Numeric / text / blank cells unchanged.
- [ ] Existing `=IFERROR(...)` wrappers nest (do not deduplicate) per spec.

### Insert CAGR (spec §3.13)

- [ ] Row has `100`, `110`, `121` in `B5:D5`. Select `E5` and click
      **Insert CAGR** → `E5` receives a worksheet formula equivalent
      to `=POWER(D5/B5,1/2)-1` and shows a percent result.
- [ ] Press **Ctrl+Z** once → both the inserted formula and the
      percent format revert in one undo step.
- [ ] Place text or a blank cell immediately to the left of the
      destination → **Insert CAGR** leaves the destination cell
      unchanged.

### Auto-color Numbers (spec §3.12)

- [ ] Number input `100` → **blue** font.
- [ ] Formula `=A1*2` (worksheet-local) → **black** font.
- [ ] Formula `=Sheet2!A1` → **green** font.
- [ ] Cell with `Insert > Hyperlink` applied → **orange** font.
- [ ] Formula `=IF(TRUE,1,0)` (no references) → **black** font (was
      mis-classified as partialInput before a Phase 2 core fix).
- [ ] Blank cell → unchanged.

### Horizontal Formula Consistency (spec §3.5)

- [ ] Type `=B$9*C10` across columns B–D. Run check.
- [ ] All three cells get a **green** fill (consistent).
- [ ] Break one formula (e.g. `=C$9*Z10`). Re-run.
- [ ] The broken cell gets a **red** fill; its neighbours stay green
      relative to each other.
- [ ] **Undo chain preservation test** — with marks visible, **click any
      other cell on the sheet** (to shift selection). Then press Ctrl+Z
      **once**. All green/red fills revert; originally-unfilled cells
      are unfilled, originally-yellow cells are yellow. If Ctrl+Z does
      nothing, or only partially reverts, we've regressed the undo chain —
      check that the handler no longer calls `saveAsync` (see CLAUDE.md
      "saveAsync breaks the Excel undo chain").
- [ ] After a successful Ctrl+Z, press Ctrl+Y → marks reappear in one redo step.
- [ ] Close the workbook without undoing. Reopen → the green/red fills are
      still there (they are regular cell fills now). There is **no Clear
      button**; to wipe them the user must either undo (if undo hasn't
      been flushed) or clear fills manually.

### Format Settings Save / Reset (spec §3.14 / §3.15)

- [ ] Open the task pane. The settings surface shows tabs for
      **Number Formats**, **Date Formats**, **Cell Formats**,
      **Text Styles**, **Auto-color Palette**, and **Trace Settings**.
- [ ] Edit a preset in the form-based editor → click **Save Settings**.
      Cycle that format type from the ribbon; the edited preset is used.
- [ ] Click **Load Saved Settings** after making unsaved edits →
      the taskpane reloads the last persisted workbook settings.
- [ ] Click **Restore Defaults** → the editor switches to built-in
      defaults and the status line explains that saving will apply them.
- [ ] After restoring defaults, click **Save Settings** → ribbon
      actions use the default presets again.
- [ ] Saving settings resets the text-style cycle index (next
      **Cycle Text Style** click starts at the first preset, not where
      it was).
- [ ] Note: saving settings still calls `saveAsync`, which is expected
      because the action updates workbook settings only; the undo chain
      for sheet mutations is not at risk here.

### Smart Fill Right (spec §3.4)

- [ ] Active cell `B5 = =A5+1`, row 4 has values in B:F, no merges → B5
      fills to F5 via the active formula.
- [ ] Row 4 has a merge in B:F, row 3 has values in B:D → B5 fills to D5.
- [ ] Active cell is not a formula → action rejected with task-pane message.
- [ ] Active cell is merged → rejected with message.

### Cycle Number Format (spec §3.7)

- [ ] Type `1234` into A1 (no specific format). Click cycle.
- [ ] A1 shows the first preset's format.
- [ ] Click again → next preset. Again → wraps back to first.
- [ ] Select A1:A3 with mixed formats → all three get the first preset.

### Cycle Date Format (spec §3.9)

- [ ] Enter today's date in A1. Click cycle.
- [ ] Format changes through each configured preset and wraps.

### Cycle Cell Format (spec §3.8)

This feature has the hardest live-Excel behavior in the whole product.
Fake-based contract tests pass trivially while live Excel exposes two
separate Office.js quirks. Work through **all** of these steps.

- [ ] Start on an empty cell. Click cycle → cell shows "Normal" (white
      fill, black font). On a fresh cell this may look identical to empty,
      so keep clicking.
- [ ] Click again → **Inputs** (yellow fill `#FFFFCC`, blue font, gray
      borders).
- [ ] Click again → **Good** (green `#C6EFCE`, dark green font `#006100`).
      **This is the canary for the null-pattern match bug** — if you see
      Good here, the fill-pattern null tolerance is working. If instead
      the cell returns to Normal white, the match logic in
      `core/cellFormatCycle.ts → doesFillMatch` has regressed.
- [ ] Click again → **Bad** (red `#FFC7CE`, dark red font `#9C0006`).
- [ ] Click again → **Important** (yellow `#FFFF00`, black bold, no
      borders). **This is the canary for the border `color-on-None` bug**
      — if the cell keeps a visible thin border that wasn't there on
      Important, `applyBorderEdge` in `excelPortLive.ts` is setting color
      on a `None` style and Office.js is upgrading the style back to
      Continuous. Same fix as before: guard the `border.color` assignment.
- [ ] Click again → wraps back to **Normal**.
- [ ] Press **Ctrl+Z** repeatedly. Each undo step reverts one cycle step;
      the whole trail back to the unfilled starting cell should undo
      without any missing intermediate state.
- [ ] Select a multi-cell range and cycle to "Inputs" → inside borders
      appear between cells, not only outside edges.

### Cycle Text Style (spec §3.10)

- [ ] Empty cell → cycle → Heading (Calibri 14, bold, gray fill, top and
      bottom borders).
- [ ] Click again → Subheading.
- [ ] Continue → Sum, Normal, Heading again (wraps).
- [ ] Normal fully resets a previously-styled cell (Calibri 11, no
      borders, white fill).
- [ ] Close and reopen the workbook → next click starts from Heading again
      (session-scoped index per spec §4.2).

### Trace Precedents / Dependents (spec §3.1, §3.2)

Trace now uses a range-aware adapter and a demand-driven dialog session.
The root and direct relationships load initially; deeper branches load only
when expanded. All Excel work is serialized through one priority scheduler.

**Initial load and direct relationships:**

- [ ] On a formula with cross-sheet precedents, click **Trace Precedents**.
      The dialog shows the root and direct relationships only. No deeper rows
      appear until a level-1 node is expanded.
- [ ] Confirm `[trace-metric]` logs show one root operation followed by one
      direct-neighbor operation. Adding unrelated worksheets does not change
      those counts.
- [ ] A constant or formula with no precedents shows only the root with a leaf
      bullet and a plain-English status; no repeated `ItemNotFound` requests.
- [ ] Trace Dependents uses the same initial shape across multiple worksheets.
      On a host below ExcelApi 1.13 it reports the requirement instead of
      attempting the unsupported API.

**Lazy expansion, caching, and references:**

- [ ] Click or ArrowRight a collapsed formula node. It briefly shows a loading
      indicator, then inserts only that node's direct children.
- [ ] Collapse and re-expand the same node. Children appear immediately and no
      new neighbor metric is emitted (session-cache hit).
- [ ] Expand one branch in a graph with many siblings. Unopened sibling branches
      remain unloaded and emit no neighbor metrics.
- [ ] In a circular/shared dependency graph, repeated cells display as ↩
      reference rows and cannot be expanded again.
- [ ] A node at configured max depth becomes a leaf without another Excel query.

**Large rectangular precedents:**

- [ ] Trace a formula such as `=SUM(Detail!A1:A10000)`. The direct result is one
      range-summary row showing 10,000 cells; logs report zero materialized
      cells for that area.
- [ ] Expand the summary. The first 100 cells appear, followed by **Load more**.
      Each click adds at most 100 cells and preserves focus. The corresponding
      `[trace-metric]` entry reports one rectangular area for a single-column
      page (and never more than three areas for a row-major multi-column page),
      rather than one Office.js proxy per cell.
- [ ] A rectangular area of exactly 50 cells expands as individual cells; an
      area of 51 cells is summarized.
- [ ] Lower the safety limit and confirm truncation occurs before additional
      cell proxies/pages are created.

**Keyboard, priority, and cancellation:**

- [ ] ArrowDown/ArrowUp navigate visible rows only; Home/End jump to the ends;
      ArrowRight expands or enters a subtree; ArrowLeft collapses or moves to
      the parent. Grid selection follows deliberate focus moves.
- [ ] Hold ArrowDown across a large visible page. Navigation stays responsive
      and Excel eventually selects the latest row rather than replaying every
      obsolete intermediate selection.
- [ ] Trigger an expansion while idle prefetch is queued. The explicit expansion
      runs first after the current Office.js boundary.
- [ ] Close the dialog during a load, then immediately open another trace. The
      old session sends no rows into the new dialog and performs no later levels.
- [ ] Enter closes the dialog and preserves the last navigated cell. Escape
      restores the original trace-root cell before closing. On Excel Desktop
      with ExcelApiDesktop 1.1, keyboard focus returns to the grid.

**Pinned formula context:**

- [ ] The complete root formula remains pinned above the scrollable tree while
      direct precedents are selected.
- [ ] ArrowDown/ArrowUp navigation in the trace tree updates the pinned formula
      highlight without adding duplicate navigation controls to the context bar.
- [ ] Selecting an A1 cell or a cell inside an A1 range highlights the matching
      token in the pinned formula. A formula with one named/table reference uses
      that token as a safe best-effort highlight.
- [ ] Expanding to a deeper level changes the pinned context to the nearest
      formula-bearing parent. In dependents mode, the selected dependent's own
      formula is shown and its parent relationship is highlighted.
- [ ] The breadcrumb shows the root-to-selection path. Clicking an ancestor
      focuses and navigates to that loaded trace row.

**Clickable formula references:**

- [ ] Trace a formula containing local cells, absolute cells, ranges, a quoted
      worksheet name, and an apostrophe-escaped worksheet name. Each supported
      A1 token is underlined and clickable in the Formula column.
- [ ] Click a formula reference. Excel selects that exact cell/range while the
      trace dialog stays open; the row itself does not fire a second navigation.
- [ ] Tab to a formula reference and press Enter. Navigation fires and the
      dialog remains open.
- [ ] A reference-looking token inside an Excel string literal (for example
      `="A1"`) remains plain text.
- [ ] Define both a workbook-scoped and worksheet-scoped named range. Each name
      is clickable from a formula and selects its resolved range. A defined name
      that resolves to a constant does not crash the dialog.
- [ ] From `Summary`, trace `=Annual_volume_growth*100` where the workbook-scoped
      name refers to a cell on `Global_Drivers`. The tooltip remains an
      unqualified named range and clicking it selects the `Global_Drivers` cell.
- [ ] Define a same-named worksheet-scoped item on `Summary`; the unqualified
      token resolves to that local item, while `Global_Drivers!Annual_volume_growth`
      resolves explicitly through the `Global_Drivers` worksheet scope.
- [ ] Create an Excel Table and test `TableName[Column]`, `TableName[#All]`,
      `TableName[#Data]`, `TableName[[#Headers],[Column]]`, and an adjacent
      column span. Each supported token selects the corresponding table range.
- [ ] Current-row structured references such as `[@Revenue]`, 3D references,
      and unsupported structured syntax remain plain text.
- [ ] An external cell/range reference is shown in the external-link style with
      a tooltip naming its workbook, sheet, and range. It is not clickable and
      does not imply that Office.js can switch workbooks.

**Undo semantics:**

- [ ] Trace, expansion, paging, and live navigation add no XLerate workbook
      mutation or undo entry.

---

## When something fails

1. Do **not** mark the feature as done.
2. Capture: what did you click, what did you expect, what happened
   instead.
3. Check `CLAUDE.md` → *Office.js gotchas we have hit* for a known pattern.
4. If new, add it to the gotchas section after you fix it so the next
   developer does not repeat the discovery.
5. Fix at the root cause, not the symptom. Use the
   `superpowers:systematic-debugging` skill if the cause is not obvious.

## Why this exists

Phase 2 shipped two Office.js-only bugs (Cycle Cell Format missing fill
pattern; Clear Consistency Marks bulk-wiping the sheet instead of
restoring originals) because every contract test passed against the
fake. Phase B added a third class (dialogs silently cannot call
`Excel.run`) discovered only in sideload. The fake port cannot model
Office.js host quirks, runtime-boundary restrictions, or Excel's
internal serialization during dialog spawn. This checklist is the
irreducible manual step that closes the gap until we have automated
live-Excel testing (Playwright-on-Excel-Online is the candidate tool
for a future phase).
