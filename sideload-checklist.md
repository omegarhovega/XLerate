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

### CAGR calculator (spec §3.13, calculator only)

- [ ] start=100, end=121, years=2 → `~0.100000`.
- [ ] start=0, end=121, years=2 → `#VALUE!`.
- [ ] Non-numeric text → `Invalid input`.

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

- [ ] Edit a preset in the Format Settings JSON editor → click **Save
      Format Settings**. Cycle that format type; the edited preset is used.
- [ ] Click **Reset Format Settings (Defaults)** → editor textarea
      repopulates with built-in defaults, status says "Format settings
      reset to defaults."
- [ ] Text-style cycle index resets (next Cycle Text Style click starts at
      the first preset, not where it was).
- [ ] Note: these actions DO call `saveAsync`, which is expected because
      they don't also mutate cells; the undo chain is not at risk here.

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

Not yet migrated through the harness (Phase 3+). If you do touch this:

- [ ] Active cell has direct precedents → tree shows root + level 1 only.
- [ ] Clicking a node jumps Excel's selection to that cell.
- [ ] Active cell has no precedents → tree shows only the root.

**Keyboard navigation (Phase A of the trace-navigation plan):**

- [ ] Run Trace Precedents against a cell with several known precedents.
      Row 0 has a visible cyan ring (`.trace-row-focused`); other rows
      do not. Hovering the tbody shows the aria-label tooltip with
      shortcut guidance in browsers that surface it.
- [ ] Click anywhere inside the taskpane (to give it document focus).
      Press **ArrowDown** — ring moves to row 1. **ArrowDown** to the
      last row; **ArrowDown** once more → ring stays on the last row
      (no wrap).
- [ ] **ArrowUp** past row 0 → ring stays on row 0 (no wrap).
- [ ] **Home** jumps ring to row 0. **End** jumps to the last row.
- [ ] **Enter** on the focused row → Excel scrolls to and selects that
      cell. The ring stays on the same row in the taskpane, so
      ArrowDown then Enter navigates cleanly down the chain.
- [ ] **Esc** → ring disappears; subsequent arrow keys in the taskpane
      do not bring it back until a new trace is run or a row is clicked.
- [ ] Click a non-first row with the mouse. Ring moves to the clicked
      row. Next **ArrowDown** advances from there, not from row 0.
- [ ] Re-run Trace Precedents on a different active cell. Ring resets
      to the new row 0.
- [ ] Run a trace that returns zero results → tbody shows the "No
      trace results." row; arrow keys do nothing; no error in the
      taskpane status.
- [ ] Sanity: Tab away from the taskpane and back. Tab order lands on
      the focused row (thanks to tabindex=0), not on the inner link
      button (which is tabindex=-1).

### Trace Dialog (Phase B of the trace-navigation plan)

Primary flow:

- [ ] The XLerate ribbon group on the Home tab has three buttons:
      **Show Task Pane**, **Trace Precedents**, **Trace Dependents**.
- [ ] With an active cell that has known precedents, click ribbon →
      **Trace Precedents**. A dialog window opens center-screen. The
      first result row has the cyan focus ring; Excel's active cell is
      unchanged from before the click (the dialog opens on that cell's
      trace tree).
- [ ] **Live-nav:** press ArrowDown. Excel's active cell jumps to the
      row-1 cell immediately; the dialog ring moves in lockstep. No
      Enter required.
- [ ] Hold ArrowDown through a long trace for ~2 seconds. The grid
      selection steps through each row without skipping; the dialog
      ring tracks. If visible lag per step appears, flag it — rAF
      coalescing is in the plan as a polish pass.
- [ ] ArrowUp walks back up symmetrically. Home / End jump to first /
      last row and Excel follows.
- [ ] **Esc** closes the dialog. Excel's active cell is wherever the
      user last arrowed to (NOT reverted to the pre-dialog cell).
      Immediately press any arrow key. If Excel's grid selection moves,
      focus-return-to-grid worked (Phase B.7 rung 1 succeeded). If the
      keypress does nothing, focus is still in the taskpane — known
      limitation, click once in the grid to recover.
- [ ] **Enter** on a mid-list row closes the dialog the same way as Esc.
- [ ] Dialog's **X button** (title-bar close) also dismisses. Taskpane
      status line does not show an error.

Idempotency + re-entry:

- [ ] With a dialog already open, click the **Trace Precedents** ribbon
      button again. The existing dialog closes and a new one opens for
      the (possibly updated) active cell.
- [ ] Click **Trace Dependents** on a cell that only has precedents →
      dialog opens, says "Trace dependents on <addr>: 1 cell" (just the
      root). Arrow keys do nothing; Esc closes; no errors.
- [ ] Click **Trace Precedents** on an empty cell with no formulas and
      no precedents → same "1 cell" end state.

Taskpane co-entry:

- [ ] The taskpane **Trace Precedents (Dialog)** and **Trace
      Dependents (Dialog)** buttons open the same dialog. Use them
      interchangeably with the ribbon buttons. State (one-dialog-only
      guard) is shared within each runtime; opening from the ribbon
      then from the taskpane closes the ribbon-opened dialog before
      opening the taskpane-opened one (each runtime has its own handle).

Cross-interaction with the taskpane trace list (Phase A):

- [ ] Run the taskpane **Trace Precedents (Active Cell)** button. List
      renders with a focused row in the taskpane. Then open the dialog
      via the ribbon for the same cell. Dialog shows the same rows.
      Closing the dialog does not affect the taskpane list.

Undo semantics:

- [ ] After live-nav selections, press Ctrl+Z in the grid. XLerate has
      not added any undo entries of its own (trace is read-only for
      the workbook). Excel's own selection history is governed by
      Excel; accept whatever it does.

Known limitation (documented, not a failure):

- [ ] On Excel Online / Mac, focus-return after Esc may not land on the
      grid; one click or keypress is needed to resume typing. The cell
      is still selected correctly.

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
restoring originals) because all 144 contract tests passed against the
fake. The fake cannot model Office.js quirks around fill rendering or
formatted-only cells. This checklist is the irreducible manual step that
closes the gap until we have automated live-Excel testing
(Playwright-on-Excel-Online is the candidate tool for a future phase).
