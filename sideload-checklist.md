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

The XLerate tab holds eleven buttons across four groups. Spec §2.1's
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
- [ ] **Formulas** group contains four buttons in this order:
      **Trace Precedents**, **Trace Dependents**, **Switch Sign**,
      **Smart Fill Right**.
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
- [ ] **Cycle Text Style** consistency check: click the ribbon
      button; then from the taskpane click **Cycle Text Style**; they
      should advance through the preset list in the same sequence
      (localStorage-backed shared index — see CLAUDE.md "Move
      text-style cycle index to localStorage").
- [ ] Click **Auto-color** on a column mixing numbers and formulas →
      each cell gets its category color per spec §3.12.
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

### Trace Dialog (Phase B + Phase C of the trace-navigation plan)

The dialog renders the precedent/dependent graph as an interactive tree,
streams rows in per BFS level as they arrive, and supports keyboard-first
navigation with live selection-following in the grid.

Primary flow — opening:

- [ ] The **XLerate** tab appears on the Excel ribbon (replaces the
      single Home-tab group from earlier builds). Opening the tab shows
      four groups in order: **Formulas**, **Auditing**, **Formatting**,
      **Settings**. (See the "Ribbon tab" section below for full
      per-button verification.)
- [ ] On a cell with precedents, click XLerate tab → Formulas group →
      **Trace Precedents**.
      Dialog opens center-screen. Excel's active cell is unchanged
      from before the click.
- [ ] Dialog shows "Loading…" briefly (≤ ~1 s on a warm bundle),
      then the root row appears with a ▼ chevron, its direct
      precedents (level 1) visible below it with either ▶ (collapsed
      parent) or a greyed leaf bullet (no own precedents).
- [ ] Status line during streaming reads
      `Trace precedents on <addr>: loading… N cells so far (depth K)`;
      changes to
      `Trace precedents on <addr>: N cells.`
      when the BFS completes.
- [ ] Row 0 (root) has the cyan focus ring. Deeper levels in `allRows`
      are present but collapsed by default — you can see there's more
      to explore from the ▶ chevrons but not the content.

Tree interaction — mouse:

- [ ] Click a ▶ chevron → that subtree expands, ▶ becomes ▼, children
      slot in below their parent.
- [ ] Click a ▼ chevron → subtree collapses.
- [ ] Click anywhere on a row EXCEPT the chevron → Excel's active cell
      jumps to that cell (live-nav). Focus ring moves to the clicked row.
- [ ] Click the chevron while a descendant is focused, then collapse
      → focus moves to the collapsing parent (focus can't be on a
      now-hidden row); Excel's selection does NOT jump during the
      collapse, only on the actual focus-change fallback.

Tree interaction — keyboard:

- [ ] **ArrowDown / ArrowUp**: next / previous *visible* row. Collapsed
      subtrees are skipped. Grid selection follows live.
- [ ] **Home / End**: first / last visible row.
- [ ] **ArrowRight** on a ▶ (collapsed parent) → **expand** in place.
      Focus stays on the same row. Excel's selection does NOT change.
- [ ] **ArrowRight** on a ▼ (expanded parent) → move focus to first
      child. Live-nav fires; Excel follows.
- [ ] **ArrowRight** on a leaf → no-op.
- [ ] **ArrowLeft** on a ▼ (expanded parent) → **collapse** in place.
      Focus stays. Excel's selection does NOT change.
- [ ] **ArrowLeft** on a ▶ or a leaf → move focus to parent. Live-nav
      fires.
- [ ] **Enter** or **Escape** → close dialog.

Progressive streaming:

- [ ] Hold ArrowDown during a long trace's streaming phase. As deeper
      levels arrive, they append under their expanded ancestors; rows
      already above your focus position don't visibly shift; your
      arrow navigation continues seamlessly.
- [ ] Collapse a mid-depth subtree manually. While the trace is still
      streaming in deeper levels below it, those additions should NOT
      re-expand your collapsed subtree. The ▶ stays ▶.

Focus-return-to-grid on close:

- [ ] **Esc** closes. Excel's active cell is wherever the user last
      navigated to (NOT reverted). Immediately press any arrow key on
      the grid — if Excel's selection moves, `Workbook.focus()` handed
      keyboard events back to the workbook successfully.
- [ ] After the same close, try a non-arrow keyboard gesture such as
      typing a digit or using **Ctrl+Arrow**. On supported Desktop hosts,
      the workbook should behave as if the sheet retained focus, not
      merely as if the selection moved.
- [ ] If focus does NOT return on a host that lacks `ExcelApiDesktop 1.1`
      support, document it as a known limitation. The current supported
      fix is Desktop-only because `Workbook.focus()` is Desktop-only.

Idempotency + re-entry:

- [ ] With a dialog open, click the **Trace Precedents** ribbon button
      again. The old dialog closes and a new one opens for the
      (possibly updated) active cell. Expansion state resets to the
      new trace's root.
- [ ] Click **Trace Dependents** on a cell that only has precedents →
      dialog opens with "1 cell" (just the root, leaf bullet). Arrow
      keys don't do anything useful; Esc closes; no errors.
- [ ] Click **Trace Precedents** on an empty cell with no formulas and
      no precedents → same "1 cell" end state.

Taskpane co-entry:

- [ ] The taskpane **Trace Precedents (Dialog)** and **Trace Dependents
      (Dialog)** buttons open the same dialog. Interchangeable with the
      ribbon buttons. In shared runtime ribbon and taskpane share a
      single `activeDialog` module variable; `openTraceDialog()` calls
      `closeActiveDialog()` before displaying the new one, so opening
      a second dialog from either entry point closes the first cleanly.

Cross-interaction with the taskpane trace list (Phase A):

- [ ] Run the taskpane **Trace Precedents (Active Cell)** button (flat
      list with level numbers, Phase A). Then open the dialog via the
      ribbon for the same cell. Dialog shows the same rows but as a
      tree. Closing the dialog does not affect the taskpane list.

Undo semantics:

- [ ] After live-nav selections, press Ctrl+Z. XLerate has not added
      any undo entries of its own (trace is read-only). Excel's own
      selection history governs this; accept whatever it does.

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
