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
- [ ] Close and reopen the workbook → marks persist.

### Clear Consistency Marks (spec §3.6)

- [ ] Apply consistency marks (above). Confirm you can see them.
- [ ] Click **Clear Consistency Marks** → confirm dialog appears.
- [ ] Click Cancel → nothing changes.
- [ ] Click again and confirm → marks are gone, **original fill colors are
      restored** (set a yellow fill on one cell before marking, then
      confirm the yellow comes back).
- [ ] Cells you did not mark are unchanged (try setting a blue fill on a
      nearby cell before clearing — it must stay blue).
- [ ] Ctrl+Z → marks reappear.

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
