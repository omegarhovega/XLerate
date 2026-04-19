# XLerate — Working with this repo

## What you are looking at

This is an Excel add-in for financial modelers, mid-migration from VBA to
Office.js + TypeScript. The legacy VBA source is at `src/`; the TypeScript
add-in is at `XLerate/`. **All TypeScript work happens under `XLerate/`.**

## Source of truth for behavior

The functional spec lives at `specs/2026-04-17-xlerate-functional-spec.md`
(local, gitignored). It defines what the product should do. When changing
behavior, update the spec first, then the code. When user-visible behavior
and the spec disagree, the spec wins.

## Architecture rules (enforced by the harness)

- `src/core/**` — pure domain logic. Takes plain data, returns plain data.
  Must not import `office-js`, must not touch the DOM, must not reach out of
  process. Every module here has unit tests.

- `src/adapters/**` — the boundary with Excel. `excelPort.ts` defines the
  interface. `excelPortLive.ts` is the ONLY place in the codebase that may
  import `office-js` and call `Excel.run`. `excelPortFake.ts` is the
  in-memory test double.

- `src/services/**` — features that compose `core` + `adapters`. A service
  takes an `ExcelPort` as a parameter; it does not construct one.

- `src/taskpane/**` — UI glue. Reads DOM, wires buttons, constructs
  `ExcelPortLive`, calls services.

**The harness enforces these boundaries via ESLint (`no-restricted-imports`),
dependency-cruiser, and a strict TypeScript config for new layers. If you
need to import `office-js` outside `adapters/`, stop and reconsider the
design — the answer is almost always "extend the port."**

**Pragmatic exception during migration:** `src/taskpane/` files use the
Office.js globals (`Excel`, `Office`) directly for features that haven't
been migrated through the port yet (`taskpane.ts`'s own `Excel.run` calls,
`traceDialogLauncher.ts`, `traceExcelNeighbors.ts`, `traceDialog.ts`,
and `ribbonActions.ts` — the ribbon ExecuteFunction handlers, which used
to live in `src/commands/commands.ts` before the shared-runtime
migration). The dependency-cruiser rule restricts `office-js` *imports*
to `adapters/`; globals-via-`/* global Excel, Office */` are not imports
and remain allowed in `taskpane/`. When a taskpane feature eventually
gets its harness migration, Office.js usage moves into the port and this
exception shrinks.

## The evaluator

`npm run ci:all` (inside `XLerate/`) is the single command that tells you
whether a change is safe. It currently runs: `typecheck:core` →
`typecheck:harness` → `lint:harness` → `arch:check` → `test:core` →
`build:dev`. Every task you complete should end with a green `ci:all`.

Individual stages:
- `npm run typecheck:core` — strict tsc against `src/core` and `tests`
- `npm run typecheck:harness` — strict tsc against `src/adapters` and `src/services`
- `npm run lint:harness` — ESLint scoped to the harness-checked adapter/service paths
- `npm run arch:check` — dependency-cruiser
- `npm run test:core` — Vitest run
- `npm run build:dev` — webpack development build

`open-issues.md` tracks the currently-known gap between this practical
gate and the stricter full integration gate we eventually want
(`lint`, manifest validation, production build).

## When making a change

1. Read the relevant section of the spec.
2. If there isn't a Vitest test covering the behavior you're about to change,
   write one first (use the fake port if the test needs Excel).
3. Implement. Keep commits small.
4. Run `npm run ci:all` and make it pass before committing.
5. Manual Excel sideload verification is still required for anything that
   touches `src/taskpane` or `src/adapters/excelPortLive.ts` — the harness
   cannot catch Office.js API misuse at the boundary.

## Things the harness does NOT check

- Whether Office.js APIs behave the way you think they do on real Excel.
- UI layout, ribbon wiring, manifest correctness.
- Cross-platform differences (Windows vs Mac).
- Performance budgets from the spec §5.4 — those are verified by manual
  profiling in Phase 3+.

For those, sideload into Excel Desktop and verify by hand. See
`sideload-checklist.md` at the repo root for the per-feature manual test
protocol. **Any change that touches `src/adapters/excelPortLive.ts` or a
taskpane handler MUST go through that checklist before the work is
considered done.** Contract tests prove your logic is correct against the
fake; they cannot prove Office.js does what you assume.

## Office.js gotchas we have hit (must-know before editing `excelPortLive.ts`)

Each entry here is a real bug we shipped and then had to fix. If you are
editing the live adapter, review this list first.

### Fill: set pattern before color

Setting `range.format.fill.color = "#..."` on a cell whose current pattern
is `"None"` (the default for unformatted cells) does **not** reliably render
a solid fill. Office.js requires the pattern to be set explicitly.

```typescript
// WRONG — fill will not render on a previously-unfilled cell
cell.format.fill.color = "#FFFFCC";

// RIGHT — pattern first, then color
cell.format.fill.pattern = "Solid";
cell.format.fill.color = "#FFFFCC";
```

In this codebase: `applyFillMutation` in `excelPortLive.ts` is the single
chokepoint that must obey this rule. If you add a new fill-touching path,
it must go through `applyFillMutation`. Clearing fills uses
`cell.format.fill.clear()` which resets both pattern and color.

### Fill: reads back as `pattern=null` even when set — match tolerantly

This is the quirk that caused Cycle Cell Format to appear broken through
multiple fix attempts. After we write `cell.format.fill.pattern = "Solid"`
and `cell.format.fill.color = "#FFFFCC"`, reading the same cell via
`format.fill.pattern` returns **literal `null`** on Excel Desktop (and
inconsistently `"None"` on Excel Online) — even though the fill is solid
and renders correctly.

Any logic that reads back a cell's formatting and compares it to a preset
(like our Cycle Cell Format cycle-detection) **must treat `fillPattern` of
`null` or `"None"` as potentially matching an expected `"Solid"` preset**,
with the color comparison breaking the tie. The canonical implementation
is `doesFillMatch` in `src/core/cellFormatCycle.ts` — copy that shape if
you write a new format-matching path.

Without this tolerance, the cycle gets stuck on whichever preset was first
applied: the read-back never matches anything, so `computeNextCellFormat`
always falls through to "apply the first preset."

### Borders: never set `border.color` on a `None`-style border

Office.js silently upgrades a border's style from `"None"` to `"Continuous"`
the moment you assign a color to it. This is obnoxious because a clearAll
loop that sets style to `"None"` and then an apply loop that tries to "also"
set the color ends up leaving borders everywhere.

```typescript
// WRONG — Office.js upgrades style to Continuous
border.style = "None";
border.color = "#000000";

// RIGHT — guard the color assignment
border.style = "None";
if (border.style !== "None") {
  border.color = "#000000";
}
```

`applyBorderEdge` in `excelPortLive.ts` enforces this rule. The pre-migration
`setRangeBorder` helper in `taskpane.ts` had the same guard — it's a
well-known Office.js surface bug, not something we discovered.

### `Office.context.document.settings.saveAsync` breaks the Excel undo chain — never call it in a cell-mutating handler

On Excel Desktop (WebView2 and COM add-in hosts), `saveAsync` commits an
out-of-band workbook save that isn't part of Excel's native undo model.
If your handler does this:

```typescript
await Excel.run(async (context) => { /* ... mutate cells ... */ });
await saveDocumentSettingsAsync(); // <-- THIS IS THE TRAP
```

...then the first time the user clicks anywhere on the sheet after the
handler finishes, Excel flushes its pending undo boundary and the
`Excel.run` mutations become unreachable via Ctrl+Z. This is unlikely to
reproduce in contract tests (the fake port has no undo model), so it will
only be caught in sideload.

**The rule:** `saveAsync` is reserved for handlers that do **not** also
mutate cells in the same invocation — the Format Settings editor's Save
and Reset buttons are the only legitimate callers in this codebase.

Do not reintroduce persistent state for features like formula-consistency
marks, cycle indices, or any "pending operation" state. Session-only
module variables (see `textStyleCycleIndex`) or state inferred from the
cells themselves (see `runCycleCellFormatService`, which reads current
formatting instead of tracking an index) are the right alternatives. If
you genuinely need persistence across a close/reopen, the feature needs
to accept that Ctrl+Z won't cover it — design for that, don't split the
difference.

### Reversible operations — prefer Ctrl+Z, not a restore snapshot

When a feature applies visual state to cells it did not own (e.g. the
formula-consistency check fills cells green/red), prefer the single Excel
undo step as the "remove" path. Do not snapshot original state via
`document.settings` — that's the gotcha above.

If Ctrl+Z is insufficient for your use case, the feature needs a redesign
discussion, not a settings-backed workaround. The historical approach
(snapshot cells into `Office.context.document.settings`, then restore
from the snapshot) shipped and was removed precisely because the save
step broke the undo chain and row/column insertions invalidated the
stored addresses. See spec §3.5 / §3.6.

Still true: never bulk-wipe with `sheet.getUsedRange(true).format.fill.clear()` —
that nukes user formatting alongside ours.

### Shared runtime: ribbon handlers live in the taskpane iframe

The manifest declares `SharedRuntime 1.1` with a long-lifetime Taskpane
runtime. That means there is **no separate commands iframe**: ribbon
`ExecuteFunction` actions dispatch into the already-loaded taskpane
JavaScript context. `Office.actions.associate(...)` calls in
`src/taskpane/ribbonActions.ts` (imported for side effects from
`taskpane.ts`) register the ribbon functions. A click on a ribbon button
invokes them in-process — no cold start, no IPC boundary crossing.

Implications:

- Module-level state (e.g. the text-style cycle index) IS shared between
  ribbon and taskpane entry points because they are the same module
  graph. Session-only `let` variables are viable again.
- DevTools-on-the-taskpane shows all `[XLerate ribbon]` errors. There is
  no separate commands DevTools to hunt for.
- Host requirement: Excel 2021+ / Microsoft 365 Desktop / Excel Online.
  Excel 2019 and earlier will not load the add-in.
- Every associated function MUST still call `event.completed()` in a
  `finally`, or the ribbon button stays in a "busy" state. The `finish()`
  wrapper in `ribbonActions.ts` is the one chokepoint that enforces this.

Do NOT reintroduce a `src/commands/` directory or a second webpack entry
for ribbon dispatch — that re-splits the runtimes and brings the
sluggishness back.

### Shared runtime: trace dialog close uses `Workbook.focus()` on Desktop

When the trace dialog closes, the supported Desktop fix is
`context.workbook.focus()` from `ExcelApiDesktop 1.1`. That API restores
keyboard events to the workbook/grid more reliably than the older
`window.blur()` heuristic. We still activate/select the current cell so
the right location is visible, then call `workbook.focus()` when the
requirement set is available.

```typescript
await Excel.run(async (ctx) => {
  const workbook = ctx.workbook;
  const cell = workbook.getActiveCell();
  cell.load("worksheet/name");
  await ctx.sync();
  cell.worksheet.activate();
  cell.select();
  if (Office.context.requirements.isSetSupported("ExcelApiDesktop", "1.1")) {
    workbook.focus();
  }
  await ctx.sync();
});
```

This still must run AFTER the browser's post-close focus return. If
called synchronously from `dialog.close()`'s caller, the browser's own
focus return happens later and overrides it. Defer via
`setTimeout(..., 50)`. See `pullFocusToGrid` in
`src/taskpane/traceDialogLauncher.ts`.

### Office Dialog API windows cannot call `Excel.run`

Documented restriction:
https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins
— *"the dialog cannot use Office host-specific APIs like Excel.run or
Word.run to interact with the host document."*

The error you'll see in the dialog's own console if you forget this:
`"You cannot perform the requested operation."` thrown out of
`Excel.run`. Swallowed into whatever try/catch surrounds it — in Phase B
this surfaced only as a generic "Trace failed" status line until we
opened DevTools on the dialog window itself.

**The pattern we use instead** (see `src/taskpane/traceDialog.ts` +
`traceDialogLauncher.ts`):

1. Dialog page is pure UI — it receives data via
   `Office.context.ui.addHandlerAsync(DialogParentMessageReceived, …)`.
2. Parent runtime (taskpane or commands) runs the `Excel.run` work and
   pushes the result via `dialog.messageChild(JSON.stringify(payload))`.
3. Dialog sends back user actions (navigate, close) via
   `Office.context.ui.messageParent`. Parent handles the Excel work
   those actions imply.
4. Handshake: dialog sends `{ action: "ready" }` after registering its
   listener; parent pushes rows on that signal (not before, or the
   message lands before the dialog can hear it).

New dialog features MUST use this pattern. Do not try to "just put an
`Excel.run` in the dialog" — it compiles, typechecks, and will fail at
runtime in a hard-to-debug way.

### Excel serializes concurrent add-in API calls during dialog spawn

Discovered in Phase B perf work (reverted in `bb396c2`): running an
`Excel.run` in parallel with `Office.context.ui.displayDialogAsync`
from the same add-in runtime does NOT overlap. Excel's add-in message
queue processes them one-at-a-time. Measured: a `getActiveCell` +
`context.sync()` that was 2 ms in serial mode became **~1,000 ms**
when fired alongside a `displayDialogAsync` call.

Implication: don't design perf optimizations around "let the dialog
spawn while compute runs." There's no parallelism benefit. The
`computeTrace` in `traceDialogLauncher.ts` runs serially after the
dialog signals `ready`, and that's the correct shape — earlier
parallel attempts only added contention overhead.

If you find a future case where this appears to hold (e.g. on a newer
Office host version), re-measure before committing — Microsoft could
change this and the cost is model-dependent.

### Progressive streaming for long operations — use `messageChild` per logical chunk

When an operation produces incremental results (e.g. BFS trace expanding
one level at a time), push each chunk via `dialog.messageChild` as it
completes rather than accumulating the full result before a single send.
The dialog repaints between chunks and the user perceives the operation
as fast even when total wall-clock time is the same.

The builder's `onProgress` callback (see `core/traceBuilder.ts`) gets
`await`ed per level so the browser has a chance to paint before the
next level's Excel sync starts. Don't fire-and-forget the callback
inside a tight loop — without the await, the updates queue up behind
the compute and the user never sees intermediate states.

### Diagnostic `[trace-perf]` logs are provisional

Several `console.log("[trace-perf] …")` calls exist in
`src/taskpane/traceDialog.ts` and `traceDialogLauncher.ts` — commit
`944029c` introduced them to diagnose dialog-open latency; subsequent
commits (`bb396c2`, `2eacc0f`) used them to validate changes. They're
harmless in production but clutter the console. When a new perf issue
isn't actively being diagnosed, remove them via a focused revert; they
live behind a `logTracePerf` helper so the cleanup is localized.

### Array-formula mutation is deferred

`ExcelPortLive.applyMutations` explicitly throws on `kind: "arrayFormula"`
with a "Phase 2" note. The pure core and the fake support array formulas
correctly. Until someone implements `Range.setArrayFormula` plumbing, any
array-formula cell routed through live Excel is an unhandled case.

### `getUsedRange(true)` excludes formatted-only cells

`valuesOnly=true` treats cells that have formatting but no value as empty.
That's the default in our read paths and is usually correct. But if you're
doing a cleanup pass over formatted-but-valueless cells, use
`getUsedRange(false)` or iterate a specific address list.

### Office.js property loading

Every property you read needs to be loaded via `range.load([...])` or
`format.load([...])` before `context.sync()`. Forgetting to load a property
typically manifests as `undefined` in the snapshot rather than an error —
which contract tests cannot catch because the fake doesn't have this
two-phase API. If a snapshot field is mysteriously undefined in sideload,
check the `load()` calls first.

### When in doubt, compare against the old working code

The pre-migration handlers in `taskpane.ts` (some still present as dead
code) were battle-tested in live Excel. When adding or changing live
behavior, find the old handler for the same feature and diff against it.
If your new implementation skips a call the old one made (especially
setting `fill.pattern`, clearing borders, or setting `font.underline` as a
string), question whether the skip is safe.
