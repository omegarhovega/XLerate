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

## The evaluator

`npm run ci:all` (inside `XLerate/`) is the single command that tells you
whether a change is safe. It runs: typecheck → lint → arch check → tests →
build. Every task you complete should end with a green `ci:all`.

Individual stages:
- `npm run typecheck:core` — strict tsc against `src/core` and `tests`
- `npm run typecheck:harness` — strict tsc against `src/adapters` and `src/services`
- `npm run lint` — ESLint
- `npm run arch:check` — dependency-cruiser
- `npm run test:core` — Vitest run
- `npm run build` — webpack production build

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

For those, sideload into Excel Desktop and verify by hand.
