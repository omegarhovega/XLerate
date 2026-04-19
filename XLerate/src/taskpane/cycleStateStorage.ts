/**
 * Session-scoped storage for the text-style cycle index, shared between
 * the taskpane and the commands runtime. Both contexts load from
 * `https://localhost:3000/...` (same origin) so `window.localStorage`
 * is a shared session store across them — the ribbon button and the
 * taskpane button stay in sync without either touching
 * `Office.context.document.settings.saveAsync` (which would break the
 * Excel undo chain on Desktop; see CLAUDE.md gotchas).
 *
 * Semantics preserved from the original `textStyleCycleIndex` module
 * variable: -1 is "before the first preset" so the next cycle applies
 * preset 0. The Format Settings editor's Save / Reset flows reset to
 * -1 via `resetTextStyleCycleIndex()`.
 */

const KEY = "xlerate.textStyleCycleIndex.v1";

export function readTextStyleCycleIndex(): number {
  try {
    const raw = window.localStorage.getItem(KEY);
    if (raw === null) return -1;
    const n = Number(raw);
    return Number.isInteger(n) ? n : -1;
  } catch {
    // localStorage can throw in sandboxed / restricted contexts. Falling
    // back to -1 means each call starts the cycle fresh; acceptable.
    return -1;
  }
}

export function writeTextStyleCycleIndex(index: number): void {
  try {
    window.localStorage.setItem(KEY, String(index));
  } catch {
    // Swallow: if localStorage is unavailable, the cycle degrades to
    // "always start from the first preset", which still produces a
    // valid visible change on each click.
  }
}

export function resetTextStyleCycleIndex(): void {
  try {
    window.localStorage.removeItem(KEY);
  } catch {
    // Swallow per above.
  }
}
