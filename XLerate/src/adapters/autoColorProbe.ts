/* global globalThis */

const AUTO_COLOR_PROBE_STORAGE_KEY = "xlerate.autocolor.probes";
const MAX_AUTO_COLOR_PROBE_ENTRIES = 200;

type AutoColorProbeEntry = {
  at: string;
  elapsedMs: number | null;
  label: string;
  payload?: unknown;
};

let autoColorProbeActive = false;
let autoColorProbeStart = 0;

function nowMs(): number {
  if (
    typeof globalThis !== "undefined" &&
    "performance" in globalThis &&
    typeof globalThis.performance?.now === "function"
  ) {
    return globalThis.performance.now();
  }
  return Date.now();
}

function appendProbeEntry(entry: AutoColorProbeEntry): void {
  if (typeof globalThis === "undefined" || !("localStorage" in globalThis)) {
    return;
  }

  try {
    const raw = globalThis.localStorage?.getItem(AUTO_COLOR_PROBE_STORAGE_KEY);
    const existing = raw ? (JSON.parse(raw) as AutoColorProbeEntry[]) : [];
    existing.push(entry);
    if (existing.length > MAX_AUTO_COLOR_PROBE_ENTRIES) {
      existing.splice(0, existing.length - MAX_AUTO_COLOR_PROBE_ENTRIES);
    }
    globalThis.localStorage?.setItem(AUTO_COLOR_PROBE_STORAGE_KEY, JSON.stringify(existing));
  } catch {
    // Ignore storage failures; console output still remains.
  }
}

export function clearAutoColorProbeLog(): void {
  if (typeof globalThis === "undefined" || !("localStorage" in globalThis)) {
    return;
  }
  try {
    globalThis.localStorage?.removeItem(AUTO_COLOR_PROBE_STORAGE_KEY);
  } catch {
    // Ignore cleanup failures.
  }
}

export function readAutoColorProbeLog(): AutoColorProbeEntry[] {
  if (typeof globalThis === "undefined" || !("localStorage" in globalThis)) {
    return [];
  }
  try {
    const raw = globalThis.localStorage?.getItem(AUTO_COLOR_PROBE_STORAGE_KEY);
    return raw ? (JSON.parse(raw) as AutoColorProbeEntry[]) : [];
  } catch {
    return [];
  }
}

export function beginAutoColorProbeSession(payload?: unknown): void {
  clearAutoColorProbeLog();
  autoColorProbeActive = true;
  autoColorProbeStart = nowMs();
  autoColorProbe("00 session-start", payload);
}

export function endAutoColorProbeSession(): void {
  if (!autoColorProbeActive) {
    return;
  }
  autoColorProbe("98 session-end");
  autoColorProbeActive = false;
}

export function autoColorProbe(label: string, payload?: unknown): void {
  if (!autoColorProbeActive) {
    return;
  }

  const entry: AutoColorProbeEntry = {
    at: new Date().toISOString(),
    elapsedMs: Number((nowMs() - autoColorProbeStart).toFixed(1)),
    label,
    payload,
  };

  appendProbeEntry(entry);
  globalThis.console?.log("[xlerate-autocolor-probe]", entry);
}
