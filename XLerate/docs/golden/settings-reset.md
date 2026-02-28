# Golden Baseline: Settings + Reset

Source logic: `src/modules/ModGlobalSettings.bas`, `src/modules/ModFormatReset.bas`

## Contract

1. Format cycles (number/date/cell/text) can read persisted workbook settings.
2. Persisted settings are validated; invalid or empty lists fall back to built-in defaults.
3. Reset action clears persisted format settings and cycle-state keys.
4. After reset, cycle commands use built-in defaults on next run.
5. Taskpane provides editor actions to load saved settings, load defaults, and save normalized settings.
