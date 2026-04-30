# AI Notes

This file is a working memory for future AI sessions in this repository. Keep it short, factual, and update it when debugging reveals project-specific behavior.

## Project Shape

- Active product is the Office.js Excel add-in in `office-addin/`.
- The legacy Excel-DNA/C# implementation in `src/XLStatUDF`, `tests/XLStatUDF.Tests`, `installer/`, and `build.ps1` is archived/reference material for parity.
- Website and docs are in `website/` and `docs/`.
- Office add-in runtime files are source-controlled under `office-addin/src/public/` and built into `office-addin/dist/`.

## Local Office Debug Workflow

- Preferred VS Code Run and Debug entry: `Evalytics: Reset Debug Session`.
- It calls `office-addin/scripts/reset-debug-session.cmd`, which runs `reset-debug-session.ps1`.
- Normal reset workflow:
  - closes Excel
  - stops the local server on port 3000
  - keeps Office WEF cache so already added local add-ins remain available
  - ensures local HTTPS cert
  - runs `npm run build`
  - regenerates local manifests
  - refreshes the trusted shared-folder catalog
  - starts the local HTTPS server
  - launches Excel
- Use `Evalytics: Hard Reset Debug Session` or pass `-ClearOfficeCache` to clear Office WEF cache after `functions.json`, manifest identity, or custom-functions metadata changes.
- Reset transcript is written to `office-addin/.reset-debug-session.log`.
- Local server logs:
  - `office-addin/.local-serve.out.log`
  - `office-addin/.local-serve.err.log`
- If add-ins are missing after cache clear, add both `Evalytics` and `Evalytics Docs` from `\\localhost\EvalyticsOfficeAddin`.

## Important Office.js Findings

- `DEBUG.MATRIX()` proved that matrix/spill custom-function output works.
- `VERSION()` proved the custom-functions runtime can load and execute scalar functions.
- Welch failures were caused by metadata for range arguments, not by the Welch calculation itself.
- For functions that accept Excel ranges, `functions.json` parameters must include `"dimensionality": "matrix"` where appropriate.
- The working fix for `WELCH.TEST.2S.G` and `DEBUG.WELCH.2S.G` was:
  - `categories`: `{ "type": "any", "dimensionality": "matrix" }`
  - `values`: `{ "type": "any", "dimensionality": "matrix" }`
- Without matrix dimensionality on range args, Excel returned `#VÝPOČET!` / `#VALUE!`-style failures before the JS function body could return useful diagnostics.

## Welch Tutorial State

- Tutorial lives in `office-addin/src/public/taskpane.js`.
- It now generates 50 + 50 cases directly in Excel:
  - `=FILL("male",50)` at `A2`
  - `=FILL("female",50)` at `A52`
  - `=GENERATE.NORM.ARRAY(50,178,8)` at `B2`
  - `=GENERATE.NORM.ARRAY(50,168,7)` at `B52`
- It inserts `=WELCH.TEST.2S.G(A1:A101,B1:B101,1,0.05,0)` at `D3`.
- Temporary debug functions `DEBUG.WELCH.2S.G`, `DEBUG.MATRIX`, and `DEBUG.RANGE.INFO` were removed after confirming the metadata issue.
- The tutorial polls calculation state instead of relying on a fixed short sleep, because custom functions can initially show `#BUSY!`.
- The sidepanel dev log has a copy button.

## Documentation Notes

- Function detail docs for the sidepanel live directly in `office-addin/src/public/function-docs.json`.
- `office-addin/scripts/build.mjs` treats that JSON file as the source of truth and only copies it to `office-addin/dist/function-docs.json`.
- The older markdown docs under `docs/` are not used by the Excel add-in build.
- Welch docs intentionally describe parameter types in the sidepanel as human labels (`řada hodnot`, `číslo`, etc.) rather than raw Office metadata (`any / matrix`, `number`).
- Enum default values are controlled by `default: true` on `enumValues` items in `function-docs.json`.
- Welch spill output now uses Unicode labels for readability, including `α` and critical labels such as `tᶜʳⁱᵗ(1−α/2)`.
- Tutorial buttons are wired in `office-addin/src/public/taskpane.js` via `tutorialDefinitions`.
- `VERSION` and `PING` remain runtime diagnostics but are hidden from sidepanel documentation.
- General-function tutorials exist for `GENERATE.NORM`, `GENERATE.NORM.ARRAY`, `GENERATE.INT`, `GENERATE.INT.ARRAY`, `FILL`, and `PARSE.NUMBER`.
- `FILL` is currently documented and registered as a stable two-argument function (`what`, `count`). `FILL.RANDOM` was removed after a repeating value/count-pair metadata attempt caused `#HODNOTA!` in Excel and the two-argument random form was not useful.
- `PARSE.NUMBER` is scalar and intended for row-by-row cleanup beside raw source cells; it strips spaces/currency/noise and accepts comma or dot decimal separators with locale-aware ambiguity handling.

## Build Notes

- Run Office add-in build from `office-addin/`:

```powershell
npm run build
```

- Build copies `src/public/function-docs.json` to `dist/` and writes runtime assets to `dist/`.
- Excel often caches custom-functions metadata aggressively. Use reset debug session after `functions.json` changes.

## Current Caution

- There are many uncommitted changes in the repository. Do not revert user changes.
- `.vscode/` and `office-addin/` may appear untracked/dirty depending on the current git state.
