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
  - keeps Office WEF cache so the already added local add-in remains available
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
- The Office add-in is merged into one manifest named `Evalytics StatLab for Excel`, combining custom functions, documentation, demos, options sheets, and wizards.
- If the add-in is missing after cache clear, add `Evalytics StatLab for Excel` from `\\localhost\EvalyticsOfficeAddin`.

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
- `PARSE.NUMBER` is scalar and intended for row-by-row cleanup beside raw source cells; it strips spaces/currency/noise and accepts comma or dot decimal separators with locale-aware ambiguity handling. Its optional second `else` argument returns a fallback string/number for unparseable values; the default fallback is 0.
- PIVOT tutorials use `runPivotTutorial` in `taskpane.js`. Pivot row/column dimensions accept blank or scalar values (for example `FALSE`) as "no split" and return TOTAL summaries instead of throwing.
- `ANOVA.RM`, `ANCOVA.G`, `CONTINGENCY.T`, `CONTINGENCY.G`, and `CORREL.MATRIX` were ported into the Office.js runtime; they should no longer be wired to `pendingFeature`.
- The active documentation source remains `office-addin/src/public/function-docs.json`; keep both `cs` and `en` entries aligned when function signatures or outputs change.
- Spill output labels can be localized from the custom-functions runtime. The taskpane stores the selected language under `evalytics.language`; custom functions read that synchronously and fall back to `navigator.language` / `Intl` locale. Existing spill outputs update after workbook recalculation, so the taskpane language switch requests a full recalculation.
- `ANOVA.RM` and `ANCOVA.G` now use localized section/label text for Czech and English spill outputs.
- Sidepanel UI uses two tabs: function docs and settings. Language controls live in the settings tab. Function documentation sections are intentionally collapsed by default; the section +/- indicator is CSS-driven from the `.collapsed` class.
- Test-function tutorials are wired through `matrixTutorialDefinitions` in `office-addin/src/public/taskpane.js` for `SHAPIRO.WILK`, `KOLMOGOROV.SMIRNOV`, `T.TEST.1S`, `PROP.TEST.1S`, `WILCOXON.PAIRED`, `MANN.WHITNEY.G`, `CHISQ.GOF`, `ANOVA.G`, and `CORREL.SPEARMAN`; `NORM.DIST.RANGE` uses the simple scalar tutorial path. Spearman tutorial data intentionally avoids perfect monotonic ordering so the test statistic stays finite.
- If range-based test tutorials show `#VÝPOČET!` while scalar functions work, re-check `functions.json`: range parameters must have `"dimensionality": "matrix"`. Changes to `functions.json` require `Evalytics: Hard Reset Debug Session` because Excel caches custom-functions metadata.
- UDF spill outputs should prefer readable statistical symbols (`α`, `μ₀`, `π₀`, `sₓ`, `χ²`, `η²`, `ω²`, `tᶜʳⁱᵗ`) and align with `function-docs.json`, which is the documentation source of truth.
- Custom functions must not return `NaN` or `Infinity`; `buildRows` and `rectangularRows` sanitize non-finite numbers to blank cells to avoid Excel turning the whole spill into `#HODNOTA!`.
- Pivot tutorials create source data with `FILL` for row/column dimensions and `GENERATE.NORM.ARRAY` for values. Pivot aggregation skips records whose value is blank and whose row or column dimension is also blank; unlabeled row/column categories are preserved when a value exists.
- Pivot docs in `function-docs.json` should describe the shared TOTAL-row/TOTAL-column output, blank/scalar dimension behavior, and blank-value omission consistently for both Czech and English.
- `OPTIONS("key", value, ...)` returns an internal `__EVALYTICS_OPTIONS__:` JSON marker from inline key/value pairs. `OPTIONS.T(range)` does the same from a two-column option/value table. `ANCOVA.G` supports `ANCOVA.G(groups, values, covariates, OPTIONS(...))` and `ANCOVA.G(groups, values, covariates, OPTIONS.T(range))`; the JS runtime still accepts the old scalar optional arguments for compatibility.
- Sidepanel renders documented `options` entries from `function-docs.json` and offers an Options sheet button. The generated sheet includes option/value columns and Excel validation lists for enum options.
- Sidepanel has a generic modal function wizard for documented functions with parameters. It captures ranges from the current Excel selection through Office.js, renders enum arguments as localized dropdowns from `function-docs.json`, localizes decimal display, pre-fills fields when the selected cell contains the same UDF call, and validates ranges/numbers before insertion. Parameter-level wizard metadata lives under `parameters[].wizard` in `function-docs.json`. Office.js does not expose the native COM-style Function Arguments range picker inside custom task panes.
- Wizard number fields use `input type="number"` with `min`/`max`/`step` from `parameters[].wizard` metadata. Programmatic numeric values use invariant decimal dots because HTML number inputs reject comma values in many WebViews; parsing still accepts comma if the WebView allows the user to type it.
- Descriptive functions that consume ranges must declare those parameters with `dimensionality: "matrix"` in `functions.json`; scalar `any` metadata causes `#VÝPOČET!` when formulas pass ranges such as `A2:A11`.
- Descriptive-function tutorials exist for weighted means/variance/stdev/CV. They create a clean data table and one result formula cell, without comments.
- `PERCENTILE.INC.IFS` and `PERCENTILE.EXC.IFS` were removed from the Office add-in; percentile-with-filters use cases are expected to move into the pivot family instead.

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
