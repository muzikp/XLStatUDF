# Evalytics StatLab for Excel

Evalytics StatLab for Excel is an `Office.js` add-in that brings statistical worksheet functions, documentation, demos, options, and function wizards into one Excel workspace.

The original `Excel-DNA` / `.xll` implementation is now treated as an archived reference and remains in the repository only to help verify feature parity during the migration.

## Active Product

The active cross-platform add-in lives in `office-addin/` and is distributed as a single manifest.

Current migration goals:

- preserve the existing statistical function surface where Office.js supports it
- remove legacy output-formatting behavior
- keep the documentation and download website in sync with the new add-in

## Archived Excel-DNA Implementation

The archived host implementation is described in `archive/EXCEL-DNA.md`.

Archived reference areas:

- `src/XLStatUDF`
- `tests/XLStatUDF.Tests`
- `installer/`
- `build.ps1`

## Documentation

Documentation remains split by language:

- Czech documentation: `docs/cs/README.md`
- English documentation: `docs/en/README.md`

## Office.js Build

```powershell
Set-Location .\office-addin
npm run build
```

This builds the custom-functions runtime, taskpane assets, and the merged Office.js manifest into `office-addin/dist/`.
