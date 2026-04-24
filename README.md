# XLStatUDF

XLStatUDF is a statistical add-in for Microsoft Excel built with C#, .NET 8, and Excel-DNA. It provides a growing set of user-defined functions for descriptive statistics, hypothesis testing, correlation analysis, contingency tables, ANOVA/ANCOVA, random-data generation, and pivot-style statistical summaries directly in worksheets.

The project is primarily intended for users who want spreadsheet-first statistical tools without leaving Excel, while still keeping the functions testable, documented, and easy to distribute.

## Add-In Builds

The repository now contains two language-specific add-in builds for manual testing in Excel:

- Czech add-in: [`artifacts/main/cs/publish/XLStatUDF-AddIn64-packed.xll`](/c:/Users/pavel/Documents/github/XLStatUDF/artifacts/main/cs/publish/XLStatUDF-AddIn64-packed.xll)
- English add-in: [`artifacts/main/en/publish/XLStatUDF-AddIn64-packed.xll`](/c:/Users/pavel/Documents/github/XLStatUDF/artifacts/main/en/publish/XLStatUDF-AddIn64-packed.xll)

## Installers

The repository contains two installer variants:

- Czech installer: [`artifacts/installer/XLStatUDF_CS_Setup.exe`](/c:/Users/pavel/Documents/github/XLStatUDF/artifacts/installer/XLStatUDF_CS_Setup.exe)
- English installer: [`artifacts/installer/XLStatUDF_EN_Setup.exe`](/c:/Users/pavel/Documents/github/XLStatUDF/artifacts/installer/XLStatUDF_EN_Setup.exe)

## Documentation

Documentation is split by language:

- Czech documentation: [`docs/cs/README.md`](/c:/Users/pavel/Documents/github/XLStatUDF/docs/cs/README.md)
- English documentation: [`docs/en/README.md`](/c:/Users/pavel/Documents/github/XLStatUDF/docs/en/README.md)

## Build

```powershell
.\build.ps1
```

This runs restore, tests, add-in packaging, and installer generation.
