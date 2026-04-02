# Framework Overview

## Purpose

XLStatUDF is an Excel add-in that provides statistical user-defined functions implemented in C# with Excel-DNA.

## Runtime

- target runtime: `.NET 8`
- Excel integration: `Excel-DNA`
- numeric/statistical library: `MathNet.Numerics`
- output for Excel: packed `.xll` add-in for 64-bit Excel

## Main Build Path

The only build file intended for regular Excel testing is:

- [`artifacts/main/publish/XLStatUDF-AddIn64-packed.xll`](/c:/Users/pavel/Documents/github/XLStatUDF/artifacts/main/publish/XLStatUDF-AddIn64-packed.xll)

## Header Mode

Relevant functions support an optional final argument `ma_zahlavi`:

| Code | Meaning |
| --- | --- |
| `0` | auto-detect header |
| `1` | header is present |
| `2` | header is not present |
